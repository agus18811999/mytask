/* MyTask shared configuration — edit values below before deployment. */
window.MYTASK_CONFIG = Object.freeze({
  firebase: {
    apiKey: 'AIzaSyBIsLSAf35f63zWjkbiSY8NmrbXVVPWBWU',
    authDomain: 'mytask-60f59.firebaseapp.com',
    projectId: 'mytask-60f59',
    appId: '1:900912234219:web:843e99a1fcb5764a15c8aeD'
  },
  gasUrl: 'https://script.google.com/macros/s/AKfycbyvXJKv7oK8h-xI-fGpBoB0G2iC_sIFmQvhfoCFw2GqNMNW5-UsvKe8D39Gxb33NVJC/exec',
  apiMode: 'proxy', // Production: Vercel serverless /api/mytask (recommended)
  apiProxy: '/api/mytask',
  apiTimeoutMs: 35000,
  cacheVersion: 'mytask-v2',
  appName: 'MyTask'
});
/* Apps Script /exec cannot be fetched reliably across origins because the
   Google redirect and missing CORS headers cannot be controlled. This bridge
   uses a POST form in a hidden iframe and a restricted-origin postMessage reply.
   Deploy GAS as 'Execute as me / Anyone', and configure ALLOWED_ORIGINS on GAS.
   ID tokens are sent only in the POST body (never query string/localStorage). */
window.MyTaskAPI = (() => {
  const cfg = window.MYTASK_CONFIG;
  const pending = new Map();
  const inflight = new Map();
  let counter = 0;
  const allowedMessageHosts = [
    /^https:\/\/script\.google\.com$/,
    /^https:\/\/([a-z0-9-]+\.)?script\.googleusercontent\.com$/,
    /^https:\/\/([a-z0-9-]+\.)?googleusercontent\.com$/
  ];
  window.addEventListener('message', event => {
    if (!allowedMessageHosts.some(re => re.test(event.origin))) return;
    const packet = event.data;
    if (!packet || packet.protocol !== 'MYTASK_BRIDGE_V1' || typeof packet.requestId !== 'string') return;
    const entry = pending.get(packet.requestId);
    if (!entry) return;
    // Apps Script HTML Service renders its result inside nested Google iframes.
    // Its message source will NOT be the outer form-target iframe. Authenticate
    // the reply with the pending request ID, a separate random nonce, and origin.
    if (packet.nonce !== entry.nonce) return;
    entry.finish(packet.response);
  });
  const error = (message, code = 'REQUEST_FAILED') => Object.assign(new Error(message), {code});
  async function rawRequest(auth, operation, payload = {}) {
    const requestStart=performance.now();
    if (!auth?.currentUser) throw error('Sesi tidak ditemukan. Silakan login.', 'UNAUTHENTICATED');
    const endpoint = cfg.gasUrl;
    if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(endpoint)) {
      throw error('Konfigurasi Apps Script belum diisi dengan URL /exec yang valid.', 'NOT_CONFIGURED');
    }
    const user = auth.currentUser;
    const token = await user.getIdToken();
    const requestId = `${Date.now().toString(36)}-${(++counter).toString(36)}-${crypto.randomUUID()}`;
    const nonce = crypto.randomUUID();
    return new Promise((resolve, reject) => {
      const iframe = document.createElement('iframe');
      iframe.name = `mytask_rpc_${counter}`;
      iframe.title = 'MyTask background request';
      iframe.tabIndex = -1;
      iframe.setAttribute('aria-hidden','true');
      iframe.style.cssText = 'position:absolute;width:0;height:0;border:0;visibility:hidden;pointer-events:none';
      const form = document.createElement('form');
      form.method = 'POST';
      form.action = endpoint;
      form.target = iframe.name;
      form.style.display = 'none';
      const entries = {
        protocol: 'MYTASK_BRIDGE_V1', requestId, operation, nonce,
        payload: JSON.stringify(payload ?? {}), token, replyOrigin: location.origin
      };
      for (const [name, value] of Object.entries(entries)) {
        const input = document.createElement('input');
        input.type = 'hidden'; input.name = name; input.value = String(value);
        form.append(input);
      }
      let done = false;
      const finish = response => {
        if (done) return; done = true;
        clearTimeout(timer); pending.delete(requestId);
        form.remove(); iframe.remove();
        console.info(`[MyTask Speed] API ${operation}: ${Math.round(performance.now()-requestStart)} ms`);
        if (!response || !response.success) {
          reject(error(response?.error?.message || 'Permintaan tidak berhasil.', response?.error?.code || 'BACKEND_ERROR'));
        } else resolve(response.data);
      };
      const timer = setTimeout(() => finish({success:false,error:{code:'TIMEOUT',message:'Server belum merespons. Periksa koneksi dan coba lagi.'}}), cfg.apiTimeoutMs);
      document.body.append(iframe, form);
      pending.set(requestId, {iframe, finish, nonce});
      form.submit();
    });
  }
  async function proxyRequest(auth,operation,payload={}){
    if(!auth?.currentUser)throw error('Sesi tidak ditemukan. Silakan login.','UNAUTHENTICATED');
    const started=performance.now();
    const token=await auth.currentUser.getIdToken();
    const controller=new AbortController();
    const timeout=setTimeout(()=>controller.abort(),cfg.apiTimeoutMs);
    try{
      const response=await fetch(cfg.apiProxy||'/api/mytask',{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        credentials:'same-origin',
        cache:'no-store',
        signal:controller.signal,
        body:JSON.stringify({operation,payload,token,requestId:crypto.randomUUID()})
      });
      let packet;
      try{packet=await response.json()}catch{throw error('Respons API tidak valid. Periksa deployment proxy Vercel.','INVALID_RESPONSE')}
      if(!response.ok||!packet.success)throw error(packet?.error?.message||'Server belum berhasil memproses permintaan.',packet?.error?.code||'BACKEND_ERROR');
      console.info(`[MyTask Speed] API proxy ${operation}: ${Math.round(performance.now()-started)} ms`);
      return packet.data;
    }catch(err){
      if(err.name==='AbortError')throw error('Koneksi server melewati batas waktu. Periksa Apps Script Executions.','TIMEOUT');
      throw err;
    }finally{clearTimeout(timeout)}
  }
  function request(auth,operation,payload={}){
    const read=operation==='bootstrap'||operation==='snapshot';
    const key=read?`${auth?.currentUser?.uid}:${operation}`:null;
    if(key&&inflight.has(key))return inflight.get(key);
    const promise=cfg.apiMode==='proxy'?proxyRequest(auth,operation,payload):rawRequest(auth,operation,payload);
    if(key){inflight.set(key,promise);promise.finally(()=>inflight.delete(key)).catch(()=>{});}
    return promise;
  }
  function cacheKey(uid, name) { return `${cfg.cacheVersion}:${uid}:${name}`; }
  function readCache(uid, name, maxAgeMs=180000) {
    try { const record = JSON.parse(sessionStorage.getItem(cacheKey(uid,name)) || 'null');
      return record && Date.now()-record.savedAt < maxAgeMs ? record.value : null;
    } catch {return null;}
  }
  function writeCache(uid, name, value) {
    try {sessionStorage.setItem(cacheKey(uid,name), JSON.stringify({savedAt:Date.now(),value}));} catch {}
  }
  function clearCache(uid) {
    try { Object.keys(sessionStorage).filter(k => k.startsWith(`${cfg.cacheVersion}:${uid}:`)).forEach(k=>sessionStorage.removeItem(k)); } catch {}
  }
  return {request, readCache, writeCache, clearCache};
})();
