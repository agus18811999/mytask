/* MyTask shared configuration — edit values below before deployment. */
window.MYTASK_CONFIG = Object.freeze({
  firebase: {
    apiKey: 'REPLACE_FIREBASE_API_KEY',
    authDomain: 'REPLACE_PROJECT_ID.firebaseapp.com',
    projectId: 'REPLACE_PROJECT_ID',
    appId: 'REPLACE_FIREBASE_APP_ID'
  },
  gasUrl: 'REPLACE_APPS_SCRIPT_WEB_APP_EXEC_URL',
  apiTimeoutMs: 35000,
  cacheVersion: 'mytask-v1',
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
    if (event.source !== entry.iframe.contentWindow) return;
    entry.finish(packet.response);
  });
  const error = (message, code = 'REQUEST_FAILED') => Object.assign(new Error(message), {code});
  async function request(auth, operation, payload = {}) {
    if (!auth?.currentUser) throw error('Sesi tidak ditemukan. Silakan login.', 'UNAUTHENTICATED');
    const endpoint = cfg.gasUrl;
    if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(endpoint)) {
      throw error('Konfigurasi Apps Script belum diisi dengan URL /exec yang valid.', 'NOT_CONFIGURED');
    }
    const user = auth.currentUser;
    const token = await user.getIdToken();
    const requestId = `${Date.now().toString(36)}-${(++counter).toString(36)}-${Math.random().toString(36).slice(2,12)}`;
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
        protocol: 'MYTASK_BRIDGE_V1', requestId, operation,
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
        if (!response || !response.success) {
          reject(error(response?.error?.message || 'Permintaan tidak berhasil.', response?.error?.code || 'BACKEND_ERROR'));
        } else resolve(response.data);
      };
      const timer = setTimeout(() => finish({success:false,error:{code:'TIMEOUT',message:'Server belum merespons. Periksa koneksi dan coba lagi.'}}), cfg.apiTimeoutMs);
      document.body.append(iframe, form);
      pending.set(requestId, {iframe, finish});
      form.submit();
    });
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
