/** MyTask server-to-server Vercel proxy.
 * Environment variable: MYTASK_GAS_EXEC_URL = https://script.google.com/macros/s/.../exec
 * Do not put private secrets or tokens in frontend code or URL parameters.
 * Firebase ID tokens are verified by Code.gs using Google Identity Toolkit.
 */
module.exports = async function mytaskHandler(req,res){
  res.setHeader('Cache-Control','no-store');
  res.setHeader('X-Content-Type-Options','nosniff');
  if(req.method!=='POST')return res.status(405).json({success:false,error:{code:'METHOD',message:'Gunakan POST.'}});
  const originalOrigin=String(req.headers.origin||'');
  const host=String(req.headers.host||'');
  if(originalOrigin){
    try{if(new URL(originalOrigin).host!==host)throw new Error('origin mismatch');}
    catch{return res.status(403).json({success:false,error:{code:'ORIGIN',message:'Request origin tidak diizinkan.'}})}
  }
  let body=req.body||{};
  if(typeof body==='string'){
    try{body=JSON.parse(body)}catch{return res.status(400).json({success:false,error:{code:'INPUT',message:'JSON invalid.'}})}
  }
  const {operation,token,requestId,payload={}}=body;
  if(typeof operation!=='string'||operation.length>60||!operation||typeof token!=='string'||!token||token.length>8000||!payload||typeof payload!=='object'||Array.isArray(payload)){
    return res.status(400).json({success:false,error:{code:'INPUT',message:'Parameter API tidak valid.'}});
  }
  const serialized=JSON.stringify(payload);
  if(serialized.length>120000)return res.status(413).json({success:false,error:{code:'INPUT',message:'Payload terlalu besar.'}});
  const endpoint=process.env.MYTASK_GAS_EXEC_URL||'';
  if(!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(endpoint)){
    return res.status(503).json({success:false,error:{code:'CONFIG',message:'Vercel Environment Variable MYTASK_GAS_EXEC_URL belum diatur.'}});
  }
  const params=new URLSearchParams({
    protocol:'MYTASK_SERVER_V2',
    operation,
    payload:serialized,
    token,
    requestId:typeof requestId==='string'?requestId.slice(0,100):'',
  });
  const controller=new AbortController();
  const timeout=setTimeout(()=>controller.abort(),32000);
  const start=Date.now();
  try{
    const upstream=await fetch(endpoint,{
      method:'POST',
      headers:{'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8'},
      body:params.toString(),
      redirect:'follow',
      cache:'no-store',
      signal:controller.signal
    });
    const text=await upstream.text();
    let result;
    try{result=JSON.parse(text)}catch{
      console.error('[MyTask Proxy] Non-JSON GAS response, status',upstream.status);
      return res.status(502).json({success:false,error:{code:'UPSTREAM_FORMAT',message:'Apps Script tidak mengembalikan JSON. Deploy ulang Code.gs versi terbaru.'}});
    }
    console.log('[MyTask Proxy Speed]',operation,'elapsedMs=',Date.now()-start,'success=',!!result.success);
    return res.status(upstream.ok?200:502).json(result);
  }catch(err){
    console.error('[MyTask Proxy]',operation,err?.name||'Error');
    return res.status(502).json({success:false,error:{code:err?.name==='AbortError'?'TIMEOUT':'NETWORK',message:'Proxy gagal menghubungi Apps Script. Periksa URL /exec dan deployment.'}});
  }finally{clearTimeout(timeout)}
};
