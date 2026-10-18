/* TSEKmate service worker: the app opens instantly and works with no signal. Bump V whenever you change any file. */
const V='tsekmate-1.0',
 SHELL=['./','index.html','manifest.webmanifest','icon.svg','icon-180.png','icon-192.png','icon-512.png','icon-maskable-512.png'],
 EXTRA=['TSEKmate-sheet-40.pdf','TSEKmate-sheet-50.pdf','TSEKmate-sheet-60.pdf'];
self.addEventListener('install',e=>e.waitUntil((async()=>{
 const c=await caches.open(V);await c.addAll(SHELL);                      /* the app itself must cache, or install fails loudly */
 await Promise.allSettled(EXTRA.map(u=>c.add(u)));self.skipWaiting()})()));  /* sheets are a bonus: a missing PDF never blocks install */
self.addEventListener('activate',e=>e.waitUntil((async()=>{
 for(const k of await caches.keys())if(k!==V)await caches.delete(k);await self.clients.claim()})()));
self.addEventListener('fetch',e=>{
 const r=e.request,u=new URL(r.url);
 if(r.method!=='GET'||u.origin!==self.location.origin)return;               /* Google Sheets sync and everything cross-origin go straight to the network */
 e.respondWith((async()=>{
  const c=await caches.open(V),hit=await c.match(r,{ignoreSearch:true});
  const net=fetch(r).then(res=>{if(res.status===200)c.put(r,res.clone());return res}).catch(()=>null);
  if(hit){e.waitUntil(net);return hit}                                      /* answer instantly from the phone, refresh quietly behind it */
  return (await net)||(r.mode==='navigate'?await c.match('index.html'):Response.error())})())});
