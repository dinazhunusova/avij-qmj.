const C='qmj-rah-2026-11-13a', P='qmj-'+C.split('-')[1]+'-';
const FILES=['./','index.html','manifest.webmanifest','icon-192.png','icon-512.png','icon-maskable-512.png','apple-touch-icon.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(FILES)));self.skipWaiting();});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C&&x.indexOf(P)===0).map(x=>caches.delete(x)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{
  const r=e.request, u=new URL(r.url), BASE=new URL('./',self.location).pathname;
  if(r.method!=='GET'||u.origin!==location.origin||!u.pathname.startsWith(BASE)||u.pathname.slice(BASE.length).includes('/')) return;
  if(r.mode==='navigate'){
    e.respondWith(fetch(r,{cache:'no-store'}).then(res=>{const cp=res.clone();caches.open(C).then(c=>c.put('index.html',cp));return res;}).catch(()=>caches.match('index.html')));
    return;
  }
  e.respondWith(caches.match(r).then(m=>m||fetch(r)));
});
