const CACHE_NAME = "eng-sheet-app-v1";
const ASSETS = [
  "./",
  "./index.html",
  "./app.js",
  "./manifest.json",
  "./icon.svg",
  "./template.xlsx",
  "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js",
];

self.addEventListener("install", (event)=>{
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache)=> cache.addAll(ASSETS)).catch(()=>{})
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event)=>{
  event.waitUntil(
    caches.keys().then((keys)=>
      Promise.all(keys.filter(k=>k!==CACHE_NAME).map(k=>caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event)=>{
  event.respondWith(
    caches.match(event.request).then((cached)=>{
      if(cached) return cached;
      return fetch(event.request).then((resp)=>{
        const copy = resp.clone();
        caches.open(CACHE_NAME).then((cache)=> cache.put(event.request, copy)).catch(()=>{});
        return resp;
      }).catch(()=> cached);
    })
  );
});
