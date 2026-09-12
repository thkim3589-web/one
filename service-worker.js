const CACHE_NAME = "eng-sheet-app-v2";
const ASSETS = [
  "./",
  "./index.html",
  "./app.js",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "./icon-180.png",
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

// 네트워크 우선: 온라인이면 항상 최신 파일을 받아오고 캐시도 갱신한다.
// 오프라인일 때만 예전에 캐시된 버전을 사용한다.
self.addEventListener("fetch", (event)=>{
  event.respondWith(
    fetch(event.request).then((resp)=>{
      const copy = resp.clone();
      caches.open(CACHE_NAME).then((cache)=> cache.put(event.request, copy)).catch(()=>{});
      return resp;
    }).catch(()=> caches.match(event.request))
  );
});
