/* Kinhfolk staff pages: keep pages available when the shop wifi drops.
   Only same-site files are cached; saving counts still needs a connection. */
var V="kf-v1";
self.addEventListener("install",function(e){self.skipWaiting()});
self.addEventListener("activate",function(e){e.waitUntil(caches.keys().then(function(k){return Promise.all(k.filter(function(x){return x!==V}).map(function(x){return caches.delete(x)}))}).then(function(){return self.clients.claim()}))});
self.addEventListener("fetch",function(e){
 var r=e.request,u=new URL(r.url);
 if(r.method!=="GET"||u.origin!==location.origin)return;
 e.respondWith(fetch(r).then(function(res){if(res&&res.ok){var c=res.clone();caches.open(V).then(function(h){h.put(r,c)})}return res}).catch(function(){return caches.match(r).then(function(m){return m||caches.match(r,{ignoreSearch:true})})}));
});
