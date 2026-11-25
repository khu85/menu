/* Kinhfolk pages: open instantly from the device and keep working when the shop wifi drops.
   - Pages: shown from the device copy right away, refreshed in the background. If a newer
     version came in, the page is told so it can reload before anyone starts typing.
   - Fonts: kept on the device (they never change under the same name).
   - Only this site's own files are kept. Data from the Google scripts is never stored here. */
var V = "kf-v2", F = "kf-fonts-1";
self.addEventListener("install", function () { self.skipWaiting(); });
self.addEventListener("activate", function (e) {
  e.waitUntil(caches.keys().then(function (k) {
    return Promise.all(k.filter(function (x) { return x !== V && x !== F; }).map(function (x) { return caches.delete(x); }));
  }).then(function () { return self.clients.claim(); }));
});
function isFont(u) { return /\/fonts\/[^\/]+\.woff2$/.test(u.pathname); }
function stamp(r) { return r && (r.headers.get("etag") || r.headers.get("last-modified") || ""); }
function keep(cache, req, res) { if (res && res.ok && !res.redirected && res.type === "basic") cache.put(req, res.clone()); return res; }
function tell(url) {
  self.clients.matchAll({ type: "window" }).then(function (cs) {
    cs.forEach(function (c) { if (c.url.split("#")[0] === url.split("#")[0]) c.postMessage({ t: "kf-new" }); });
  });
}
self.addEventListener("fetch", function (e) {
  var r = e.request, u = new URL(r.url);
  if (r.method !== "GET" || u.origin !== location.origin) return;
  if (isFont(u)) {
    e.respondWith(caches.open(F).then(function (c) {
      return c.match(r).then(function (m) { return m || fetch(r).then(function (res) { return keep(c, r, res); }); });
    }));
    return;
  }
  var page = r.mode === "navigate" || /\.html?$|\/$/.test(u.pathname);
  if (!page) {   /* other files (pdf, json): network first, device copy if offline */
    e.respondWith(fetch(r).then(function (res) { caches.open(V).then(function (c) { keep(c, r, res); }); return res; })
      .catch(function () { return caches.match(r, { ignoreSearch: true }); }));
    return;
  }
  e.respondWith(caches.open(V).then(function (c) {
    return c.match(r, { ignoreSearch: true }).then(function (m) {
      var net = fetch(r).then(function (res) {
        if (res && res.ok && !res.redirected) {
          var a = stamp(m), b = stamp(res);
          var put = c.put(new Request(u.origin + u.pathname), res.clone());
          if (m && a && b && a !== b) put.then(function () { tell(r.url); });
        }
        return res;
      });
      if (m) { e.waitUntil(net.catch(function () {})); return m; }
      return net;
    });
  }));
});
/* Pages ask for the other tabs to be fetched ahead of time, so switching tabs is instant. */
self.addEventListener("message", function (e) {
  var d = e.data || {};
  if (d.t !== "warm" || !Array.isArray(d.urls)) return;
  e.waitUntil(caches.open(V).then(function (c) {
    return Promise.all(d.urls.slice(0, 12).map(function (s) {
      var u = new URL(s, location.href);
      if (u.origin !== location.origin) return null;
      var key = u.origin + u.pathname;
      return c.match(key).then(function (m) {
        if (m) return null;
        return fetch(key, { credentials: "same-origin" }).then(function (res) { keep(c, new Request(key), res); }).catch(function () {});
      });
    }));
  }));
});
