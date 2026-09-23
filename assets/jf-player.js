/* Jordan Fein — project page player (Mux HLS, site-styled controls) */
(function () {
  var MONO = "'Courier Prime','Courier New',monospace";
  var CREAM = "#f2efe9";
  function loadHls(cb) {
    var v = document.createElement("video");
    if (v.canPlayType("application/vnd.apple.mpegurl") && /iphone|ipad|^((?!chrome|android|crios|fxios).)*safari/i.test(navigator.userAgent)) return cb(null);
    if (window.Hls) return cb(window.Hls);
    var s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/hls.js@1/dist/hls.min.js";
    s.onload = function () { cb(window.Hls); };
    s.onerror = function () { cb(null); };
    document.head.appendChild(s);
  }
  function fmt(t) { t = Math.max(0, Math.floor(t || 0)); var m = Math.floor(t / 60), s = t % 60; return (m < 10 ? "0" : "") + m + ":" + (s < 10 ? "0" : "") + s; }
  function el(tag, css, html) { var e = document.createElement(tag); if (css) e.style.cssText = css; if (html != null) e.innerHTML = html; return e; }
  function btnCss() { return "cursor:pointer;background:transparent;border:0;padding:10px 0;color:" + CREAM + ";font-family:" + MONO + ";font-size:11px;letter-spacing:0.18em;text-transform:uppercase;white-space:nowrap;min-height:44px;"; }

  function mount(box) {
    var id = box.getAttribute("data-mux"), title = box.getAttribute("data-title") || "";
    var poster = "https://image.mux.com/" + id + "/thumbnail.webp?width=1920&time=" + (box.getAttribute("data-time") || "0");
    if (getComputedStyle(box).position === "static") box.style.position = "relative"; box.style.background = "#000"; box.style.overflow = "hidden";
    var v = el("video", "position:absolute;inset:0;width:100%;height:100%;object-fit:contain;background:#000;display:block;");
    v.setAttribute("playsinline", ""); v.setAttribute("preload", "metadata"); v.poster = poster; v.setAttribute("aria-label", title);
    box.appendChild(v);

    var big = el("button", "position:absolute;inset:0;margin:auto;width:112px;height:112px;border-radius:50%;border:1px solid rgba(242,239,233,0.8);background:rgba(10,10,11,0.25);backdrop-filter:blur(4px);-webkit-backdrop-filter:blur(4px);color:" + CREAM + ";font-family:" + MONO + ";font-size:11px;letter-spacing:0.2em;text-transform:uppercase;cursor:pointer;transition:transform 0.4s cubic-bezier(0.22,1,0.36,1),opacity 0.4s ease,background 0.3s ease;", "Play");
    big.setAttribute("aria-label", "Play " + title);
    big.onmouseenter = function () { big.style.transform = "scale(1.06)"; big.style.background = "rgba(242,239,233,0.12)"; };
    big.onmouseleave = function () { big.style.transform = "none"; big.style.background = "rgba(10,10,11,0.25)"; };
    box.appendChild(big);

    var bar = el("div", "position:absolute;left:0;right:0;bottom:0;padding:40px clamp(14px,3vw,24px) 8px;background:linear-gradient(0deg,rgba(10,10,11,0.8),rgba(10,10,11,0));display:flex;flex-direction:column;gap:6px;opacity:0;pointer-events:none;transition:opacity 0.5s ease;");
    var track = el("div", "position:relative;height:14px;cursor:pointer;display:flex;align-items:center;");
    var rail = el("div", "position:relative;width:100%;height:1px;background:rgba(242,239,233,0.25);transition:height 0.25s ease;");
    var buf = el("div", "position:absolute;left:0;top:0;bottom:0;width:0;background:rgba(242,239,233,0.35);");
    var fill = el("div", "position:absolute;left:0;top:0;bottom:0;width:0;background:" + CREAM + ";");
    rail.appendChild(buf); rail.appendChild(fill); track.appendChild(rail);
    track.onmouseenter = function () { rail.style.height = "3px"; };
    track.onmouseleave = function () { rail.style.height = "1px"; };
    var row = el("div", "display:flex;align-items:center;gap:clamp(14px,3vw,26px);");
    var play = el("button", btnCss(), "Play");
    var time = el("span", "font-family:" + MONO + ";font-size:11px;letter-spacing:0.12em;color:#b8b2a9;font-variant-numeric:tabular-nums;", "00:00 / 00:00");
    var sp = el("span", "flex:1;");
    var snd = el("button", btnCss(), "Sound off");
    var fs = el("button", btnCss(), "Full screen");
    row.appendChild(play); row.appendChild(time); row.appendChild(sp); row.appendChild(snd); row.appendChild(fs);
    bar.appendChild(track); bar.appendChild(row); box.appendChild(bar);

    var started = false, hideT = null;
    function showBar() { if (!started) return; bar.style.opacity = "1"; bar.style.pointerEvents = "auto"; box.style.cursor = "default"; clearTimeout(hideT); hideT = setTimeout(function () { if (!v.paused) { bar.style.opacity = "0"; bar.style.pointerEvents = "none"; box.style.cursor = "none"; } }, 2600); }
    function attach(cb) {
      if (v._ready) return cb();
      var url = "https://stream.mux.com/" + id + ".m3u8";
      loadHls(function (Hls) {
        if (Hls && Hls.isSupported()) { var h = new Hls({ capLevelToPlayerSize: true, startLevel: -1 }); h.loadSource(url); h.attachMedia(v); }
        else v.src = url;
        v._ready = true; cb();
      });
    }
    function start() {
      attach(function () {
        started = true; v.muted = false; big.style.opacity = "0"; big.style.pointerEvents = "none";
        var p = v.play(); if (p && p.catch) p.catch(function () { v.muted = true; v.play(); });
        if (window.gtag) gtag("event", "video_play", { video_title: title, video_id: id, page: "project" });
        showBar();
      });
    }
    function toggle() { if (!started) return start(); if (v.paused) v.play(); else v.pause(); showBar(); }
    big.onclick = function (e) { e.stopPropagation(); start(); };
    v.onclick = toggle;
    play.onclick = function (e) { e.stopPropagation(); toggle(); };
    snd.onclick = function (e) { e.stopPropagation(); v.muted = !v.muted; showBar(); };
    fs.onclick = function (e) {
      e.stopPropagation();
      var d = document;
      if (d.fullscreenElement || d.webkitFullscreenElement) { (d.exitFullscreen || d.webkitExitFullscreen).call(d); return; }
      var req = box.requestFullscreen || box.webkitRequestFullscreen;
      if (req) req.call(box); else if (v.webkitEnterFullscreen) v.webkitEnterFullscreen();
    };
    function seek(e) { var r = track.getBoundingClientRect(); var x = ((e.touches ? e.touches[0].clientX : e.clientX) - r.left) / r.width; if (v.duration) v.currentTime = Math.max(0, Math.min(1, x)) * v.duration; showBar(); }
    track.onclick = function (e) { e.stopPropagation(); seek(e); };
    box.addEventListener("mousemove", showBar);
    box.addEventListener("touchstart", showBar, { passive: true });
    v.addEventListener("play", function () { play.textContent = "Pause"; showBar(); });
    v.addEventListener("pause", function () { play.textContent = "Play"; bar.style.opacity = "1"; bar.style.pointerEvents = "auto"; box.style.cursor = "default"; });
    v.addEventListener("volumechange", function () { snd.textContent = v.muted ? "Sound on" : "Sound off"; });
    v.addEventListener("timeupdate", function () { var d = v.duration || 0; fill.style.width = d ? (v.currentTime / d * 100) + "%" : "0"; time.textContent = fmt(v.currentTime) + " / " + fmt(d); });
    v.addEventListener("progress", function () { try { var d = v.duration; if (d && v.buffered.length) buf.style.width = (v.buffered.end(v.buffered.length - 1) / d * 100) + "%"; } catch (e) {} });
    v.addEventListener("ended", function () { play.textContent = "Replay"; });
    document.addEventListener("fullscreenchange", function () { fs.textContent = document.fullscreenElement === box ? "Exit full screen" : "Full screen"; v.style.objectFit = "contain"; });
    box.tabIndex = 0;
    box.addEventListener("keydown", function (e) { if (e.code === "Space" || e.key === "k") { e.preventDefault(); toggle(); } else if (e.key === "f") fs.onclick(e); else if (e.key === "m") snd.onclick(e); });
  }
  function init() { Array.prototype.forEach.call(document.querySelectorAll("[data-jf-player]"), mount); }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
