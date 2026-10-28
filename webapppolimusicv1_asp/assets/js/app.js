/* ==========================================================================
   PoliMusic - Logica del reproductor y utilidades de interfaz
   ========================================================================== */
(function () {
  "use strict";

  var audio      = document.getElementById("audio");
  var btnPlay    = document.getElementById("btnPlay");
  var btnPrev    = document.getElementById("btnPrev");
  var btnNext    = document.getElementById("btnNext");
  var seek       = document.getElementById("seek");
  var seekFill   = document.getElementById("seekFill");
  var tCur       = document.getElementById("tCur");
  var tDur       = document.getElementById("tDur");
  var npTitle    = document.getElementById("npTitle");
  var npArtist   = document.getElementById("npArtist");
  var npArt      = document.getElementById("npArt");
  var vol        = document.getElementById("vol");
  var btnMute    = document.getElementById("btnMute");
  var btnShuffle = document.getElementById("btnShuffle");
  var btnRepeat  = document.getElementById("btnRepeat");
  var btnPlayAll = document.getElementById("btnPlayAll");

  if (!audio) { return; }

  var queue   = [];   // [{id,title,artist,url,cover}]
  var current = -1;
  var shuffle = false;
  var repeat  = false;
  var registered = {};   // ids ya contabilizados en esta sesion de pagina

  /* ---------------------------------------------------------------------
     Utilidades
     --------------------------------------------------------------------- */
  function fmt(sec) {
    if (!isFinite(sec) || sec < 0) { sec = 0; }
    var m = Math.floor(sec / 60);
    var s = Math.floor(sec % 60);
    return m + ":" + (s < 10 ? "0" : "") + s;
  }

  function buildQueue() {
    queue = [];
    var nodes = document.querySelectorAll("[data-song-url]");
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i];
      var url = n.getAttribute("data-song-url");
      if (!url) { continue; }
      queue.push({
        id:     n.getAttribute("data-song-id")     || "",
        title:  n.getAttribute("data-song-title")  || "Sin titulo",
        artist: n.getAttribute("data-song-artist") || "Artista desconocido",
        cover:  n.getAttribute("data-song-cover")  || "",
        url:    url,
        el:     n
      });
    }
  }

  function markPlaying(id) {
    var rows = document.querySelectorAll(".tracklist tbody tr");
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].getAttribute("data-row-song") === id) {
        rows[i].classList.add("playing");
      } else {
        rows[i].classList.remove("playing");
      }
    }
  }

  function registerPlay(id) {
    if (!id || registered[id]) { return; }
    registered[id] = true;
    try {
      var x = new XMLHttpRequest();
      x.open("GET", "play.asp?id=" + encodeURIComponent(id), true);
      x.send();
    } catch (e) { /* silencioso */ }
  }

  function setNowPlaying(item) {
    if (!item) {
      npTitle.textContent = "Nada en reproduccion";
      npArtist.textContent = "Elige una cancion";
      if (npArt) { npArt.innerHTML = ""; }
      return;
    }
    npTitle.textContent  = item.title;
    npArtist.textContent = item.artist;
    if (npArt) {
      if (item.cover) {
        npArt.innerHTML = '<img src="' + item.cover + '" alt="">';
      } else {
        npArt.innerHTML = '<div class="cover-ph sm">&#9835;</div>';
      }
    }
    document.title = item.title + " - " + item.artist + " | PoliMusic";
  }

  /* ---------------------------------------------------------------------
     Control de reproduccion
     --------------------------------------------------------------------- */
  function playIndex(i) {
    if (i < 0 || i >= queue.length) { return; }
    var item = queue[i];
    if (current !== i || audio.src.indexOf(item.url) === -1) {
      audio.src = item.url;
      current = i;
    }
    audio.play().catch(function () { /* autoplay bloqueado */ });
    setNowPlaying(item);
    markPlaying(item.id);
    registerPlay(item.id);
    if (btnPlay) { btnPlay.innerHTML = "&#10074;&#10074;"; }
  }

  function togglePlay() {
    if (current === -1) {
      if (queue.length) { playIndex(0); }
      return;
    }
    if (audio.paused) {
      audio.play().catch(function () {});
      if (btnPlay) { btnPlay.innerHTML = "&#10074;&#10074;"; }
    } else {
      audio.pause();
      if (btnPlay) { btnPlay.innerHTML = "&#9654;"; }
    }
  }

  function nextIndex() {
    if (!queue.length) { return -1; }
    if (shuffle) { return Math.floor(Math.random() * queue.length); }
    return (current + 1) % queue.length;
  }

  function prevIndex() {
    if (!queue.length) { return -1; }
    if (shuffle) { return Math.floor(Math.random() * queue.length); }
    return (current - 1 + queue.length) % queue.length;
  }

  /* ---------------------------------------------------------------------
     Eventos
     --------------------------------------------------------------------- */
  if (btnPlay) {
    btnPlay.addEventListener("click", togglePlay);
  }
  if (btnNext) {
    btnNext.addEventListener("click", function () {
      var i = nextIndex();
      if (i >= 0) { playIndex(i); }
    });
  }
  if (btnPrev) {
    btnPrev.addEventListener("click", function () {
      if (audio.currentTime > 3) { audio.currentTime = 0; return; }
      var i = prevIndex();
      if (i >= 0) { playIndex(i); }
    });
  }
  if (btnShuffle) {
    btnShuffle.addEventListener("click", function () {
      shuffle = !shuffle;
      btnShuffle.classList.toggle("on", shuffle);
    });
  }
  if (btnRepeat) {
    btnRepeat.addEventListener("click", function () {
      repeat = !repeat;
      btnRepeat.classList.toggle("on", repeat);
    });
  }

  audio.addEventListener("timeupdate", function () {
    if (!audio.duration) { return; }
    var pct = (audio.currentTime / audio.duration) * 100;
    if (seekFill) { seekFill.style.width = pct + "%"; }
    if (tCur) { tCur.textContent = fmt(audio.currentTime); }
  });
  audio.addEventListener("loadedmetadata", function () {
    if (tDur) { tDur.textContent = fmt(audio.duration); }
  });
  audio.addEventListener("play", function () {
    if (btnPlay) { btnPlay.innerHTML = "&#10074;&#10074;"; }
  });
  audio.addEventListener("pause", function () {
    if (btnPlay) { btnPlay.innerHTML = "&#9654;"; }
  });
  audio.addEventListener("ended", function () {
    if (repeat) {
      audio.currentTime = 0;
      audio.play().catch(function () {});
      return;
    }
    var i = nextIndex();
    if (i >= 0 && queue.length > 1) { playIndex(i); }
  });
  audio.addEventListener("error", function () {
    if (npArtist) { npArtist.textContent = "No se pudo cargar el archivo de audio"; }
  });

  if (seek) {
    seek.addEventListener("click", function (ev) {
      if (!audio.duration) { return; }
      var rect = seek.getBoundingClientRect();
      var ratio = (ev.clientX - rect.left) / rect.width;
      if (ratio < 0) { ratio = 0; }
      if (ratio > 1) { ratio = 1; }
      audio.currentTime = ratio * audio.duration;
    });
  }

  if (vol) {
    audio.volume = parseFloat(vol.value);
    vol.addEventListener("input", function () {
      audio.volume = parseFloat(vol.value);
      if (btnMute) { btnMute.classList.toggle("on", audio.volume === 0); }
    });
  }
  if (btnMute) {
    btnMute.addEventListener("click", function () {
      audio.muted = !audio.muted;
      btnMute.classList.toggle("on", audio.muted);
    });
  }

  /* ---------------------------------------------------------------------
     Botones "reproducir" repartidos por la pagina
     --------------------------------------------------------------------- */
  document.addEventListener("click", function (ev) {
    var el = ev.target.closest ? ev.target.closest("[data-play-index]") : null;
    if (el) {
      ev.preventDefault();
      buildQueue();
      var idx = parseInt(el.getAttribute("data-play-index"), 10);
      if (!isNaN(idx)) { playIndex(idx); }
      return;
    }
    el = ev.target.closest ? ev.target.closest("[data-play-id]") : null;
    if (el) {
      ev.preventDefault();
      buildQueue();
      var sid = el.getAttribute("data-play-id");
      for (var i = 0; i < queue.length; i++) {
        if (queue[i].id === sid) { playIndex(i); return; }
      }
    }
  });

  if (btnPlayAll) {
    btnPlayAll.addEventListener("click", function (ev) {
      ev.preventDefault();
      buildQueue();
      if (queue.length) { playIndex(0); }
    });
  }

  /* Confirmaciones para acciones destructivas */
  document.addEventListener("submit", function (ev) {
    var f = ev.target;
    if (f.getAttribute && f.getAttribute("data-confirm")) {
      if (!window.confirm(f.getAttribute("data-confirm"))) {
        ev.preventDefault();
      }
    }
  });
  document.addEventListener("click", function (ev) {
    var a = ev.target.closest ? ev.target.closest("a[data-confirm]") : null;
    if (a && !window.confirm(a.getAttribute("data-confirm"))) {
      ev.preventDefault();
    }
  });

  /* Cierre automatico de alertas */
  var alerts = document.querySelectorAll(".alert");
  for (var k = 0; k < alerts.length; k++) {
    (function (a) {
      setTimeout(function () {
        a.style.transition = "opacity .5s";
        a.style.opacity = "0";
        setTimeout(function () { if (a.parentNode) { a.parentNode.removeChild(a); } }, 520);
      }, 6000);
    })(alerts[k]);
  }

  buildQueue();
})();
