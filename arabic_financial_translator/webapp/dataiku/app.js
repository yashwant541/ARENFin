/* Dataiku standard web app — JavaScript tab (modern redesign).
   Same backend contract; richer UI: animated progress, SVG pills, hero stats. */
(function () {
  "use strict";

  var BK = (typeof getWebAppBackendUrl === "function")
    ? getWebAppBackendUrl
    : function (p) { return p; };

  var $ = function (id) { return document.getElementById(id); };
  var selected = [];
  var lastJob = null;

  var el = {
    drop: $("afxDrop"), input: $("afxInput"),
    queueWrap: $("afxQueueWrap"), list: $("afxList"), count: $("afxCount"),
    convert: $("afxConvert"), clear: $("afxClear"),
    summary: $("afxSummary"), sumDone: $("afxSumDone"), sumFlip: $("afxSumFlip"),
    sumReview: $("afxSumReview"), sumFail: $("afxSumFail"), sumFailWrap: $("afxSumFailWrap"),
    downloadAll: $("afxDownloadAll"),
    health: $("afxHealth"), healthText: $("afxHealthText"),
    statTerms: $("afxStatTerms"), statVariants: $("afxStatVariants"), statCats: $("afxStatCats"),
  };

  function icon(id) { return '<svg><use href="#' + id + '"/></svg>'; }
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function humanSize(n) {
    if (n < 1024) return n + " B";
    if (n < 1048576) return (n / 1024).toFixed(1) + " KB";
    return (n / 1048576).toFixed(1) + " MB";
  }
  function extOf(name) { var m = /\.([^.]+)$/.exec(name || ""); return m ? m[1].toLowerCase() : ""; }
  function keyOf(f) { return f.name + "|" + f.size; }
  var ALLOWED = ["xlsx", "xlsm", "xls", "csv", "xltx"];

  // ---- health / hero stats ----------------------------------------------
  fetch(BK("/api/health"))
    .then(function (r) { return r.json(); })
    .then(function (h) {
      var g = h.glossary || {};
      el.health.classList.add("ok");
      el.healthText.textContent = "glossary ready";
      countUp(el.statTerms, g.entries || 0);
      countUp(el.statVariants, g.arabic_variants || 0);
      countUp(el.statCats, g.categories || 0);
    })
    .catch(function () {
      el.health.classList.add("bad");
      el.healthText.textContent = "backend offline";
    });

  function countUp(node, target) {
    if (!node) return;
    var start = performance.now(), dur = 700;
    function step(t) {
      var p = Math.min(1, (t - start) / dur);
      node.textContent = Math.round(target * (1 - Math.pow(1 - p, 3))).toLocaleString();
      if (p < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
  }

  // ---- selection ---------------------------------------------------------
  function addFiles(fileList) {
    Array.prototype.forEach.call(fileList, function (f) {
      if (ALLOWED.indexOf(extOf(f.name)) === -1) return;
      var k = keyOf(f);
      if (selected.some(function (s) { return s.key === k; })) return;
      selected.push({ key: k, file: f });
    });
    render();
  }
  function removeKey(k) { selected = selected.filter(function (s) { return s.key !== k; }); render(); }

  function render() {
    el.count.textContent = selected.length;
    el.queueWrap.hidden = selected.length === 0;
    el.summary.hidden = true;
    el.downloadAll.hidden = true;
    el.list.innerHTML = "";
    selected.forEach(function (s) {
      var ext = extOf(s.file.name);
      var li = document.createElement("li");
      li.className = "afx-row " + ext;
      li.dataset.key = s.key;
      li.dataset.name = s.file.name;
      li.innerHTML =
        '<div class="afx-row-icon">' + esc(ext.toUpperCase().slice(0, 4)) + '</div>' +
        '<div class="afx-row-body">' +
          '<div class="afx-row-name" title="' + esc(s.file.name) + '">' + esc(s.file.name) + '</div>' +
          '<div class="afx-row-meta" data-role="meta">' +
            '<span>' + humanSize(s.file.size) + '</span>' +
            '<span class="afx-dotsep">·</span>' +
            '<span>ready</span>' +
          '</div>' +
        '</div>' +
        '<div class="afx-row-actions" data-role="actions">' +
          '<button class="afx-icon-btn" title="Remove" data-remove="' + esc(s.key) + '">' + icon("i-close") + '</button>' +
        '</div>';
      el.list.appendChild(li);
    });
  }

  // ---- convert -----------------------------------------------------------
  function startProcessingUI() {
    Array.prototype.forEach.call(el.list.querySelectorAll(".afx-row"), function (li) {
      var meta = li.querySelector('[data-role="meta"]');
      var size = humanSize(sizeFor(li.dataset.name));
      meta.innerHTML =
        '<span>' + size + '</span><span class="afx-dotsep">·</span>' +
        '<span class="afx-pill proc"><span class="afx-spin"></span> converting…</span>' +
        '<div class="afx-prog"><div class="afx-prog-bar indet"></div></div>';
      var actions = li.querySelector('[data-role="actions"]');
      actions.innerHTML = "";
    });
  }

  function sizeFor(name) {
    var s = selected.find(function (x) { return x.file.name === name; });
    return s ? s.file.size : 0;
  }

  function applyResults(job) {
    lastJob = job.job_id;
    var byName = {};
    (job.results || []).forEach(function (r) { byName[r.name] = r; });

    Array.prototype.forEach.call(el.list.querySelectorAll(".afx-row"), function (li) {
      var name = li.dataset.name;
      var r = byName[name];
      var meta = li.querySelector('[data-role="meta"]');
      var actions = li.querySelector('[data-role="actions"]');
      var size = humanSize(sizeFor(name));
      if (!r) return;

      if (r.status === "done") {
        var badges =
          '<span class="afx-pill done">' + icon("i-check") + ' done</span>' +
          '<span class="afx-badge">' + icon("i-sheet") + r.sheets + ' sheet' + (r.sheets === 1 ? '' : 's') + '</span>' +
          (r.flipped ? '<span class="afx-badge flip">' + icon("i-flip") + r.flipped + ' RTL flipped</span>' : '') +
          (r.review ? '<span class="afx-badge review">' + icon("i-flag") + r.review + ' to review</span>'
                    : '<span class="afx-badge clean">' + icon("i-check") + ' all matched</span>');
        meta.innerHTML = '<span>' + size + '</span><span class="afx-dotsep">·</span>' + badges +
          '<div class="afx-prog"><div class="afx-prog-bar" style="width:100%"></div></div>';
        if (r.sample_terms && r.sample_terms.length) {
          var terms = document.createElement("div");
          terms.className = "afx-terms";
          terms.innerHTML = "Add to glossary: " +
            r.sample_terms.slice(0, 5).map(function (t) { return "<code>" + esc(t) + "</code>"; }).join(" ");
          li.querySelector(".afx-row-body").appendChild(terms);
        }
        actions.innerHTML =
          '<a class="afx-dl" href="' + BK("/api/download/" + job.job_id + "/" + r.fid) + '" download>' +
          icon("i-download") + 'Download</a>';
      } else {
        meta.innerHTML = '<span>' + size + '</span><span class="afx-dotsep">·</span>' +
          '<span class="afx-pill err">' + icon("i-alert") + esc(r.message || "failed") + '</span>';
        actions.innerHTML = "";
      }
    });

    var flips = (job.results || []).reduce(function (a, r) { return a + (r.flipped || 0); }, 0);
    countUp(el.sumDone, job.succeeded);
    countUp(el.sumFlip, flips);
    countUp(el.sumReview, job.total_review);
    countUp(el.sumFail, job.failed);
    el.sumFailWrap.hidden = !job.failed;
    el.summary.hidden = false;
    el.downloadAll.hidden = job.succeeded === 0;
    el.convert.disabled = false;
    el.convert.querySelector(".afx-btn-label").textContent = "Convert all";
  }

  function convertAll() {
    if (!selected.length) return;
    el.convert.disabled = true;
    el.convert.querySelector(".afx-btn-label").textContent = "Converting…";
    startProcessingUI();

    var fd = new FormData();
    selected.forEach(function (s) { fd.append("files", s.file, s.file.name); });

    fetch(BK("/api/convert"), { method: "POST", body: fd })
      .then(function (r) { return r.json(); })
      .then(function (job) {
        if (job.error) throw new Error(job.error);
        applyResults(job);
      })
      .catch(function (e) {
        Array.prototype.forEach.call(el.list.querySelectorAll('[data-role="meta"]'), function (m) {
          m.innerHTML = '<span class="afx-pill err">' + icon("i-alert") + esc(e.message || "request failed") + '</span>';
        });
        el.convert.disabled = false;
        el.convert.querySelector(".afx-btn-label").textContent = "Convert all";
      });
  }

  // ---- wiring ------------------------------------------------------------
  el.drop.addEventListener("click", function () { el.input.click(); });
  el.drop.addEventListener("keydown", function (e) {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); el.input.click(); }
  });
  el.input.addEventListener("change", function (e) { addFiles(e.target.files); el.input.value = ""; });

  ["dragenter", "dragover"].forEach(function (ev) {
    el.drop.addEventListener(ev, function (e) { e.preventDefault(); el.drop.classList.add("drag"); });
  });
  ["dragleave", "drop"].forEach(function (ev) {
    el.drop.addEventListener(ev, function (e) { e.preventDefault(); el.drop.classList.remove("drag"); });
  });
  el.drop.addEventListener("drop", function (e) {
    if (e.dataTransfer && e.dataTransfer.files) addFiles(e.dataTransfer.files);
  });

  el.list.addEventListener("click", function (e) {
    var btn = e.target.closest ? e.target.closest("[data-remove]") : null;
    if (btn) removeKey(btn.getAttribute("data-remove"));
  });
  el.clear.addEventListener("click", function () { selected = []; render(); });
  el.convert.addEventListener("click", convertAll);
  el.downloadAll.addEventListener("click", function () {
    if (lastJob) window.location = BK("/api/download_all/" + lastJob);
  });
})();
