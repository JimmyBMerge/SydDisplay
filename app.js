(function () {
  "use strict";

  const CONFIG = {
    dataUrl: "https://mergeitgroup.sharepoint.com/sites/TeamMergeTechnologies/SiteAssets/sydplan.json",
    refreshMs: 30 * 60 * 1000,
    managerOne: "Chris Burton",
    managerTwo: "Matt Bowcock"
  };

  function el(id) { return document.getElementById(id); }
  function esc(v) {
    return String(v == null ? "" : v)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }
  function val(row, key, fallback) {
    if (row && row[key] != null) return row[key];
    for (let i = 0; i < fallback.length; i++) if (row && row[fallback[i]] != null) return row[fallback[i]];
    return "";
  }
  function date(v) { const d = new Date(v); return isNaN(d.getTime()) ? null : d; }
  function day(d) { return Math.floor(d.getTime() / 86400000); }
  function clamp(v, min, max) { return Math.max(min, Math.min(max, v)); }
  function fmtDate(d) {
    if (!d) return "--/--/--";
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yy = String(d.getFullYear()).slice(-2);
    return `${dd}/${mm}/${yy}`;
  }
  function endSortAsc(a, b) {
    const da = date(val(a, "end", ["DueDate", "EndDate"]));
    const db = date(val(b, "end", ["DueDate", "EndDate"]));
    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;
    return da.getTime() - db.getTime();
  }
  function progress(v) {
    if (v == null || v === "") return null;
    const raw = String(v).trim();
    const n = Number(raw.replace("%", ""));
    if (isNaN(n)) return null;
    return clamp(raw.includes("%") ? n : (n <= 1 ? n * 100 : n), 0, 100);
  }
  function colorProgress(p) {
    if (p == null) return "#8a8886";
    if (p >= 80) return "#107c10";
    if (p >= 40) return "#005a9e";
    return "#a4262c";
  }
  function normalize(payload) {
    if (Array.isArray(payload)) return payload;
    if (payload && Array.isArray(payload.rows)) return payload.rows;
    if (payload && Array.isArray(payload.value)) return payload.value;
    return [];
  }

  function timeline(rows, managerName) {
    let min = null, max = null;
    rows.forEach(r => {
      let s = date(val(r, "start", ["StartDate", "Start"]));
      let e = date(val(r, "end", ["DueDate", "EndDate"]));
      if (!s && e) s = e;
      if (!e && s) e = s;
      if (!s || !e) return;
      if (!min || s < min) min = s;
      if (!max || e > max) max = e;
    });
    if (!min || !max) return `<div class="timeline-card"><h3>${esc(managerName)}</h3><div>No valid dates.</div></div>`;

    const minDay = day(min), total = Math.max(1, day(max) - minDay + 1);
    const lines = rows.map(r => {
      const t = val(r, "task", ["Task", "Title"]) || "(Untitled)";
      const p = progress(val(r, "progress", ["Progress", "PercentComplete"]));
      let s = date(val(r, "start", ["StartDate", "Start"]));
      let e = date(val(r, "end", ["DueDate", "EndDate"]));
      if (!s && e) s = e;
      if (!e && s) e = s;
      if (!s || !e) return "";
      const left = clamp(((day(s) - minDay) / total) * 100, 0, 100);
      const right = clamp(((day(e) - minDay + 1) / total) * 100, 0, 100);
      const width = Math.max(1, right - left);
      return `<div class="timeline-row"><div class="task-name">${esc(t)} <span class="task-dates">(${esc(fmtDate(s))} - ${esc(fmtDate(e))})</span></div><div class="task-sub">Progress: ${esc(p == null ? "N/A" : p.toFixed(0) + "%")}</div><div class="timeline-track"><div class="timeline-bar" style="left:${left.toFixed(2)}%;width:${width.toFixed(2)}%;background:${colorProgress(p)};"></div></div></div>`;
    }).join("");

    return `<div class="timeline-card"><h3>${esc(managerName)}</h3><div class="timeline-axis"><span>${esc(fmtDate(min))}</span><span>${esc(fmtDate(max))}</span></div>${lines}</div>`;
  }

  async function render() {
    const status = el("status"), timelines = el("timelines");
    try {
      status.textContent = "Loading...";
      const url = `${CONFIG.dataUrl}${CONFIG.dataUrl.includes("?") ? "&" : "?"}t=${Date.now()}`;
      const res = await fetch(url, { cache: "no-store" });
      if (!res.ok) throw new Error(`HTTP ${res.status} ${res.statusText}`);
      const rows = normalize(await res.json());
      const m1 = CONFIG.managerOne.toLowerCase();
      const m2 = CONFIG.managerTwo.toLowerCase();
      const byManager = rows.filter(r => {
        const m = String(val(r, "manager", ["ProjectManager", "Manager"])).toLowerCase();
        return m === m1 || m === m2;
      });
      const rows1 = byManager.filter(r => String(val(r, "manager", ["ProjectManager", "Manager"])).toLowerCase() === m1);
      const rows2 = byManager.filter(r => String(val(r, "manager", ["ProjectManager", "Manager"])).toLowerCase() === m2);
      rows1.sort(endSortAsc);
      rows2.sort(endSortAsc);

      status.textContent = `${fmtDate(new Date())}`;
      timelines.innerHTML = timeline(rows1, CONFIG.managerOne) + timeline(rows2, CONFIG.managerTwo);
    } catch (e) {
      status.innerHTML = `<span class="error">Failed to load data: ${esc(e && e.message ? e.message : e)}</span>`;
    }
  }

  render();
  window.setInterval(render, CONFIG.refreshMs);
})();
