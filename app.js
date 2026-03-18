(function () {
  "use strict";

  const CONFIG = {
    dataUrl: "/sydplan.json",
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
  function progress(v) {
    if (v == null || v === "") return null;
    const raw = String(v).trim();
    const n = Number(raw.replace("%", ""));
    if (isNaN(n)) return null;
    return clamp(raw.includes("%") ? n : (n <= 1 ? n * 100 : n), 0, 100);
  }
  function band(p) {
    const n = Number(p);
    if (isNaN(n)) return "Unspecified";
    if (n === 1) return "Urgent";
    if (n <= 4) return "Important";
    if (n <= 7) return "Medium";
    if (n <= 10) return "Low";
    return "Unspecified";
  }
  function colorPriority(b) {
    if (b === "Urgent") return "#a4262c";
    if (b === "Important") return "#ff8c00";
    if (b === "Medium") return "#005a9e";
    if (b === "Low") return "#107c10";
    return "#8a8886";
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
      const b = band(val(r, "priority", ["Priority"]));
      let s = date(val(r, "start", ["StartDate", "Start"]));
      let e = date(val(r, "end", ["DueDate", "EndDate"]));
      if (!s && e) s = e;
      if (!e && s) e = s;
      if (!s || !e) return "";
      const left = clamp(((day(s) - minDay) / total) * 100, 0, 100);
      const right = clamp(((day(e) - minDay + 1) / total) * 100, 0, 100);
      const width = Math.max(1, right - left);
      return `<div class="timeline-row"><div class="task-name">${esc(t)}</div><div class="task-sub">Progress: ${esc(p == null ? "N/A" : p.toFixed(0) + "%")} | Priority: ${esc(b)}</div><div class="timeline-track"><div class="timeline-bar" style="left:${left.toFixed(2)}%;width:${width.toFixed(2)}%;background:${colorPriority(b)};"></div></div></div>`;
    }).join("");

    return `<div class="timeline-card"><h3>${esc(managerName)}</h3><div class="timeline-axis"><span>${esc(min.toLocaleDateString())}</span><span>${esc(max.toLocaleDateString())}</span></div>${lines}</div>`;
  }

  function chart(rows, managerName) {
    const counts = { Urgent: 0, Important: 0, Medium: 0, Low: 0, Unspecified: 0 };
    let total = 0, items = 0;
    rows.forEach(r => {
      const b = band(val(r, "priority", ["Priority"]));
      counts[b] = (counts[b] || 0) + 1;
      const p = progress(val(r, "progress", ["Progress", "PercentComplete"]));
      if (p != null) { total += p; items++; }
    });
    const labels = Object.keys(counts).filter(k => counts[k] > 0);
    const max = Math.max(1, ...labels.map(k => counts[k]));
    const avg = items ? Math.round(total / items) : null;
    const bars = labels.map(k => `<div class="bar-row"><div>${esc(k)}</div><div class="bar-track"><div class="bar-fill" style="width:${((counts[k] / max) * 100).toFixed(2)}%;background:${colorPriority(k)};"></div></div><div>${counts[k]}</div></div>`).join("");
    return `<div class="chart-card"><h3>${esc(managerName)}</h3><div class="avg-label">Average Progress: ${esc(avg == null ? "N/A" : avg + "%")}</div><div class="avg-track"><div class="avg-fill" style="width:${avg == null ? 0 : avg}%;background:${colorProgress(avg)};"></div></div>${bars || "<div>No rows.</div>"}</div>`;
  }

  function table(rows) {
    const tableEl = el("dataTable");
    const keys = {};
    rows.forEach(r => Object.keys(r || {}).forEach(k => { keys[k] = true; }));
    const cols = Object.keys(keys);
    const thead = `<thead><tr>${cols.map(c => `<th>${esc(c)}</th>`).join("")}</tr></thead>`;
    const tbody = `<tbody>${rows.map(r => `<tr>${cols.map(c => {
      const raw = r[c];
      let out = raw;
      if (typeof raw === "object") { try { out = JSON.stringify(raw); } catch (e) { out = "[object]"; } }
      return `<td>${esc(out)}</td>`;
    }).join("")}</tr>`).join("")}</tbody>`;
    tableEl.innerHTML = thead + tbody;
  }

  async function render() {
    const meta = el("meta"), timelines = el("timelines"), charts = el("charts");
    try {
      meta.textContent = "Loading...";
      const res = await fetch(CONFIG.dataUrl, { cache: "no-store" });
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

      meta.textContent = `Filtered rows: ${byManager.length} | Last updated: ${new Date().toLocaleString()}`;
      timelines.innerHTML = timeline(rows1, CONFIG.managerOne) + timeline(rows2, CONFIG.managerTwo);
      charts.innerHTML = chart(rows1, CONFIG.managerOne) + chart(rows2, CONFIG.managerTwo);
      table(byManager);
    } catch (e) {
      meta.innerHTML = `<span class="error">Failed to load data: ${esc(e && e.message ? e.message : e)}</span>`;
    }
  }

  render();
  window.setInterval(render, CONFIG.refreshMs);
})();
