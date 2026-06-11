(function () {
  const data = window.CORE_DEADLINE_DATA || { generated_at: "", sheets: [] };
  const config = window.CORE_SITE_CONFIG || {};
  const urgentDays = Number.isFinite(Number(config.urgentDays)) ? Number(config.urgentDays) : 10;
  const soonDays = Number.isFinite(Number(config.soonDays)) ? Number(config.soonDays) : 30;
  const rankOrder = new Map([
    ["A*", 0],
    ["A", 1],
    ["B", 2],
    ["Australasian B", 2],
    ["C", 3],
    ["Australasian C", 3],
  ]);

  const els = {
    generatedAt: document.getElementById("generatedAt"),
    sheetSelect: document.getElementById("sheetSelect"),
    searchInput: document.getElementById("searchInput"),
    rankSelect: document.getElementById("rankSelect"),
    deadlineSelect: document.getElementById("deadlineSelect"),
    sortSelect: document.getElementById("sortSelect"),
    visibleCount: document.getElementById("visibleCount"),
    urgentCount: document.getElementById("urgentCount"),
    urgentLabel: document.getElementById("urgentLabel"),
    upcomingCount: document.getElementById("upcomingCount"),
    missingCount: document.getElementById("missingCount"),
    topDeadline: document.getElementById("topDeadline"),
    tableBody: document.getElementById("tableBody"),
  };

  const fields = {
    conference: "Conference",
    acronym: "Acronym",
    rank: "Rank",
    next: "Next Deadline",
    abstract: "Abstract Deadline",
    submission: "Submission Deadline",
    pages: "Page Limit / Format",
    url: "URL",
  };

  function parseDeadline(value) {
    if (!value) return null;
    const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return null;
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 23, 59, 59);
  }

  function countdown(value) {
    const deadline = parseDeadline(value);
    if (!deadline) return "";

    const now = new Date();
    let seconds = Math.floor((deadline.getTime() - now.getTime()) / 1000);
    const overdue = seconds < 0;
    seconds = Math.abs(seconds);

    const days = Math.floor(seconds / 86400);
    const hours = Math.floor((seconds % 86400) / 3600);
    const minutes = Math.floor((seconds % 3600) / 60);

    let label = "";
    if (days > 0) label = `${days} days ${hours} hours`;
    else if (hours > 0) label = `${hours} hours ${minutes} minutes`;
    else label = `${minutes} minutes`;

    return overdue ? `overdue by ${label}` : label;
  }

  function deadlineState(value) {
    const deadline = parseDeadline(value);
    if (!deadline) return "missing";
    const now = new Date();
    const days = Math.floor((deadline.getTime() - now.getTime()) / 86400000);
    if (days < 0) return "overdue";
    if (days < urgentDays) return "urgent";
    if (days <= soonDays) return "soon";
    return "upcoming";
  }

  function daysUntil(value) {
    const deadline = parseDeadline(value);
    if (!deadline) return null;
    return Math.floor((deadline.getTime() - new Date().getTime()) / 86400000);
  }

  function deadlineSortKey(value) {
    const deadline = parseDeadline(value);
    if (!deadline) return [2, 0];

    const timestamp = deadline.getTime();
    if (timestamp >= Date.now()) return [0, timestamp];
    return [1, -timestamp];
  }

  function currentSheet() {
    return data.sheets.find((sheet) => sheet.name === els.sheetSelect.value) || data.sheets[0] || { rows: [] };
  }

  function rankScore(row) {
    const rank = row[fields.rank] || "";
    return rankOrder.has(rank) ? rankOrder.get(rank) : 99;
  }

  function normalizedRows() {
    const query = els.searchInput.value.trim().toLowerCase();
    const rank = els.rankSelect.value;
    const deadlineFilter = els.deadlineSelect.value;

    return currentSheet().rows.filter((row) => {
      const state = deadlineState(row[fields.next]);
      const haystack = [
        row[fields.conference],
        row[fields.acronym],
        row[fields.rank],
        row[fields.pages],
        row[fields.submission],
        row[fields.abstract],
      ]
        .join(" ")
        .toLowerCase();

      if (query && !haystack.includes(query)) return false;
      if (rank && row[fields.rank] !== rank) return false;
      if (deadlineFilter === "upcoming" && !["upcoming", "soon", "urgent"].includes(state)) return false;
      if (deadlineFilter === "missing" && state !== "missing") return false;
      if (deadlineFilter === "overdue" && state !== "overdue") return false;
      return true;
    });
  }

  function sortRows(rows) {
    const mode = els.sortSelect.value;
    rows.sort((a, b) => {
      if (mode === "name") return String(a[fields.conference] || "").localeCompare(String(b[fields.conference] || ""));
      if (mode === "rank") {
        return rankScore(a) - rankScore(b) || String(a[fields.conference] || "").localeCompare(String(b[fields.conference] || ""));
      }

      const aKey = deadlineSortKey(a[fields.next]);
      const bKey = deadlineSortKey(b[fields.next]);
      return (
        aKey[0] - bKey[0] ||
        aKey[1] - bKey[1] ||
        String(a[fields.conference] || "").localeCompare(String(b[fields.conference] || ""))
      );
    });
    return rows;
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function renderTable(rows) {
    if (!rows.length) {
      els.tableBody.innerHTML = `<tr><td class="empty" colspan="9">No conferences match the current filters.</td></tr>`;
      return;
    }

    els.tableBody.innerHTML = rows
      .map((row) => {
        const state = deadlineState(row[fields.next]);
        const days = daysUntil(row[fields.next]);
        const url = row[fields.url] || "";
        const countdownText = countdown(row[fields.next]);
        const stateLabel = {
          urgent: "Urgent",
          soon: "Soon",
          upcoming: "Upcoming",
          overdue: "Overdue",
          missing: "Missing",
        }[state];
        const urgencyMeta = days === null ? "" : days < 0 ? "Past deadline" : `${days + 1} calendar days left`;
        const urlHtml = url ? `<a class="open-link" href="${escapeHtml(url)}" target="_blank" rel="noreferrer">Open</a>` : `<span class="muted">-</span>`;
        return `
          <tr class="row-${state}">
            <td class="conf">${escapeHtml(row[fields.conference])}</td>
            <td>${escapeHtml(row[fields.acronym]) || `<span class="muted">-</span>`}</td>
            <td><span class="rank-pill rank-${escapeHtml(String(row[fields.rank] || "other").replace(/[^A-Za-z0-9]+/g, "-"))}">${escapeHtml(row[fields.rank]) || "-"}</span></td>
            <td class="deadline-cell">
              <span class="deadline ${state}">${escapeHtml(row[fields.next]) || "-"}</span>
              <span class="deadline-state ${state}">${stateLabel}</span>
            </td>
            <td>
              ${
                countdownText
                  ? `<span class="countdown ${state}">${escapeHtml(countdownText)}</span><span class="countdown-meta">${escapeHtml(urgencyMeta)}</span>`
                  : `<span class="muted">-</span>`
              }
            </td>
            <td>${escapeHtml(row[fields.submission]) || `<span class="muted">-</span>`}</td>
            <td>${escapeHtml(row[fields.abstract]) || `<span class="muted">-</span>`}</td>
            <td>${row[fields.pages] ? `<span class="format-chip">${escapeHtml(row[fields.pages])}</span>` : `<span class="muted">-</span>`}</td>
            <td>${urlHtml}</td>
          </tr>
        `;
      })
      .join("");
  }

  function renderStats(rows) {
    const urgent = rows.filter((row) => deadlineState(row[fields.next]) === "urgent").length;
    const upcoming = rows.filter((row) => ["upcoming", "soon", "urgent"].includes(deadlineState(row[fields.next]))).length;
    const missing = rows.filter((row) => deadlineState(row[fields.next]) === "missing").length;
    const nearest = rows.find((row) => parseDeadline(row[fields.next]));

    els.visibleCount.textContent = rows.length;
    els.urgentCount.textContent = urgent;
    els.urgentLabel.textContent = `under ${urgentDays} days`;
    els.upcomingCount.textContent = upcoming;
    els.missingCount.textContent = missing;
    els.topDeadline.textContent = nearest ? nearest[fields.next] : "-";
  }

  function refreshRanks() {
    const sheet = currentSheet();
    const ranks = Array.from(new Set(sheet.rows.map((row) => row[fields.rank]).filter(Boolean)));
    ranks.sort((a, b) => (rankOrder.get(a) ?? 99) - (rankOrder.get(b) ?? 99) || a.localeCompare(b));

    const selected = els.rankSelect.value;
    els.rankSelect.innerHTML = `<option value="">All ranks</option>` + ranks.map((rank) => `<option value="${escapeHtml(rank)}">${escapeHtml(rank)}</option>`).join("");
    if (ranks.includes(selected)) els.rankSelect.value = selected;
  }

  function render() {
    refreshRanks();
    const rows = sortRows(normalizedRows());
    renderStats(rows);
    renderTable(rows);
  }

  function init() {
    els.generatedAt.textContent = data.generated_at ? `Generated ${data.generated_at}` : "No generated timestamp";
    els.sheetSelect.innerHTML = data.sheets.map((sheet) => `<option value="${escapeHtml(sheet.name)}">${escapeHtml(sheet.name)}</option>`).join("");

    ["change", "input"].forEach((eventName) => {
      els.sheetSelect.addEventListener(eventName, render);
      els.searchInput.addEventListener(eventName, render);
      els.rankSelect.addEventListener(eventName, render);
      els.deadlineSelect.addEventListener(eventName, render);
      els.sortSelect.addEventListener(eventName, render);
    });

    render();
    window.setInterval(render, 60000);
  }

  init();
})();
