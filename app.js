(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const fileInput = $("fileInput");
  const dropzone = $("dropzone");
  const messageEl = $("message");
  const monthSelect = $("monthSelect");
  const categoryFilter = $("categoryFilter");
  const periodLabel = $("periodLabel");
  const clearBtn = $("clearBtn");
  const tabs = $("tabs");
  const sections = {
    summary: $("dashboard"),
    shared: $("sharedSection"),
    tips: $("tipsSection"),
  };
  const partnerInput = $("partnerName");
  const sharedList = $("sharedList");
  const sharedSummary = $("sharedSummary");
  const tipsList = $("tipsList");

  let allRows = [];
  let currentPage = "summary";

  // Everything is stored in this browser only (localStorage), so it never leaves the device.
  const STORAGE_KEY = "expense-summary-v1";
  const SHARED_KEY = "expense-shared-v1";

  const fmt = (n) => "₪" + Math.round(n).toLocaleString("en-US");
  const MONTHS = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"];

  // ---------- storage ----------

  function loadSaved() {
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
      return saved && Array.isArray(saved.rows) ? saved : null;
    } catch (e) {
      return null;
    }
  }

  function rowFromSaved(r) {
    const date = new Date(r.date);
    return { ...r, date, monthKey: date.getFullYear() * 100 + date.getMonth() };
  }

  function rowToSaved(r) {
    return { ...r, date: r.date.toISOString() };
  }

  function rowKey(r) {
    return [r.date.toISOString(), r.description, r.amount, r.category].join("|");
  }

  function saveRows(rows, fileName) {
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({ updatedAt: new Date().toISOString(), fileName, rows: rows.map(rowToSaved) })
      );
    } catch (e) {
      showMessage("לא הצלחתי לשמור את הנתונים בדפדפן, הם יישארו רק עד שתסגרו את הדף.");
    }
  }

  function loadShared() {
    try {
      const s = JSON.parse(localStorage.getItem(SHARED_KEY) || "null");
      return s && typeof s === "object" ? { partner: s.partner || "", shares: s.shares || {} } : { partner: "", shares: {} };
    } catch (e) {
      return { partner: "", shares: {} };
    }
  }

  let shared = loadShared();

  function saveShared() {
    try {
      localStorage.setItem(SHARED_KEY, JSON.stringify(shared));
    } catch (e) { /* storage unavailable: keep working in memory */ }
  }

  // ---------- parsing ----------

  function showMessage(text) {
    messageEl.textContent = text;
    messageEl.hidden = !text;
  }

  // Find a column by name: exact match first (in the order of `words`), then partial match.
  function findKey(keys, words) {
    for (const w of words) {
      const exact = keys.find((k) => String(k).trim() === w);
      if (exact !== undefined) return exact;
    }
    return keys.find((k) => words.some((w) => String(k).includes(w)));
  }

  function parseAmount(value) {
    if (typeof value === "number") return value;
    const cleaned = String(value || "").replace(/[₪,\s]/g, "");
    const n = parseFloat(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  function parseDate(value) {
    if (value instanceof Date && !isNaN(value)) return value;
    if (typeof value === "number") {
      const d = new Date(Math.round((value - 25569) * 86400 * 1000));
      return isNaN(d) ? null : d;
    }
    const m = String(value || "").trim().match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
    if (m) {
      const year = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
      return new Date(year, Number(m[2]) - 1, Number(m[1]));
    }
    return null;
  }

  // Reads every sheet. In each sheet, the header row is the first row that has a date column
  // ("תאריך" or "תאריך עסקה"). Bank exports put a few title rows above it, so this skips them.
  function readWorkbook(buffer) {
    const wb = XLSX.read(buffer, { type: "array", cellDates: true });
    const records = [];
    for (const name of wb.SheetNames) {
      const grid = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "", raw: true });
      const headerIdx = grid.findIndex((row) =>
        row.some((c) => ["תאריך", "תאריך עסקה"].includes(String(c).trim()))
      );
      if (headerIdx < 0) continue;
      const headers = grid[headerIdx].map((h) => String(h).trim());
      for (const row of grid.slice(headerIdx + 1)) {
        const obj = {};
        headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
        records.push(obj);
      }
    }
    return records;
  }

  function normalize(rawRows) {
    const rows = [];
    for (const r of rawRows) {
      const keys = Object.keys(r);
      const kDate = findKey(keys, ["תאריך עסקה", "תאריך"]);
      const kAmount = findKey(keys, ["סכום חיוב", "סכום הוצאה", "סכום"]);
      const kDesc = findKey(keys, ["תיאור", "שם בית העסק"]);
      const kCat = findKey(keys, ["סוג הוצאה", "קטגוריה"]);
      const kPay = findKey(keys, ["אמצעי תשלום"]);

      if (!kDate || !kAmount) continue;
      const date = parseDate(r[kDate]);
      const amount = parseAmount(r[kAmount]);
      // Skips total rows, blank rows and rows without a valid date or amount.
      if (!date || !Number.isFinite(amount)) continue;

      rows.push({
        date,
        monthKey: date.getFullYear() * 100 + date.getMonth(),
        description: String(kDesc ? r[kDesc] : "").trim(),
        amount,
        category: String(kCat ? r[kCat] : "").trim() || "אחר",
        payment: String(kPay ? r[kPay] : "").trim(),
      });
    }
    return { rows };
  }

  // ---------- helpers ----------

  function prevMonthKey(key) {
    const y = Math.floor(key / 100);
    const m = key % 100;
    return m === 0 ? (y - 1) * 100 + 11 : key - 1;
  }

  function currentRows() {
    const key = Number(monthSelect.value);
    return allRows.filter((r) => r.monthKey === key);
  }

  function groupByCategory(rows) {
    const map = new Map();
    for (const r of rows) map.set(r.category, (map.get(r.category) || 0) + r.amount);
    return [...map.entries()].map(([name, total]) => ({ name, total })).sort((a, b) => b.total - a.total);
  }

  // ---------- month selector and pages ----------

  function fillMonths() {
    const keys = [...new Set(allRows.map((r) => r.monthKey))].sort((a, b) => b - a);
    const previous = monthSelect.value;
    monthSelect.innerHTML = "";
    for (const key of keys) {
      const opt = document.createElement("option");
      opt.value = String(key);
      opt.textContent = MONTHS[key % 100] + " " + Math.floor(key / 100);
      monthSelect.appendChild(opt);
    }
    if (keys.map(String).includes(previous)) monthSelect.value = previous;
    monthSelect.hidden = keys.length === 0;
  }

  function refreshVisibility() {
    const has = allRows.length > 0;
    tabs.hidden = !has;
    sections.summary.hidden = !has || currentPage !== "summary";
    sections.shared.hidden = !has || currentPage !== "shared";
    sections.tips.hidden = !has || currentPage !== "tips";
    document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("active", t.dataset.page === currentPage));
  }

  function showPage(name) {
    currentPage = name;
    refreshVisibility();
    render();
  }

  // ---------- summary page ----------

  function renderKpis(rows) {
    const total = rows.reduce((s, r) => s + r.amount, 0);
    const byCat = groupByCategory(rows);
    $("kpiTotal").textContent = fmt(total);
    $("kpiCount").textContent = rows.length.toLocaleString("en-US");
    $("kpiTop").textContent = byCat[0] ? byCat[0].name : "-";
    $("kpiAvg").textContent = rows.length ? fmt(total / rows.length) : "₪0";
  }

  function renderBars(rows) {
    const container = $("categoryBars");
    container.innerHTML = "";
    const cats = groupByCategory(rows);
    const max = cats.length ? cats[0].total : 1;
    for (const c of cats) {
      const row = document.createElement("div");
      row.className = "bar-row";

      const label = document.createElement("span");
      label.textContent = c.name;

      const track = document.createElement("div");
      track.className = "track";
      const fill = document.createElement("div");
      fill.className = "fill";
      fill.style.width = ((c.total / max) * 100).toFixed(1) + "%";
      track.appendChild(fill);

      const amount = document.createElement("span");
      amount.className = "bar-amount";
      amount.textContent = fmt(c.total);

      row.append(label, track, amount);
      container.appendChild(row);
    }
  }

  function fillCategoryFilter(rows) {
    const current = categoryFilter.value;
    const cats = groupByCategory(rows).map((c) => c.name);
    categoryFilter.innerHTML = "";
    const all = document.createElement("option");
    all.value = "";
    all.textContent = "כל הקטגוריות";
    categoryFilter.appendChild(all);
    for (const name of cats) {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      categoryFilter.appendChild(opt);
    }
    categoryFilter.value = cats.includes(current) ? current : "";
  }

  function renderTransactions(rows) {
    const body = $("txBody");
    body.innerHTML = "";
    const selected = categoryFilter.value;
    const list = rows
      .filter((r) => !selected || r.category === selected)
      .sort((a, b) => b.date - a.date);

    for (const r of list) {
      const tr = document.createElement("tr");

      const tdDate = document.createElement("td");
      tdDate.textContent = r.date.toLocaleDateString("he-IL");

      const tdDesc = document.createElement("td");
      tdDesc.textContent = r.description;

      const tdCat = document.createElement("td");
      const tag = document.createElement("span");
      tag.className = "tag";
      tag.textContent = r.category;
      tdCat.appendChild(tag);

      const tdAmt = document.createElement("td");
      tdAmt.className = "num";
      tdAmt.textContent = fmt(r.amount);

      tr.append(tdDate, tdDesc, tdCat, tdAmt);
      body.appendChild(tr);
    }
  }

  // ---------- shared expenses page ----------

  function renderShared(rows) {
    if (document.activeElement !== partnerInput) partnerInput.value = shared.partner;
    sharedList.innerHTML = "";

    const list = rows.slice().sort((a, b) => b.date - a.date);
    for (const r of list) {
      const key = rowKey(r);
      const pct = shared.shares[key];

      const line = document.createElement("div");
      line.className = "shared-row";

      const box = document.createElement("input");
      box.type = "checkbox";
      box.checked = pct !== undefined;
      box.id = "s" + sharedList.children.length;

      const label = document.createElement("label");
      label.htmlFor = box.id;
      label.className = "shared-main";
      const name = document.createElement("span");
      name.textContent = r.description || r.category;
      const meta = document.createElement("span");
      meta.className = "muted small";
      meta.textContent = r.date.toLocaleDateString("he-IL") + " · " + r.category;
      label.append(name, meta);

      const amt = document.createElement("span");
      amt.className = "num";
      amt.textContent = fmt(r.amount);

      const share = document.createElement("input");
      share.type = "number";
      share.min = "0";
      share.max = "100";
      share.step = "5";
      share.className = "pct";
      share.value = pct !== undefined ? pct : 50;
      share.disabled = !box.checked;
      share.title = "חלקו של השותף/ה באחוזים";

      box.addEventListener("change", () => {
        if (box.checked) shared.shares[key] = Number(share.value) || 50;
        else delete shared.shares[key];
        share.disabled = !box.checked;
        saveShared();
        renderSharedSummary(rows);
      });

      share.addEventListener("change", () => {
        const v = Math.min(100, Math.max(0, Number(share.value) || 0));
        share.value = v;
        if (box.checked) shared.shares[key] = v;
        saveShared();
        renderSharedSummary(rows);
      });

      line.append(box, label, amt, share);
      sharedList.appendChild(line);
    }
    renderSharedSummary(rows);
  }

  function renderSharedSummary(rows) {
    const marked = rows.filter((r) => shared.shares[rowKey(r)] !== undefined);
    const total = marked.reduce((s, r) => s + r.amount, 0);
    const owed = marked.reduce((s, r) => s + (r.amount * shared.shares[rowKey(r)]) / 100, 0);
    const who = shared.partner.trim() || "השותף/ה";
    sharedSummary.innerHTML = "";

    const lines = [
      ["הוצאות משותפות החודש", fmt(total) + " (" + marked.length + " עסקאות)"],
      [who + " חייב/ת לך", fmt(owed)],
    ];
    for (const [label, value] of lines) {
      const row = document.createElement("div");
      row.className = "sum-row";
      const l = document.createElement("span");
      l.className = "muted";
      l.textContent = label;
      const v = document.createElement("span");
      v.className = "kpi-value small-value";
      v.textContent = value;
      row.append(l, v);
      sharedSummary.appendChild(row);
    }
    const note = document.createElement("p");
    note.className = "muted small";
    note.textContent = "החישוב מניח שאתם שילמתם את כל ההוצאות שסומנו. ההתחשבנות נשמרת בדפדפן הזה.";
    sharedSummary.appendChild(note);
  }

  partnerInput.addEventListener("input", () => {
    shared.partner = partnerInput.value;
    saveShared();
    renderSharedSummary(currentRows());
  });

  // ---------- recommendations page ----------

  const SUB_WORDS = [
    "netflix", "spotify", "apple", "google", "youtube", "patreon", "disney", "hbo",
    "chatgpt", "openai", "claude", "adobe", "microsoft", "icloud", "amazon prime",
  ];

  function isSubscription(r) {
    const desc = r.description.toLowerCase();
    return r.category.includes("מנוי") || SUB_WORDS.some((w) => desc.includes(w));
  }

  function buildTips(rows) {
    const tips = [];
    const prevKey = prevMonthKey(Number(monthSelect.value));
    const prev = allRows.filter((r) => r.monthKey === prevKey);
    const total = rows.reduce((s, r) => s + r.amount, 0);

    // 1. Categories that rose a lot compared with the previous month.
    if (prev.length) {
      const prevCats = new Map(groupByCategory(prev).map((c) => [c.name, c.total]));
      for (const c of groupByCategory(rows)) {
        const before = prevCats.get(c.name) || 0;
        const diff = c.total - before;
        if (before > 0 && diff >= 200 && c.total > before * 1.2) {
          const pct = Math.round((diff / before) * 100);
          tips.push({
            title: "עלייה ב" + c.name,
            text: "ההוצאה על " + c.name + " עלתה ב-" + pct + "% (" + fmt(diff) + ") לעומת החודש הקודם. כדאי לבדוק מה השתנה.",
          });
        }
      }
    }

    // 2. Subscriptions.
    const subs = rows.filter(isSubscription);
    if (subs.length) {
      const subTotal = subs.reduce((s, r) => s + r.amount, 0);
      tips.push({
        title: "מנויים",
        text: "זוהו " + subs.length + " עסקאות של מנויים או חיובים חוזרים, בסך " + fmt(subTotal) + " החודש. כדאי לבדוק אילו מהם בשימוש בפועל.",
      });
    }

    // 3. Merchants charged several times in the same month.
    const byMerchant = new Map();
    for (const r of rows) {
      if (!r.description) continue;
      const m = byMerchant.get(r.description) || { count: 0, total: 0 };
      m.count += 1;
      m.total += r.amount;
      byMerchant.set(r.description, m);
    }
    const repeated = [...byMerchant.entries()].filter(([, m]) => m.count >= 3).sort((a, b) => b[1].total - a[1].total);
    if (repeated.length) {
      const [name, m] = repeated[0];
      tips.push({
        title: "חיובים חוזרים",
        text: name + " חויב/ה " + m.count + " פעמים החודש, בסך " + fmt(m.total) + ". כדאי לבדוק אם זו הוצאה שאפשר לצמצם.",
      });
    }

    // 4. Concentration on one merchant.
    const topMerchant = [...byMerchant.entries()].sort((a, b) => b[1].total - a[1].total)[0];
    if (topMerchant && topMerchant[1].total >= 300 && total > 0 && topMerchant[1].total / total >= 0.15) {
      const pct = Math.round((topMerchant[1].total / total) * 100);
      tips.push({
        title: "עסק מרכזי",
        text: topMerchant[0] + " מהווה " + pct + "% מכלל ההוצאות החודש (" + fmt(topMerchant[1].total) + ").",
      });
    }

    // 5. A concrete savings target: 10% of the largest category that is not housing.
    const SKIP_FOR_SAVINGS = ["דיור", "חשבונות", "העברת כספים"];
    const candidate = groupByCategory(rows).find((c) => !SKIP_FOR_SAVINGS.includes(c.name));
    if (candidate && candidate.total >= 500) {
      const saving = Math.round(candidate.total * 0.1);
      tips.push({
        title: "יעד לחיסכון",
        text: "קיצוץ של 10% ב" + candidate.name + " יחסוך " + fmt(saving) + " בחודש, כ-" + fmt(saving * 12) + " בשנה.",
      });
    }

    if (!tips.length) {
      tips.push({ title: "אין חריגות", text: "לא זיהיתי חריגות בחודש הזה לפי הכללים הפשוטים שלנו." });
    }
    return tips;
  }

  function renderTips(rows) {
    tipsList.innerHTML = "";
    for (const t of buildTips(rows)) {
      const card = document.createElement("div");
      card.className = "tip";
      const title = document.createElement("div");
      title.className = "tip-title";
      title.textContent = t.title;
      const text = document.createElement("div");
      text.className = "muted";
      text.textContent = t.text;
      card.append(title, text);
      tipsList.appendChild(card);
    }
    const note = document.createElement("p");
    note.className = "muted small";
    note.textContent = "ההמלצות מבוססות על כללים פשוטים מהנתונים שלכם, ואינן ייעוץ פיננסי.";
    tipsList.appendChild(note);
  }

  // ---------- render ----------

  function render() {
    const rows = currentRows();
    renderKpis(rows);
    renderBars(rows);
    fillCategoryFilter(rows);
    renderTransactions(rows);
    renderShared(rows);
    renderTips(rows);
    const label = monthSelect.options[monthSelect.selectedIndex];
    periodLabel.textContent = label ? "סיכום ל" + label.textContent : "";
    refreshVisibility();
  }

  // ---------- upload ----------

  async function handleFile(file) {
    showMessage("");
    try {
      const buffer = await file.arrayBuffer();
      const raw = readWorkbook(buffer);
      if (!raw.length) {
        throw new Error("לא מצאתי בקובץ שורת כותרות עם עמודת תאריך.");
      }
      const { rows } = normalize(raw);
      if (!rows.length) {
        throw new Error("לא נמצאו שורות הוצאה תקינות בקובץ.");
      }
      // Merge with what is already saved, so each monthly upload adds to the previous ones.
      const saved = loadSaved();
      const existing = saved ? saved.rows.map(rowFromSaved) : [];
      const seen = new Set(existing.map(rowKey));
      const fresh = rows.filter((r) => !seen.has(rowKey(r)));
      allRows = existing.concat(fresh);
      saveRows(allRows, file.name);

      fillMonths();
      render();
      showMessage(
        fresh.length
          ? "נוספו " + fresh.length + " עסקאות חדשות. סה״כ שמורות: " + allRows.length + "."
          : "לא נוספו עסקאות חדשות, כולן כבר שמורות."
      );
    } catch (err) {
      showMessage(err.message || "לא הצלחתי לקרוא את הקובץ.");
    }
  }

  clearBtn.addEventListener("click", () => {
    if (!confirm("למחוק את כל ההוצאות השמורות בדפדפן הזה?")) return;
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(SHARED_KEY);
    shared = { partner: "", shares: {} };
    allRows = [];
    clearBtn.hidden = true;
    monthSelect.innerHTML = "";
    monthSelect.hidden = true;
    render();
    showMessage("הנתונים השמורים נמחקו.");
  });

  document.querySelectorAll(".tab").forEach((t) =>
    t.addEventListener("click", () => showPage(t.dataset.page))
  );

  // On load, show the saved expenses if there are any.
  (function restoreSaved() {
    const saved = loadSaved();
    if (!saved || !saved.rows.length) return;
    allRows = saved.rows.map(rowFromSaved);
    clearBtn.hidden = false;
    fillMonths();
    render();
  })();

  fileInput.addEventListener("change", (e) => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
  });
  monthSelect.addEventListener("change", render);
  categoryFilter.addEventListener("change", () => renderTransactions(currentRows()));

  ["dragenter", "dragover"].forEach((evt) =>
    dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.add("drag"); })
  );
  ["dragleave", "drop"].forEach((evt) =>
    dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.remove("drag"); })
  );
  dropzone.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
})();
