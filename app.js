(function () {
  "use strict";

  const fileInput = document.getElementById("fileInput");
  const dropzone = document.getElementById("dropzone");
  const messageEl = document.getElementById("message");
  const dashboard = document.getElementById("dashboard");
  const monthSelect = document.getElementById("monthSelect");
  const categoryFilter = document.getElementById("categoryFilter");
  const periodLabel = document.getElementById("periodLabel");

  let allRows = [];

  const fmt = (n) => "₪" + Math.round(n).toLocaleString("en-US");
  const MONTHS = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"];

  function showMessage(text) {
    messageEl.textContent = text;
    messageEl.hidden = !text;
  }

  // Map a header row to our columns by matching Hebrew keywords, so small wording differences still work.
  function findKey(keys, words) {
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
      // Excel serial date
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

  function readWorkbook(buffer) {
    const wb = XLSX.read(buffer, { type: "array", cellDates: true });
    const sheetName = wb.SheetNames.includes("הוצאות") ? "הוצאות" : wb.SheetNames[0];
    return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });
  }

  function normalize(rawRows) {
    if (!rawRows.length) return { rows: [], skipped: 0 };
    const keys = Object.keys(rawRows[0]);
    const kDate = findKey(keys, ["תאריך"]);
    const kDesc = findKey(keys, ["תיאור"]);
    const kAmount = findKey(keys, ["סכום"]);
    const kCat = findKey(keys, ["סוג", "קטגוריה"]);
    const kPay = findKey(keys, ["אמצעי"]);

    if (!kDate || !kAmount) {
      throw new Error("לא מצאתי עמודות של תאריך וסכום בקובץ. בדקו את שורת הכותרות.");
    }

    const rows = [];
    let skipped = 0;
    for (const r of rawRows) {
      const date = parseDate(r[kDate]);
      const amount = parseAmount(r[kAmount]);
      if (!date || !Number.isFinite(amount)) {
        skipped++;
        continue;
      }
      rows.push({
        date,
        monthKey: date.getFullYear() * 100 + date.getMonth(),
        description: String(kDesc ? r[kDesc] : "").trim(),
        amount,
        category: String(kCat ? r[kCat] : "").trim() || "אחר",
        payment: String(kPay ? r[kPay] : "").trim(),
      });
    }
    return { rows, skipped };
  }

  function fillMonths() {
    const keys = [...new Set(allRows.map((r) => r.monthKey))].sort((a, b) => b - a);
    monthSelect.innerHTML = "";
    for (const key of keys) {
      const opt = document.createElement("option");
      const year = Math.floor(key / 100);
      const month = key % 100;
      opt.value = String(key);
      opt.textContent = MONTHS[month] + " " + year;
      monthSelect.appendChild(opt);
    }
    monthSelect.hidden = keys.length === 0;
  }

  function currentRows() {
    const key = Number(monthSelect.value);
    return allRows.filter((r) => r.monthKey === key);
  }

  function renderKpis(rows) {
    const total = rows.reduce((s, r) => s + r.amount, 0);
    const byCat = groupByCategory(rows);
    const top = byCat[0];
    document.getElementById("kpiTotal").textContent = fmt(total);
    document.getElementById("kpiCount").textContent = rows.length.toLocaleString("en-US");
    document.getElementById("kpiTop").textContent = top ? top.name : "-";
    document.getElementById("kpiAvg").textContent = rows.length ? fmt(total / rows.length) : "₪0";
  }

  function groupByCategory(rows) {
    const map = new Map();
    for (const r of rows) map.set(r.category, (map.get(r.category) || 0) + r.amount);
    return [...map.entries()].map(([name, total]) => ({ name, total })).sort((a, b) => b.total - a.total);
  }

  function renderBars(rows) {
    const container = document.getElementById("categoryBars");
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
    const body = document.getElementById("txBody");
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

  function render() {
    const rows = currentRows();
    renderKpis(rows);
    renderBars(rows);
    fillCategoryFilter(rows);
    renderTransactions(rows);
    const label = monthSelect.options[monthSelect.selectedIndex];
    periodLabel.textContent = label ? "סיכום ל" + label.textContent : "";
  }

  async function handleFile(file) {
    showMessage("");
    try {
      const buffer = await file.arrayBuffer();
      const raw = readWorkbook(buffer);
      const { rows, skipped } = normalize(raw);
      if (!rows.length) {
        throw new Error("לא נמצאו שורות תקינות בקובץ.");
      }
      allRows = rows;
      fillMonths();
      dashboard.hidden = false;
      render();
      if (skipped) {
        showMessage("דילגתי על " + skipped + " שורות שאין בהן תאריך או סכום תקינים.");
      }
    } catch (err) {
      dashboard.hidden = true;
      showMessage(err.message || "לא הצלחתי לקרוא את הקובץ.");
    }
  }

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
