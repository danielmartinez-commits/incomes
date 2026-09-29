/*
 * We Capital | Identificación de ingresos
 * Production application logic.
 *
 * The application is intentionally read-only:
 * it identifies possible matches but never modifies AR.
 */

(() => {
  "use strict";

  const CONFIG = {
    dbName: "weCapitalIncomeIdentification",
    dbVersion: 1,
    storeName: "datasets",
    activeKey: "active",
    defaultReport: "data/AR_Report.xlsx",
    maxCombinationCandidates: 80,
    maxCombinationItems: 6,
    maxCombinationStates: 75000,
    dateWeight: 0.08,
    amountToleranceCents: 1
  };

  const state = {
    rows: [],
    brokers: [],
    clients: [],
    source: "",
    loadedAt: null,
    reportMeta: null,
    chart: null
  };

  const money = new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2
  });

  const numbers = new Intl.NumberFormat("en-US");
  const $ = (id) => document.getElementById(id);

  /* =====================================================
     GENERAL HELPERS
     ===================================================== */

  function normalize(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .replace(/\s+/g, " ")
      .toLowerCase();
  }

  function clean(value) {
    return String(value ?? "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, (char) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    }[char]));
  }

  function moneyFormat(value) {
    return money.format(Number(value || 0));
  }

  function numberFormat(value) {
    return numbers.format(Number(value || 0));
  }

  function cents(value) {
    return Math.round(Number(value || 0) * 100);
  }

  function shorten(value, length = 28) {
    return value.length > length
      ? `${value.slice(0, length - 1)}…`
      : value;
  }

  let toastTimer;

  function toast(message) {
    const element = $("toast");
    element.textContent = message;
    element.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => element.classList.remove("show"), 3200);
  }

  /* =====================================================
     MONEY / DATE PARSING
     ===================================================== */

  function parseMoney(value) {
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : NaN;
    }

    let text = String(value ?? "").trim();
    if (!text) return NaN;

    const negative = /^\(.*\)$/.test(text);
    text = text.replace(/[$,\s]/g, "").replace(/[()]/g, "");

    const result = Number(text);
    if (!Number.isFinite(result)) return NaN;

    return negative ? -result : result;
  }

  function parseDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return value;
    }

    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) {
        return new Date(parsed.y, parsed.m - 1, parsed.d);
      }
    }

    const text = String(value ?? "").trim();
    if (!text) return null;

    const us = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
    if (us) {
      const date = new Date(Number(us[3]), Number(us[1]) - 1, Number(us[2]));
      return Number.isNaN(date.getTime()) ? null : date;
    }

    const date = new Date(text);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function formatDate(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "—";
    return new Intl.DateTimeFormat("en-US", {
      month: "2-digit",
      day: "2-digit",
      year: "numeric"
    }).format(date);
  }

  function dateForInput(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
  }

  function dateDistanceDays(first, second) {
    if (!first || !second) return 0;
    return Math.abs(first.getTime() - second.getTime()) / 86400000;
  }

  /* =====================================================
     INDEXEDDB
     ===================================================== */

  function openDatabase() {
    return new Promise((resolve, reject) => {
      if (!("indexedDB" in window)) {
        reject(new Error("This browser does not support IndexedDB."));
        return;
      }

      const request = indexedDB.open(CONFIG.dbName, CONFIG.dbVersion);

      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(CONFIG.storeName)) {
          db.createObjectStore(CONFIG.storeName);
        }
      };

      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  function serializeRows(rows) {
    return rows.map((row) => ({
      ...row,
      date: row.date ? row.date.toISOString() : null
    }));
  }

  function deserializeRows(rows) {
    return rows.map((row) => ({
      ...row,
      date: row.date ? new Date(row.date) : null
    }));
  }

  async function saveDataset() {
    const db = await openDatabase();

    return new Promise((resolve, reject) => {
      const transaction = db.transaction(CONFIG.storeName, "readwrite");

      transaction.objectStore(CONFIG.storeName).put({
        rows: serializeRows(state.rows),
        source: state.source,
        loadedAt: state.loadedAt.toISOString(),
        reportMeta: state.reportMeta
      }, CONFIG.activeKey);

      transaction.oncomplete = () => {
        db.close();
        resolve();
      };

      transaction.onerror = () => {
        db.close();
        reject(transaction.error);
      };
    });
  }

  async function readDataset() {
    const db = await openDatabase();

    return new Promise((resolve, reject) => {
      const transaction = db.transaction(CONFIG.storeName, "readonly");
      const request = transaction.objectStore(CONFIG.storeName).get(CONFIG.activeKey);

      request.onsuccess = () => {
        db.close();
        resolve(request.result || null);
      };

      request.onerror = () => {
        db.close();
        reject(request.error);
      };
    });
  }

  /* =====================================================
     EXCEL PARSER
     ===================================================== */

  const HEADER_ALIASES = {
    date: ["transaction date", "date", "transaction_date"],
    type: ["transaction type", "type", "transaction_type"],
    num: ["num", "number", "invoice", "invoice number", "reference"],
    client: ["cliente", "client", "name", "customer", "customer name"],
    broker: ["broker", "factor", "debtor", "customer broker"],
    amount: ["amount", "transaction amount", "debit", "credit"],
    balance: ["balance", "running balance"]
  };

  function normalizeHeader(value) {
    return normalize(value).replace(/[_-]+/g, " ");
  }

  function findHeader(matrix) {
    const aliases = new Set(
      Object.values(HEADER_ALIASES)
        .flat()
        .map(normalizeHeader)
    );

    let best = null;

    for (let index = 0; index < Math.min(matrix.length, 50); index++) {
      const row = matrix[index] || [];
      const score = row.reduce(
        (total, cell) => total + (aliases.has(normalizeHeader(cell)) ? 1 : 0),
        0
      );

      if (!best || score > best.score) {
        best = { index, row, score };
      }
    }

    if (!best || best.score < 4) {
      throw new Error(
        "The report structure could not be recognized. Broker, Amount and Transaction Date were not found."
      );
    }

    return best;
  }

  function buildColumnMap(header) {
    const map = {};

    for (const [field, aliases] of Object.entries(HEADER_ALIASES)) {
      const set = new Set(aliases.map(normalizeHeader));
      map[field] = header.row.findIndex((cell) => set.has(normalizeHeader(cell)));
    }

    const required = ["date", "broker", "amount"];
    const missing = required.filter((field) => map[field] < 0);

    if (missing.length) {
      throw new Error(`Required columns are missing: ${missing.join(", ")}.`);
    }

    return map;
  }

  function parseWorkbook(buffer, sourceName) {
    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true,
      dense: true
    });

    if (!workbook.SheetNames.length) {
      throw new Error("The Excel file does not contain a worksheet.");
    }

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: null,
      blankrows: false
    });

    const header = findHeader(matrix);
    const columns = buildColumnMap(header);
    const rows = [];
    let rejectedRows = 0;

    for (let index = header.index + 1; index < matrix.length; index++) {
      const row = matrix[index] || [];

      const date = parseDate(row[columns.date]);
      const broker = clean(row[columns.broker]);
      const amount = parseMoney(row[columns.amount]);

      if (!date || !broker || !Number.isFinite(amount)) {
        rejectedRows++;
        continue;
      }

      const client = columns.client >= 0 ? clean(row[columns.client]) : "";
      const num = columns.num >= 0 ? clean(row[columns.num]) : "";
      const type = columns.type >= 0 ? clean(row[columns.type]) : "";
      const balanceValue = columns.balance >= 0 ? parseMoney(row[columns.balance]) : NaN;

      rows.push({
        id: [index + 1, normalize(broker), normalize(num), amount, date.getTime()].join("|"),
        rowNumber: index + 1,
        date,
        type,
        num,
        client,
        clientKey: normalize(client),
        broker,
        brokerKey: normalize(broker),
        amount,
        balance: Number.isFinite(balanceValue) ? balanceValue : null
      });
    }

    if (!rows.length) {
      throw new Error("The report was read, but no valid transaction rows were found.");
    }

    const seen = new Set();
    const duplicates = new Set();

    rows.forEach((row) => {
      if (seen.has(row.id)) duplicates.add(row.id);
      seen.add(row.id);
    });

    const dates = rows.map((row) => row.date.getTime());

    return {
      rows,
      source: sourceName,
      loadedAt: new Date(),
      reportMeta: {
        sheetName: workbook.SheetNames[0],
        headerRow: header.index + 1,
        rejectedRows,
        duplicateKeys: duplicates.size,
        totalAmount: rows.reduce((sum, row) => sum + row.amount, 0),
        firstDate: new Date(Math.min(...dates)),
        lastDate: new Date(Math.max(...dates))
      }
    };
  }

  /* =====================================================
     DATASET / INDEXES
     ===================================================== */

  function rebuildIndexes() {
    const brokerMap = new Map();
    const clients = new Set();

    state.rows.forEach((row) => {
      if (!brokerMap.has(row.brokerKey)) {
        brokerMap.set(row.brokerKey, row.broker);
      }
      if (row.clientKey) clients.add(row.client);
    });

    state.brokers = [...brokerMap.entries()]
      .map(([key, name]) => ({ key, name }))
      .sort((a, b) => a.name.localeCompare(b.name));

    state.clients = [...clients].sort((a, b) => a.localeCompare(b));
  }

  async function activateDataset(dataset) {
    state.rows = dataset.rows;
    state.source = dataset.source;
    state.loadedAt = dataset.loadedAt;
    state.reportMeta = dataset.reportMeta;

    rebuildIndexes();
    renderAll();

    try {
      await saveDataset();
    } catch (error) {
      console.error(error);
      toast("Report loaded, but local browser storage could not be updated.");
    }
  }

  async function restoreDataset() {
    try {
      const saved = await readDataset();
      if (!saved || !saved.rows?.length) return false;

      state.rows = deserializeRows(saved.rows);
      state.source = saved.source || "Local dataset";
      state.loadedAt = saved.loadedAt ? new Date(saved.loadedAt) : new Date();
      state.reportMeta = saved.reportMeta || null;

      rebuildIndexes();
      renderAll();
      return true;
    } catch (error) {
      console.warn(error);
      return false;
    }
  }

  async function loadFile(file) {
    if (!file) return;

    $("identifyBtn").disabled = true;

    try {
      const buffer = await file.arrayBuffer();
      const parsed = parseWorkbook(buffer, file.name);

      await activateDataset(parsed);

      toast(`Report loaded: ${numberFormat(parsed.rows.length)} valid movements.`);
      switchView("dashboard");
    } catch (error) {
      console.error(error);
      toast(error.message || "The report could not be loaded.");
    } finally {
      $("identifyBtn").disabled = false;
    }
  }

  async function autoLoad() {
    $("sidebarStatus").textContent = "Loading...";

    if (await restoreDataset()) return;

    try {
      const response = await fetch(CONFIG.defaultReport, { cache: "no-store" });
      if (!response.ok) throw new Error("Default report not found.");

      const parsed = parseWorkbook(
        await response.arrayBuffer(),
        CONFIG.defaultReport
      );

      await activateDataset(parsed);
      toast("Default report loaded.");
    } catch (error) {
      console.warn(error);
      $("sidebarStatus").textContent = "Waiting for report";
      renderAll();
    }
  }

  /* =====================================================
     BROKER AGGREGATION
     ===================================================== */

  function aggregateBrokers() {
    const map = new Map();

    state.rows.forEach((row) => {
      const current = map.get(row.brokerKey) || {
        key: row.brokerKey,
        name: row.broker,
        amount: 0,
        count: 0,
        clients: new Set(),
        oldest: null
      };

      current.amount += row.amount;
      current.count++;
      if (row.clientKey) current.clients.add(row.clientKey);
      if (!current.oldest || row.date < current.oldest) current.oldest = row.date;

      map.set(row.brokerKey, current);
    });

    return [...map.values()].sort((a, b) => b.amount - a.amount);
  }

  /* =====================================================
     MATCHING ENGINE
     ===================================================== */

  function paymentDate() {
    const value = $("dateInput").value;
    if (!value) return null;

    const date = new Date(`${value}T00:00:00`);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function paymentAmount() {
    const amount = parseMoney($("amountInput").value);
    return Number.isFinite(amount) ? amount : NaN;
  }

  function candidateScore(row, target, targetDate) {
    const amountDifference = Math.abs(row.amount - target);
    const relativeDifference = amountDifference / Math.max(Math.abs(target), 1);
    const days = targetDate
      ? Math.min(dateDistanceDays(row.date, targetDate), 365)
      : 0;

    return 600 - relativeDifference * 1000 - days * CONFIG.dateWeight;
  }

  function rankCandidates(rows, target, targetDate) {
    return [...rows]
      .map((row) => ({
        row,
        score: candidateScore(row, target, targetDate)
      }))
      .sort((a, b) => b.score - a.score);
  }

  /*
   * Bounded subset-sum.
   * It finds an exact combination without allowing an
   * uncontrolled combinatorial search.
   */
  function findExactCombination(rows, targetAmount, targetDate) {
    const target = cents(targetAmount);
    if (target <= 0) return null;

    const candidates = rankCandidates(rows, targetAmount, targetDate)
      .filter(({ row }) => row.amount > 0 && cents(row.amount) <= target)
      .slice(0, CONFIG.maxCombinationCandidates);

    const states = new Map([[0, []]]);

    for (const { row } of candidates) {
      const value = cents(row.amount);
      const snapshot = [...states.entries()];

      for (const [sum, combination] of snapshot) {
        const next = sum + value;

        if (next > target) continue;
        if (combination.length >= CONFIG.maxCombinationItems) continue;
        if (states.has(next)) continue;

        const nextCombination = [...combination, row];
        states.set(next, nextCombination);

        if (next === target) return nextCombination;
        if (states.size >= CONFIG.maxCombinationStates) return null;
      }
    }

    return null;
  }

  function identify() {
    const broker = normalize($("brokerInput").value);
    const amount = paymentAmount();
    const date = paymentDate();

    if (!broker) {
      toast("Enter the broker.");
      return;
    }

    if (!Number.isFinite(amount) || amount <= 0) {
      toast("Enter a valid payment amount.");
      return;
    }

    const rows = state.rows.filter((row) =>
      row.brokerKey === broker ||
      row.brokerKey.includes(broker) ||
      broker.includes(row.brokerKey)
    );

    $("identifyEmpty").classList.add("hidden");
    $("identifyResults").classList.remove("hidden");

    if (!rows.length) {
      $("identifyResults").innerHTML = `
        <div class="result-head">
          <div>
            <h2>No candidate found</h2>
            <p>No AR movement was found for the selected broker.</p>
          </div>
          <span class="badge red">NO MATCH</span>
        </div>`;
      return;
    }

    const ranked = rankCandidates(rows, amount, date);

    const exact = ranked
      .filter(({ row }) =>
        Math.abs(cents(row.amount) - cents(amount)) <= CONFIG.amountToleranceCents
      )
      .map(({ row }) => row)
      .slice(0, 10);

    const combination = exact.length === 0
      ? findExactCombination(rows, amount, date)
      : null;

    let html = `
      <div class="result-head">
        <div>
          <h2>Identification result</h2>
          <p>${numberFormat(rows.length)} movements found for <strong>${escapeHtml(rows[0].broker)}</strong>.</p>
        </div>
      </div>`;

    if (exact.length) {
      html += renderMatchCard(
        "Exact amount match",
        exact,
        amount,
        date,
        "green",
        "EXACT MATCH"
      );
    }

    if (combination) {
      html += renderMatchCard(
        "Exact batch match",
        combination,
        amount,
        date,
        "green",
        "EXACT COMBINATION"
      );
    }

    if (!exact.length && !combination) {
      html += renderReviewCard(
        ranked.slice(0, 12).map(({ row }) => row),
        amount,
        date
      );
    }

    $("identifyResults").innerHTML = html;
  }

  function renderMatchCard(title, rows, payment, date, badgeClass, badgeText) {
    const selected = rows.reduce((sum, row) => sum + row.amount, 0);
    const difference = payment - selected;

    return `
      <article class="match-card">
        <div class="match-top">
          <div class="match-title">
            <h3>${escapeHtml(title)}</h3>
            <p>${numberFormat(rows.length)} movement${rows.length === 1 ? "" : "s"} · ${date ? `payment date ${formatDate(date)}` : "payment date not supplied"}</p>
          </div>
          <span class="badge ${badgeClass}">${badgeText}</span>
        </div>
        <div class="match-summary">
          <div class="summary-item"><small>Selected</small><strong>${moneyFormat(selected)}</strong></div>
          <div class="summary-item"><small>Payment</small><strong>${moneyFormat(payment)}</strong></div>
          <div class="summary-item"><small>Difference</small><strong>${moneyFormat(difference)}</strong></div>
        </div>
        ${movementTable(rows)}
      </article>`;
  }

  function renderReviewCard(rows, payment, date) {
    return `
      <article class="match-card">
        <div class="match-top">
          <div class="match-title">
            <h3>Closest candidates</h3>
            <p>No exact amount or exact combination was found. Review these candidates manually.</p>
          </div>
          <span class="badge yellow">REVIEW</span>
        </div>
        <div class="match-summary">
          <div class="summary-item"><small>Payment</small><strong>${moneyFormat(payment)}</strong></div>
          <div class="summary-item"><small>Date</small><strong>${date ? formatDate(date) : "—"}</strong></div>
          <div class="summary-item"><small>Candidates</small><strong>${numberFormat(rows.length)}</strong></div>
        </div>
        ${movementTable(rows, true, payment)}
      </article>`;
  }

  function movementTable(rows, showDifference = false, payment = 0) {
    return `
      <table class="match-table">
        <thead>
          <tr>
            <th>Date</th>
            <th>Num</th>
            <th>Client</th>
            <th>Broker</th>
            <th>Amount</th>
            ${showDifference ? "<th>Difference</th>" : ""}
          </tr>
        </thead>
        <tbody>
          ${rows.map((row) => `
            <tr>
              <td>${formatDate(row.date)}</td>
              <td>${escapeHtml(row.num || "—")}</td>
              <td>${escapeHtml(row.client || "—")}</td>
              <td>${escapeHtml(row.broker)}</td>
              <td class="amount">${moneyFormat(row.amount)}</td>
              ${showDifference ? `<td class="amount">${moneyFormat(row.amount - payment)}</td>` : ""}
            </tr>`).join("")}
        </tbody>
      </table>`;
  }

  /* =====================================================
     DASHBOARD
     ===================================================== */

  function renderDashboard() {
    const total = state.rows.reduce((sum, row) => sum + row.amount, 0);
    const brokers = aggregateBrokers().slice(0, 10);

    $("kpiMovements").textContent = numberFormat(state.rows.length);
    $("kpiAmount").textContent = moneyFormat(total);
    $("kpiBrokers").textContent = numberFormat(state.brokers.length);
    $("kpiClients").textContent = numberFormat(state.clients.length);

    if (state.chart) {
      state.chart.destroy();
      state.chart = null;
    }

    if (!state.rows.length) {
      $("topBrokersTable").innerHTML = "<p>No report loaded.</p>";
      return;
    }

    state.chart = new Chart($("brokerChart"), {
      type: "bar",
      data: {
        labels: brokers.map((broker) => shorten(broker.name, 25)),
        datasets: [{
          data: brokers.map((broker) => broker.amount),
          backgroundColor: "#6D4AFF",
          borderRadius: 5,
          barThickness: 18
        }]
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (context) => moneyFormat(context.raw)
            }
          }
        },
        scales: {
          x: {
            grid: { color: "#EEF0F4" },
            ticks: {
              font: { size: 9 },
              callback: (value) => moneyFormat(value)
            }
          },
          y: {
            grid: { display: false },
            ticks: { font: { size: 9 } }
          }
        }
      }
    });

    $("topBrokersTable").innerHTML = `
      <table class="mini-table">
        <thead><tr><th>Broker</th><th>Items</th><th>Amount</th></tr></thead>
        <tbody>
          ${brokers.map((broker) => `
            <tr>
              <td>${escapeHtml(broker.name)}</td>
              <td>${numberFormat(broker.count)}</td>
              <td>${moneyFormat(broker.amount)}</td>
            </tr>`).join("")}
        </tbody>
      </table>`;
  }

  /* =====================================================
     BROKERS
     ===================================================== */

  function renderBrokers(filter = "") {
    const query = normalize(filter);
    const brokers = aggregateBrokers().filter((broker) =>
      !query || normalize(broker.name).includes(query)
    );

    $("brokerCount").textContent = `${numberFormat(brokers.length)} brokers`;

    $("brokerCards").innerHTML = brokers.map((broker) => `
      <article class="broker-card" data-broker="${escapeHtml(broker.key)}">
        <h3 title="${escapeHtml(broker.name)}">${escapeHtml(broker.name)}</h3>
        <p>${numberFormat(broker.clients.size)} clients · ${numberFormat(broker.count)} movements</p>
        <div class="amount">${moneyFormat(broker.amount)}</div>
        <div class="broker-meta"><span>Oldest ${formatDate(broker.oldest)}</span><span>View →</span></div>
      </article>`).join("") || `
        <div class="empty-state"><h2>No brokers found</h2><p>Try another search.</p></div>`;

    document.querySelectorAll(".broker-card").forEach((card) => {
      card.addEventListener("click", () => openBroker(card.dataset.broker));
    });
  }

  function openBroker(key) {
    const rows = state.rows
      .filter((row) => row.brokerKey === key)
      .sort((a, b) => b.date - a.date);

    if (!rows.length) return;

    const total = rows.reduce((sum, row) => sum + row.amount, 0);
    const clients = new Set(rows.map((row) => row.clientKey).filter(Boolean));

    $("modalContent").innerHTML = `
      <div class="detail-title">
        <h2>${escapeHtml(rows[0].broker)}</h2>
        <p>Movement history available in the current AR report.</p>
      </div>
      <div class="detail-stats">
        <div class="detail-stat"><small>Total amount</small><strong>${moneyFormat(total)}</strong></div>
        <div class="detail-stat"><small>Movements</small><strong>${numberFormat(rows.length)}</strong></div>
        <div class="detail-stat"><small>Clients</small><strong>${numberFormat(clients.size)}</strong></div>
      </div>
      <table class="detail-table">
        <thead><tr><th>Date</th><th>Num</th><th>Client</th><th>Type</th><th>Amount</th><th>Balance</th></tr></thead>
        <tbody>
          ${rows.map((row) => `
            <tr>
              <td>${formatDate(row.date)}</td>
              <td>${escapeHtml(row.num || "—")}</td>
              <td>${escapeHtml(row.client || "—")}</td>
              <td>${escapeHtml(row.type || "—")}</td>
              <td>${moneyFormat(row.amount)}</td>
              <td>${row.balance === null ? "—" : moneyFormat(row.balance)}</td>
            </tr>`).join("")}
        </tbody>
      </table>`;

    $("modal").classList.remove("hidden");
    $("modal").setAttribute("aria-hidden", "false");
  }

  /* =====================================================
     DATA VIEW
     ===================================================== */

  function renderData() {
    if (!state.rows.length) {
      $("dataDescription").textContent = "No report loaded.";
      $("dataStats").innerHTML = "";
      return;
    }

    const dates = state.rows.map((row) => row.date.getTime());
    const total = state.rows.reduce((sum, row) => sum + row.amount, 0);
    const meta = state.reportMeta || {};

    $("dataDescription").textContent = `${state.source} · loaded ${new Intl.DateTimeFormat("en-US", { dateStyle: "medium", timeStyle: "short" }).format(state.loadedAt)}`;

    const stats = [
      ["Rows", numberFormat(state.rows.length)],
      ["Brokers", numberFormat(state.brokers.length)],
      ["Clients", numberFormat(state.clients.length)],
      ["Amount", moneyFormat(total)],
      ["First date", formatDate(new Date(Math.min(...dates)))],
      ["Last date", formatDate(new Date(Math.max(...dates)))],
      ["Rejected rows", numberFormat(meta.rejectedRows || 0)],
      ["Duplicate keys", numberFormat(meta.duplicateKeys || 0)]
    ];

    $("dataStats").innerHTML = stats.map(([label, value]) => `
      <div class="data-stat"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>
    `).join("");
  }

  /* =====================================================
     SEARCH / NAVIGATION
     ===================================================== */

  function switchView(view) {
    document.querySelectorAll(".nav-item").forEach((button) => {
      button.classList.toggle("active", button.dataset.view === view);
    });

    document.querySelectorAll(".view").forEach((section) => {
      section.classList.remove("active-view");
    });

    $(`view-${view}`).classList.add("active-view");
  }

  function globalSearch(value) {
    const query = normalize(value);
    if (!query) return;

    const transaction = state.rows.find((row) => normalize(row.num) === query);

    if (transaction) {
      $("brokerInput").value = transaction.broker;
      $("amountInput").value = transaction.amount;
      $("dateInput").value = dateForInput(transaction.date);
      switchView("identify");
      identify();
      return;
    }

    const broker = state.brokers.find((item) => normalize(item.name).includes(query));

    if (broker) {
      $("brokerFilter").value = broker.name;
      renderBrokers(broker.name);
      switchView("brokers");
      return;
    }

    const client = state.rows.find((row) => row.clientKey.includes(query));

    if (client) {
      $("brokerInput").value = client.broker;
      $("amountInput").value = client.amount;
      $("dateInput").value = dateForInput(client.date);
      switchView("identify");
      identify();
      return;
    }

    const amount = parseMoney(query);

    if (Number.isFinite(amount) && amount > 0) {
      $("amountInput").value = amount;
      switchView("identify");
      toast("Amount found. Add the broker to identify it.");
      return;
    }

    toast("No broker, client, invoice or amount matched that search.");
  }

  /* =====================================================
     RENDER ALL
     ===================================================== */

  function renderAll() {
    $("sidebarStatus").textContent = state.rows.length
      ? `${numberFormat(state.rows.length)} rows ready`
      : "Waiting for report";

    $("reportDate").textContent = state.loadedAt
      ? `Updated ${new Intl.DateTimeFormat("en-US", { month: "short", day: "2-digit", year: "numeric", hour: "numeric", minute: "2-digit" }).format(state.loadedAt)}`
      : "No report";

    $("brokerList").innerHTML = state.brokers.map((broker) =>
      `<option value="${escapeHtml(broker.name)}"></option>`
    ).join("");

    renderDashboard();
    renderBrokers();
    renderData();
  }

  /* =====================================================
     EVENTS
     ===================================================== */

  function closeModal() {
    $("modal").classList.add("hidden");
    $("modal").setAttribute("aria-hidden", "true");
  }

  function bindEvents() {
    document.querySelectorAll(".nav-item").forEach((button) => {
      button.addEventListener("click", () => switchView(button.dataset.view));
    });

    $("identifyBtn").addEventListener("click", identify);

    ["brokerInput", "amountInput"].forEach((id) => {
      $(id).addEventListener("keydown", (event) => {
        if (event.key === "Enter") identify();
      });
    });

    $("brokerFilter").addEventListener("input", (event) => {
      renderBrokers(event.target.value);
    });

    $("globalSearch").addEventListener("keydown", (event) => {
      if (event.key === "Enter") globalSearch(event.target.value);
    });

    $("fileInput").addEventListener("change", (event) => {
      loadFile(event.target.files[0]);
      event.target.value = "";
    });

    $("refreshBtn").addEventListener("click", async () => {
      const restored = await restoreDataset();
      toast(restored ? "Saved report reloaded." : "No saved report found.");
    });

    $("modalClose").addEventListener("click", closeModal);
    document.querySelector(".modal-backdrop").addEventListener("click", closeModal);

    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") closeModal();
    });

    const dropZone = $("dropZone");

    ["dragenter", "dragover"].forEach((eventName) => {
      dropZone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropZone.classList.add("dragover");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      dropZone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropZone.classList.remove("dragover");
      });
    });

    dropZone.addEventListener("drop", (event) => {
      const file = event.dataTransfer.files[0];
      if (file) loadFile(file);
    });
  }

  bindEvents();
  autoLoad();
})();
