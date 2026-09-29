/*
 * We Capital | Identificación de ingresos
 *
 * The application is intentionally read-only.
 * It loads an AR report, validates it, normalizes it and builds indexes
 * that make broker / client / invoice searches fast even when the report grows.
 */

const CONFIG = {
    defaultReport: "./data/AR_Report.xlsx",
    xlsxWorkerUrl: "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
    databaseName: "weCapitalIncomeIdentification",
    databaseVersion: 1,
    databaseStore: "datasets",
    maxCombinationCandidates: 70,
    maxCombinationStates: 60000,
    currency: "USD"
};

const HEADER_ALIASES = {
    date: ["transaction date", "date", "transaction_date"],
    type: ["transaction type", "type", "transaction_type"],
    invoice: ["invoice", "invoice number", "number", "num", "reference"],
    client: ["cliente", "client", "customer", "customer name", "name"],
    broker: ["broker", "factor", "debtor", "customer broker"],
    amount: ["amount", "transaction amount", "debit", "credit"],
    balance: ["balance", "running balance"]
};

const state = {
    dataset: null,
    brokerIndex: new Map(),
    clientIndex: new Map(),
    invoiceIndex: new Map(),
    chart: null,
    currentFileName: "",
    currentObjectUrl: null
};

const $ = (id) => document.getElementById(id);

const elements = {
    globalSearch: $("globalSearch"),
    refreshBtn: $("refreshBtn"),
    dataStatus: $("dataStatus"),
    systemStatus: $("systemStatus"),
    brokerInput: $("brokerInput"),
    amountInput: $("amountInput"),
    paymentDateInput: $("paymentDateInput"),
    identifyBtn: $("identifyBtn"),
    identificationResults: $("identificationResults"),
    brokerSearch: $("brokerSearch"),
    brokerCount: $("brokerCount"),
    brokerGrid: $("brokerGrid"),
    fileInput: $("fileInput"),
    dropZone: $("dropZone"),
    uploadProgress: $("uploadProgress"),
    progressTitle: $("progressTitle"),
    progressPercent: $("progressPercent"),
    progressBar: $("progressBar"),
    progressMessage: $("progressMessage"),
    progressRows: $("progressRows"),
    dataFileName: $("dataFileName"),
    dataRows: $("dataRows"),
    dataBrokers: $("dataBrokers"),
    dataClients: $("dataClients"),
    dataStartDate: $("dataStartDate"),
    dataEndDate: $("dataEndDate"),
    dataBeginningBalance: $("dataBeginningBalance"),
    dataLastBalance: $("dataLastBalance"),
    datasetMessage: $("datasetMessage"),
    kpiMovements: $("kpiMovements"),
    kpiOpenAR: $("kpiOpenAR"),
    kpiBrokers: $("kpiBrokers"),
    kpiClients: $("kpiClients"),
    brokerSummary: $("brokerSummary"),
    brokerChart: $("brokerChart"),
    toast: $("toast"),
    brokerModal: $("brokerModal"),
    brokerModalContent: $("brokerModalContent"),
    closeBrokerModal: $("closeBrokerModal")
};

/* =========================================================
   Formatting and normalization
   ========================================================= */

function normalizeText(value) {
    return String(value ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .trim();
}

function searchKey(value) {
    return normalizeText(value).toLowerCase();
}

function displayName(value) {
    return normalizeText(value);
}

function money(value) {
    const amount = Number(value) || 0;

    return new Intl.NumberFormat("en-US", {
        style: "currency",
        currency: CONFIG.currency,
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(amount);
}

function integer(value) {
    return new Intl.NumberFormat("en-US").format(Number(value) || 0);
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function parseNumber(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
        return value;
    }

    if (value === null || value === undefined || value === "") {
        return null;
    }

    const text = String(value)
        .trim()
        .replace(/\$/g, "")
        .replace(/,/g, "");

    const parsed = Number(text);
    return Number.isFinite(parsed) ? parsed : null;
}

function parseDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return value;
    }

    if (typeof value === "number" && window.XLSX?.SSF) {
        const date = XLSX.SSF.parse_date_code(value);

        if (date) {
            return new Date(date.y, date.m - 1, date.d);
        }
    }

    const text = normalizeText(value);

    if (!text) {
        return null;
    }

    const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

    if (slashMatch) {
        const month = Number(slashMatch[1]);
        const day = Number(slashMatch[2]);
        const year = Number(slashMatch[3]);
        const date = new Date(year, month - 1, day);

        return Number.isNaN(date.getTime()) ? null : date;
    }

    const date = new Date(text);
    return Number.isNaN(date.getTime()) ? null : date;
}

function dateKey(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
        return "";
    }

    return [
        date.getFullYear(),
        String(date.getMonth() + 1).padStart(2, "0"),
        String(date.getDate()).padStart(2, "0")
    ].join("-");
}

function formatDate(value) {
    const date = value instanceof Date ? value : parseDate(value);

    if (!date) {
        return "—";
    }

    return new Intl.DateTimeFormat("en-US", {
        month: "2-digit",
        day: "2-digit",
        year: "numeric"
    }).format(date);
}

function yieldToBrowser() {
    return new Promise((resolve) => setTimeout(resolve, 0));
}

/* =========================================================
   File parsing
   ========================================================= */

function normalizeHeader(value) {
    return searchKey(value).replace(/[^a-z0-9_ ]/g, "");
}

function findHeaderRow(rows) {
    const limit = Math.min(rows.length, 50);
    let best = null;

    for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
        const row = rows[rowIndex] || [];
        const normalized = row.map(normalizeHeader);
        const found = {};

        for (const [field, aliases] of Object.entries(HEADER_ALIASES)) {
            const aliasSet = aliases.map(normalizeHeader);
            const columnIndex = normalized.findIndex((header) => aliasSet.includes(header));

            if (columnIndex !== -1) {
                found[field] = columnIndex;
            }
        }

        const score = Object.keys(found).length;

        if (!best || score > best.score) {
            best = {
                rowIndex,
                columns: found,
                score
            };
        }
    }

    if (!best || best.score < 3) {
        throw new Error("The report header could not be detected. Required fields: Transaction date, Broker and Amount.");
    }

    if (best.columns.date === undefined || best.columns.amount === undefined) {
        throw new Error("The report must contain Transaction date and Amount columns.");
    }

    return best;
}

function parseWorkbookWithWorker(arrayBuffer) {
    return new Promise((resolve, reject) => {
        const workerCode = `
            importScripts(${JSON.stringify(CONFIG.xlsxWorkerUrl)});

            self.onmessage = function(event) {
                try {
                    const workbook = XLSX.read(event.data, {
                        type: "array",
                        cellDates: true,
                        raw: true
                    });

                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(firstSheet, {
                        header: 1,
                        defval: null,
                        raw: true,
                        blankrows: true
                    });

                    self.postMessage({
                        ok: true,
                        sheetName: workbook.SheetNames[0],
                        rows: rows
                    });
                } catch (error) {
                    self.postMessage({
                        ok: false,
                        error: error && error.message ? error.message : String(error)
                    });
                }
            };
        `;

        const blob = new Blob([workerCode], { type: "application/javascript" });
        const workerUrl = URL.createObjectURL(blob);
        const worker = new Worker(workerUrl);

        worker.onmessage = (event) => {
            worker.terminate();
            URL.revokeObjectURL(workerUrl);

            if (event.data.ok) {
                resolve(event.data);
            } else {
                reject(new Error(event.data.error));
            }
        };

        worker.onerror = (event) => {
            worker.terminate();
            URL.revokeObjectURL(workerUrl);
            reject(new Error(event.message || "The Excel worker failed."));
        };

        worker.postMessage(arrayBuffer, [arrayBuffer]);
    });
}

async function parseWorkbook(arrayBuffer) {
    try {
        return await parseWorkbookWithWorker(arrayBuffer);
    } catch (workerError) {
        // Fallback for environments that block Web Workers or CDN workers.
        if (!window.XLSX) {
            throw new Error("The Excel engine could not be loaded.");
        }

        const workbook = XLSX.read(arrayBuffer, {
            type: "array",
            cellDates: true,
            raw: true
        });

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(firstSheet, {
            header: 1,
            defval: null,
            raw: true,
            blankrows: true
        });

        return {
            ok: true,
            sheetName: workbook.SheetNames[0],
            rows
        };
    }
}

function validateRows(rows) {
    const header = findHeaderRow(rows);
    const dataRows = rows.slice(header.rowIndex + 1);

    let beginningBalance = null;
    let skippedBlankRows = 0;
    let skippedInvalidRows = 0;
    const prepared = [];

    for (let index = 0; index < dataRows.length; index += 1) {
        const raw = dataRows[index] || [];
        const dateValue = raw[header.columns.date];
        const typeValue = raw[header.columns.type];
        const invoiceValue = header.columns.invoice !== undefined ? raw[header.columns.invoice] : "";
        const clientValue = header.columns.client !== undefined ? raw[header.columns.client] : "";
        const brokerValue = header.columns.broker !== undefined ? raw[header.columns.broker] : "";
        const amountValue = raw[header.columns.amount];
        const balanceValue = header.columns.balance !== undefined ? raw[header.columns.balance] : null;

        const allBlank = raw.every((cell) => cell === null || cell === undefined || String(cell).trim() === "");

        if (allBlank) {
            skippedBlankRows += 1;
            continue;
        }

        const dateText = normalizeText(dateValue);

        if (searchKey(dateText) === "beginning balance") {
            const balance = parseNumber(balanceValue);
            if (balance !== null) {
                beginningBalance = balance;
            }
            continue;
        }

        const date = parseDate(dateValue);
        const amount = parseNumber(amountValue);

        if (!date || amount === null) {
            skippedInvalidRows += 1;
            continue;
        }

        prepared.push({
            sourceRow: header.rowIndex + index + 2,
            date,
            dateKey: dateKey(date),
            type: displayName(typeValue),
            invoice: displayName(invoiceValue),
            client: displayName(clientValue),
            broker: displayName(brokerValue),
            brokerKey: searchKey(brokerValue),
            clientKey: searchKey(clientValue),
            invoiceKey: searchKey(invoiceValue),
            amount,
            balance: parseNumber(balanceValue)
        });
    }

    if (prepared.length === 0) {
        throw new Error("No valid transactions were found after the header. Check the report format.");
    }

    return {
        rows: prepared,
        beginningBalance,
        skippedBlankRows,
        skippedInvalidRows,
        headerRow: header.rowIndex + 1,
        columns: header.columns
    };
}

async function processRows(prepared, fileName, metadata) {
    const rows = prepared.rows;
    const total = rows.length;
    const brokerIndex = new Map();
    const clientIndex = new Map();
    const invoiceIndex = new Map();
    const brokerTotals = new Map();
    const clientTotals = new Map();

    let minDate = null;
    let maxDate = null;
    let lastBalance = null;

    for (let start = 0; start < total; start += 250) {
        const end = Math.min(start + 250, total);

        for (let index = start; index < end; index += 1) {
            const row = rows[index];

            if (row.brokerKey) {
                addIndexValue(brokerIndex, row.brokerKey, index);
                brokerTotals.set(
                    row.brokerKey,
                    (brokerTotals.get(row.brokerKey) || 0) + row.amount
                );
            }

            if (row.clientKey) {
                addIndexValue(clientIndex, row.clientKey, index);
                clientTotals.set(
                    row.clientKey,
                    (clientTotals.get(row.clientKey) || 0) + row.amount
                );
            }

            if (row.invoiceKey) {
                addIndexValue(invoiceIndex, row.invoiceKey, index);
            }

            if (!minDate || row.date < minDate) {
                minDate = row.date;
            }

            if (!maxDate || row.date > maxDate) {
                maxDate = row.date;
            }

            if (row.balance !== null) {
                lastBalance = row.balance;
            }
        }

        const percent = Math.round((end / total) * 100);
        updateProgress(percent, `Indexing transactions...`, `${integer(end)} / ${integer(total)} rows`);
        await yieldToBrowser();
    }

    return {
        rows,
        metadata: {
            fileName,
            sheetName: metadata.sheetName,
            headerRow: metadata.headerRow,
            beginningBalance: prepared.beginningBalance,
            lastBalance,
            startDate: minDate,
            endDate: maxDate,
            skippedBlankRows: prepared.skippedBlankRows,
            skippedInvalidRows: prepared.skippedInvalidRows,
            rowCount: rows.length,
            brokerCount: brokerIndex.size,
            clientCount: clientIndex.size
        },
        brokerIndex,
        clientIndex,
        invoiceIndex,
        brokerTotals,
        clientTotals
    };
}

function addIndexValue(index, key, rowIndex) {
    const existing = index.get(key);

    if (existing) {
        existing.push(rowIndex);
    } else {
        index.set(key, [rowIndex]);
    }
}

/* =========================================================
   IndexedDB persistence
   ========================================================= */

function openDatabase() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(CONFIG.databaseName, CONFIG.databaseVersion);

        request.onupgradeneeded = () => {
            const db = request.result;

            if (!db.objectStoreNames.contains(CONFIG.databaseStore)) {
                db.createObjectStore(CONFIG.databaseStore);
            }
        };

        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function saveDataset(dataset) {
    try {
        const db = await openDatabase();

        await new Promise((resolve, reject) => {
            const transaction = db.transaction(CONFIG.databaseStore, "readwrite");
            transaction.objectStore(CONFIG.databaseStore).put(dataset, "active");
            transaction.oncomplete = resolve;
            transaction.onerror = () => reject(transaction.error);
        });

        db.close();
    } catch (error) {
        console.warn("IndexedDB save skipped:", error);
    }
}

async function loadSavedDataset() {
    try {
        const db = await openDatabase();

        const result = await new Promise((resolve, reject) => {
            const transaction = db.transaction(CONFIG.databaseStore, "readonly");
            const request = transaction.objectStore(CONFIG.databaseStore).get("active");
            request.onsuccess = () => resolve(request.result || null);
            request.onerror = () => reject(request.error);
        });

        db.close();
        return result;
    } catch (error) {
        console.warn("IndexedDB load skipped:", error);
        return null;
    }
}

/* =========================================================
   Dataset activation
   ========================================================= */

function rebuildIndexesFromDataset(dataset) {
    const brokerIndex = new Map();
    const clientIndex = new Map();
    const invoiceIndex = new Map();
    const brokerTotals = new Map();
    const clientTotals = new Map();

    dataset.rows.forEach((row, index) => {
        if (row.brokerKey) {
            addIndexValue(brokerIndex, row.brokerKey, index);
            brokerTotals.set(row.brokerKey, (brokerTotals.get(row.brokerKey) || 0) + row.amount);
        }

        if (row.clientKey) {
            addIndexValue(clientIndex, row.clientKey, index);
            clientTotals.set(row.clientKey, (clientTotals.get(row.clientKey) || 0) + row.amount);
        }

        if (row.invoiceKey) {
            addIndexValue(invoiceIndex, row.invoiceKey, index);
        }
    });

    state.dataset = dataset;
    state.brokerIndex = brokerIndex;
    state.clientIndex = clientIndex;
    state.invoiceIndex = invoiceIndex;
    state.currentFileName = dataset.metadata.fileName;

    dataset.brokerTotals = brokerTotals;
    dataset.clientTotals = clientTotals;
}

async function activatePreparedDataset(prepared, fileName, parsed) {
    const dataset = await processRows(prepared, fileName, parsed);

    rebuildIndexesFromDataset(dataset);
    updateDataPanel();
    renderDashboard();
    renderBrokers();

    await saveDataset(dataset);

    setSystemStatus("Operational");
    elements.dataStatus.textContent = `${integer(dataset.rows.length)} transactions loaded`;
    elements.datasetMessage.textContent = "Report validated and ready for identification.";
}

/* =========================================================
   Upload workflow
   ========================================================= */

function updateProgress(percent, message, rowsMessage) {
    elements.uploadProgress.classList.remove("hidden");
    elements.progressPercent.textContent = `${Math.round(percent)}%`;
    elements.progressBar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
    elements.progressMessage.textContent = message;
    elements.progressRows.textContent = rowsMessage || "";
}

function resetProgress() {
    elements.uploadProgress.classList.add("hidden");
    elements.progressPercent.textContent = "0%";
    elements.progressBar.style.width = "0%";
    elements.progressMessage.textContent = "Preparing...";
    elements.progressRows.textContent = "0 rows";
}

async function handleFile(file) {
    if (!file) {
        return;
    }

    const validExtension = /\.(xlsx|xls|csv)$/i.test(file.name);

    if (!validExtension) {
        showToast("Please select an Excel or CSV file.", "error");
        return;
    }

    elements.progressTitle.textContent = "Processing report";
    elements.fileInput.disabled = true;
    setSystemStatus("Processing");

    try {
        updateProgress(5, "Reading file...", "Preparing file");
        await yieldToBrowser();

        const arrayBuffer = await file.arrayBuffer();

        updateProgress(15, "Reading workbook...", "Opening first worksheet");
        await yieldToBrowser();

        const parsedWorkbook = await parseWorkbook(arrayBuffer);

        updateProgress(35, "Validating structure...", `Sheet: ${parsedWorkbook.sheetName}`);
        await yieldToBrowser();

        const prepared = validateRows(parsedWorkbook.rows);

        updateProgress(50, "Normalizing transactions...", `${integer(prepared.rows.length)} valid transactions`);
        await yieldToBrowser();

        await activatePreparedDataset(prepared, file.name, {
            sheetName: parsedWorkbook.sheetName,
            headerRow: prepared.headerRow
        });

        updateProgress(100, "Report ready", `${integer(prepared.rows.length)} transactions indexed`);
        showToast("The report was validated and loaded successfully.", "success");

        setTimeout(resetProgress, 1400);
    } catch (error) {
        console.error(error);
        setSystemStatus(state.dataset ? "Operational" : "Waiting for data");
        updateProgress(0, "Upload failed", "The active dataset was not changed");
        showToast(error.message || "The report could not be processed.", "error");
    } finally {
        elements.fileInput.disabled = false;
        elements.fileInput.value = "";
    }
}

async function loadDefaultReport() {
    try {
        updateProgress(5, "Loading bundled report...", "Fetching AR_Report.xlsx");

        const response = await fetch(CONFIG.defaultReport, { cache: "no-store" });

        if (!response.ok) {
            throw new Error(`Default report returned HTTP ${response.status}.`);
        }

        const arrayBuffer = await response.arrayBuffer();
        const parsedWorkbook = await parseWorkbook(arrayBuffer);
        const prepared = validateRows(parsedWorkbook.rows);

        await activatePreparedDataset(prepared, "AR_Report.xlsx", {
            sheetName: parsedWorkbook.sheetName,
            headerRow: prepared.headerRow
        });

        resetProgress();
    } catch (error) {
        console.warn("Bundled report not loaded:", error);
        resetProgress();
        setSystemStatus("Waiting for data");
        elements.dataStatus.textContent = "No report loaded";
        elements.datasetMessage.textContent = "Upload an AR report to begin.";
    }
}

async function loadInitialDataset() {
    try {
        const saved = await loadSavedDataset();

        if (saved && Array.isArray(saved.rows) && saved.rows.length) {
            rebuildIndexesFromDataset(saved);
            updateDataPanel();
            renderDashboard();
            renderBrokers();
            setSystemStatus("Operational");
            elements.dataStatus.textContent = `${integer(saved.rows.length)} transactions loaded`;
            elements.datasetMessage.textContent = "Using the last validated report saved in this browser.";
            return;
        }
    } catch (error) {
        console.warn(error);
    }

    await loadDefaultReport();
}

/* =========================================================
   Identification engine
   ========================================================= */

function getRowsByIndexes(indexes) {
    if (!state.dataset) {
        return [];
    }

    return indexes.map((index) => state.dataset.rows[index]).filter(Boolean);
}

function rankRows(rows, paymentDate) {
    const targetDate = paymentDate ? parseDate(paymentDate) : null;

    return [...rows].sort((a, b) => {
        if (!targetDate) {
            return Math.abs(a.amount) - Math.abs(b.amount);
        }

        const aDays = Math.abs(a.date - targetDate);
        const bDays = Math.abs(b.date - targetDate);

        if (aDays !== bDays) {
            return aDays - bDays;
        }

        return Math.abs(a.amount) - Math.abs(b.amount);
    });
}

function findSubsetMatches(rows, target) {
    const candidates = rows
        .filter((row) => row.amount > 0 && row.amount <= target + 0.01)
        .sort((a, b) => b.amount - a.amount)
        .slice(0, CONFIG.maxCombinationCandidates);

    if (!candidates.length) {
        return [];
    }

    const states = new Map();
    states.set("0.00", []);

    for (let index = 0; index < candidates.length; index += 1) {
        const row = candidates[index];
        const currentStates = Array.from(states.entries());

        for (const [sumKey, selectedIndexes] of currentStates) {
            if (states.size >= CONFIG.maxCombinationStates) {
                break;
            }

            const nextSum = Number(sumKey) + row.amount;

            if (nextSum > target + 0.01) {
                continue;
            }

            const nextIndexes = [...selectedIndexes, index];

            if (nextIndexes.length > 6) {
                continue;
            }

            const nextKey = nextSum.toFixed(2);

            if (!states.has(nextKey)) {
                states.set(nextKey, nextIndexes);
            }

            if (Math.abs(nextSum - target) <= 0.01) {
                return [nextIndexes.map((candidateIndex) => candidates[candidateIndex])];
            }
        }
    }

    const closest = [...states.entries()]
        .map(([sum, selectedIndexes]) => ({
            sum: Number(sum),
            rows: selectedIndexes.map((index) => candidates[index])
        }))
        .filter((match) => match.rows.length > 1)
        .sort((a, b) => Math.abs(target - a.sum) - Math.abs(target - b.sum))
        .slice(0, 5);

    return closest.map((match) => match.rows);
}

function identifyPayment() {
    if (!state.dataset) {
        showToast("Load an AR report first.", "error");
        return;
    }

    const broker = searchKey(elements.brokerInput.value);
    const amount = Number(elements.amountInput.value);
    const paymentDate = elements.paymentDateInput.value;

    if (!broker) {
        showToast("Enter a broker.", "error");
        elements.brokerInput.focus();
        return;
    }

    if (!Number.isFinite(amount) || amount <= 0) {
        showToast("Enter a valid payment amount.", "error");
        elements.amountInput.focus();
        return;
    }

    const brokerIndexes = state.brokerIndex.get(broker) || [];

    if (!brokerIndexes.length) {
        const partial = findBrokerKeys(broker).flatMap((key) => state.brokerIndex.get(key) || []);

        if (!partial.length) {
            renderNoResults(`No transactions were found for “${elements.brokerInput.value.trim()}”.`);
            return;
        }

        renderIdentification(partial, amount, paymentDate, elements.brokerInput.value.trim());
        return;
    }

    renderIdentification(brokerIndexes, amount, paymentDate, elements.brokerInput.value.trim());
}

function findBrokerKeys(query) {
    return [...state.brokerIndex.keys()]
        .filter((key) => key.includes(query))
        .slice(0, 20);
}

function renderIdentification(indexes, targetAmount, paymentDate, brokerDisplay) {
    const rows = getRowsByIndexes([...new Set(indexes)]);
    const exactAmountRows = rows.filter((row) => Math.abs(row.amount - targetAmount) <= 0.01);
    const exactRows = rankRows(exactAmountRows, paymentDate).slice(0, 25);
    const combinationRows = findSubsetMatches(rows, targetAmount);
    const closestRows = rankRows(rows, paymentDate)
        .sort((a, b) => Math.abs(a.amount - targetAmount) - Math.abs(b.amount - targetAmount))
        .slice(0, 12);

    elements.identificationResults.innerHTML = `
        <div class="result-head">
            <div>
                <h2>Identification results</h2>
                <p>${integer(rows.length)} transactions available for ${escapeHtml(brokerDisplay)}.</p>
            </div>
            <span class="badge ${exactRows.length ? "green" : combinationRows.length ? "yellow" : "red"}">
                ${exactRows.length ? "Exact amount found" : combinationRows.length ? "Combination found" : "Review candidates"}
            </span>
        </div>
        ${renderExactSection(exactRows, targetAmount, paymentDate)}
        ${renderCombinationSection(combinationRows, targetAmount, paymentDate)}
        ${renderClosestSection(closestRows, targetAmount, paymentDate)}
    `;
}

function renderNoResults(message) {
    elements.identificationResults.innerHTML = `
        <div class="empty-state">
            <div class="empty-icon">!</div>
            <h2>No matching broker</h2>
            <p>${escapeHtml(message)}</p>
        </div>
    `;
}

function renderExactSection(rows, targetAmount, paymentDate) {
    if (!rows.length) {
        return "";
    }

    return `
        <div class="match-card">
            <div class="match-top">
                <div class="match-title">
                    <h3>Exact transaction matches</h3>
                    <p>One transaction has the same amount as the payment.</p>
                </div>
                <span class="badge green">Exact</span>
            </div>
            ${renderMatchTable(rows, targetAmount, paymentDate)}
        </div>
    `;
}

function renderCombinationSection(combinations, targetAmount, paymentDate) {
    if (!combinations.length) {
        return "";
    }

    return combinations.map((rows, index) => {
        const total = rows.reduce((sum, row) => sum + row.amount, 0);
        const difference = total - targetAmount;
        const exact = Math.abs(difference) <= 0.01;

        return `
            <div class="match-card">
                <div class="match-top">
                    <div class="match-title">
                        <h3>${exact ? "Exact combination" : `Closest combination #${index + 1}`}</h3>
                        <p>${rows.length} transactions combined against the payment amount.</p>
                    </div>
                    <span class="badge ${exact ? "green" : "yellow"}">${exact ? "Exact" : "Review"}</span>
                </div>
                <div class="match-summary">
                    <div class="summary-item"><small>Payment</small><strong>${money(targetAmount)}</strong></div>
                    <div class="summary-item"><small>Transactions</small><strong>${rows.length}</strong></div>
                    <div class="summary-item"><small>Combined</small><strong>${money(total)}</strong></div>
                    <div class="summary-item"><small>Difference</small><strong>${money(difference)}</strong></div>
                </div>
                ${renderMatchTable(rows, targetAmount, paymentDate)}
            </div>
        `;
    }).join("");
}

function renderClosestSection(rows, targetAmount, paymentDate) {
    if (!rows.length) {
        return `
            <div class="empty-state">
                <h2>No candidates available</h2>
                <p>The broker exists, but there are no valid transactions to compare.</p>
            </div>
        `;
    }

    return `
        <div class="match-card">
            <div class="match-top">
                <div class="match-title">
                    <h3>Closest candidates</h3>
                    <p>Useful when the payment does not exactly match a single AR movement.</p>
                </div>
                <span class="badge yellow">Review</span>
            </div>
            ${renderMatchTable(rows, targetAmount, paymentDate, true)}
        </div>
    `;
}

function renderMatchTable(rows, targetAmount, paymentDate, showDifference = false) {
    return `
        <table class="match-table">
            <thead>
                <tr>
                    <th>Date</th>
                    <th>Type</th>
                    <th>Invoice</th>
                    <th>Client</th>
                    <th>Broker</th>
                    <th>Amount</th>
                    <th>${showDifference ? "Difference" : "Balance"}</th>
                </tr>
            </thead>
            <tbody>
                ${rows.map((row) => {
                    const difference = row.amount - targetAmount;

                    return `
                        <tr>
                            <td>${formatDate(row.date)}</td>
                            <td>${escapeHtml(row.type || "—")}</td>
                            <td>${escapeHtml(row.invoice || "—")}</td>
                            <td>${escapeHtml(row.client || "—")}</td>
                            <td>${escapeHtml(row.broker || "—")}</td>
                            <td class="amount">${money(row.amount)}</td>
                            <td class="amount">${showDifference ? money(difference) : money(row.balance)}</td>
                        </tr>
                    `;
                }).join("")}
            </tbody>
        </table>
    `;
}

/* =========================================================
   Dashboard
   ========================================================= */

function renderDashboard() {
    if (!state.dataset) {
        return;
    }

    const { rows, metadata } = state.dataset;
    const lastBalance = metadata.lastBalance ?? 0;

    elements.kpiMovements.textContent = integer(rows.length);
    elements.kpiOpenAR.textContent = money(lastBalance);
    elements.kpiBrokers.textContent = integer(metadata.brokerCount);
    elements.kpiClients.textContent = integer(metadata.clientCount);

    renderBrokerSummary();
    renderBrokerChart();
}

function getTopBrokers(limit = 10) {
    if (!state.dataset) {
        return [];
    }

    return [...state.dataset.brokerTotals.entries()]
        .map(([key, total]) => ({
            key,
            name: findDisplayBroker(key),
            total,
            count: state.brokerIndex.get(key)?.length || 0
        }))
        .sort((a, b) => Math.abs(b.total) - Math.abs(a.total))
        .slice(0, limit);
}

function findDisplayBroker(key) {
    const index = state.brokerIndex.get(key)?.[0];
    return index === undefined ? key : state.dataset.rows[index].broker;
}

function renderBrokerSummary() {
    const brokers = getTopBrokers(10);

    if (!brokers.length) {
        elements.brokerSummary.innerHTML = `<div class="empty-state"><p>No broker data available.</p></div>`;
        return;
    }

    elements.brokerSummary.innerHTML = `
        <table class="mini-table">
            <thead>
                <tr>
                    <th>Broker</th>
                    <th>Rows</th>
                    <th>Total</th>
                </tr>
            </thead>
            <tbody>
                ${brokers.map((broker) => `
                    <tr>
                        <td>${escapeHtml(broker.name)}</td>
                        <td>${integer(broker.count)}</td>
                        <td>${money(broker.total)}</td>
                    </tr>
                `).join("")}
            </tbody>
        </table>
    `;
}

function renderBrokerChart() {
    if (!window.Chart || !elements.brokerChart) {
        return;
    }

    const brokers = getTopBrokers(10);

    if (state.chart) {
        state.chart.destroy();
    }

    state.chart = new Chart(elements.brokerChart, {
        type: "bar",
        data: {
            labels: brokers.map((broker) => broker.name.length > 24 ? `${broker.name.slice(0, 24)}…` : broker.name),
            datasets: [{
                label: "AR activity",
                data: brokers.map((broker) => broker.total),
                backgroundColor: "rgba(109, 74, 255, 0.72)",
                borderRadius: 5
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: "y",
            plugins: {
                legend: { display: false }
            },
            scales: {
                x: {
                    ticks: {
                        callback: (value) => money(value)
                    },
                    grid: {
                        color: "rgba(20, 28, 52, 0.06)"
                    }
                },
                y: {
                    grid: {
                        display: false
                    }
                }
            }
        }
    });
}

/* =========================================================
   Broker explorer
   ========================================================= */

function renderBrokers() {
    if (!state.dataset) {
        elements.brokerCount.textContent = "0 brokers";
        elements.brokerGrid.innerHTML = `<div class="empty-state"><p>Load a report to see brokers.</p></div>`;
        return;
    }

    const query = searchKey(elements.brokerSearch.value);
    const allBrokers = [...state.brokerIndex.keys()]
        .map((key) => ({
            key,
            name: findDisplayBroker(key),
            total: state.dataset.brokerTotals.get(key) || 0,
            count: state.brokerIndex.get(key)?.length || 0
        }))
        .filter((broker) => !query || broker.key.includes(query))
        .sort((a, b) => Math.abs(b.total) - Math.abs(a.total));

    elements.brokerCount.textContent = `${integer(allBrokers.length)} broker${allBrokers.length === 1 ? "" : "s"}`;

    if (!allBrokers.length) {
        elements.brokerGrid.innerHTML = `<div class="empty-state"><div class="empty-icon">◎</div><h2>No brokers found</h2><p>Try another search.</p></div>`;
        return;
    }

    elements.brokerGrid.innerHTML = allBrokers.slice(0, 300).map((broker) => `
        <article class="broker-card" data-broker-key="${escapeHtml(broker.key)}">
            <h3>${escapeHtml(broker.name)}</h3>
            <p>Normalized search key: ${escapeHtml(broker.key)}</p>
            <div class="amount">${money(broker.total)}</div>
            <div class="broker-meta">
                <span>${integer(broker.count)} movements</span>
                <span>View details →</span>
            </div>
        </article>
    `).join("");
}

function openBrokerModal(key) {
    if (!state.dataset) {
        return;
    }

    const indexes = state.brokerIndex.get(key) || [];
    const rows = getRowsByIndexes(indexes);
    const brokerName = findDisplayBroker(key);
    const total = rows.reduce((sum, row) => sum + row.amount, 0);

    elements.brokerModalContent.innerHTML = `
        <div class="detail-title">
            <h2 id="brokerModalTitle">${escapeHtml(brokerName)}</h2>
            <p>${integer(rows.length)} movements associated with this normalized broker.</p>
        </div>
        <div class="detail-stats">
            <div class="detail-stat"><small>Transactions</small><strong>${integer(rows.length)}</strong></div>
            <div class="detail-stat"><small>Total activity</small><strong>${money(total)}</strong></div>
            <div class="detail-stat"><small>Latest date</small><strong>${formatDate(rows.reduce((latest, row) => !latest || row.date > latest ? row.date : latest, null))}</strong></div>
        </div>
        <table class="detail-table">
            <thead>
                <tr>
                    <th>Date</th>
                    <th>Type</th>
                    <th>Invoice</th>
                    <th>Client</th>
                    <th>Amount</th>
                    <th>Balance</th>
                </tr>
            </thead>
            <tbody>
                ${rows.slice().sort((a, b) => b.date - a.date).slice(0, 500).map((row) => `
                    <tr>
                        <td>${formatDate(row.date)}</td>
                        <td>${escapeHtml(row.type || "—")}</td>
                        <td>${escapeHtml(row.invoice || "—")}</td>
                        <td>${escapeHtml(row.client || "—")}</td>
                        <td>${money(row.amount)}</td>
                        <td>${money(row.balance)}</td>
                    </tr>
                `).join("")}
            </tbody>
        </table>
    `;

    elements.brokerModal.classList.remove("hidden");
    elements.brokerModal.setAttribute("aria-hidden", "false");
}

function closeBrokerModal() {
    elements.brokerModal.classList.add("hidden");
    elements.brokerModal.setAttribute("aria-hidden", "true");
}

/* =========================================================
   Data panel
   ========================================================= */

function updateDataPanel() {
    if (!state.dataset) {
        return;
    }

    const { metadata } = state.dataset;

    elements.dataFileName.textContent = metadata.fileName || "—";
    elements.dataRows.textContent = integer(metadata.rowCount);
    elements.dataBrokers.textContent = integer(metadata.brokerCount);
    elements.dataClients.textContent = integer(metadata.clientCount);
    elements.dataStartDate.textContent = formatDate(metadata.startDate);
    elements.dataEndDate.textContent = formatDate(metadata.endDate);
    elements.dataBeginningBalance.textContent = money(metadata.beginningBalance);
    elements.dataLastBalance.textContent = money(metadata.lastBalance);
}

/* =========================================================
   Global search
   ========================================================= */

function globalSearch() {
    if (!state.dataset) {
        return;
    }

    const query = searchKey(elements.globalSearch.value);

    if (query.length < 2) {
        return;
    }

    const invoiceMatches = state.invoiceIndex.get(query) || [];

    if (invoiceMatches.length) {
        switchView("identification");
        const row = state.dataset.rows[invoiceMatches[0]];
        elements.brokerInput.value = row.broker;
        elements.amountInput.value = row.amount.toFixed(2);
        identifyPayment();
        return;
    }

    const brokerKey = [...state.brokerIndex.keys()].find((key) => key.includes(query));

    if (brokerKey) {
        switchView("brokers");
        elements.brokerSearch.value = query;
        renderBrokers();
        return;
    }

    const clientKey = [...state.clientIndex.keys()].find((key) => key.includes(query));

    if (clientKey) {
        switchView("identification");
        const index = state.clientIndex.get(clientKey)?.[0];
        const row = state.dataset.rows[index];

        if (row?.broker) {
            elements.brokerInput.value = row.broker;
            elements.amountInput.value = row.amount.toFixed(2);
            identifyPayment();
        }
        return;
    }

    const amount = Number(query.replace(/,/g, ""));

    if (Number.isFinite(amount) && amount > 0) {
        switchView("identification");
        elements.brokerInput.value = "";
        elements.amountInput.value = amount.toFixed(2);
        showToast("Enter a broker to narrow the amount search.");
    }
}

/* =========================================================
   Navigation and UI helpers
   ========================================================= */

function switchView(viewName) {
    document.querySelectorAll(".view").forEach((view) => {
        view.classList.toggle("active-view", view.id === `view-${viewName}`);
    });

    document.querySelectorAll(".nav-item").forEach((button) => {
        button.classList.toggle("active", button.dataset.view === viewName);
    });

    if (viewName === "dashboard") {
        renderDashboard();
    }

    if (viewName === "brokers") {
        renderBrokers();
    }
}

function setSystemStatus(status) {
    elements.systemStatus.textContent = status;
}

function showToast(message, type = "") {
    elements.toast.textContent = message;
    elements.toast.className = `toast show ${type}`.trim();

    clearTimeout(showToast.timer);

    showToast.timer = setTimeout(() => {
        elements.toast.className = "toast";
    }, 4200);
}

/* =========================================================
   Events
   ========================================================= */

function bindEvents() {
    document.querySelectorAll(".nav-item").forEach((button) => {
        button.addEventListener("click", () => switchView(button.dataset.view));
    });

    elements.identifyBtn.addEventListener("click", identifyPayment);

    [elements.brokerInput, elements.amountInput, elements.paymentDateInput].forEach((input) => {
        input.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
                identifyPayment();
            }
        });
    });

    elements.fileInput.addEventListener("change", (event) => {
        handleFile(event.target.files[0]);
    });

    ["dragenter", "dragover"].forEach((eventName) => {
        elements.dropZone.addEventListener(eventName, (event) => {
            event.preventDefault();
            elements.dropZone.classList.add("dragover");
        });
    });

    ["dragleave", "drop"].forEach((eventName) => {
        elements.dropZone.addEventListener(eventName, (event) => {
            event.preventDefault();
            elements.dropZone.classList.remove("dragover");
        });
    });

    elements.dropZone.addEventListener("drop", (event) => {
        const file = event.dataTransfer.files[0];
        handleFile(file);
    });

    elements.brokerSearch.addEventListener("input", renderBrokers);

    elements.brokerGrid.addEventListener("click", (event) => {
        const card = event.target.closest("[data-broker-key]");

        if (card) {
            openBrokerModal(card.dataset.brokerKey);
        }
    });

    elements.globalSearch.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
            globalSearch();
        }
    });

    elements.refreshBtn.addEventListener("click", async () => {
        if (state.dataset) {
            await loadInitialDataset();
        } else {
            await loadDefaultReport();
        }
    });

    elements.closeBrokerModal.addEventListener("click", closeBrokerModal);

    elements.brokerModal.addEventListener("click", (event) => {
        if (event.target.matches("[data-close-modal]")) {
            closeBrokerModal();
        }
    });

    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape") {
            closeBrokerModal();
        }
    });
}

/* =========================================================
   Boot
   ========================================================= */

document.addEventListener("DOMContentLoaded", async () => {
    bindEvents();
    await loadInitialDataset();
});
