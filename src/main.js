import ExcelJS from "exceljs";
import JSZip from "jszip";
import "./styles.css";

const STORAGE_KEY = "trip-claims-state-v1";
const DB_NAME = "trip-claims-receipts-v1";
const DB_STORE = "receipts";
const TEMPLATE_FILE = `${import.meta.env.BASE_URL}template.xlsx`;
const SHEET_NAME = "Expenses Reimbursement -Travel";
const VND_TO_MYR_DEFAULT = 0.00015;
const EXPENSE_START_ROW = 14;
const EXPENSE_END_ROW = 81;
const CASH_RETURN_START_ROW = 91;
const CASH_RETURN_END_ROW = 99;

const CURRENCY_META = {
  USD: { label: "USD", symbol: "$", column: 5 },
  VND: { label: "VND", symbol: "₫", column: 6 },
  MYR: { label: "MYR", symbol: "RM", column: 7 },
};

const defaultState = () => ({
  profile: {
    name: "",
    department: "",
    destination: "",
    startDate: "",
    endDate: "",
    attendees: "",
  },
  rates: {
    vndToMyr: VND_TO_MYR_DEFAULT,
    usdToMyr: "",
  },
  advance: { USD: "", VND: "", MYR: "" },
  expenses: [],
  cashReturn: [],
});

let state = loadState();
let activeTab = "expenses";
let editingExpenseId = null;
let pendingReceipt = null;
let removeReceiptOnSave = false;
let toastTimer = null;
let pendingExportFile = null;
let pendingDownloadUrl = null;

function loadState() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
    if (!saved) return defaultState();
    return {
      ...defaultState(),
      ...saved,
      profile: { ...defaultState().profile, ...(saved.profile || {}) },
      rates: { ...defaultState().rates, ...(saved.rates || {}) },
      advance: { ...defaultState().advance, ...(saved.advance || {}) },
      expenses: Array.isArray(saved.expenses) ? saved.expenses : [],
      cashReturn: Array.isArray(saved.cashReturn) ? saved.cashReturn : [],
    };
  } catch {
    return defaultState();
  }
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function uid(prefix = "id") {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function escapeHtml(value = "") {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function numberValue(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function optionalNumber(value) {
  if (value === "" || value === null || value === undefined) return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function money(value, currency = "MYR") {
  const n = numberValue(value);
  return new Intl.NumberFormat("en-MY", {
    style: "currency",
    currency,
    currencyDisplay: currency === "VND" ? "code" : "symbol",
    maximumFractionDigits: currency === "VND" ? 0 : 2,
  }).format(n);
}

function formatDate(value) {
  if (!value) return "";
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat("en-GB", {
    day: "numeric",
    month: "short",
    year: "numeric",
  }).format(date);
}

function formatDateRange(profile) {
  const start = profile.startDate;
  const end = profile.endDate;
  if (!start && !end) return "";
  if (!end || start === end) return formatDate(start || end);
  const startDate = new Date(`${start}T00:00:00`);
  const endDate = new Date(`${end}T00:00:00`);
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) return `${start} - ${end}`;
  const sameMonth = startDate.getFullYear() === endDate.getFullYear() && startDate.getMonth() === endDate.getMonth();
  if (sameMonth) {
    return `${startDate.getDate()}-${endDate.getDate()} ${new Intl.DateTimeFormat("en-GB", { month: "short" }).format(startDate)}, ${endDate.getFullYear()}`;
  }
  return `${formatDate(start)} - ${formatDate(end)}`;
}

function getRate(currency) {
  if (currency === "MYR") return 1;
  if (currency === "VND") return numberValue(state.rates.vndToMyr);
  return optionalNumber(state.rates.usdToMyr);
}

function expenseMyr(expense) {
  const rate = getRate(expense.currency);
  return rate === null ? null : numberValue(expense.amount) * rate;
}

function totals() {
  const byCurrency = { USD: 0, VND: 0, MYR: 0 };
  let myrEquivalent = 0;
  let knownEquivalent = true;
  state.expenses.forEach((expense) => {
    byCurrency[expense.currency] += numberValue(expense.amount);
    const converted = expenseMyr(expense);
    if (converted === null) knownEquivalent = false;
    else myrEquivalent += converted;
  });
  const advances = { USD: numberValue(state.advance.USD), VND: numberValue(state.advance.VND), MYR: numberValue(state.advance.MYR) };
  const advanceMyr = Object.entries(advances).reduce((sum, [currency, amount]) => {
    const rate = getRate(currency);
    return sum + (rate === null ? 0 : amount * rate);
  }, 0);
  return { byCurrency, myrEquivalent, knownEquivalent, advances, advanceMyr, netMyr: myrEquivalent - advanceMyr };
}

function openReceiptDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(DB_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveReceipt(id, file) {
  const db = await openReceiptDb();
  await new Promise((resolve, reject) => {
    const transaction = db.transaction(DB_STORE, "readwrite");
    transaction.objectStore(DB_STORE).put({ blob: file, name: file.name, type: file.type }, id);
    transaction.oncomplete = resolve;
    transaction.onerror = () => reject(transaction.error);
  });
  db.close();
}

async function getReceipt(id) {
  if (!id) return null;
  const db = await openReceiptDb();
  const result = await new Promise((resolve, reject) => {
    const request = db.transaction(DB_STORE, "readonly").objectStore(DB_STORE).get(id);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
  db.close();
  return result;
}

async function deleteReceipt(id) {
  if (!id) return;
  const db = await openReceiptDb();
  await new Promise((resolve, reject) => {
    const transaction = db.transaction(DB_STORE, "readwrite");
    transaction.objectStore(DB_STORE).delete(id);
    transaction.oncomplete = resolve;
    transaction.onerror = () => reject(transaction.error);
  });
  db.close();
}

function render() {
  const root = document.querySelector("#app");
  const summary = totals();
  root.innerHTML = `
    <div class="app-shell">
      <header class="topbar">
        <div>
          <div class="eyebrow">SOUTH ISLAND GARMENT</div>
          <h1>Trip Claims</h1>
        </div>
        <div class="status-pill"><span class="status-dot"></span><span id="connection-label">Offline-ready</span></div>
      </header>
      <nav class="tab-bar" aria-label="Trip sections">
        ${tabButton("expenses", "Expenses", `${state.expenses.length}`)}
        ${tabButton("trip", "Trip details")}
        ${tabButton("cash", "Cash")}
        ${tabButton("export", "Export")}
      </nav>
      <main class="content">
        ${activeTab === "expenses" ? renderExpenses(summary) : ""}
        ${activeTab === "trip" ? renderTrip() : ""}
        ${activeTab === "cash" ? renderCash(summary) : ""}
        ${activeTab === "export" ? renderExport(summary) : ""}
      </main>
    </div>
    <div id="modal-root"></div>
    <div id="toast" class="toast" role="status" aria-live="polite"></div>
  `;
  bindEvents();
  updateConnectionLabel();
}

function tabButton(id, label, count = "") {
  return `<button class="tab-button ${activeTab === id ? "is-active" : ""}" data-tab="${id}">${escapeHtml(label)}${count ? `<span class="tab-count">${escapeHtml(count)}</span>` : ""}</button>`;
}

function renderExpenses(summary) {
  const profileLabel = state.profile.destination || "Set up your trip details";
  return `
    <section class="hero-card">
      <div>
        <p class="section-kicker">CURRENT TRIP</p>
        <h2>${escapeHtml(profileLabel)}</h2>
        <p class="muted">${escapeHtml(formatDateRange(state.profile) || "Add dates before your first expense")}</p>
      </div>
      <button class="icon-button" data-action="open-trip" aria-label="Edit trip details">✎</button>
    </section>
    <section class="summary-card">
      <div class="summary-label">Estimated claim in MYR</div>
      <div class="summary-value">${summary.knownEquivalent ? money(summary.myrEquivalent, "MYR") : "Set USD rate"}</div>
      <div class="summary-subline">${state.expenses.length} expense${state.expenses.length === 1 ? "" : "s"} · Advance ${money(summary.advanceMyr, "MYR")}</div>
      <div class="currency-strip">
        ${["MYR", "VND", "USD"].map((currency) => `<div><span>${currency}</span><strong>${money(summary.byCurrency[currency], currency)}</strong></div>`).join("")}
      </div>
    </section>
    <div class="section-heading">
      <div><p class="section-kicker">EXPENSE LOG</p><h2>Recent expenses</h2></div>
      <button class="primary-button compact" data-action="add-expense">＋ Add expense</button>
    </div>
    ${state.expenses.length ? `<div class="expense-list">${state.expenses.slice().sort(sortExpenses).map(renderExpenseCard).join("")}</div>` : renderEmptyExpenses()}
    <p class="privacy-note">Saved on this device. You can keep adding expenses without a signal.</p>
  `;
}

function sortExpenses(a, b) {
  return String(b.date || "").localeCompare(String(a.date || "")) || String(b.createdAt || "").localeCompare(String(a.createdAt || ""));
}

function renderEmptyExpenses() {
  return `
    <section class="empty-card">
      <div class="empty-icon">＋</div>
      <h3>No expenses yet</h3>
      <p>Capture each receipt while it is still in your hand. Photos stay with the trip on this device.</p>
      <button class="primary-button" data-action="add-expense">Add your first expense</button>
    </section>
  `;
}

function renderExpenseCard(expense) {
  const converted = expenseMyr(expense);
  return `
    <article class="expense-card">
      <div class="expense-card-main">
        <div class="expense-date">${escapeHtml(formatDate(expense.date) || "No date")}</div>
        <h3>${escapeHtml(expense.details || "Untitled expense")}</h3>
        <p>${escapeHtml(expense.description || "No description")}</p>
        <div class="tag-row"><span class="tag">${escapeHtml(expense.payee || "Unassigned")}</span>${expense.receiptId ? `<span class="tag receipt-tag">Receipt attached</span>` : ""}</div>
      </div>
      <div class="expense-card-side">
        <strong>${money(expense.amount, expense.currency)}</strong>
        <span>${converted === null ? "MYR rate needed" : `≈ ${money(converted, "MYR")}`}</span>
        <div class="card-actions"><button data-action="edit-expense" data-id="${expense.id}">Edit</button><button data-action="delete-expense" data-id="${expense.id}" class="danger-link">Delete</button></div>
      </div>
    </article>
  `;
}

function renderTrip() {
  const p = state.profile;
  return `
    <section class="page-intro"><p class="section-kicker">TRIP SETUP</p><h2>Trip details</h2><p class="muted">These fields fill the top of the Excel claim report.</p></section>
    <form class="form-card" data-form="trip">
      <div class="field-grid">
        ${field("Your name", "name", p.name, "text", "Your full name", "required")}
        ${field("Department", "department", p.department, "text", "Your department", "required")}
        ${field("Travel to", "destination", p.destination, "text", "Destination city", "required", "full")}
        ${field("Start date", "startDate", p.startDate, "date", "", "required")}
        ${field("End date", "endDate", p.endDate, "date", "", "required")}
        ${field("Travelling with", "attendees", p.attendees, "text", "Optional", "", "full")}
      </div>
      <div class="form-actions"><button class="primary-button" type="submit">Save trip details</button><button class="secondary-button" type="button" data-action="reset-trip">Start a new trip</button></div>
    </form>
    <section class="info-card"><strong>Currency reference</strong><p>VND → MYR uses the template rate of 0.00015. Add a USD → MYR rate only when your company gives you the rate to use.</p><button class="text-button" data-action="open-rates">Edit currency rates</button></section>
  `;
}

function renderCash(summary) {
  const denominations = [500000, 200000, 100000, 50000, 20000, 10000, 5000, 2000, 1000];
  const rows = denominations.map((denomination) => {
    const existing = state.cashReturn.find((item) => numberValue(item.denomination) === denomination);
    return { denomination, quantity: existing?.quantity ?? "" };
  });
  return `
    <section class="page-intro"><p class="section-kicker">CASH RECONCILIATION</p><h2>Advance and cash returned</h2><p class="muted">The export fills the “Less: Advance” and “Cash returned” sections.</p></section>
    <form class="form-card" data-form="cash">
      <div class="field-grid">
        ${currencyField("Cash advance (USD)", "USD", state.advance.USD)}
        ${currencyField("Cash advance (VND)", "VND", state.advance.VND)}
        ${currencyField("Cash advance (MYR)", "MYR", state.advance.MYR)}
      </div>
      <div class="subsection-title"><h3>Cash returned in VND notes</h3><span class="muted">${money(cashReturnTotal(), "VND")}</span></div>
      <div class="cash-table"><div class="cash-table-head"><span>Note</span><span>Quantity</span><span>Value</span></div>${rows.map((row) => `<label class="cash-row"><span>${money(row.denomination, "VND")}</span><input inputmode="numeric" min="0" step="1" type="number" name="note-${row.denomination}" value="${escapeHtml(row.quantity)}" placeholder="0"><output>${money(row.denomination * numberValue(row.quantity), "VND")}</output></label>`).join("")}</div>
      <div class="form-actions"><button class="primary-button" type="submit">Save cash details</button></div>
    </form>
    <section class="summary-card light"><div class="summary-label">Estimated net claim</div><div class="summary-value">${summary.knownEquivalent ? money(summary.netMyr, "MYR") : "Set USD rate"}</div><div class="summary-subline">Expenses ${money(summary.myrEquivalent, "MYR")} less advance ${money(summary.advanceMyr, "MYR")}</div></section>
  `;
}

function currencyField(label, currency, value) {
  return `<label class="field"><span>${escapeHtml(label)}</span><div class="currency-input"><span>${currency}</span><input inputmode="decimal" min="0" step="0.01" type="number" name="advance-${currency}" value="${escapeHtml(value ?? "")}" placeholder="0"></div></label>`;
}

function cashReturnTotal() {
  return state.cashReturn.reduce((sum, item) => sum + numberValue(item.denomination) * numberValue(item.quantity), 0);
}

function renderExport(summary) {
  const missing = [];
  if (!state.profile.name) missing.push("your name");
  if (!state.profile.destination) missing.push("travel destination");
  if (!state.profile.startDate || !state.profile.endDate) missing.push("trip dates");
  if (!state.expenses.length) missing.push("at least one expense");
  if (state.expenses.some((expense) => expense.currency === "USD") && getRate("USD") === null) missing.push("USD → MYR rate");
  return `
    <section class="page-intro"><p class="section-kicker">CLAIM HANDOFF</p><h2>Export your claim</h2><p class="muted">Download the workbook when the trip is complete and submit it with your receipts.</p></section>
    <section class="export-card">
      <div class="export-status ${missing.length ? "needs-attention" : "ready"}"><span class="status-icon">${missing.length ? "!" : "✓"}</span><div><strong>${missing.length ? "Almost ready" : "Ready to export"}</strong><p>${missing.length ? `Complete ${escapeHtml(missing.join(", "))}.` : "Your Excel file will keep the company template layout."}</p></div></div>
      <div class="export-metrics"><div><span>Expenses</span><strong>${state.expenses.length}</strong></div><div><span>Gross claim</span><strong>${summary.knownEquivalent ? money(summary.myrEquivalent, "MYR") : "—"}</strong></div><div><span>Net after advance</span><strong>${summary.knownEquivalent ? money(summary.netMyr, "MYR") : "—"}</strong></div></div>
      <button class="primary-button wide" data-action="export-xlsx">Download Excel claim</button>
      <button class="secondary-button wide" data-action="export-receipts">Download receipt pack</button>
      <p class="small-note">The workbook uses the supplied “Expenses Reimbursement -Travel” sheet and preserves its totals, cash return and signature areas.</p>
    </section>
    <section class="info-card"><strong>On iPhone</strong><p>Use Safari’s Share button and choose “Add to Home Screen”. Your trip remains available without a signal after the first load.</p></section>
    <button class="text-button danger-link" data-action="reset-trip">Clear this trip and start again</button>
  `;
}

function field(label, name, value, type = "text", placeholder = "", required = "", extraClass = "") {
  return `<label class="field ${extraClass}"><span>${escapeHtml(label)}</span><input type="${type}" name="${escapeHtml(name)}" value="${escapeHtml(value ?? "")}" placeholder="${escapeHtml(placeholder)}" ${required}></label>`;
}

function renderRatesModal() {
  const modalRoot = document.querySelector("#modal-root");
  modalRoot.innerHTML = `<div class="modal-backdrop" data-action="close-modal"><section class="modal-sheet" role="dialog" aria-modal="true" aria-labelledby="rates-title"><div class="modal-handle"></div><div class="modal-header"><div><p class="section-kicker">SETTINGS</p><h2 id="rates-title">Currency rates</h2></div><button class="icon-button" data-action="close-modal" aria-label="Close">×</button></div><form data-form="rates"><label class="field"><span>VND → MYR</span><input inputmode="decimal" type="number" step="0.000001" min="0" name="vndToMyr" value="${escapeHtml(state.rates.vndToMyr)}"></label><label class="field"><span>USD → MYR <em>optional</em></span><input inputmode="decimal" type="number" step="0.0001" min="0" name="usdToMyr" value="${escapeHtml(state.rates.usdToMyr ?? "")}" placeholder="Enter company rate"></label><p class="small-note">The VND rate starts at the rate used in your Excel template. USD stays blank until you enter the company-approved rate.</p><button class="primary-button wide" type="submit">Save rates</button></form></section></div>`;
  bindModalEvents();
}

function renderExpenseModal(expense = null) {
  const value = expense || { date: state.profile.startDate || new Date().toISOString().slice(0, 10), details: "", description: "", currency: "VND", amount: "", payee: "", receiptId: "" };
  editingExpenseId = expense?.id || null;
  pendingReceipt = null;
  removeReceiptOnSave = false;
  document.querySelector("#modal-root").innerHTML = `<div class="modal-backdrop" data-action="close-modal"><section class="modal-sheet" role="dialog" aria-modal="true" aria-labelledby="expense-title"><div class="modal-handle"></div><div class="modal-header"><div><p class="section-kicker">EXPENSE LOG</p><h2 id="expense-title">${expense ? "Edit expense" : "Add expense"}</h2></div><button class="icon-button" data-action="close-modal" aria-label="Close">×</button></div><form data-form="expense"><div class="field-grid"><label class="field"><span>Date</span><input type="date" name="date" value="${escapeHtml(value.date)}" required></label><label class="field"><span>Currency</span><select name="currency">${["VND", "MYR", "USD"].map((currency) => `<option ${value.currency === currency ? "selected" : ""}>${currency}</option>`).join("")}</select></label><label class="field full"><span>Details</span><input type="text" name="details" value="${escapeHtml(value.details)}" placeholder="Dinner, transport, flight ticket" required></label><label class="field full"><span>Description / entertainment to</span><textarea name="description" rows="2" placeholder="Who was it for? What was it for?">${escapeHtml(value.description)}</textarea></label><label class="field"><span>Amount</span><input inputmode="decimal" min="0" step="0.01" type="number" name="amount" value="${escapeHtml(value.amount)}" placeholder="0.00" required></label><label class="field"><span>Payee</span><input type="text" name="payee" value="${escapeHtml(value.payee)}" placeholder="Gavin, Yvonne, Cash advance"></label></div><label class="receipt-picker"><span class="receipt-picker-icon">⌕</span><span><strong>${value.receiptId ? "Receipt attached" : "Attach receipt photo"}</strong><small>${value.receiptId ? "Choose another photo to replace it" : "Use your iPhone camera or Photos"}</small></span><input type="file" name="receipt" accept="image/*" capture="environment"></label>${value.receiptId ? `<label class="check-row"><input type="checkbox" name="removeReceipt"><span>Remove current receipt</span></label>` : ""}<div class="form-actions"><button class="primary-button wide" type="submit">${expense ? "Save changes" : "Add expense"}</button></div></form></section></div>`;
  bindModalEvents();
}

function bindEvents() {
  document.querySelectorAll("[data-tab]").forEach((button) => button.addEventListener("click", () => { activeTab = button.dataset.tab; render(); }));
  document.querySelectorAll("[data-action]").forEach((element) => element.addEventListener("click", handleAction));
  document.querySelectorAll("form[data-form]").forEach((form) => form.addEventListener("submit", handleSubmit));
  const receiptInput = document.querySelector('input[name="receipt"]');
  if (receiptInput) receiptInput.addEventListener("change", () => { pendingReceipt = receiptInput.files?.[0] || null; });
  window.addEventListener("online", updateConnectionLabel, { once: true });
  window.addEventListener("offline", updateConnectionLabel, { once: true });
}

function bindModalEvents() {
  document.querySelectorAll("#modal-root [data-action]").forEach((element) => element.addEventListener("click", handleAction));
  document.querySelectorAll("#modal-root form[data-form]").forEach((form) => form.addEventListener("submit", handleSubmit));
  const receiptInput = document.querySelector('#modal-root input[name="receipt"]');
  if (receiptInput) receiptInput.addEventListener("change", () => { pendingReceipt = receiptInput.files?.[0] || null; });
}

function updateConnectionLabel() {
  const label = document.querySelector("#connection-label");
  if (label) label.textContent = navigator.onLine ? "Synced on device" : "Offline-ready";
}

async function handleAction(event) {
  const action = event.currentTarget.dataset.action;
  if (action === "add-expense") renderExpenseModal();
  if (action === "edit-expense") renderExpenseModal(state.expenses.find((item) => item.id === event.currentTarget.dataset.id));
  if (action === "delete-expense") await removeExpense(event.currentTarget.dataset.id);
  if (action === "open-trip") { activeTab = "trip"; render(); }
  if (action === "open-rates") renderRatesModal();
  if (action === "close-modal" && event.target === event.currentTarget) closeModal();
  if (action === "reset-trip") resetTrip();
  if (action === "export-xlsx") {
    try { await exportWorkbook(); }
    catch (error) { console.error("Excel claim export failed:", error?.message || error); showToast("Could not create the Excel claim. Please try again."); }
  }
  if (action === "export-receipts") await exportReceiptPack();
  if (action === "share-export") sharePendingExport();
}

function closeModal() {
  document.querySelector("#modal-root").innerHTML = "";
  editingExpenseId = null;
  pendingReceipt = null;
}

async function handleSubmit(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const data = new FormData(form);
  if (form.dataset.form === "trip") {
    state.profile = { name: data.get("name").trim(), department: data.get("department").trim(), destination: data.get("destination").trim(), startDate: data.get("startDate"), endDate: data.get("endDate"), attendees: data.get("attendees").trim() };
    persist(); render(); showToast("Trip details saved");
  }
  if (form.dataset.form === "rates") {
    state.rates.vndToMyr = numberValue(data.get("vndToMyr")) || VND_TO_MYR_DEFAULT;
    state.rates.usdToMyr = data.get("usdToMyr") === "" ? "" : numberValue(data.get("usdToMyr"));
    persist(); closeModal(); render(); showToast("Currency rates saved");
  }
  if (form.dataset.form === "cash") {
    state.advance = { USD: data.get("advance-USD") || "", VND: data.get("advance-VND") || "", MYR: data.get("advance-MYR") || "" };
    state.cashReturn = [500000, 200000, 100000, 50000, 20000, 10000, 5000, 2000, 1000].map((denomination) => ({ denomination, quantity: data.get(`note-${denomination}`) || "" })).filter((item) => numberValue(item.quantity) > 0);
    persist(); render(); showToast("Cash details saved");
  }
  if (form.dataset.form === "expense") await saveExpense(data);
}

async function saveExpense(data) {
  const id = editingExpenseId || uid("expense");
  const existing = state.expenses.find((item) => item.id === id);
  const next = { id, createdAt: existing?.createdAt || new Date().toISOString(), date: data.get("date"), details: data.get("details").trim(), description: data.get("description").trim(), currency: data.get("currency"), amount: numberValue(data.get("amount")), payee: data.get("payee").trim(), receiptId: existing?.receiptId || "" };
  const removeCurrent = data.get("removeReceipt") === "on" || removeReceiptOnSave;
  if (removeCurrent && next.receiptId) { await deleteReceipt(next.receiptId); next.receiptId = ""; }
  if (pendingReceipt) {
    const previousReceiptId = next.receiptId;
    next.receiptId = uid("receipt");
    await saveReceipt(next.receiptId, pendingReceipt);
    if (previousReceiptId) await deleteReceipt(previousReceiptId);
  }
  state.expenses = existing ? state.expenses.map((item) => item.id === id ? next : item) : [...state.expenses, next];
  persist(); closeModal(); render(); showToast(existing ? "Expense updated" : "Expense added");
}

async function removeExpense(id) {
  const expense = state.expenses.find((item) => item.id === id);
  if (!expense || !window.confirm("Delete this expense?")) return;
  if (expense.receiptId) await deleteReceipt(expense.receiptId);
  state.expenses = state.expenses.filter((item) => item.id !== id);
  persist(); render(); showToast("Expense deleted");
}

function resetTrip() {
  if (!window.confirm("Clear this trip from this device? This cannot be undone.")) return;
  const receiptIds = state.expenses.map((expense) => expense.receiptId).filter(Boolean);
  Promise.all(receiptIds.map(deleteReceipt)).catch(() => {});
  state = defaultState();
  persist();
  activeTab = "trip";
  render();
  showToast("New trip started");
}

function showToast(message) {
  const toast = document.querySelector("#toast");
  if (!toast) return;
  toast.textContent = message;
  toast.classList.add("is-visible");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("is-visible"), 2600);
}

function setCellValue(worksheet, address, value) {
  worksheet.getCell(address).value = value ?? null;
}

function clearTemplateRange(worksheet, startRow, endRow, startCol = 1, endCol = 10) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) worksheet.getRow(row).getCell(col).value = null;
  }
}

function exportFile(buffer, filename, type = "application/octet-stream") {
  const file = new File([buffer], filename, { type });
  if (pendingDownloadUrl) URL.revokeObjectURL(pendingDownloadUrl);
  pendingDownloadUrl = null;
  document.querySelectorAll(".export-download").forEach((link) => link.remove());
  const url = URL.createObjectURL(file);
  pendingExportFile = null;
  document.querySelectorAll('[data-action="share-export"]').forEach((button) => button.remove());
  const exportCard = document.querySelector(".export-card");
  if (exportCard && typeof navigator.share === "function" && navigator.canShare?.({ files: [file] })) {
    pendingExportFile = file;
    const button = document.createElement("button");
    button.className = "secondary-button wide";
    button.dataset.action = "share-export";
    button.textContent = `Save or share ${filename}`;
    button.addEventListener("click", sharePendingExport);
    const downloadLink = document.createElement("a");
    downloadLink.className = "secondary-button wide export-download";
    downloadLink.href = url;
    downloadLink.download = filename;
    downloadLink.textContent = "Download Excel file instead";
    pendingDownloadUrl = url;
    exportCard.appendChild(button);
    exportCard.appendChild(downloadLink);
    showToast("File ready. Share it, or tap Download Excel file instead.");
    return true;
  }
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
  return false;
}

function sharePendingExport() {
  if (!pendingExportFile) return;
  const file = pendingExportFile;
  try {
    navigator.share({ files: [file], title: file.name }).then(() => {
      pendingExportFile = null;
      document.querySelector('[data-action="share-export"]')?.remove();
      showToast("File shared or saved");
    }).catch((error) => {
      showToast(error?.name === "AbortError" ? "Sharing cancelled. Use Download Excel file instead." : "Sharing failed. Use Download Excel file instead.");
    });
  } catch {
    showToast("Sharing failed. Use Download Excel file instead.");
  }
}

function safeFilePart(value) {
  return String(value || "expense").replace(/[^a-z0-9-_]+/gi, "-").replace(/^-|-$/g, "").slice(0, 50) || "expense";
}

function toExcelDate(dateValue) {
  const [year, month, day] = String(dateValue).split("-").map(Number);
  return new Date(Date.UTC(year, month - 1, day));
}

async function exportWorkbook() {
  if (!state.expenses.length) { showToast("Add at least one expense first"); return; }
  const buffer = await fetch(TEMPLATE_FILE).then((response) => response.arrayBuffer());
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.getWorksheet(SHEET_NAME) || workbook.worksheets[0];
  const profile = state.profile;
  const name = profile.attendees ? `${profile.name}  (together with ${profile.attendees})` : profile.name;
  setCellValue(worksheet, "C4", name);
  setCellValue(worksheet, "C5", profile.department);
  setCellValue(worksheet, "C6", profile.destination);
  setCellValue(worksheet, "C7", formatDateRange(profile));
  clearTemplateRange(worksheet, EXPENSE_START_ROW, EXPENSE_END_ROW);
  const sortedExpenses = state.expenses.slice().sort(sortExpenses).reverse();
  if (sortedExpenses.length > EXPENSE_END_ROW - EXPENSE_START_ROW + 1) {
    showToast("The template supports up to 68 expense rows");
    return;
  }
  let previousDate = "";
  sortedExpenses.forEach((expense, index) => {
    const rowNo = EXPENSE_START_ROW + index;
    const row = worksheet.getRow(rowNo);
    const showDate = expense.date !== previousDate;
    row.getCell(1).value = showDate ? toExcelDate(expense.date) : null;
    row.getCell(2).value = expense.details;
    row.getCell(4).value = expense.description;
    row.getCell(5).value = expense.currency === "USD" ? expense.amount : null;
    row.getCell(6).value = expense.currency === "VND" ? expense.amount : null;
    row.getCell(7).value = expense.currency === "MYR" ? expense.amount : null;
    row.getCell(8).value = null;
    const rate = getRate(expense.currency);
    if (expense.currency === "VND" && rate !== null) row.getCell(9).value = { formula: `F${rowNo}*${rate}`, result: expense.amount * rate };
    else if (expense.currency === "USD" && rate !== null) row.getCell(9).value = { formula: `E${rowNo}*${rate}`, result: expense.amount * rate };
    else row.getCell(9).value = null;
    row.getCell(10).value = expense.payee;
    previousDate = expense.date;
  });
  for (let rowNo = EXPENSE_START_ROW; rowNo <= EXPENSE_END_ROW; rowNo += 1) {
    const row = worksheet.getRow(rowNo);
    if (row.getCell(1).value instanceof Date) row.getCell(1).numFmt = "d-mmm-yy";
  }
  setCellValue(worksheet, "E82", { formula: `SUM(E${EXPENSE_START_ROW}:E${EXPENSE_END_ROW})`, result: totals().byCurrency.USD });
  setCellValue(worksheet, "F82", { formula: `SUM(F${EXPENSE_START_ROW}:F${EXPENSE_END_ROW})`, result: totals().byCurrency.VND });
  setCellValue(worksheet, "G82", { formula: `SUM(G${EXPENSE_START_ROW}:G${EXPENSE_END_ROW})`, result: totals().byCurrency.MYR });
  setCellValue(worksheet, "I82", { formula: `SUM(I${EXPENSE_START_ROW}:I${EXPENSE_END_ROW})`, result: state.expenses.reduce((sum, expense) => sum + (expense.currency === "MYR" ? 0 : expenseMyr(expense) || 0), 0) });
  ["E", "F", "G", "I"].forEach((column) => {
    setCellValue(worksheet, `${column}83`, { formula: `${column}82`, result: worksheet.getCell(`${column}82`).value.result });
  });
  setCellValue(worksheet, "E84", state.advance.USD === "" ? null : numberValue(state.advance.USD));
  setCellValue(worksheet, "F84", state.advance.VND === "" ? null : numberValue(state.advance.VND));
  setCellValue(worksheet, "G84", state.advance.MYR === "" ? null : numberValue(state.advance.MYR));
  setCellValue(worksheet, "I84", null);
  ["E", "F", "G", "I"].forEach((column) => {
    const expenseTotal = numberValue(worksheet.getCell(`${column}83`).value?.result);
    const advance = numberValue(worksheet.getCell(`${column}84`).value);
    setCellValue(worksheet, `${column}86`, { formula: `${column}83-${column}84`, result: expenseTotal - advance });
  });
  setCellValue(worksheet, "E88", null);
  setCellValue(worksheet, "G88", null);
  setCellValue(worksheet, "E89", null);
  setCellValue(worksheet, "G89", null);
  setCellValue(worksheet, "A89", "Cash returned");
  clearTemplateRange(worksheet, CASH_RETURN_START_ROW, CASH_RETURN_END_ROW, 2, 4);
  const denominations = [500000, 200000, 100000, 50000, 20000, 10000, 5000, 2000, 1000];
  denominations.forEach((denomination, index) => {
    const rowNo = CASH_RETURN_START_ROW + index;
    const entry = state.cashReturn.find((item) => numberValue(item.denomination) === denomination);
    worksheet.getCell(`B${rowNo}`).value = denomination;
    worksheet.getCell(`C${rowNo}`).value = entry ? numberValue(entry.quantity) : null;
    worksheet.getCell(`D${rowNo}`).value = entry ? { formula: `C${rowNo}*B${rowNo}`, result: denomination * numberValue(entry.quantity) } : null;
  });
  worksheet.getCell("D101").value = { formula: `SUM(D${CASH_RETURN_START_ROW}:D${CASH_RETURN_END_ROW})`, result: cashReturnTotal() };
  const output = await workbook.xlsx.writeBuffer();
  const filename = `${safeFilePart(profile.destination || "business-trip")}-expense-claim.xlsx`;
  if (!exportFile(output, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) showToast("Excel claim downloaded");
}

async function exportReceiptPack() {
  const expensesWithReceipts = state.expenses.filter((expense) => expense.receiptId);
  if (!expensesWithReceipts.length) { showToast("No receipt photos attached yet"); return; }
  const zip = new JSZip();
  for (const [index, expense] of expensesWithReceipts.entries()) {
    const receipt = await getReceipt(expense.receiptId);
    if (!receipt) continue;
    const extension = receipt.type?.split("/")[1]?.replace("jpeg", "jpg") || "jpg";
    zip.file(`${String(index + 1).padStart(2, "0")}-${safeFilePart(expense.date)}-${safeFilePart(expense.details)}.${extension}`, receipt.blob);
  }
  const output = await zip.generateAsync({ type: "blob" });
  if (!exportFile(output, `${safeFilePart(state.profile.destination || "business-trip")}-receipts.zip`, "application/zip")) showToast("Receipt pack downloaded");
}

if ("serviceWorker" in navigator) navigator.serviceWorker.register(`${import.meta.env.BASE_URL}sw.js`).catch(() => {});
render();
