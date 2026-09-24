import { JSDOM } from "jsdom";
import ExcelJS from "exceljs";
import fs from "node:fs/promises";
import path from "node:path";

const assetName = (await fs.readdir("dist/assets")).find((name) => name.endsWith(".js"));
const script = await fs.readFile(path.join("dist/assets", assetName), "utf8");
const templateBytes = await fs.readFile("public/template.xlsx");

const dom = new JSDOM("<!doctype html><html><body><div id=app></div></body></html>", {
  url: "http://127.0.0.1:5173/",
  pretendToBeVisual: true,
  runScripts: "outside-only",
});
const { window } = dom;
let lastDownload = null;
const clickedAnchors = [];
window.confirm = () => true;
window.fetch = async () => ({ arrayBuffer: async () => new window.Uint8Array(templateBytes).buffer });
window.URL.createObjectURL = (blob) => { lastDownload = { blob, filename: null }; return "blob:trip-claims-test"; };
window.URL.revokeObjectURL = () => {};
window.HTMLAnchorElement.prototype.click = function click() { clickedAnchors.push(this); if (lastDownload) lastDownload.filename = this.download; };
window.eval(script);

const wait = (ms = 30) => new Promise((resolve) => setTimeout(resolve, ms));
const assert = (condition, message) => { if (!condition) throw new Error(message); };

assert(window.document.querySelector("h1")?.textContent === "Trip Claims", "app shell did not render in DOM smoke test");
assert(window.document.body.textContent.includes("No expenses yet"), "empty state did not render in DOM smoke test");

window.document.querySelector('[data-action="add-expense"]').click();
const expenseForm = window.document.querySelector('form[data-form="expense"]');
expenseForm.querySelector('[name="date"]').value = "2026-05-04";
expenseForm.querySelector('[name="details"]').value = "Sample meal";
expenseForm.querySelector('[name="description"]').value = "Sample traveler";
expenseForm.querySelector('[name="currency"]').value = "VND";
expenseForm.querySelector('[name="amount"]').value = "120000";
expenseForm.querySelector('[name="payee"]').value = "Cash advance";
expenseForm.dispatchEvent(new window.Event("submit", { bubbles: true, cancelable: true }));
await wait();
assert(window.document.body.textContent.includes("Sample meal"), "expense did not save in DOM smoke test");

window.document.querySelector('[data-tab="trip"]').click();
const tripForm = window.document.querySelector('form[data-form="trip"]');
tripForm.querySelector('[name="name"]').value = "Sample User";
tripForm.querySelector('[name="destination"]').value = "Sample City";
tripForm.querySelector('[name="startDate"]').value = "2026-05-04";
tripForm.querySelector('[name="endDate"]').value = "2026-05-05";
tripForm.dispatchEvent(new window.Event("submit", { bubbles: true, cancelable: true }));
await wait();

window.document.querySelector('[data-tab="cash"]').click();
const cashForm = window.document.querySelector('form[data-form="cash"]');
cashForm.querySelector('[name="advance-VND"]').value = "4000000";
cashForm.querySelector('[name="note-500000"]').value = "2";
cashForm.dispatchEvent(new window.Event("submit", { bubbles: true, cancelable: true }));
await wait();

window.document.querySelector('[data-tab="export"]').click();
assert(window.document.body.textContent.includes("Ready to export"), "export state did not become ready");
window.document.querySelector('[data-action="export-xlsx"]').click();
for (let attempt = 0; attempt < 80 && !lastDownload?.filename; attempt += 1) await wait(50);
assert(lastDownload?.blob, "Excel export did not produce a Blob");
const exportedBytes = Buffer.from(await lastDownload.blob.arrayBuffer());
await fs.writeFile("/private/tmp/trip-claims-dom-smoke.xlsx", exportedBytes);

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.load(exportedBytes);
const sheet = workbook.getWorksheet("Expenses Reimbursement -Travel");
assert(sheet, "exported workbook is missing the template sheet");
assert(sheet.getCell("C6").value === "Sample City", "destination did not export");
assert(sheet.getCell("B14").value === "Sample meal", "expense details did not export");
assert(sheet.getCell("A14").value instanceof Date && sheet.getCell("A14").value.toISOString().startsWith("2026-05-04"), "expense date shifted during export");
assert(sheet.getCell("F14").value === 120000, "VND amount did not export");
assert(sheet.getCell("I14").value.result === 18, "VND-to-MYR reference did not export");
assert(sheet.getCell("F82").value.result === 120000, "expense subtotal did not export");
assert(sheet.getCell("F84").value === 4000000, "advance did not export");
assert(sheet.getCell("D91").value.result === 1000000, "cash return did not export");
assert(sheet.getCell("E88").value === null && sheet.getCell("G88").value === null, "stale template cash-return allocation remained");
const directDownloadFilename = lastDownload.filename;
const anchorsAfterDirectDownload = clickedAnchors.length;

let sharedFile = null;
Object.defineProperty(window.navigator, "canShare", { configurable: true, value: () => true });
Object.defineProperty(window.navigator, "share", { configurable: true, value: async ({ files }) => { sharedFile = files[0]; } });
lastDownload = null;
window.document.querySelector('[data-action="export-xlsx"]').click();
for (let attempt = 0; attempt < 80 && !window.document.querySelector('[data-action="share-export"]'); attempt += 1) await wait(50);
const shareButton = window.document.querySelector('[data-action="share-export"]');
assert(shareButton, "share-capable browser did not get a user-activated save button");
assert(clickedAnchors.length === anchorsAfterDirectDownload, "share-capable browser auto-clicked a download before user activation");
const downloadLink = window.document.querySelector(".export-download");
assert(downloadLink?.download === "Sample-City-expense-claim.xlsx", "share-capable browser did not get a direct Excel download link");
shareButton.click();
await wait();
assert(sharedFile?.name === "Sample-City-expense-claim.xlsx", "share action did not receive the generated Excel file");
assert(sharedFile.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "share action received an incorrect file type");
downloadLink.click();
assert(clickedAnchors.at(-1) === downloadLink, "direct Excel download link was not user-clickable");

console.log(JSON.stringify({
  rendered: true,
  expenseSaved: true,
  exportFilename: directDownloadFilename,
  shareFilename: sharedFile.name,
  exportedCells: { C6: sheet.getCell("C6").value, B14: sheet.getCell("B14").value, F14: sheet.getCell("F14").value, F82: sheet.getCell("F82").value.result, F84: sheet.getCell("F84").value, D91: sheet.getCell("D91").value.result },
}, null, 2));
window.close();
