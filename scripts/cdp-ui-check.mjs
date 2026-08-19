import { mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

const outputDir = resolve("tmp/ui-smoke");
await mkdir(outputDir, { recursive: true });

const deadline = Date.now() + 15_000;
let pages;
while (!pages && Date.now() < deadline) {
  try { pages = await fetch("http://127.0.0.1:9223/json").then((response) => response.json()); }
  catch { await new Promise((resolveWait) => setTimeout(resolveWait, 100)); }
}
const page = pages?.find((item) => item.type === "page" && item.url.includes("127.0.0.1:4173"));
if (!page) throw new Error("Local Chrome page was not available.");

const socket = new WebSocket(page.webSocketDebuggerUrl);
await new Promise((resolveOpen, rejectOpen) => {
  socket.addEventListener("open", resolveOpen, { once: true });
  socket.addEventListener("error", rejectOpen, { once: true });
});
let nextId = 1;
const pending = new Map();
socket.addEventListener("message", (event) => {
  const message = JSON.parse(event.data);
  if (!message.id || !pending.has(message.id)) return;
  const { resolveMessage, rejectMessage } = pending.get(message.id);
  pending.delete(message.id);
  if (message.error) rejectMessage(new Error(message.error.message));
  else resolveMessage(message.result);
});
function command(method, params = {}) {
  const id = nextId++;
  socket.send(JSON.stringify({ id, method, params }));
  return new Promise((resolveMessage, rejectMessage) => pending.set(id, { resolveMessage, rejectMessage }));
}
async function evaluate(expression) {
  const result = await command("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true, userGesture: true });
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.text);
  return result.result.value;
}
async function waitFor(expression, timeout = 90_000) {
  const end = Date.now() + timeout;
  while (Date.now() < end) {
    if (await evaluate(`Boolean(${expression})`)) return;
    await new Promise((resolveWait) => setTimeout(resolveWait, 100));
  }
  throw new Error(`Timed out waiting for ${expression}`);
}
async function screenshot(name, fullPage = true) {
  const metrics = await command("Page.getLayoutMetrics");
  const width = Math.ceil(metrics.cssContentSize?.width ?? 1440);
  const height = Math.ceil(metrics.cssContentSize?.height ?? 1000);
  const result = await command("Page.captureScreenshot", { format: "png", captureBeyondViewport: fullPage, clip: fullPage ? { x: 0, y: 0, width, height, scale: 1 } : undefined });
  await writeFile(resolve(outputDir, name), Buffer.from(result.data, "base64"));
}

await command("Runtime.enable");
await command("Page.enable");
await command("Emulation.setDeviceMetricsOverride", { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false });
await command("Page.navigate", { url: "http://127.0.0.1:4173" });
await waitFor(`document.querySelector('[data-testid="guided-sample-button"]')`);
const home = await evaluate(`({
  button: document.querySelector('[data-testid="guided-sample-button"]')?.textContent.trim(),
  note: document.querySelector('.hero-note')?.textContent.trim(),
  remoteLinks: [...document.querySelectorAll('a')].filter(a => /^https?:/i.test(a.href) && new URL(a.href).origin !== location.origin).length
})`);
if (home.button !== "Try the guided sample" || home.remoteLinks !== 0) throw new Error(`Unexpected home state: ${JSON.stringify(home)}`);
await screenshot("01-home.png", false);

await evaluate(`document.querySelector('[data-testid="guided-sample-button"]').click()`);
await waitFor(`document.querySelectorAll('.exhibit-review-card').length === 6 || document.querySelector('.error-toast')`, 90_000);
const review = await evaluate(`({
  cards: document.querySelectorAll('.exhibit-review-card').length,
  error: document.querySelector('.error-toast span')?.textContent || null,
  heading: document.querySelector('.workspace-header h1')?.textContent.trim(),
  marks: [...document.querySelectorAll('.exhibit-review-card .mark-badge')].map(x => x.textContent.trim()),
  paragraphThree: [...document.querySelectorAll('.exhibit-review-card .citation-preview strong')].filter(x => x.textContent.includes('Paragraph 3')).map(x => x.textContent.trim()),
  citationMetric: document.querySelector('.metric-row article:first-child small')?.textContent.trim(),
  metricLabels: [...document.querySelectorAll('.metric-row article > span:first-child')].map(item => item.textContent.trim()),
  optionalAuditClosed: [...document.querySelectorAll('.review-card-details')].every(item => !item.open),
  unusedFileAction: document.querySelector('.metric-link')?.textContent.trim()
})`);
if (review.error || review.cards !== 6 || review.heading !== 'Check the proposed exhibits' || review.paragraphThree.length !== 3 || !review.marks.every(mark => /^AH\s*\d+$/.test(mark)) || review.citationMetric !== 'Across 4 statement paragraphs' || JSON.stringify(review.metricLabels) !== JSON.stringify(['Citations found', 'References matched', 'References confirmed', 'Unused source files']) || !review.optionalAuditClosed || review.unusedFileAction !== 'Review these files') throw new Error(`Unexpected guided review: ${JSON.stringify(review)}`);
await screenshot("02-guided-review.png", true);

await evaluate(`document.querySelector('.statement-safety-note > summary').click()`);
await waitFor(`document.querySelector('[data-testid="witness-setting"]')`);
const witnessSettings = await evaluate(`({
  summary: document.querySelector('.statement-safety-note > summary')?.textContent.replace(/\\s+/g, ' ').trim(),
  filename: document.querySelector('[data-testid="witness-setting"] legend')?.textContent.trim(),
  labels: [...document.querySelectorAll('[data-testid="witness-setting"] label > span')].map(item => item.textContent.trim()),
  values: [...document.querySelectorAll('[data-testid="witness-setting"] input')].map(item => item.value),
  example: document.querySelector('[data-testid="witness-setting"] .witness-initials small')?.textContent.trim()
})`);
if (!witnessSettings.summary.includes('Project name, page numbering, templates and witness details') || witnessSettings.filename !== '01_GUIDED_SAMPLE_Witness_Statement.docx' || JSON.stringify(witnessSettings.labels) !== JSON.stringify(['Witness name', 'Exhibit initials']) || JSON.stringify(witnessSettings.values) !== JSON.stringify(['Guided Sample', 'AH']) || !witnessSettings.example.includes('AH1, AH2')) throw new Error(`Witness settings are unclear: ${JSON.stringify(witnessSettings)}`);
await screenshot("02a-witness-settings.png", true);
await command("Emulation.setDeviceMetricsOverride", { width: 1024, height: 900, deviceScaleFactor: 1, mobile: false });
await screenshot("02a-witness-settings-1024.png", true);
await command("Emulation.setDeviceMetricsOverride", { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false });
await evaluate(`document.querySelector('.statement-safety-note > summary').click()`);

await evaluate(`document.querySelector('.metric-link').click()`);
await waitFor(`document.querySelector('.reconcile-grid')`);
const unusedFiles = await evaluate(`({ heading: document.querySelector('.workspace-header h1')?.textContent.trim(), back: document.querySelector('.workspace-header button')?.textContent.trim(), count: document.querySelectorAll('.file-review-card').length })`);
if (unusedFiles.heading !== 'Reconcile the evidence inbox' || unusedFiles.back !== 'Back to exhibit review' || unusedFiles.count < 1) throw new Error(`Unused-file review is unclear: ${JSON.stringify(unusedFiles)}`);
await screenshot("02b-unused-files.png", true);
await evaluate(`document.querySelector('.workspace-header button').click()`);
await waitFor(`document.querySelectorAll('.exhibit-review-card').length === 6`);

const targets = [
  "01_SAMPLE_Agreement.pdf",
  "02_SAMPLE_Invoice.pdf",
  "03_SAMPLE_Claimant_Confirmation_Email.pdf",
  "04_SAMPLE_Claimant_Delivery_Email.pdf",
  "05_SAMPLE_Claimant_Payment_Email.pdf",
  "01_SAMPLE_Agreement.pdf",
];
await evaluate(`(() => {
  const targets = ${JSON.stringify(targets)};
  const cards = [...document.querySelectorAll('.exhibit-review-card')];
  cards.forEach((card, index) => {
    const select = card.querySelector('select[aria-label^="Matched file"]');
    const option = [...select.options].find(item => item.textContent.includes(targets[index]));
    if (!option) throw new Error('Missing source option ' + targets[index]);
    select.value = option.value;
    select.dispatchEvent(new Event('change', { bubbles: true }));
  });
  return true;
})()`);
await new Promise((resolveWait) => setTimeout(resolveWait, 500));
await evaluate(`(() => { document.querySelectorAll('.document-confirm input:not(:disabled):not(:checked)').forEach(input => input.click()); return true; })()`);
await new Promise((resolveWait) => setTimeout(resolveWait, 300));
await evaluate(`(() => { const same=[...document.querySelectorAll('.repeat-panel button')].find(button => button.textContent.trim()==='Same exhibit'); if(same) same.click(); return Boolean(same); })()`);
await new Promise((resolveWait) => setTimeout(resolveWait, 300));
await evaluate(`(() => { document.querySelectorAll('.repeat-confirm-button:not(:disabled)').forEach(button => button.click()); return true; })()`);
await waitFor(`[...document.querySelectorAll('button')].some(button => (button.textContent.includes('Continue to workbook sheets') || button.textContent.includes('Continue to finalise')) && !button.disabled)`, 15_000);
await evaluate(`[...document.querySelectorAll('button')].find(button => button.textContent.includes('Continue to workbook sheets') || button.textContent.includes('Continue to finalise')).click()`);
await waitFor(`document.querySelector('.finalise-card')`);

await evaluate(`document.querySelector('.add-exhibit-button').click()`);
await waitFor(`document.querySelector('.manual-exhibit-panel')`);
const finalise = await evaluate(`({
  heading: document.querySelector('.finalise-card h2')?.textContent.trim(),
  orderHeading: document.querySelector('.order-list-heading')?.textContent.trim(),
  warning: document.querySelector('.manual-exhibit-note')?.textContent.replace(/\\s+/g,' ').trim(),
  addLabel: document.querySelector('.manual-exhibit-panel .primary-button')?.textContent.trim()
})`);
if (finalise.heading !== "Choose the exhibit order" || finalise.orderHeading !== "Current bundle order" || !finalise.warning.includes("not cited in the statement") || finalise.addLabel !== "Add exhibit") throw new Error(`Unexpected finalise state: ${JSON.stringify(finalise)}`);
await screenshot("03-finalise-add-exhibit.png", true);

await evaluate(`(() => { const select=document.querySelector('.manual-exhibit-panel select'); const option=[...select.options].find(item => item.textContent.includes('06_SAMPLE_Unreferenced_Checklist.pdf')); select.value=option.value; select.dispatchEvent(new Event('change',{bubbles:true})); return true; })()`);
await new Promise((resolveWait) => setTimeout(resolveWait, 200));
await evaluate(`document.querySelector('.manual-exhibit-panel .primary-button').click()`);
await waitFor(`document.querySelector('.manual-exhibit-badge')`);
const manual = await evaluate(`({
  badge: document.querySelector('.manual-exhibit-badge')?.textContent.trim(),
  count: document.querySelectorAll('.finalise-order-list > li').length,
  suggestions: document.querySelectorAll('.manual-exhibit-badge').length
})`);
if (!manual.badge.includes("not cited in statement") || manual.count !== 6) throw new Error(`Manual addition failed: ${JSON.stringify(manual)}`);

await evaluate(`document.querySelector('.order-toolbar .primary-button').click()`);
await waitFor(`document.querySelector('.order-preview-banner')`);
const preview = await evaluate(`({
  label: document.querySelector('.order-preview-banner strong')?.textContent.trim(),
  heading: document.querySelector('.order-list-heading')?.textContent.trim(),
  actions: [...document.querySelectorAll('.order-preview-banner button')].map(button => button.textContent.trim())
})`);
if (preview.label !== "Preview only" || preview.heading !== "Proposed order" || !preview.actions.includes("Use this order") || !preview.actions.includes("Cancel preview")) throw new Error(`Order preview failed: ${JSON.stringify(preview)}`);
await screenshot("04-order-preview.png", true);

console.log(JSON.stringify({ home, review, finalise, manual, preview }, null, 2));
socket.close();
