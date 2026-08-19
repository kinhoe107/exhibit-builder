import { mkdir, readdir, stat } from "node:fs/promises";
import { resolve } from "node:path";

const downloadDir = resolve("tmp", `guided-roundtrip-${Date.now()}`);
await mkdir(downloadDir, { recursive: true });

const deadline = Date.now() + 15_000;
let pages;
while (!pages && Date.now() < deadline) {
  try { pages = await fetch("http://127.0.0.1:9223/json").then((response) => response.json()); }
  catch { await new Promise((resolveWait) => setTimeout(resolveWait, 100)); }
}
const page = pages?.find((item) => item.type === "page" && item.url.startsWith("http://127.0.0.1:"));
if (!page) throw new Error("Local browser page was not available.");
console.log("Packaged smoke: application page ready");

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

await command("Runtime.enable");
await command("Page.enable");
await command("Browser.setDownloadBehavior", { behavior: "allow", downloadPath: downloadDir, eventsEnabled: true });
await command("Page.navigate", { url: "http://127.0.0.1:4173" });
await waitFor(`document.querySelector('[data-testid="guided-sample-button"]')`);
await evaluate(`document.querySelector('[data-testid="guided-sample-button"]').click()`);
await waitFor(`document.querySelectorAll('.exhibit-review-card').length === 6 || document.querySelector('.error-toast')`);
console.log("Packaged smoke: guided analysis ready");

const editedDescription = "Round-trip saved guided description";
await evaluate(`(() => {
  const input = document.querySelector('.exhibit-review-card .description-input');
  const setValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
  setValue.call(input, ${JSON.stringify(editedDescription)});
  input.dispatchEvent(new Event('input', { bubbles: true }));
  return input.value;
})()`);
await waitFor(`document.querySelector('.exhibit-review-card .description-input')?.value === ${JSON.stringify(editedDescription)}`);

const before = await evaluate(`({
  marks: [...document.querySelectorAll('.exhibit-review-card .mark-badge')].map((item) => item.textContent.trim()),
  description: document.querySelector('.exhibit-review-card .description-input')?.value
})`);
await evaluate(`[...document.querySelectorAll('button')].find((button) => button.textContent.trim() === 'Save exhibit project').click()`);
console.log("Packaged smoke: save requested");

let archivePath;
const downloadDeadline = Date.now() + 30_000;
while (!archivePath && Date.now() < downloadDeadline) {
  const names = await readdir(downloadDir);
  const archive = names.find((name) => name.endsWith(".bundle-project"));
  if (archive && !names.some((name) => name.endsWith(".crdownload"))) {
    const candidate = resolve(downloadDir, archive);
    if ((await stat(candidate)).size > 0) archivePath = candidate;
  }
  if (!archivePath) await new Promise((resolveWait) => setTimeout(resolveWait, 100));
}
if (!archivePath) throw new Error("The guided project archive was not downloaded.");
console.log("Packaged smoke: project archive downloaded");

await evaluate(`window.confirm = () => true`);
await evaluate(`[...document.querySelectorAll('button')].find((button) => button.textContent.trim() === 'New exhibit project').click()`);
await waitFor(`document.querySelector('[data-testid="guided-sample-button"]')`);
console.log("Packaged smoke: returned to home screen");

const documentNode = await command("DOM.getDocument");
const inputNode = await command("DOM.querySelector", {
  nodeId: documentNode.root.nodeId,
  selector: 'input[accept=".bundle-project,.zip"]',
});
if (!inputNode.nodeId) throw new Error("The project-open file input was not found.");
await command("DOM.setFileInputFiles", { nodeId: inputNode.nodeId, files: [archivePath] });
await waitFor(`document.querySelectorAll('.exhibit-review-card').length === 6 || document.querySelector('.error-toast')`);
console.log("Packaged smoke: saved project reopened");

const after = await evaluate(`({
  error: document.querySelector('.error-toast span')?.textContent || null,
  marks: [...document.querySelectorAll('.exhibit-review-card .mark-badge')].map((item) => item.textContent.trim()),
  description: document.querySelector('.exhibit-review-card .description-input')?.value
})`);
if (after.error || after.description !== editedDescription || after.marks.length !== 6 || !after.marks.every((mark) => /^AH\s*\d+$/.test(mark))) {
  throw new Error(`Guided project did not round-trip consistently: ${JSON.stringify({ before, after })}`);
}

console.log(JSON.stringify({ archivePath, before, after }, null, 2));
socket.close();
process.exit(0);
