import assert from "node:assert/strict";
import test from "node:test";

import {
  PDF_PREVIEW_LIMITS,
  openPreviewDocument,
  previewCanvasPlan,
  renderPreviewPage,
} from "../app/lib/original-pdf-preview.ts";

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

test("preview canvas is bounded for very tall admitted PDF pages", () => {
  const plan = previewCanvasPlan({ width: 625, height: 10_000 }, 2);
  assert.ok(plan.width <= PDF_PREVIEW_LIMITS.maxCanvasDimension);
  assert.ok(plan.height <= PDF_PREVIEW_LIMITS.maxCanvasDimension);
  assert.ok(plan.width * plan.height <= PDF_PREVIEW_LIMITS.maxCanvasPixels);
});

test("closing during file reading never creates a PDF loading task", async () => {
  const bytes = deferred();
  const controller = new AbortController();
  let taskCount = 0;
  const opening = openPreviewDocument(
    { arrayBuffer: () => bytes.promise },
    () => { taskCount += 1; throw new Error("must not run"); },
    controller.signal,
    1_000,
  );
  controller.abort();
  await assert.rejects(opening, { name: "AbortError" });
  bytes.resolve(new ArrayBuffer(4));
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(taskCount, 0);
});

test("closing during document opening destroys the loading task exactly once", async () => {
  const loaded = deferred();
  const controller = new AbortController();
  let destroys = 0;
  let documentDestroys = 0;
  const opening = openPreviewDocument(
    { arrayBuffer: async () => new ArrayBuffer(4) },
    () => ({ promise: loaded.promise, destroy: async () => { destroys += 1; } }),
    controller.signal,
    1_000,
  );
  await new Promise((resolve) => setImmediate(resolve));
  controller.abort();
  await assert.rejects(opening, { name: "AbortError" });
  assert.equal(destroys, 1);
  loaded.resolve({ numPages: 1, getPage: async () => { throw new Error("unused"); }, destroy: async () => { documentDestroys += 1; } });
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(documentDestroys, 1, "a document that resolves after the abort remains owned and is destroyed");
});

test("a document resolving after the opening deadline is destroyed exactly once", async () => {
  const loaded = deferred();
  const controller = new AbortController();
  let taskDestroys = 0;
  let documentDestroys = 0;
  const opening = openPreviewDocument(
    { arrayBuffer: async () => new ArrayBuffer(4) },
    () => ({ promise: loaded.promise, destroy: async () => { taskDestroys += 1; } }),
    controller.signal,
    5,
  );
  await assert.rejects(opening, /took too long and was stopped/);
  assert.equal(taskDestroys, 1);
  loaded.resolve({ numPages: 1, getPage: async () => { throw new Error("unused"); }, destroy: async () => { documentDestroys += 1; } });
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(documentDestroys, 1);
});

test("a page resolving after close is still cleaned and never rendered", async () => {
  const pageReady = deferred();
  const controller = new AbortController();
  let cleanups = 0;
  let renders = 0;
  const document = { numPages: 1, getPage: () => pageReady.promise, destroy: async () => undefined };
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ({}) };
  const rendering = renderPreviewPage(document, 1, canvas, 1, controller.signal);
  controller.abort();
  await assert.rejects(rendering, { name: "AbortError" });
  pageReady.resolve({
    getViewport: ({ scale }) => ({ width: 600 * scale, height: 800 * scale }),
    render: () => { renders += 1; return { promise: Promise.resolve(), cancel() {} }; },
    cleanup: () => { cleanups += 1; },
  });
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(renders, 0);
  assert.equal(cleanups, 1);
});

test("closing during rendering cancels and cleans exactly once", async () => {
  const rendered = deferred();
  const controller = new AbortController();
  let cancels = 0;
  let cleanups = 0;
  const page = {
    getViewport: ({ scale }) => ({ width: 600 * scale, height: 800 * scale }),
    render: () => ({ promise: rendered.promise, cancel: () => { cancels += 1; } }),
    cleanup: () => { cleanups += 1; },
  };
  const document = { numPages: 1, getPage: async () => page, destroy: async () => undefined };
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ({}) };
  const rendering = renderPreviewPage(document, 1, canvas, 1, controller.signal);
  await new Promise((resolve) => setImmediate(resolve));
  controller.abort();
  await assert.rejects(rendering, { name: "AbortError" });
  assert.equal(cancels, 1);
  assert.equal(cleanups, 1);
});
