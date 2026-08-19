export const PDF_PREVIEW_LIMITS = Object.freeze({
  openTimeoutMs: 30_000,
  pageTimeoutMs: 10_000,
  renderTimeoutMs: 30_000,
  maxCanvasDimension: 8_192,
  maxCanvasPixels: 20_000_000,
});

export type PreviewViewport = { width: number; height: number };
export type PreviewRenderTask = { promise: Promise<void>; cancel(): void };
export type PreviewPage = {
  getViewport(options: { scale: number }): PreviewViewport;
  render(options: {
    canvas: HTMLCanvasElement;
    canvasContext: CanvasRenderingContext2D;
    viewport: PreviewViewport;
    transform?: number[];
  }): PreviewRenderTask;
  cleanup(): void;
};
export type PreviewDocument = {
  numPages: number;
  getPage(pageNumber: number): Promise<PreviewPage>;
  destroy(): Promise<void>;
};
export type PreviewLoadingTask = {
  promise: Promise<PreviewDocument>;
  destroy(): Promise<void>;
};

type ArrayBufferSource = { arrayBuffer(): Promise<ArrayBuffer> };

function abortError() {
  return new DOMException("The PDF preview was closed.", "AbortError");
}

function callOnce(action: () => void | Promise<void>) {
  let called = false;
  return () => {
    if (called) return;
    called = true;
    void action();
  };
}

export function withinPreviewDeadline<T>(
  promise: Promise<T>,
  signal: AbortSignal,
  timeoutMs: number,
  label: string,
  cancel: () => void | Promise<void> = () => undefined,
) {
  return new Promise<T>((resolve, reject) => {
    if (signal.aborted) {
      void cancel();
      reject(abortError());
      return;
    }
    let settled = false;
    const cancelOnce = callOnce(cancel);
    const finish = (action: () => void) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal.removeEventListener("abort", onAbort);
      action();
    };
    const onAbort = () => finish(() => {
      cancelOnce();
      reject(abortError());
    });
    const timer = setTimeout(() => finish(() => {
      cancelOnce();
      reject(new Error(`${label} took too long and was stopped.`));
    }), timeoutMs);
    signal.addEventListener("abort", onAbort, { once: true });
    promise.then(
      (value) => finish(() => resolve(value)),
      (error) => finish(() => reject(error instanceof Error ? error : new Error(String(error)))),
    );
  });
}

export async function openPreviewDocument(
  file: ArrayBufferSource,
  createLoadingTask: (bytes: Uint8Array) => PreviewLoadingTask,
  signal: AbortSignal,
  timeoutMs = PDF_PREVIEW_LIMITS.openTimeoutMs,
) {
  const bytes = await withinPreviewDeadline(
    file.arrayBuffer(),
    signal,
    timeoutMs,
    "Opening the PDF file",
  );
  if (signal.aborted) throw abortError();
  const task = createLoadingTask(new Uint8Array(bytes));
  const destroyTask = callOnce(() => task.destroy().catch(() => undefined));
  let returnedDocument: PreviewDocument | null = null;
  let ownershipAbandoned = false;
  // If PDF.js resolves despite an abort/destroy request, retain ownership of
  // that late document. A successfully returned document belongs to the UI.
  void task.promise.then((lateDocument) => {
    if (!returnedDocument && ownershipAbandoned) void lateDocument.destroy().catch(() => undefined);
  }, () => undefined);
  try {
    const document = await withinPreviewDeadline(
      task.promise,
      signal,
      timeoutMs,
      "Opening the PDF preview",
      destroyTask,
    );
    if (signal.aborted) {
      await document.destroy().catch(() => undefined);
      throw abortError();
    }
    returnedDocument = document;
    return document;
  } catch (error) {
    ownershipAbandoned = true;
    destroyTask();
    throw error;
  }
}

export function previewCanvasPlan(base: PreviewViewport, devicePixelRatio: number) {
  if (!Number.isFinite(base.width) || !Number.isFinite(base.height) || base.width <= 0 || base.height <= 0) {
    throw new Error("This PDF page has invalid dimensions and cannot be previewed.");
  }
  const outputScale = Math.min(Math.max(devicePixelRatio || 1, 1), 2);
  const preferredScale = Math.min(1.6, 1_000 / Math.max(1, base.width));
  const dimensionScale = Math.min(
    PDF_PREVIEW_LIMITS.maxCanvasDimension / (base.width * outputScale),
    PDF_PREVIEW_LIMITS.maxCanvasDimension / (base.height * outputScale),
  );
  const pixelScale = Math.sqrt(
    PDF_PREVIEW_LIMITS.maxCanvasPixels / (base.width * base.height * outputScale * outputScale),
  );
  const scale = Math.min(preferredScale, dimensionScale, pixelScale);
  if (!Number.isFinite(scale) || scale <= 0) {
    throw new Error("This PDF page is too large to preview safely.");
  }
  const width = Math.max(1, Math.floor(base.width * scale * outputScale));
  const height = Math.max(1, Math.floor(base.height * scale * outputScale));
  if (
    width > PDF_PREVIEW_LIMITS.maxCanvasDimension ||
    height > PDF_PREVIEW_LIMITS.maxCanvasDimension ||
    width * height > PDF_PREVIEW_LIMITS.maxCanvasPixels
  ) {
    throw new Error("This PDF page is too large to preview safely.");
  }
  return {
    scale,
    outputScale,
    width,
    height,
    cssWidth: Math.max(1, Math.floor(base.width * scale)),
    cssHeight: Math.max(1, Math.floor(base.height * scale)),
  };
}

export async function renderPreviewPage(
  document: PreviewDocument,
  pageNumber: number,
  canvas: HTMLCanvasElement,
  devicePixelRatio: number,
  signal: AbortSignal,
) {
  const pagePromise = document.getPage(pageNumber);
  let page: PreviewPage;
  try {
    page = await withinPreviewDeadline(
      pagePromise,
      signal,
      PDF_PREVIEW_LIMITS.pageTimeoutMs,
      `Loading PDF page ${pageNumber}`,
    );
  } catch (error) {
    // A page may resolve after the dialog has closed. Retain ownership of that
    // late value so PDF.js resources are still released deterministically.
    void pagePromise.then((latePage) => latePage.cleanup(), () => undefined);
    throw error;
  }
  const cleanupPage = callOnce(() => page.cleanup());
  let renderTask: PreviewRenderTask | null = null;
  try {
    if (signal.aborted) throw abortError();
    const base = page.getViewport({ scale: 1 });
    const plan = previewCanvasPlan(base, devicePixelRatio);
    const viewport = page.getViewport({ scale: plan.scale });
    const context = canvas.getContext("2d", { alpha: false });
    if (!context) throw new Error("The PDF preview canvas is unavailable.");
    canvas.width = plan.width;
    canvas.height = plan.height;
    canvas.style.width = `${plan.cssWidth}px`;
    canvas.style.height = `${plan.cssHeight}px`;
    renderTask = page.render({
      canvas,
      canvasContext: context,
      viewport,
      transform: plan.outputScale === 1 ? undefined : [plan.outputScale, 0, 0, plan.outputScale, 0, 0],
    });
    const cancelRender = callOnce(() => renderTask?.cancel());
    await withinPreviewDeadline(
      renderTask.promise,
      signal,
      PDF_PREVIEW_LIMITS.renderTimeoutMs,
      `Rendering PDF page ${pageNumber}`,
      cancelRender,
    );
    if (signal.aborted) throw abortError();
    return plan;
  } finally {
    cleanupPage();
  }
}
