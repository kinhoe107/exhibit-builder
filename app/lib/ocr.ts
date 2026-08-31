export type OcrResult = {
  pages: Array<{ text: string; confidence: number }>;
};

const OCR_LIMITS = {
  fileBytes: 100 * 1024 * 1024,
  pages: 200,
  dimensionPixels: 7_000,
  pixels: 20_000_000,
  cumulativePixels: 1_500_000_000,
  pageMilliseconds: 45_000,
  totalMilliseconds: 20 * 60 * 1_000,
} as const;

import Tesseract from "tesseract.js";

/** Runs only in the desktop renderer. Every asset is packaged under /ocr. */
export async function ocrPdfLocally(file: File): Promise<OcrResult> {
  if (typeof document === "undefined") {
    throw new Error("OCR is available only in the local desktop application.");
  }
  if (!file.size || file.size > OCR_LIMITS.fileBytes) throw new Error("PDF exceeds the 100 MB OCR safety limit.");
  const startedAt = Date.now();
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const task = pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()), isEvalSupported: false, useWorkerFetch: false, verbosity: 0 });
  let openTimer: ReturnType<typeof setTimeout> | undefined;
  const pdf = await Promise.race([
    task.promise,
    new Promise<never>((_resolve, reject) => {
      openTimer = setTimeout(() => {
        void task.destroy().catch(() => undefined);
        reject(new Error("PDF did not open within the 30-second OCR safety limit."));
      }, 30_000);
    }),
  ]).finally(() => { if (openTimer) clearTimeout(openTimer); });
  if (pdf.numPages > OCR_LIMITS.pages) {
    await pdf.destroy();
    throw new Error(`PDF exceeds the ${OCR_LIMITS.pages}-page OCR safety limit.`);
  }
  let worker: Awaited<ReturnType<typeof Tesseract.createWorker>> | undefined;
  worker = await Tesseract.createWorker("eng", 1, {
    workerPath: "/ocr/worker.min.js",
    corePath: "/ocr",
    langPath: "/ocr",
    workerBlobURL: false,
    cacheMethod: "none",
    gzip: true,
  });
  try {
    const pages: OcrResult["pages"] = [];
    let cumulativePixels = 0;
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const remaining = OCR_LIMITS.totalMilliseconds - (Date.now() - startedAt);
      if (remaining <= 0) throw new Error("OCR exceeded the 20-minute total safety limit.");
      const page = await pdf.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 300 / 72 });
      if (viewport.width > OCR_LIMITS.dimensionPixels || viewport.height > OCR_LIMITS.dimensionPixels || viewport.width * viewport.height > OCR_LIMITS.pixels) throw new Error(`Page ${pageNumber} is too large for safe 300 dpi OCR.`);
      cumulativePixels += viewport.width * viewport.height;
      if (cumulativePixels > OCR_LIMITS.cumulativePixels) throw new Error("PDF exceeds the cumulative OCR pixel safety limit.");
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      const context = canvas.getContext("2d", { alpha: false });
      if (!context) throw new Error("OCR canvas could not be created.");
      const renderTask = page.render({ canvas, canvasContext: context, viewport });
      let timer: ReturnType<typeof setTimeout> | undefined;
      await Promise.race([
        renderTask.promise,
        new Promise<never>((_resolve, reject) => {
          timer = setTimeout(() => {
            renderTask.cancel();
            reject(new Error(`OCR page ${pageNumber} rendering exceeded the 45-second safety limit.`));
          }, Math.min(OCR_LIMITS.pageMilliseconds, remaining));
        }),
      ]).finally(() => { if (timer) clearTimeout(timer); });
      const recognitionRemaining = OCR_LIMITS.totalMilliseconds - (Date.now() - startedAt);
      if (recognitionRemaining <= 0) throw new Error("OCR exceeded the 20-minute total safety limit.");
      const result = await Promise.race([
        worker.recognize(canvas),
        new Promise<never>((_resolve, reject) => {
          timer = setTimeout(() => {
            void worker?.terminate().catch(() => undefined);
            reject(new Error(`OCR page ${pageNumber} recognition exceeded the 45-second safety limit.`));
          }, Math.min(OCR_LIMITS.pageMilliseconds, recognitionRemaining));
        }),
      ]).finally(() => { if (timer) clearTimeout(timer); });
      pages.push({ text: result.data.text.trim(), confidence: result.data.confidence });
      canvas.width = 1;
      canvas.height = 1;
    }
    return { pages };
  } finally {
    await worker?.terminate().catch(() => undefined);
    await pdf.destroy().catch(() => undefined);
  }
}
