import { useEffect, useRef, useState } from "react";
import { openPreviewDocument, renderPreviewPage, type PreviewDocument, type PreviewLoadingTask } from "./lib/original-pdf-preview.ts";

export type OriginalPdfReviewPurpose = "source-review" | "template-preview";

export type OriginalPdfReviewProps = {
  file: File;
  name: string;
  accessibleText?: string;
  onPageRendered: () => void;
  purpose?: OriginalPdfReviewPurpose;
  errorFallback?: string;
};

function readablePreviewError(error: unknown, errorFallback: string) {
  const message = error instanceof Error ? error.message : String(error);
  if (/password/i.test(message)) return "This PDF is password protected and cannot be displayed here.";
  if (/invalid|corrupt|format/i.test(message)) return "This PDF could not be displayed because its structure is invalid or damaged.";
  return errorFallback;
}

/**
 * Renders the original PDF through the bundled PDF renderer. This deliberately
 * avoids Chromium's optional PDF-viewer plug-in: an iframe can report that it
 * loaded even when the plug-in has produced only a blank viewer shell.
 */
export function OriginalPdfReview({
  file,
  name,
  accessibleText = "",
  onPageRendered,
  purpose = "source-review",
  errorFallback = purpose === "template-preview"
    ? "The template PDF could not be displayed. Check the file and try again."
    : "The original PDF could not be displayed. Do not approve the OCR exception unless the preview loads successfully.",
}: OriginalPdfReviewProps) {
  const isTemplatePreview = purpose === "template-preview";
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const wrapRef = useRef<HTMLDivElement>(null);
  const previousRef = useRef<HTMLButtonElement>(null);
  const nextRef = useRef<HTMLButtonElement>(null);
  const documentRef = useRef<PreviewDocument | null>(null);
  const onPageRenderedRef = useRef(onPageRendered);
  const [pdf, setPdf] = useState<PreviewDocument | null>(null);
  const [pageNumber, setPageNumber] = useState(1);
  const [status, setStatus] = useState<"loading-document" | "loading-page" | "ready" | "error">("loading-document");
  const [errorMessage, setErrorMessage] = useState("");

  useEffect(() => {
    onPageRenderedRef.current = onPageRendered;
  }, [onPageRendered]);

  useEffect(() => {
    const controller = new AbortController();
    let document: PreviewDocument | null = null;
    setPdf(null);
    setPageNumber(1);
    setErrorMessage("");
    setStatus("loading-document");

    void (async () => {
      const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
      if (controller.signal.aborted) return;
      if (!pdfjs.GlobalWorkerOptions.workerSrc) {
        pdfjs.GlobalWorkerOptions.workerSrc = new URL(
          "../node_modules/pdfjs-dist/legacy/build/pdf.worker.min.mjs",
          import.meta.url,
        ).toString();
      }
      const openedDocument = await openPreviewDocument(file, (bytes) => pdfjs.getDocument({
        data: bytes,
        isEvalSupported: false,
        useWorkerFetch: false,
        verbosity: 0,
      }) as unknown as PreviewLoadingTask, controller.signal);
      documentRef.current = openedDocument;
      if (controller.signal.aborted || documentRef.current !== openedDocument) {
        if (documentRef.current === openedDocument) {
          documentRef.current = null;
          await openedDocument.destroy().catch(() => undefined);
        }
        return;
      }
      document = openedDocument;
      setPdf(document);
      setStatus("loading-page");
    })().catch((error: unknown) => {
      if (controller.signal.aborted) return;
      setErrorMessage(readablePreviewError(error, errorFallback));
      setStatus("error");
    });

    return () => {
      controller.abort();
      const opened = document ?? documentRef.current;
      documentRef.current = null;
      if (opened) void opened.destroy().catch(() => undefined);
    };
  }, [errorFallback, file]);

  useEffect(() => {
    if (!pdf) return;
    const controller = new AbortController();
    setErrorMessage("");
    setStatus("loading-page");

    void (async () => {
      const canvas = canvasRef.current;
      if (!canvas || controller.signal.aborted) return;
      await renderPreviewPage(pdf, pageNumber, canvas, window.devicePixelRatio || 1, controller.signal);
      if (controller.signal.aborted) return;
      setStatus("ready");
      onPageRenderedRef.current();
    })().catch((error: unknown) => {
      if (controller.signal.aborted || (error instanceof Error && ["AbortError", "RenderingCancelledException"].includes(error.name))) return;
      setErrorMessage(readablePreviewError(error, errorFallback));
      setStatus("error");
    });

    return () => {
      controller.abort();
    };
  }, [errorFallback, pageNumber, pdf]);

  useEffect(() => {
    const active = document.activeElement;
    const previousDisabled = previousRef.current?.disabled;
    const nextDisabled = nextRef.current?.disabled;
    if (active === previousRef.current && previousDisabled) {
      (nextRef.current && !nextRef.current.disabled ? nextRef.current : wrapRef.current)?.focus();
    } else if (active === nextRef.current && nextDisabled) {
      (previousRef.current && !previousRef.current.disabled ? previousRef.current : wrapRef.current)?.focus();
    }
  }, [pageNumber, pdf, status]);

  return <div className="original-pdf-review" data-preview-status={status} data-preview-purpose={purpose}>
    <div className="original-pdf-toolbar" role="toolbar" aria-label={isTemplatePreview ? "Template PDF page controls" : "Original PDF page controls"}>
      <button ref={previousRef} className="secondary-button compact-action" type="button" disabled={!pdf || pageNumber <= 1 || status === "loading-page"} onClick={() => setPageNumber((current) => Math.max(1, current - 1))}>Previous page</button>
      <strong>{pdf ? `Page ${pageNumber} of ${pdf.numPages}` : "Opening PDF…"}</strong>
      <button ref={nextRef} className="secondary-button compact-action" type="button" disabled={!pdf || pageNumber >= pdf.numPages || status === "loading-page"} onClick={() => setPageNumber((current) => Math.min(pdf?.numPages ?? current, current + 1))}>Next page</button>
    </div>
    <div ref={wrapRef} className="original-pdf-canvas-wrap" tabIndex={0} role="region" aria-label={isTemplatePreview ? `Template PDF page ${pageNumber}` : `Original PDF page ${pageNumber}`} aria-busy={status === "loading-document" || status === "loading-page"}>
      {(status === "loading-document" || status === "loading-page") && <p className="original-pdf-status" role="status">{status === "loading-document" ? `Opening ${name}…` : `Rendering page ${pageNumber}…`}</p>}
      {status === "error" && <div className="original-pdf-error" role="alert"><strong>Preview failed</strong><p>{errorMessage}</p></div>}
      <canvas ref={canvasRef} role="img" aria-label={isTemplatePreview ? `Visual rendering of page ${pageNumber} of the template PDF ${name}` : `Visual rendering of page ${pageNumber} of the original PDF ${name}`} aria-describedby={isTemplatePreview ? undefined : "original-pdf-accessibility-note"} hidden={status === "error"} />
    </div>
    {status === "ready" && <p className="original-pdf-confirmation" role="status">{isTemplatePreview ? "The template page is displayed. Review all relevant pages before confirming." : "The original PDF page is displayed. Review the relevant pages before closing this window."}</p>}
    {!isTemplatePreview && <section className="original-pdf-accessible-review" id="original-pdf-accessibility-note">
      {accessibleText.trim()
        ? <details><summary>Read extracted document text</summary><p>This text is supplied as an accessible review aid. Check the visual page as well because layout, signatures and images may not appear here.</p><pre>{accessibleText.trim().slice(0, 10_000)}</pre></details>
        : <p role="note"><strong>No accessible text is available for this document.</strong> A person must visually review the original page before any OCR exception is approved.</p>}
    </section>}
  </div>;
}
