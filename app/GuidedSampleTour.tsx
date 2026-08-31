"use client";

import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { TOUR_COPY, type TourStepId } from "./lib/guided-tour.ts";

type CardPosition = { top: number; left: number };

function cardPositionFor(target: DOMRect): CardPosition {
  const width = 320;
  const cardHeight = 108;
  const left = Math.min(Math.max(16, target.left), window.innerWidth - width - 16);
  const below = target.bottom + 12;
  if (below + cardHeight < window.innerHeight) return { top: below, left };
  return { top: Math.max(76, target.top - cardHeight - 12), left };
}

function escapeIsForOtherUi(event: KeyboardEvent, dialogOpen: boolean): boolean {
  if (dialogOpen || event.defaultPrevented) return true;
  const target = event.target;
  if (target instanceof HTMLElement) {
    if (target instanceof HTMLInputElement && target.type === "file") return true;
    if (target.closest("dialog, [aria-modal='true'], .document-picker, .confirmation-backdrop")) return true;
  }
  return Boolean(document.querySelector("dialog[open], [aria-modal='true'], .document-picker[open]"));
}

export function GuidedSampleTour({
  stepId,
  dialogOpen,
  onSkip,
}: {
  stepId: TourStepId | null;
  dialogOpen: boolean;
  onSkip: () => void;
}) {
  const skipRef = useRef<HTMLButtonElement>(null);
  const [targetRect, setTargetRect] = useState<DOMRect | null>(null);

  useEffect(() => {
    skipRef.current?.focus({ preventScroll: true });
  }, []);

  useEffect(() => {
    let ignoreFileDialogEscape = false;
    let releaseTimer = 0;
    const blockEscapeForNativeDialog = () => {
      ignoreFileDialogEscape = true;
      window.clearTimeout(releaseTimer);
    };
    const releaseEscapeAfterNativeDialog = () => {
      window.clearTimeout(releaseTimer);
      releaseTimer = window.setTimeout(() => {
        ignoreFileDialogEscape = false;
      }, 400);
    };
    const onKey = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      if (ignoreFileDialogEscape) return;
      if (escapeIsForOtherUi(event, dialogOpen)) return;
      onSkip();
    };
    window.addEventListener("blur", blockEscapeForNativeDialog);
    window.addEventListener("focus", releaseEscapeAfterNativeDialog);
    window.addEventListener("keydown", onKey);
    return () => {
      window.clearTimeout(releaseTimer);
      window.removeEventListener("blur", blockEscapeForNativeDialog);
      window.removeEventListener("focus", releaseEscapeAfterNativeDialog);
      window.removeEventListener("keydown", onKey);
    };
  }, [dialogOpen, onSkip]);

  useLayoutEffect(() => {
    if (!stepId || dialogOpen) {
      document.documentElement.removeAttribute("data-tour-step");
      document.querySelectorAll(".tour-current-target").forEach((node) => node.classList.remove("tour-current-target"));
      setTargetRect(null);
      return;
    }
    document.documentElement.setAttribute("data-tour-step", stepId);
    const measure = () => {
      document.querySelectorAll(".tour-current-target").forEach((node) => node.classList.remove("tour-current-target"));
      const target = document.querySelector<HTMLElement>(`[data-tour="${stepId}"]`);
      if (target) target.classList.add("tour-current-target");
      setTargetRect(target?.getBoundingClientRect() ?? null);
    };
    measure();
    const onScroll = () => measure();
    window.addEventListener("resize", onScroll);
    window.addEventListener("scroll", onScroll, true);
    const timer = window.setInterval(measure, 250);
    return () => {
      window.removeEventListener("resize", onScroll);
      window.removeEventListener("scroll", onScroll, true);
      window.clearInterval(timer);
      document.documentElement.removeAttribute("data-tour-step");
      document.querySelectorAll(".tour-current-target").forEach((node) => node.classList.remove("tour-current-target"));
    };
  }, [stepId, dialogOpen]);

  if (!stepId) return null;
  const copy = TOUR_COPY[stepId];
  const card = targetRect && !dialogOpen ? cardPositionFor(targetRect) : null;

  return (
    <div className="guided-tour-layer">
      <button ref={skipRef} className="guided-tour-skip" type="button" onClick={onSkip}>
        Skip walkthrough
      </button>
      {card ? (
        <div
          className="guided-tour-card"
          style={{ top: card.top, left: card.left }}
          role="status"
          aria-live="polite"
          aria-atomic="true"
        >
          <p className="guided-tour-card-title">{copy.title}</p>
          <p>{copy.body}</p>
        </div>
      ) : null}
    </div>
  );
}
