export type ViewportAnchor = {
  scrollY: number;
  viewportTop: number;
};

export function cssAttrEscape(value: string) {
  if (typeof CSS !== "undefined" && typeof CSS.escape === "function") return CSS.escape(value);
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

export function reviewCardSelector(candidateId: string) {
  return `[data-candidate-id="${cssAttrEscape(candidateId)}"]`;
}

export function confirmFocusSelector(candidateId: string) {
  return `${reviewCardSelector(candidateId)} [data-confirm-focus]`;
}

export function confirmDocumentButtonSelector(candidateId: string) {
  return `${reviewCardSelector(candidateId)} [data-confirm-document][data-confirm-action="confirm"]:not(:disabled)`;
}

export function reviewCardListSelector() {
  return ".exhibit-card-list";
}

export function emailAttachmentsSelector(candidateId: string) {
  return `${reviewCardSelector(candidateId)} [data-email-attachments]`;
}

export function probeSelectorUntilFound<T>(
  query: (selector: string) => T | null | undefined,
  selector: string,
  schedule: (callback: () => void) => void,
  onFound: (element: T) => void,
  attempts = 8,
) {
  let cancelled = false;
  const probe = (remaining: number) => {
    schedule(() => {
      if (cancelled) return;
      const target = query(selector);
      if (target) {
        onFound(target);
        return;
      }
      if (remaining > 1) probe(remaining - 1);
    });
  };
  probe(Math.max(1, attempts));
  return { cancel() { cancelled = true; } };
}

export function firstActionableControl(root: { querySelector(selector: string): HTMLElement | null } | HTMLElement | null | undefined) {
  if (!root) return null;
  if ("querySelector" in root) {
    return root.querySelector("button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), summary")
      ?? (root as HTMLElement);
  }
  return root;
}

export function emailChildDescriptionSelector(childExhibitId: string) {
  return `[data-email-child-exhibit="${cssAttrEscape(childExhibitId)}"] input`;
}

export function probeUntilStable(
  query: () => HTMLElement | null | undefined,
  schedule: (callback: () => void) => void,
  onReady: (element: HTMLElement) => void,
  attempts = 8,
) {
  let cancelled = false;
  let lastTop: number | null = null;
  const probe = (remaining: number) => {
    schedule(() => {
      if (cancelled) return;
      const target = query();
      if (!target) {
        if (remaining > 1) probe(remaining - 1);
        return;
      }
      const top = target.getBoundingClientRect().top;
      if (lastTop !== null && Math.abs(top - lastTop) < 1) {
        onReady(target);
        return;
      }
      lastTop = top;
      if (remaining > 1) probe(remaining - 1);
      else onReady(target);
    });
  };
  probe(Math.max(1, attempts));
  return { cancel() { cancelled = true; } };
}

export function captureViewportAnchor(
  element: { getBoundingClientRect(): { top: number } } | null | undefined,
  scrollY: number,
): ViewportAnchor | null {
  if (!element) return null;
  return { scrollY, viewportTop: element.getBoundingClientRect().top };
}

export function restoredScrollY(anchor: ViewportAnchor, nextViewportTop: number) {
  return Math.max(0, anchor.scrollY + (nextViewportTop - anchor.viewportTop));
}

export function restoreViewportAnchor(
  element: { getBoundingClientRect(): { top: number } } | null | undefined,
  anchor: ViewportAnchor | null,
  scrollToY: (top: number) => void,
) {
  if (!element || !anchor) return;
  const next = restoredScrollY(anchor, element.getBoundingClientRect().top);
  if (Math.abs(next - anchor.scrollY) >= 1) scrollToY(next);
}

export function firstVisibleReviewCardId(
  cards: Array<{ getBoundingClientRect(): { top: number; bottom: number }; getAttribute(name: string): string | null }>,
  viewportHeight: number,
) {
  const visible = cards.find((card) => {
    const box = card.getBoundingClientRect();
    return box.bottom > 0 && box.top < viewportHeight;
  });
  return visible?.getAttribute("data-candidate-id") ?? cards[0]?.getAttribute("data-candidate-id") ?? null;
}

function isExcludedReviewCard(card: { className?: string; classList?: { contains(token: string): boolean } }) {
  return Boolean(card.classList?.contains("excluded") || (typeof card.className === "string" && /\bexcluded\b/.test(card.className)));
}

export function nextReviewCardId(
  cards: Array<{ getAttribute(name: string): string | null; className?: string; classList?: { contains(token: string): boolean } }>,
  confirmedId: string,
) {
  const index = cards.findIndex((card) => card.getAttribute("data-candidate-id") === confirmedId);
  if (index < 0) return null;
  const next = cards.slice(index + 1).find((card) => !isExcludedReviewCard(card));
  return next?.getAttribute("data-candidate-id") ?? confirmedId;
}

export function nextPendingConfirmCardId(
  cards: Array<{ getAttribute(name: string): string | null; className?: string; classList?: { contains(token: string): boolean } }>,
  confirmedId: string,
) {
  const index = cards.findIndex((card) => card.getAttribute("data-candidate-id") === confirmedId);
  if (index < 0) return null;
  const next = cards.slice(index + 1).find((card) => {
    if (isExcludedReviewCard(card)) return false;
    if (card.getAttribute("data-included") === "false") return false;
    if (card.getAttribute("data-confirmed") === "true") return false;
    if (card.getAttribute("data-confirmable") === "false") return false;
    return true;
  });
  return next?.getAttribute("data-candidate-id") ?? null;
}

export function restoreWindowScrollY(
  savedY: number,
  currentY: number,
  scrollToY: (top: number) => void,
) {
  if (currentY !== savedY) scrollToY(savedY);
}
