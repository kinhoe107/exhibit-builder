import assert from "node:assert/strict";
import test from "node:test";
import {
  captureViewportAnchor,
  confirmDocumentButtonSelector,
  confirmFocusSelector,
  cssAttrEscape,
  emailAttachmentsSelector,
  firstVisibleReviewCardId,
  nextPendingConfirmCardId,
  nextReviewCardId,
  probeSelectorUntilFound,
  probeUntilStable,
  restoreViewportAnchor,
  restoreWindowScrollY,
  restoredScrollY,
  reviewCardListSelector,
  reviewCardSelector,
} from "../app/lib/review-viewport.ts";

test("confirm focus stays on the compact review control without using scrollIntoView", () => {
  assert.equal(cssAttrEscape('id"x'), 'id\\"x');
  assert.equal(reviewCardSelector("candidate-1"), '[data-candidate-id="candidate-1"]');
  assert.equal(confirmFocusSelector("candidate-1"), '[data-candidate-id="candidate-1"] [data-confirm-focus]');
});

test("document confirm restores the captured window scroll rather than following a collapsing card", () => {
  const scrolled = [];
  restoreWindowScrollY(640, 500, (top) => scrolled.push(top));
  assert.deepEqual(scrolled, [640]);
  restoreWindowScrollY(640, 640, (top) => scrolled.push(top));
  assert.deepEqual(scrolled, [640]);
  restoreWindowScrollY(640.75, 640, (top) => scrolled.push(top));
  assert.deepEqual(scrolled, [640, 640.75]);
});

test("restores an element-relative viewport only for non-confirm paths", () => {
  const anchor = captureViewportAnchor({ getBoundingClientRect: () => ({ top: 180 }) }, 640);
  assert.deepEqual(anchor, { scrollY: 640, viewportTop: 180 });
  assert.equal(restoredScrollY(anchor, 40), 500);
  const scrolled = [];
  restoreViewportAnchor({ getBoundingClientRect: () => ({ top: 40 }) }, anchor, (top) => scrolled.push(top));
  assert.deepEqual(scrolled, [500]);
  restoreViewportAnchor({ getBoundingClientRect: () => ({ top: 180 }) }, anchor, (top) => scrolled.push(top));
  assert.deepEqual(scrolled, [500]);
});

test("bulk confirmation anchors the first visible review card", () => {
  const cards = [
    { getBoundingClientRect: () => ({ top: -120, bottom: -20 }), getAttribute: () => "a" },
    { getBoundingClientRect: () => ({ top: 40, bottom: 220 }), getAttribute: () => "b" },
    { getBoundingClientRect: () => ({ top: 240, bottom: 420 }), getAttribute: () => "c" },
  ];
  assert.equal(firstVisibleReviewCardId(cards, 700), "b");
});

test("confirm anchors the next review card after the collapsing card", () => {
  const cards = [
    { getAttribute: () => "a" },
    { getAttribute: () => "b" },
    { getAttribute: () => "c" },
  ];
  assert.equal(nextReviewCardId(cards, "a"), "b");
  assert.equal(nextReviewCardId(cards, "c"), "c");
  assert.equal(nextReviewCardId(cards, "missing"), null);
});

test("confirm skips excluded and already-confirmed cards when choosing the next pending control", () => {
  const cards = [
    { getAttribute: (name) => ({ "data-candidate-id": "a", "data-included": "true", "data-confirmed": "true", "data-confirmable": "true" }[name]), className: "exhibit-review-card compact" },
    { getAttribute: (name) => ({ "data-candidate-id": "b", "data-included": "false", "data-confirmed": "false", "data-confirmable": "false" }[name]), className: "exhibit-review-card excluded" },
    { getAttribute: (name) => ({ "data-candidate-id": "c", "data-included": "true", "data-confirmed": "false", "data-confirmable": "false" }[name]), className: "exhibit-review-card" },
    { getAttribute: (name) => ({ "data-candidate-id": "d", "data-included": "true", "data-confirmed": "false", "data-confirmable": "true" }[name]), className: "exhibit-review-card" },
  ];
  assert.equal(nextPendingConfirmCardId(cards, "a"), "d");
  assert.equal(nextPendingConfirmCardId(cards, "d"), null);
  assert.equal(nextPendingConfirmCardId(cards, "missing"), null);
  assert.equal(confirmDocumentButtonSelector("d"), '[data-candidate-id="d"] [data-confirm-document][data-confirm-action="confirm"]:not(:disabled)');
  assert.equal(reviewCardListSelector(), ".exhibit-card-list");
});

test("email attachment focus selector is on the parent card panel", () => {
  assert.equal(emailAttachmentsSelector("email-1"), '[data-candidate-id="email-1"] [data-email-attachments]');
});

test("blocker focus retries until the expanded panel exists", () => {
  let remainingMisses = 3;
  const found = [];
  const scheduled = [];
  probeSelectorUntilFound(
    () => {
      if (remainingMisses > 0) {
        remainingMisses -= 1;
        return null;
      }
      return { id: "panel" };
    },
    "[data-email-attachments]",
    (callback) => scheduled.push(callback),
    (element) => found.push(element),
    8,
  );
  assert.equal(found.length, 0);
  while (scheduled.length) scheduled.shift()();
  assert.deepEqual(found, [{ id: "panel" }]);
  assert.equal(remainingMisses, 0);
});

test("overlapping selector probes abort the earlier probe", () => {
  const found = [];
  const scheduled = [];
  const first = probeSelectorUntilFound(
    () => ({ id: "late" }),
    "[data-email-attachments]",
    (callback) => scheduled.push(callback),
    (element) => found.push(element),
    4,
  );
  first.cancel();
  probeSelectorUntilFound(
    () => ({ id: "current" }),
    "[data-email-attachments]",
    (callback) => scheduled.push(callback),
    (element) => found.push(element),
    4,
  );
  while (scheduled.length) scheduled.shift()();
  assert.deepEqual(found, [{ id: "current" }]);
});

test("stable probe waits until the anchored card top stops moving", () => {
  const tops = [40, 120, 120];
  const found = [];
  const scheduled = [];
  probeUntilStable(
    () => ({ getBoundingClientRect: () => ({ top: tops.shift() ?? 120 }) }),
    (callback) => scheduled.push(callback),
    (element) => found.push(element.getBoundingClientRect().top),
    8,
  );
  while (scheduled.length) scheduled.shift()();
  assert.deepEqual(found, [120]);
});
