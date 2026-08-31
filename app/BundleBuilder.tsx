"use client";

import { ChangeEvent, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import {
  AnalysisResult,
  analyseBundleStatements,
  applyWitnessDetails,
  analyseEvidenceFiles,
  BundleStatementInput,
  BuildResult,
  EvidenceRecord,
  ExhibitCandidate,
  analyseFiles,
  attachDerivedEmailEvidence,
  buildBundle,
  buildStatementUpdateSuggestions,
  formatStatementUpdateSuggestionText,
  partitionStatementUpdateSuggestions,
  checkSamplePackAvailability,
  finalizeBuildAudit,
  downloadBytes,
  downloadJson,
  setSavePathProtectionReader,
  loadSamplePack,
  verifiedStatementSnapshots,
} from "./lib/bundle-engine.ts";
import { BUNDLE_PROFILES, DEFAULT_BUNDLE_LAYOUT, DEFAULT_PAGINATION, BundleLayoutSettings, NonA4PageHandling, PageNumberSettings, TemplateFile, TemplateSlot, BuildResolution, PreflightCheck, BuildResolutionAction, TemplateDiscrepancyConfirmation, StoredTemplateReview, updatePaginationDraft, commitPaginationChange, paginationDraftAfterWarningAccept, paginationDraftAfterWarningCancel, coverPrintsPageNumber, coverPrintsVolumeLabel, coverWritesMatterText, countsOptionalPagesInReferences, lockPagination } from "./lib/bundle-types.ts";
import { createProjectArchive, openProjectArchive } from "./lib/project-archive.ts";
import { runPreflight } from "./lib/preflight.ts";
import { AUTOMATIC_MATCH_REVIEW_THRESHOLD, bulkConfirmableCandidates as selectBulkConfirmableCandidates, deriveExhibitGroups, emailChildInsertAfter, exhibitGroupLookup, formatRepeatExhibitNote, isAutomaticLowConfidenceMatch, isReviewerSelectedSource, orderExhibitGroups, pendingReviewCandidateIds, reconcileExhibitArrangement, repeatExhibitCount, REVIEWER_SELECTED_RATIONALE, reviewCandidatesForDisplay, reviewItemNumbers } from "./lib/exhibit-groups.ts";
import { convertWordTemplate } from "./lib/template-converter.ts";
import { compareTemplateMatterReviews, matterDraftFromReview, matterValuesEqual, matterValuesFromConfirmation, parseMatterDraft, reviewTemplateMatterPdf, type MatterDraft } from "./lib/template-matter-review.ts";
import { restoreProjectTemplates } from "./lib/template-persistence.ts";
import { buildBlockers, type BuildBlocker } from "./lib/build-readiness.ts";
import { applyBuildResolutions, isOcrCheck, templateFallbackSlots } from "./lib/build-resolutions.ts";
import { compareSubstantiveBuilds, createSubstantiveBuildSnapshot, type RebuildDifference, type SubstantiveBuildSnapshot } from "./lib/rebuild-comparison.ts";
import { mergeRecoveryProjectDeltas, type RecoveryProjectPayload, type RecoverySourceDescriptor } from "./lib/recovery-restore.ts";
import { restoreCitedCandidateDecision } from "./lib/candidate-restore.ts";
import { captureRetainedBundle, dropStaleRetainedBuild, restoredBundleFromRetain, retainedBuildInputsFrom, retainBuildReadiness, type RetainedBundle } from "./lib/retained-build.ts";
import { applyCandidateReviewChange } from "./lib/candidate-review-change.ts";
import { captureViewportAnchor, confirmDocumentButtonSelector, confirmFocusSelector, emailAttachmentsSelector, emailChildDescriptionSelector, firstActionableControl, firstVisibleReviewCardId, nextPendingConfirmCardId, probeSelectorUntilFound, probeUntilStable, restoreViewportAnchor, restoreWindowScrollY, reviewCardListSelector, reviewCardSelector } from "./lib/review-viewport.ts";
import { createBuildReportPayload, formatBuildReportText } from "./lib/build-report.ts";
import { emailChildrenForDisposition, unresolvedEmailAttachments, type EmailAttachmentChild, type EmailChildDisposition } from "./lib/email-attachments.ts";
import type { WorkbookAnalysis } from "./lib/xlsx.ts";
import { workbookPlanCheckCopy } from "./lib/workbook-print-copy.ts";
import { OriginalPdfReview } from "./OriginalPdfReview.tsx";
import { GuidedSampleTour } from "./GuidedSampleTour.tsx";
import { analysisWithGuidedMapping, applyGuidedSampleMapping, hasGuidedSampleEvidence } from "./lib/guided-sample.ts";
import { resolveTourStep, tourWorkspaceView } from "./lib/guided-tour.ts";
import {
  addArrangementSection,
  bundleArrangementFromLegacyOrder,
  cloneBundleArrangement,
  deleteArrangementSectionKeepItems,
  exhibitContainerLocation,
  flattenBundleArrangement,
  moveArrangementExhibit,
  moveArrangementExhibitInContainer,
  moveArrangementSection,
  moveArrangementSectionBefore,
  renameArrangementSection,
  sortBundleArrangementWithinSections,
  type ArrangementExhibitNode,
  type ArrangementSectionNode,
  type BundleArrangement,
} from "./lib/bundle-arrangement.ts";

type WorkspaceView = "sources" | "sheets" | "review" | "reconcile" | "build";
type ManualAddOrigin = "review" | "reconcile" | "finalise";
type BuildProgress = { stage: string; detail?: string; startedAt: number };
function confidenceLabel(score: number) {
  if (score >= 90) return "High";
  if (score >= AUTOMATIC_MATCH_REVIEW_THRESHOLD) return "Medium";
  return "Review";
}

function matchStrengthLabel(score: number, rationale?: string) {
  if (rationale === REVIEWER_SELECTED_RATIONALE) return "Selected by you";
  if (score >= 90) return "Strong suggested match";
  if (score >= AUTOMATIC_MATCH_REVIEW_THRESHOLD) return "Possible suggested match";
  return "Low-confidence suggested match — check the statement wording against the selected document.";
}

function isEmailAttachmentBlocker(blocker: Pick<BuildBlocker, "code" | "id" | "label">) {
  return Boolean(blocker.code?.startsWith("email.attachment_"));
}

function emailChildDispositionResult(disposition: EmailChildDisposition | undefined) {
  if (disposition === "print-with-email") return "Prints after this email.";
  if (disposition === "add-as-exhibit") return "Separate index item, kept with this email.";
  if (disposition === "leave-out") return "Left out.";
  return "Choose how this attachment is treated.";
}

function shortHash(value: string) {
  return `${value.slice(0, 8)}...${value.slice(-6)}`;
}

function numberingDiffersFromPdfOrder(settings: PageNumberSettings) {
  return Boolean(settings.prefix || settings.suffix || settings.padding > 0 || settings.startAt !== 1);
}

function numberingDifferenceExample(settings: PageNumberSettings) {
  const first = Number.isFinite(settings.startAt) ? Math.max(1, Math.floor(settings.startAt)) : 1;
  const numeric = settings.padding ? String(first).padStart(settings.padding, "0") : String(first);
  return `PDF page 1 is printed as ${settings.prefix}${numeric}${settings.suffix}. The index and suggested statement references use that same printed label.`;
}

const recoveryFileHashes = new WeakMap<File, Promise<string>>();
function fileSha256(file: File) {
  const existing = recoveryFileHashes.get(file);
  if (existing) return existing;
  const pending = file.arrayBuffer().then((buffer) => crypto.subtle.digest("SHA-256", buffer)).then((digest) => Array.from(new Uint8Array(digest)).map((part) => part.toString(16).padStart(2, "0")).join(""));
  recoveryFileHashes.set(file, pending);
  return pending;
}

async function copyPlainText(value: string): Promise<{ copied: boolean; detail?: string }> {
  try {
    if (window.bundleBuilderDesktop?.copyText) {
      const result = await window.bundleBuilderDesktop.copyText(value);
      if (!result?.copied) return { copied: false, detail: "The desktop clipboard did not accept the suggestions. Download the .txt file instead." };
      return { copied: true };
    }
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(value);
      return { copied: true };
    }
    const area = document.createElement("textarea");
    area.value = value;
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.append(area);
    area.select();
    const ok = document.execCommand("copy");
    area.remove();
    if (!ok) return { copied: false, detail: "Copy is unavailable in this workspace. Download the .txt file instead." };
    return { copied: true };
  } catch {
    return { copied: false, detail: "Copy failed. Download the .txt file instead." };
  }
}

export default function BundleBuilder() {
  const [analysis, setAnalysis] = useState<AnalysisResult | null>(null);
  const [candidates, setCandidates] = useState<ExhibitCandidate[]>([]);
  const [build, setBuild] = useState<BuildResult | null>(null);
  const [view, setView] = useState<WorkspaceView>("review");
  const [busy, setBusy] = useState<"sample" | "analyse" | "manual" | "build" | "template" | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [statementFile, setStatementFile] = useState<File | null>(null);
  const [evidenceFiles, setEvidenceFiles] = useState<File[]>([]);
  const [statements, setStatements] = useState<BundleStatementInput[]>([]);
  const [pagination, setPagination] = useState<PageNumberSettings>(DEFAULT_PAGINATION);
  const [paginationDraft, setPaginationDraft] = useState<PageNumberSettings>(DEFAULT_PAGINATION);
  const [paginationConfirmation, setPaginationConfirmation] = useState<PageNumberSettings | null>(null);
  const [paginationPendingChange, setPaginationPendingChange] = useState<Partial<PageNumberSettings> | null>(null);
  const [volumeNumberingConfirmation, setVolumeNumberingConfirmation] = useState(false);
  const [layout, setLayout] = useState<BundleLayoutSettings>(DEFAULT_BUNDLE_LAYOUT);
  const [pageSizeChoices, setPageSizeChoices] = useState<Record<string, NonA4PageHandling>>({});
  const [templates, setTemplates] = useState<TemplateFile[]>([]);
  const [resolutions, setResolutions] = useState<BuildResolution[]>([]);
  const [projectName, setProjectName] = useState("Untitled exhibit project");
  const [showPendingOnly, setShowPendingOnly] = useState(false);
  const [showDuplicatesOnly, setShowDuplicatesOnly] = useState(false);
  const [showSourcePreview, setShowSourcePreview] = useState<string | null>(null);
  const [openDocumentPickerId, setOpenDocumentPickerId] = useState<string | null>(null);
  const [expandedConfirmedCards, setExpandedConfirmedCards] = useState<Set<string>>(new Set());
  const [openEmailAttachmentsId, setOpenEmailAttachmentsId] = useState<string | null>(null);
  const [templatePreview, setTemplatePreview] = useState<{ slot: TemplateSlot; name: string; file: File } | null>(null);
  const [templatePreviewLoaded, setTemplatePreviewLoaded] = useState(false);
  const [matterDraft, setMatterDraft] = useState<MatterDraft | null>(null);
  const [ocrSourcePreview, setOcrSourcePreview] = useState<{ sourceId: string; sourceSha256: string; name: string; file: File; extractedText: string } | null>(null);
  const [visuallyReviewedSourceHashes, setVisuallyReviewedSourceHashes] = useState<Set<string>>(new Set());
  const [templateDiscrepancyConfirmation, setTemplateDiscrepancyConfirmation] = useState<TemplateDiscrepancyConfirmation | null>(null);
  const [draggingExhibitId, setDraggingExhibitId] = useState<string | null>(null);
  const [draggingSectionId, setDraggingSectionId] = useState<string | null>(null);
  const [headingDropTargetKey, setHeadingDropTargetKey] = useState<string | null>(null);
  const [collapsedSectionIds, setCollapsedSectionIds] = useState<Set<string>>(new Set());
  const [orderChangeConfirmation, setOrderChangeConfirmation] = useState(false);
  const [buildProgress, setBuildProgress] = useState<BuildProgress | null>(null);
  const [analysisProgress, setAnalysisProgress] = useState<BuildProgress | null>(null);
  const [analysisProgressAnnouncement, setAnalysisProgressAnnouncement] = useState("");
  const [arrangement, setArrangement] = useState<BundleArrangement>(() => bundleArrangementFromLegacyOrder([]));
  const [orderHistory, setOrderHistory] = useState<BundleArrangement[]>([]);
  const [orderSort, setOrderSort] = useState<"statement" | "date" | "filename" | "description">("statement");
  const [orderPreview, setOrderPreview] = useState<{ arrangement: BundleArrangement; label: string } | null>(null);
  const [newSectionHeading, setNewSectionHeading] = useState("");
  const [arrangementStatus, setArrangementStatus] = useState("");
  const [reviewActionStatus, setReviewActionStatus] = useState("");
  const [copyStatus, setCopyStatus] = useState<{ kind: "success" | "failure"; message: string } | null>(null);
  const [manualAddOpen, setManualAddOpen] = useState(false);
  const [manualAddOrigin, setManualAddOrigin] = useState<ManualAddOrigin>("finalise");
  const [manualEvidenceId, setManualEvidenceId] = useState("");
  const [manualDescription, setManualDescription] = useState("");
  const [manualDate, setManualDate] = useState("Date not stated");
  const [manualUploadedEvidence, setManualUploadedEvidence] = useState<EvidenceRecord | null>(null);
  const [lastBuildSnapshot, setLastBuildSnapshot] = useState<SubstantiveBuildSnapshot | null>(null);
  const [rebuildComparison, setRebuildComparison] = useState<RebuildDifference | null>(null);
  const [statementDrafts, setStatementDrafts] = useState<Record<string, { witnessName: string; witnessInitials: string }>>({});
  const [initialsConfirmation, setInitialsConfirmation] = useState<{ statementId: string; existing: string[]; proposed: string } | null>(null);
  const [reorderReturn, setReorderReturn] = useState<RetainedBundle<BuildResult, BundleArrangement, ExhibitCandidate> | null>(null);
  const [recoveryOffer, setRecoveryOffer] = useState<{ recoveryId: string; revision: number; projectName: string } | null>(null);
  const [recoveryIssues, setRecoveryIssues] = useState<string[]>([]);
  const [guidedSampleAvailability, setGuidedSampleAvailability] = useState<"checking" | "available" | "unavailable">("checking");
  const [guidedSampleHidden, setGuidedSampleHiddenState] = useState(false);
  const [guidedSamplePreferenceReady, setGuidedSamplePreferenceReady] = useState(false);
  const [tourActive, setTourActive] = useState(false);
  const [tourOpenedFolder, setTourOpenedFolder] = useState(false);
  const [tourDownloaded, setTourDownloaded] = useState(false);
  const [tourSaved, setTourSaved] = useState(false);
  const useSamplePackRef = useRef<() => Promise<void>>(async () => {});
  const [recoveryDataDialogOpen, setRecoveryDataDialogOpen] = useState(false);
  const [recoveryDataStored, setRecoveryDataStored] = useState(false);
  const [recoveryDeleteAcknowledged, setRecoveryDeleteAcknowledged] = useState(false);
  const [bulkConfirmationOpen, setBulkConfirmationOpen] = useState(false);
  const [bulkConfirmationAcknowledged, setBulkConfirmationAcknowledged] = useState(false);
  const recoveryId = useRef<string | null>(null);
  const recoveryRevision = useRef(0);
  const recoveryRestoring = useRef(false);
  const recoverySourceDescriptors = useRef<RecoverySourceDescriptor[]>([]);
  const confirmationDialogOrigin = useRef<HTMLElement | null>(null);
  const statementInput = useRef<HTMLInputElement>(null);
  const evidenceInput = useRef<HTMLInputElement>(null);
  const projectInput = useRef<HTMLInputElement>(null);
  const manualEvidenceInput = useRef<HTMLInputElement>(null);
  const manualPanelHeading = useRef<HTMLHeadingElement>(null);
  const manualAddTrigger = useRef<HTMLElement | null>(null);
  const templatePreviewDialog = useRef<HTMLDialogElement>(null);
  const templatePreviewCloseButton = useRef<HTMLButtonElement>(null);
  const templatePreviewOrigin = useRef<HTMLElement | null>(null);
  const ocrPreviewDialog = useRef<HTMLDialogElement>(null);
  const ocrPreviewCloseButton = useRef<HTMLButtonElement>(null);
  const ocrPreviewOrigin = useRef<HTMLElement | null>(null);
  const announcedAnalysisStage = useRef("");
  const reviewHeading = useRef<HTMLHeadingElement>(null);
  const viewportProbe = useRef<{ cancel(): void } | null>(null);
  const confirmViewport = useRef<{ id: number; scrollY: number; confirmedId: string; nextPendingId: string | null } | null>(null);
  const confirmViewportSeq = useRef(0);
  const focusReviewAfterAnalysis = useRef(false);
  const buildCancelRequested = useRef(false);
  const analysisCancelRequested = useRef(false);

  const profile = BUNDLE_PROFILES[0];
  const workspaceAvailable = Boolean(analysis);
  const confirmationDialogKey = recoveryOffer ? "recovery"
    : orderChangeConfirmation ? "order"
      : paginationConfirmation ? "pagination"
        : volumeNumberingConfirmation ? "volume-numbering"
          : initialsConfirmation ? "initials"
            : bulkConfirmationOpen ? "bulk"
              : recoveryDataDialogOpen ? "recovery-data"
                : "";
  const finalOrder = useMemo(() => flattenBundleArrangement(arrangement), [arrangement]);
  const rawPreflight = useMemo(
    () => (analysis ? runPreflight(analysis, candidates, profile) : []),
    [analysis, candidates, profile],
  );
  const preflight = useMemo(
    () => applyBuildResolutions(rawPreflight, resolutions),
    [rawPreflight, resolutions],
  );
  const rawExhibitGroups = useMemo(
    () => (analysis ? deriveExhibitGroups(analysis, candidates) : []),
    [analysis, candidates],
  );
  const exhibitGroups = useMemo(
    () => orderExhibitGroups(rawExhibitGroups, finalOrder),
    [rawExhibitGroups, finalOrder],
  );
  const displayedArrangement = orderPreview?.arrangement ?? arrangement;
  const hasIndexHeadings = arrangement.nodes.some((node) => node.type === "section");
  const exhibitGroupLookupByKind = useMemo(() => exhibitGroupLookup(exhibitGroups), [exhibitGroups]);
  const evidenceById = useMemo(() => new Map((analysis?.evidence ?? []).map((record) => [record.id, record])), [analysis]);
  const evidenceHashCounts = useMemo(() => {
    const counts = new Map<string, number>();
    for (const record of analysis?.evidence ?? []) counts.set(record.sha256, (counts.get(record.sha256) ?? 0) + 1);
    return counts;
  }, [analysis]);
  const displayedOrderNumbers = useMemo(() => new Map(flattenBundleArrangement(displayedArrangement).map((id, index) => [id, index + 1])), [displayedArrangement]);

  const includedCount = useMemo(
    () => candidates.filter((candidate) => candidate.included && !candidate.parentEmailProvenance).length,
    [candidates],
  );
  const confirmedCount = useMemo(
    () =>
      candidates.filter(
        (candidate) =>
          candidate.included && !candidate.parentEmailProvenance && candidate.evidenceId && candidate.confirmed,
      ).length,
    [candidates],
  );
  const citedCandidates = useMemo(
    () => candidates.filter((candidate) => !candidate.manualAddition),
    [candidates],
  );
  const matchedCitationCount = useMemo(
    () => citedCandidates.filter((candidate) => candidate.included && candidate.evidenceId).length,
    [citedCandidates],
  );
  const confirmedCitationCount = useMemo(
    () => citedCandidates.filter((candidate) => candidate.included && candidate.evidenceId && candidate.confirmed).length,
    [citedCandidates],
  );
  const repeatExhibits = useMemo(() => repeatExhibitCount(exhibitGroups), [exhibitGroups]);
  const repeatExhibitNote = formatRepeatExhibitNote(repeatExhibits);
  const addedExhibitCount = useMemo(
    () => candidates.filter((candidate) => candidate.manualAddition && candidate.included).length,
    [candidates],
  );
  const pendingCandidateIds = useMemo(
    () => pendingReviewCandidateIds(candidates, exhibitGroups, analysis?.evidence ?? []),
    [analysis, candidates, exhibitGroups],
  );
  const needsDecisionCount = pendingCandidateIds.size;
  const reviewCandidates = useMemo(() => {
    return reviewCandidatesForDisplay(candidates, exhibitGroups);
  }, [candidates, exhibitGroups]);
  const reviewItemNumberByCandidate = useMemo(
    () => reviewItemNumbers(reviewCandidates),
    [reviewCandidates],
  );
  const bulkConfirmableCandidates = useMemo(
    () => selectBulkConfirmableCandidates(candidates, exhibitGroups),
    [candidates, exhibitGroups],
  );
  const bulkConfirmableCount = bulkConfirmableCandidates.length;
  const includedReferenceMarks = useMemo(() => new Set(candidates
    .filter((candidate) => candidate.included && !candidate.manualAddition)
    .map((candidate) => `${(candidate.exhibitInitials ?? candidate.witnessInitials ?? "EX").replace(/\s+/g, "").toUpperCase()}${candidate.exhibitSequence ?? 1}`)), [candidates]);
  const referenceMarkConflict = includedReferenceMarks.size > 1;
  const fallbackReferenceSequence = candidates.find((candidate) => !candidate.manualAddition)?.exhibitSequence ?? 1;
  const fallbackReferenceBundleMark = `${analysis?.witnessInitials ?? statements[0]?.witnessInitials ?? "EX"}${fallbackReferenceSequence}`.replace(/\s+/g, "");
  const referenceBundleMark = includedReferenceMarks.size === 1 ? [...includedReferenceMarks][0] : fallbackReferenceBundleMark;
  const selectedEvidenceIds = useMemo(
    () =>
      new Set(
        candidates.filter((candidate) => candidate.included && candidate.evidenceId).map((candidate) => candidate.evidenceId!),
      ),
    [candidates],
  );
  const selectedEvidenceHashes = useMemo(
    () => new Set((analysis?.evidence ?? []).filter((record) => selectedEvidenceIds.has(record.id)).map((record) => record.sha256)),
    [analysis, selectedEvidenceIds],
  );
  const possibleDuplicateEvidenceIds = useMemo(() => {
    const ids = new Set<string>();
    for (const check of preflight) {
      if (check.code !== "source.near_duplicate") continue;
      // Do not infer filtering state from a human-readable filename or
      // warning message; the preflight relationship carries explicit IDs.
      for (const id of check.relatedSourceIds ?? []) ids.add(id);
    }
    return ids;
  }, [preflight]);
  const unreferencedEvidence = useMemo(
    () =>
      analysis?.evidence.filter(
        (record) => !record.derivedFromEmail && !selectedEvidenceIds.has(record.id),
      ) ?? [],
    [analysis, selectedEvidenceIds],
  );
  const nonA4Exhibits = useMemo(() => {
    const seen = new Set<string>();
    return exhibitGroups.flatMap((group) => {
      const items: Array<{ choiceKey: string; name: string; nonA4: NonNullable<EvidenceRecord["pageSizes"]>; marginCount: number; annotatedCount: number }> = [];
      const record = group.evidence;
      if (!group.canonical.parentEmailProvenance && !seen.has(record.id) && record.pageSizes?.some((page) => !page.isA4)) {
        seen.add(record.id);
        const nonA4 = record.pageSizes.filter((page) => !page.isA4);
        items.push({ choiceKey: record.id, name: record.name, nonA4, marginCount: nonA4.filter((page) => page.wouldAddMarginsOnA4).length, annotatedCount: nonA4.filter((page) => page.hasAnnotations).length });
      }
      for (const child of emailChildrenForDisposition(record.emailAttachments, group.canonical.emailAttachmentDispositions, "print-with-email")) {
        if (child.extension !== "pdf" || seen.has(child.identity) || !child.pageSizes?.some((page) => !page.isA4)) continue;
        seen.add(child.identity);
        const nonA4 = child.pageSizes.map((page, index) => ({
          ...page,
          page: index + 1,
          orientation: page.width > page.height ? "landscape" as const : "portrait" as const,
        })).filter((page) => !page.isA4);
        items.push({ choiceKey: child.identity, name: `${child.name} (printed with ${record.name})`, nonA4, marginCount: nonA4.filter((page) => page.wouldAddMarginsOnA4).length, annotatedCount: nonA4.filter((page) => page.hasAnnotations).length });
      }
      return items;
    });
  }, [exhibitGroups]);
  useEffect(() => {
    if (!confirmationDialogKey) {
      confirmationDialogOrigin.current?.focus({ preventScroll: true });
      confirmationDialogOrigin.current = null;
      return;
    }
    confirmationDialogOrigin.current ??= document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const dialog = document.querySelector<HTMLElement>(".confirmation-backdrop .confirmation-dialog");
    if (!dialog) return;
    const focusableSelector = 'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';
    const focusable = () => Array.from(dialog.querySelectorAll<HTMLElement>(focusableSelector)).filter((item) => item.offsetParent !== null);
    window.requestAnimationFrame(() => focusable()[0]?.focus());
    const keepFocusInside = (event: KeyboardEvent) => {
      if (event.key !== "Tab") return;
      const controls = focusable();
      if (!controls.length) { event.preventDefault(); dialog.focus(); return; }
      const first = controls[0];
      const last = controls[controls.length - 1];
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
    };
    document.addEventListener("keydown", keepFocusInside);
    return () => document.removeEventListener("keydown", keepFocusInside);
  }, [confirmationDialogKey]);

  useLayoutEffect(() => {
    if (!workspaceAvailable) return;
    window.requestAnimationFrame(() => document.querySelector<HTMLElement>(".workspace-main h1")?.focus({ preventScroll: true }));
  }, [view, workspaceAvailable]);

  useEffect(() => {
    if (manualAddOpen) {
      window.requestAnimationFrame(() => manualPanelHeading.current?.focus({ preventScroll: true }));
      return;
    }
    manualAddTrigger.current?.focus({ preventScroll: true });
    manualAddTrigger.current = null;
  }, [manualAddOpen]);

  useEffect(() => {
    setPageSizeChoices((current) => {
      let changed = false;
      const next = { ...current };
      for (const { choiceKey, annotatedCount } of nonA4Exhibits) {
        if (!annotatedCount || next[choiceKey] === "keep-original") continue;
        next[choiceKey] = "keep-original";
        changed = true;
      }
      return changed ? next : current;
    });
  }, [nonA4Exhibits]);
  useEffect(() => {
    if (!openDocumentPickerId) return;
    const close = (event: Event) => {
      if (event instanceof KeyboardEvent && event.key !== "Escape") return;
      if (event instanceof PointerEvent) {
        const target = event.target;
        if (target instanceof Element && target.closest(".document-picker")) return;
      } else if (event instanceof KeyboardEvent) {
        restoreDocumentPickerSummary(document.activeElement);
      }
      setOpenDocumentPickerId(null);
    };
    document.addEventListener("keydown", close);
    document.addEventListener("pointerdown", close);
    return () => {
      document.removeEventListener("keydown", close);
      document.removeEventListener("pointerdown", close);
    };
  }, [openDocumentPickerId]);
  const fallbackTemplateSlots = useMemo(() => templateFallbackSlots(resolutions, templates), [resolutions, templates]);
  const usedTemplateSlots = useMemo(() => new Set<TemplateSlot>(["cover", "index", ...(layout.includeDividerPages ? ["divider" as const] : []), ...(layout.includeExhibitCoverPages ? ["exhibitCover" as const] : [])]), [layout.includeDividerPages, layout.includeExhibitCoverPages]);
  const activeCustomTemplates = useMemo(() => templates.filter((template) => usedTemplateSlots.has(template.slot) && !fallbackTemplateSlots.has(template.slot)), [fallbackTemplateSlots, templates, usedTemplateSlots]);
  const previewedTemplate = templatePreview ? templates.find((template) => template.slot === templatePreview.slot) : undefined;
  const templateMatterDiscrepancies = useMemo(() => compareTemplateMatterReviews(activeCustomTemplates.flatMap((template) => template.reviewState?.matterReview ? [{
    templateId: template.slot,
    role: template.slot,
    sourceName: template.file.name,
    review: template.reviewState.matterReview,
    confirmedValues: template.reviewState.matterConfirmation?.pdfSha256 === template.pdfSha256
      ? matterValuesFromConfirmation(template.reviewState.matterConfirmation)
      : undefined,
  }] : [])), [activeCustomTemplates]);
  const templateDiscrepancyFingerprint = useMemo(() => JSON.stringify(templateMatterDiscrepancies.map((item) => ({ field: item.field, evidence: item.evidence.map((evidence) => ({ templateId: evidence.templateId, pdfSha256: evidence.pdfSha256, normalizedValues: evidence.normalizedValues })) }))), [templateMatterDiscrepancies]);
  const templateReviewFailures = useMemo(() => activeCustomTemplates.flatMap((template) => {
    const pdfSha256 = template.pdfSha256;
    const review = template.reviewState;
    const failures: string[] = [];
    if (!pdfSha256 || review?.matterReview?.pdfSha256 !== pdfSha256) failures.push("read the exact PDF");
    if (review?.matterConfirmation?.pdfSha256 !== pdfSha256) failures.push("confirm matter details and party names");
    if ((template.sourceFormat ?? "pdf") !== "pdf" && review?.appearanceConfirmation?.pdfSha256 !== pdfSha256) failures.push("confirm the converted appearance");
    if (review?.matterReview?.placeholders.length && review?.placeholderConfirmation?.pdfSha256 !== pdfSha256) failures.push("confirm visible placeholders");
    return failures.length ? [`${template.file.name}: ${failures.join(", ")}`] : [];
  }), [activeCustomTemplates]);
  const templateDiscrepancyPending = templateMatterDiscrepancies.length > 0 && templateDiscrepancyConfirmation?.fingerprint !== templateDiscrepancyFingerprint;
  const templateReviewPending = templateReviewFailures.length > 0 || templateDiscrepancyPending;
  const buildBlockerList = useMemo(
    () => {
      const blockers = buildBlockers({
      includedCount,
      confirmedCount,
      pendingApprovalCount: pendingCandidateIds.size,
      templateReviewPending,
      unapprovedTemplateNames: [...templateReviewFailures, ...(templateDiscrepancyPending ? ["selected templates: acknowledge conflicting possible matter details"] : [])],
        preflight,
      });
      if (referenceMarkConflict) blockers.unshift({
        id: "reference-mark-conflict",
        label: "Statement reference marks need review",
        detail: `Included references use ${[...includedReferenceMarks].join(" and ")}. Choose one statement-reference mark in Optional settings before building.`,
        kind: "state",
        target: "templates",
        actionLabel: "Review statement-reference mark",
      });
      if (orderPreview) blockers.unshift({
        id: "order-preview-pending",
        label: "Automatic order preview is still open",
        detail: "Use this order or cancel the preview before building. The displayed proposal has not changed the saved bundle arrangement.",
        kind: "state",
        target: "finalise",
        actionLabel: "Review proposed order",
      });
      return blockers;
    },
    [confirmedCount, includedCount, includedReferenceMarks, orderPreview, pendingCandidateIds, preflight, referenceMarkConflict, templateDiscrepancyPending, templateReviewFailures, templateReviewPending],
  );
  const technicalBuildBlockers = buildBlockerList.filter((blocker) => blocker.kind !== "approval");
  const readyToBuild = buildBlockerList.length === 0;
  const retainReadyToBuild = retainBuildReadiness(buildBlockerList);
  const includedWorkbookCount = useMemo(() => {
    if (!analysis) return 0;
    const top = analysis.evidence.filter((record) => record.extension === "xlsx" && candidates.some((candidate) => candidate.included && candidate.confirmed && candidate.evidenceId === record.id)).length;
    const printedChildren = analysis.evidence.reduce((count, record) => {
      const owners = candidates.filter((candidate) => candidate.included && candidate.confirmed && candidate.evidenceId === record.id);
      if (!owners.length) return count;
      return count + owners.reduce((inner, owner) => inner + emailChildrenForDisposition(record.emailAttachments, owner.emailAttachmentDispositions, "print-with-email").filter((child) => child.extension === "xlsx").length, 0);
    }, 0);
    return top + printedChildren;
  }, [analysis, candidates]);
  const includedWorkbookInMatter = useMemo(() => {
    if (!analysis) return false;
    if (analysis.evidence.some((record) => record.extension === "xlsx" && candidates.some((candidate) => candidate.included && candidate.evidenceId === record.id))) return true;
    return analysis.evidence.some((record) => {
      const owners = candidates.filter((candidate) => candidate.included && candidate.evidenceId === record.id);
      return owners.some((owner) => emailChildrenForDisposition(record.emailAttachments, owner.emailAttachmentDispositions, "print-with-email").some((child) => child.extension === "xlsx"));
    });
  }, [analysis, candidates]);
  const tourAttachmentPending = Boolean(analysis && candidates.some((candidate) => {
    if (!candidate.included) return false;
    const record = analysis.evidence.find((item) => item.id === candidate.evidenceId);
    return unresolvedEmailAttachments(record?.emailAttachments, candidate.emailAttachmentDispositions).length > 0;
  }));
  const tourStep = resolveTourStep({
    active: tourActive,
    openedFolder: tourOpenedFolder,
    statementSelected: Boolean(statementFile),
    evidenceSelected: hasGuidedSampleEvidence(evidenceFiles.map((file) => file.name)),
    analysisReady: Boolean(analysis),
    bulkConfirmableCount,
    attachmentPending: tourAttachmentPending,
    attachmentChoicesOpen: Boolean(openEmailAttachmentsId),
    repeatPending: candidates.some((candidate) => candidate.included && candidate.repeatDecision === "same" && !candidate.confirmed),
    printWithEmailVisible: Boolean(analysis && candidates.some((candidate) => {
      if (!candidate.included) return false;
      const record = analysis.evidence.find((item) => item.id === candidate.evidenceId);
      return Boolean(record?.emailAttachments?.length);
    })),
    view: tourWorkspaceView(Boolean(analysis), view),
    includedWorkbook: includedWorkbookInMatter,
    hasBuild: Boolean(build),
    downloaded: tourDownloaded,
    saved: tourSaved,
  });
  const workbookSourceCount = useMemo(() => {
    if (!analysis) return 0;
    return analysis.evidence.filter((record) => record.extension === "xlsx").length
      + analysis.evidence.reduce((count, record) => count + (record.emailAttachments ?? []).filter((child) => child.extension === "xlsx").length, 0);
  }, [analysis]);
  const confirmedWorkbookExhibits = useMemo(() => {
    if (!analysis) return [] as Array<{ key: string; name: string; evidenceId: string; childIdentity?: string; workbook?: WorkbookAnalysis; sheetSelections: Array<{ name: string; included: boolean; range: string }> }>;
    const top = analysis.evidence
      .filter((record) => record.extension === "xlsx" && candidates.some((candidate) => candidate.included && candidate.confirmed && candidate.evidenceId === record.id))
      .map((record) => ({ key: record.id, name: record.name, evidenceId: record.id, childIdentity: undefined, workbook: record.workbook, sheetSelections: record.sheetSelections ?? [] }));
    const children = analysis.evidence.flatMap((record) => {
      const owners = candidates.filter((candidate) => candidate.included && candidate.confirmed && candidate.evidenceId === record.id);
      if (!owners.length) return [];
      return emailChildrenForDisposition(record.emailAttachments, owners[0]?.emailAttachmentDispositions, "print-with-email").flatMap((child) => {
        if (child.extension !== "xlsx") return [];
        return [{ key: child.identity, name: `${child.name} (printed with ${record.name})`, evidenceId: record.id, childIdentity: child.identity, workbook: child.workbook, sheetSelections: child.sheetSelections ?? [] }];
      });
    });
    return [...top, ...children];
  }, [analysis, candidates]);
  const reviewBlockingCodes = new Set(["workbook.no_sheet", "workbook.sheet_unreadable"]);
  const readyToLeaveReview =
    includedCount > 0 &&
    confirmedCount === includedCount &&
    pendingCandidateIds.size === 0 &&
    !referenceMarkConflict &&
    !templateReviewPending &&
    !preflight.some((check) => check.severity === "blocking" && !reviewBlockingCodes.has(check.code ?? ""));
  const suppliedCoverOmitsPageNumber = layout.coverInsertion === "exact" && !coverPrintsPageNumber(layout);
  const suppliedCoverOmitsVolumeLabel = layout.coverInsertion === "exact" && !coverPrintsVolumeLabel(layout);
  const otherRequirementCount = technicalBuildBlockers.length;
  const firstOtherRequirement = technicalBuildBlockers[0];
  const nextReviewBlocker = pendingCandidateIds.size
    ? buildBlockerList.find((blocker) => blocker.kind === "approval") ?? buildBlockerList[0]
    : technicalBuildBlockers[0] ?? buildBlockerList[0];
  const reviewContinueReason = readyToLeaveReview
    ? null
    : otherRequirementCount
      ? `Cannot continue — ${otherRequirementCount} other requirement${otherRequirementCount === 1 ? "" : "s"} remaining (${firstOtherRequirement?.label ?? "see Other requirements"}${firstOtherRequirement?.fileName ? ` for ${firstOtherRequirement.fileName}` : ""})`
      : pendingCandidateIds.size
        ? null
        : "Cannot continue until remaining review requirements are complete";
  const statementSuggestions = useMemo(
    () => build ? buildStatementUpdateSuggestions(build.volumes ? build.volumes.flatMap((volume) => volume.records) : build.records) : [],
    [build],
  );
  const statementSuggestionSections = useMemo(
    () => partitionStatementUpdateSuggestions(statementSuggestions),
    [statementSuggestions],
  );
  const statementUpdateText = useMemo(
    () => formatStatementUpdateSuggestionText(statementSuggestions),
    [statementSuggestions],
  );
  useEffect(() => {
    setCopyStatus(null);
  }, [statementUpdateText]);
  const preflightWarningCount = preflight.filter((check) => check.severity === "warning").length;
  const retainedWarningCount = build
    ? build.checks.filter((check) => check.status !== "pass").length
    : preflightWarningCount;
  const paginationDraftChanged = useMemo(
    () => JSON.stringify(paginationDraft) !== JSON.stringify(pagination),
    [paginationDraft, pagination],
  );
  const retainedBuildInputs = useMemo(() => {
    if (!analysis) return null;
    return retainedBuildInputsFrom({
      analysis,
      candidates,
      templates,
      layout,
      pagination,
      pageSizeChoices,
      resolutions,
      templateDiscrepancyConfirmation,
    });
  }, [analysis, candidates, templates, layout, pagination, pageSizeChoices, resolutions, templateDiscrepancyConfirmation]);


  useEffect(() => {
    if (!analysis || view !== "review" || !focusReviewAfterAnalysis.current) return;
    focusReviewAfterAnalysis.current = false;
    requestAnimationFrame(() => reviewHeading.current?.focus({ preventScroll: true }));
  }, [analysis, view]);

  useLayoutEffect(() => {
    const txn = confirmViewport.current;
    if (!txn) return;
    if (view !== "review") {
      confirmViewport.current = null;
      return;
    }
    confirmViewport.current = null;
    const nextConfirm = txn.nextPendingId
      ? document.querySelector<HTMLElement>(confirmDocumentButtonSelector(txn.nextPendingId))
      : null;
    const currentReview = document.querySelector<HTMLElement>(`${reviewCardSelector(txn.confirmedId)} [data-confirm-focus]`);
    const showAll = document.querySelector<HTMLElement>(".pending-empty button");
    const list = document.querySelector<HTMLElement>(reviewCardListSelector());
    const focusTarget = nextConfirm ?? currentReview ?? showAll ?? list;
    focusTarget?.focus({ preventScroll: true });
    restoreWindowScrollY(txn.scrollY, window.scrollY, (top) => window.scrollTo({ top, behavior: "auto" }));
  }, [candidates, expandedConfirmedCards, openEmailAttachmentsId, showPendingOnly, view]);

  useEffect(() => {
    let active = true;
    void (async () => {
      try {
        const stored = window.bundleBuilderDesktop
          ? await window.bundleBuilderDesktop.readPreferences()
          : { hideGuidedSample: window.localStorage.getItem("exhibit-builder-hide-guided-sample") === "true" };
        if (active) setGuidedSampleHiddenState(stored.hideGuidedSample);
      } catch {
        if (active) setGuidedSampleHiddenState(false);
      } finally {
        if (active) setGuidedSamplePreferenceReady(true);
      }
      const available = await checkSamplePackAvailability();
      if (active) setGuidedSampleAvailability(available ? "available" : "unavailable");
    })();
    const handleAnalyseGuidedSample = () => { void useSamplePackRef.current(); };
    window.addEventListener("exhibit-builder:analyse-guided-sample", handleAnalyseGuidedSample);
    return () => {
      active = false;
      window.removeEventListener("exhibit-builder:analyse-guided-sample", handleAnalyseGuidedSample);
    };
  }, []);

  useEffect(() => {
    if (tourActive && !tourStep) skipGuidedSampleTour();
  }, [tourActive, tourStep]);

  useEffect(() => {
    if (!ocrSourcePreview) {
      ocrPreviewOrigin.current?.focus();
      ocrPreviewOrigin.current = null;
      return;
    }
    const dialog = ocrPreviewDialog.current;
    if (!dialog) return;
    if (!dialog.open) dialog.showModal();
    ocrPreviewCloseButton.current?.focus();
    const handleCancel = (event: Event) => { event.preventDefault(); setOcrSourcePreview(null); };
    dialog.addEventListener("cancel", handleCancel);
    return () => dialog.removeEventListener("cancel", handleCancel);
  }, [ocrSourcePreview]);

  useEffect(() => {
    if (!templatePreview) {
      setTemplatePreviewLoaded(false);
      templatePreviewOrigin.current?.focus();
      templatePreviewOrigin.current = null;
      return;
    }
    const dialog = templatePreviewDialog.current;
    if (!dialog) return;
    if (!dialog.open) dialog.showModal();
    templatePreviewCloseButton.current?.focus();
    const handleCancel = (event: Event) => {
      event.preventDefault();
      setTemplatePreview(null);
    };
    dialog.addEventListener("cancel", handleCancel);
    return () => dialog.removeEventListener("cancel", handleCancel);
  }, [templatePreview]);

  useEffect(() => {
    if (!templatePreview || !previewedTemplate?.reviewState?.matterReview) {
      setMatterDraft(null);
      return;
    }
    const review = previewedTemplate.reviewState.matterReview;
    const confirmation = previewedTemplate.reviewState.matterConfirmation?.pdfSha256 === previewedTemplate.pdfSha256
      ? previewedTemplate.reviewState.matterConfirmation
      : undefined;
    setMatterDraft(matterDraftFromReview(review, confirmation));
  }, [previewedTemplate?.pdfSha256, previewedTemplate?.reviewState?.matterReview, templatePreview]);

  useEffect(() => {
    const insertAfter = emailChildInsertAfter(rawExhibitGroups);
    setArrangement((current) => {
      const reconciled = reconcileExhibitArrangement(current, rawExhibitGroups, insertAfter);
      return JSON.stringify(reconciled) === JSON.stringify(current) ? current : reconciled;
    });
  }, [rawExhibitGroups]);

  useEffect(() => {
    const desktop = window.bundleBuilderDesktop;
    if (!desktop) return;
    let active = true;
    void desktop.recoveryStatus().then(async (status) => {
      if (!active) return;
      setRecoveryDataStored(status.stored);
      if (status.corrupt && status.issue) setRecoveryIssues([status.issue]);
      if (status.available && status.recoveryId && status.revision !== undefined) {
        setRecoveryOffer({ recoveryId: status.recoveryId, revision: status.revision, projectName: status.projectName ?? "Recovered exhibit project" });
        return;
      }
      const begun = await desktop.beginRecovery();
      if (!active) return;
      recoveryId.current = begun.recoveryId;
      recoveryRevision.current = begun.revision;
    }).catch(() => {});
    return () => { active = false; };
  }, []);

  useEffect(() => {
    setSavePathProtectionReader(() => {
      const paths = new Set<string>();
      const desktop = window.bundleBuilderDesktop;
      const addFilePath = (file?: File) => {
        if (!file || !desktop) return;
        try {
          const sourcePath = desktop.sourcePath(file);
          if (sourcePath) paths.add(sourcePath);
        } catch {
          // A file without a recoverable native path cannot be protected by path identity.
        }
      };
      for (const statement of statements) addFilePath(statement.file);
      for (const record of analysis?.evidence ?? []) {
        if (!record.derivedFromEmail) addFilePath(record.file);
      }
      for (const template of templates) addFilePath(template.file);
      let allowedOverwritePath: string | undefined;
      for (const source of recoverySourceDescriptors.current) {
        if (!source.path) continue;
        if (source.id === "saved-project-archive" || source.role === "project") {
          allowedOverwritePath = source.path;
          paths.add(source.path);
          continue;
        }
        if (source.role === "statement" || source.role === "evidence" || source.role === "template") paths.add(source.path);
      }
      return { protectedSourcePaths: [...paths], allowedOverwritePath };
    });
    return () => setSavePathProtectionReader(() => ({ protectedSourcePaths: [] }));
  });

  useEffect(() => {
    const desktop = window.bundleBuilderDesktop;
    if (!desktop || !analysis || recoveryOffer || recoveryRestoring.current) return;
    const timer = window.setTimeout(() => {
      void (async () => {
        // Do not advance or replace the last valid journal until every current
        // statement still matches the exact bytes used for citation analysis.
        await verifiedStatementSnapshots(analysis, statements);
        const statementHashById = new Map(analysis.statementSources.map((source) => [source.statementId, source.sha256]));
        if (!recoveryId.current) {
          const begun = await desktop.beginRecovery();
          recoveryId.current = begun.recoveryId;
          recoveryRevision.current = begun.revision;
        }
        const direct: RecoverySourceDescriptor[] = [];
        const sourcePathOf = (file: File) => {
          try {
            return desktop.sourcePath(file) || "";
          } catch {
            return "";
          }
        };
        for (const statement of statements) {
          const sourcePath = sourcePathOf(statement.file);
          const analysisSha256 = statementHashById.get(statement.id);
          if (!analysisSha256) throw new Error(`${statement.file.name} is no longer bound to the current analysis. Re-analyse before automatic recovery can update.`);
          if (sourcePath) direct.push({ id: statement.id, role: "statement", name: statement.file.name, path: sourcePath, sha256: analysisSha256, size: statement.file.size });
        }
        for (const evidence of analysis.evidence) {
          if (evidence.derivedFromEmail) continue;
          const sourcePath = sourcePathOf(evidence.file);
          if (sourcePath) direct.push({ id: evidence.id, role: "evidence", name: evidence.name, path: sourcePath, sha256: evidence.sha256, size: evidence.file.size });
        }
        for (const template of templates) {
          const sourcePath = sourcePathOf(template.file);
          if (sourcePath) direct.push({ id: `template-${template.slot}`, role: "template", name: template.file.name, path: sourcePath, sha256: template.sha256, size: template.file.size });
        }
        const currentSourceIds = new Set([
          ...statements.map((statement) => statement.id),
          ...analysis.evidence.filter((evidence) => !evidence.derivedFromEmail).map((evidence) => evidence.id),
          ...templates.map((template) => `template-${template.slot}`),
          "saved-project-archive",
        ]);
        // Do not retain paths for sources the user removed. In-memory files
        // restored from a journal or saved project keep their still-current
        // descriptor because Electron cannot recover a native path from them.
        const byId = new Map(recoverySourceDescriptors.current
          .filter((source) => currentSourceIds.has(source.id))
          .map((source) => [source.id, source]));
        for (const source of direct) byId.set(source.id, source);
        const sources = [...byId.values()];
        if (!sources.length) {
          recoverySourceDescriptors.current = [];
          return;
        }
        recoverySourceDescriptors.current = sources;
        const evidenceHashById = new Map(analysis.evidence.map((record) => [record.id, record.sha256]));
        const payload = {
          project: { name: projectName, profileId: profile.id },
          candidates: candidates.map((candidate) => ({ id: candidate.id, evidenceId: candidate.evidenceId, sourceSha256: candidate.evidenceId ? evidenceHashById.get(candidate.evidenceId) ?? candidate.parentEmailProvenance?.childSha256 ?? null : candidate.parentEmailProvenance?.childSha256 ?? null, statementSha256: candidate.manualAddition ? null : statementHashById.get(candidate.statementId ?? "") ?? null, included: candidate.included, confirmed: candidate.confirmed, confirmationMethod: candidate.confirmationMethod, confirmedAt: candidate.confirmedAt, description: candidate.description, aliases: candidate.aliases ?? [], reviewNote: candidate.reviewNote ?? "", date: candidate.date, paragraph: candidate.paragraph, citation: candidate.citation, citationResolution: candidate.citationResolution, discoverySignals: candidate.discoverySignals, confidence: candidate.confidence, rationale: candidate.rationale, mark: candidate.mark, provisionalNumber: candidate.provisionalNumber, exhibitInitials: candidate.exhibitInitials, exhibitSequence: candidate.exhibitSequence, witnessInitials: candidate.witnessInitials, witnessKey: candidate.witnessKey, statementId: candidate.statementId, statementName: candidate.statementName, pageStart: candidate.pageStart ?? null, pageEnd: candidate.pageEnd ?? null, sequenceOrder: candidate.sequenceOrder ?? candidate.provisionalNumber, repeatDecision: candidate.repeatDecision ?? null, manualAddition: candidate.manualAddition ?? false, manualAddedAt: candidate.manualAddedAt, manualWarningAcknowledgedAt: candidate.manualWarningAcknowledgedAt, emailAttachmentDispositions: candidate.emailAttachmentDispositions, parentEmailProvenance: candidate.parentEmailProvenance })),
          arrangement,
          layout,
          pagination,
          pageSizeChoices,
          resolutions,
          statements: statements.map((statement) => ({ id: statement.id, witnessName: statement.witnessName, witnessInitials: statement.witnessInitials, sourceId: statement.id, sourceSha256: statementHashById.get(statement.id)! })),
          templates: templates.map((template) => ({ sourceId: `template-${template.slot}`, slot: template.slot, sourceFormat: template.sourceFormat ?? "pdf" })),
          templateReviews: templates.flatMap((template) => template.pdfSha256 && template.reviewState ? [{ slot: template.slot, sourceId: `template-${template.slot}`, renderedSourceId: template.sourceFormat === "pdf" ? undefined : `template-rendered-${template.slot}`, sourceFormat: template.sourceFormat ?? "pdf", sourceSha256: template.sha256, pdfSha256: template.pdfSha256, reviewState: template.reviewState }] : []),
          templateDiscrepancyConfirmation: templateDiscrepancyPending ? undefined : templateDiscrepancyConfirmation ?? undefined,
          sources,
          fingerprint: lastBuildSnapshot?.fingerprint ?? null,
        };
        const revision = recoveryRevision.current + 1;
        recoveryRevision.current = revision;
        await desktop.writeRecovery(recoveryId.current!, revision, payload);
      })().catch((caught) => {
        if (!(caught instanceof Error) || !/stale/i.test(caught.message)) {
          if (caught instanceof Error && /statement.*(?:changed|analysis)|changed after it was selected/i.test(caught.message)) {
            const issue = "Automatic recovery kept its last valid copy because the witness statement no longer matches the analysed version. Re-analyse the statement before continuing.";
            setRecoveryIssues((issues) => issues.includes(issue) ? issues : [issue]);
            return;
          }
          const detail = caught instanceof Error && caught.message ? ` Automatic recovery remains unavailable: ${caught.message}` : "";
          const issue = `The automatic recovery copy could not be updated.${detail} Your source files and current bundle review are unaffected; save the exhibit project manually before closing.`;
          setRecoveryIssues((issues) => issues.includes(issue) ? issues : [issue]);
        }
      });
    }, 600);
    return () => window.clearTimeout(timer);
  }, [analysis, arrangement, candidates, layout, pagination, pageSizeChoices, projectName, recoveryOffer, resolutions, statements, templates, templateDiscrepancyConfirmation, templateDiscrepancyPending, lastBuildSnapshot, profile.id]);

  function makeStatement(file: File, index: number): BundleStatementInput {
    const stem = file.name.replace(/\.[^.]+$/, "").replace(/^\d+[_ -]*/, "");
    const initials = stem.split(/[_ -]+/).filter(Boolean).slice(-2).map((part) => part[0]).join("").toUpperCase() || "WS";
    return { id: `${Date.now()}-${index}-${file.name}`, file, witnessName: stem.replace(/_/g, " "), witnessInitials: initials };
  }

  function reportAnalysisProgress(stage: string, detail?: string) {
    setAnalysisProgress((current) => ({ stage, detail, startedAt: current?.startedAt ?? Date.now() }));
    const count = detail?.match(/^(\d+) of (\d+):/);
    const current = count ? Number(count[1]) : 0;
    const total = count ? Number(count[2]) : 0;
    if (announcedAnalysisStage.current !== stage || !count || current === 1 || current === total || current % 25 === 0) {
      announcedAnalysisStage.current = stage;
      setAnalysisProgressAnnouncement(count ? `${stage}: ${current} of ${total}` : stage);
    }
  }

  async function runAnalysis(
    statement: File,
    evidence: File[],
    requestedStatements?: BundleStatementInput[],
  ) {
    analysisCancelRequested.current = false;
    await dismissRecoveryOffer();
    recoverySourceDescriptors.current = recoverySourceDescriptors.current.filter((source) => source.id !== "saved-project-archive");
    setBusy("analyse");
    setError(null);
    setBuild(null);
    setAnalysisProgress({ stage: "Preparing the analysis", detail: "Checking the selected files", startedAt: Date.now() });
    announcedAnalysisStage.current = "Preparing the analysis";
    setAnalysisProgressAnnouncement("Preparing the analysis");
    try {
      const activeStatements = requestedStatements?.length
        ? requestedStatements
        : statements.length
          ? statements
          : [makeStatement(statement, 0)];
      const result = analysisWithGuidedMapping(
        await analyseBundleStatements(activeStatements, evidence, reportAnalysisProgress, () => analysisCancelRequested.current),
        statement.name,
        evidence.map((file) => file.name),
      );
      setAnalysis(result);
      setCandidates(result.candidates);
      if (result.statementId && result.witnessInitials) {
        setStatements(activeStatements.map((item) => item.id === result.statementId
          ? { ...item, witnessInitials: result.witnessInitials! }
          : item));
      }
      setResolutions([]);
      focusReviewAfterAnalysis.current = true;
      setView("review");
      setAnalysisProgressAnnouncement("Analysis complete. Review the proposed exhibit matches.");
    } catch (caught) {
      setError(
        caught instanceof Error ? caught.message : "Analysis could not complete.",
      );
    } finally {
      analysisCancelRequested.current = false;
      setBusy(null);
      setAnalysisProgress(null);
    }
  }

  async function useSamplePack() {
    analysisCancelRequested.current = false;
    await dismissRecoveryOffer();
    recoverySourceDescriptors.current = recoverySourceDescriptors.current.filter((source) => source.id !== "saved-project-archive");
    setBusy("sample");
    setError(null);
    setAnalysisProgress({ stage: "Opening the guided sample", detail: "Reading its bundled tutorial files", startedAt: Date.now() });
    announcedAnalysisStage.current = "Opening the guided sample";
    setAnalysisProgressAnnouncement("Opening the guided sample");
    try {
      const sample = await loadSamplePack();
      const guidedStatements: BundleStatementInput[] = [{
        id: "guided-sample",
        file: sample.statement,
        witnessName: "Guided Sample",
        witnessInitials: "AH",
      }];
      setStatementFile(sample.statement);
      setStatements(guidedStatements);
      setEvidenceFiles(sample.evidence);
      const result = applyGuidedSampleMapping(await analyseBundleStatements(guidedStatements, sample.evidence, reportAnalysisProgress, () => analysisCancelRequested.current));
      setAnalysis(result);
      setCandidates(result.candidates);
      setTemplates([]);
      setTemplateDiscrepancyConfirmation(null);
      setProjectName("Guided sample exhibit bundle");
      setLayout(DEFAULT_BUNDLE_LAYOUT);
      setPagination(DEFAULT_PAGINATION);
      setPaginationDraft(DEFAULT_PAGINATION);
      setResolutions([]);
      setBuild(null);
      focusReviewAfterAnalysis.current = true;
      setView("review");
      setGuidedSampleAvailability("available");
      setBusy(null);
      setAnalysisProgress(null);
      setAnalysisProgressAnnouncement("Guided sample analysis complete. Review the proposed exhibit matches.");
    } catch (caught) {
      const detail = caught instanceof Error ? caught.message : "";
      if (/Analysis cancelled/i.test(detail)) setError(detail);
      else {
        setGuidedSampleAvailability("unavailable");
        setError("The optional guided sample is unavailable. You can still build a bundle from your own files.");
      }
      setBusy(null);
      setAnalysisProgress(null);
      setAnalysisProgressAnnouncement("The guided sample could not be opened.");
    } finally {
      analysisCancelRequested.current = false;
    }
  }
  useSamplePackRef.current = useSamplePack;

  async function openGuidedSampleFolder() {
    if (!window.bundleBuilderDesktop?.openGuidedSampleFolder) {
      setError("The guided sample folder can only be opened in the desktop app.");
      return;
    }
    setError(null);
    try {
      await window.bundleBuilderDesktop.openGuidedSampleFolder();
    } catch {
      setError("The guided sample folder could not be opened. You can still build a bundle from your own files.");
    }
  }

  async function setGuidedSampleHidden(hidden: boolean) {
    setGuidedSampleHiddenState(hidden);
    try {
      if (window.bundleBuilderDesktop) {
        await window.bundleBuilderDesktop.setGuidedSampleHidden(hidden);
      } else {
        window.localStorage.setItem("exhibit-builder-hide-guided-sample", String(hidden));
      }
    } catch {
      setError("The guided-sample display preference could not be saved. The rest of the workspace is unaffected.");
    }
  }

  function startGuidedSampleTour() {
    setTourOpenedFolder(false);
    setTourDownloaded(false);
    setTourSaved(false);
    setTourActive(true);
  }

  function skipGuidedSampleTour() {
    setTourActive(false);
    setTourOpenedFolder(false);
    setTourDownloaded(false);
    setTourSaved(false);
  }

  async function retryGuidedSampleAvailability() {
    setGuidedSampleAvailability("checking");
    const available = await checkSamplePackAvailability();
    setGuidedSampleAvailability(available ? "available" : "unavailable");
  }

  function prepareManualEvidence(evidenceId: string) {
    setManualEvidenceId(evidenceId);
    const record = analysis?.evidence.find((item) => item.id === evidenceId) ?? (manualUploadedEvidence?.id === evidenceId ? manualUploadedEvidence : undefined);
    if (!record) return;
    setManualDescription(record.name.replace(/\.[^.]+$/, "").replace(/[_-]+/g, " "));
    setManualDate("Date not stated");
  }

  function openManualAdd(origin: ManualAddOrigin, evidenceId?: string) {
    manualAddTrigger.current = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    setOrderPreview(null);
    setManualAddOrigin(origin);
    setManualAddOpen(true);
    if (evidenceId) prepareManualEvidence(evidenceId);
  }

  function closeManualAdd() {
    setManualAddOpen(false);
    setManualEvidenceId("");
    setManualDescription("");
    setManualDate("Date not stated");
    setManualUploadedEvidence(null);
  }

  function eligibleUnusedAdd(record: EvidenceRecord) {
    return !selectedEvidenceHashes.has(record.sha256) && record.marker !== "N/A";
  }

  async function uploadManualEvidence(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file || !analysis) return;
    setBusy("manual");
    setError(null);
    try {
      const [record] = await analyseEvidenceFiles([file], analysis.evidence.length);
      setManualUploadedEvidence(record);
      setManualEvidenceId(record.id);
      setManualDescription(record.name.replace(/\.[^.]+$/, "").replace(/[_-]+/g, " "));
      setManualDate("Date not stated");
      setBuild(null);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The additional exhibit could not be read.");
    } finally {
      setBusy(null);
      event.target.value = "";
    }
  }

  function addManualExhibit() {
    const record = analysis?.evidence.find((item) => item.id === manualEvidenceId) ?? (manualUploadedEvidence?.id === manualEvidenceId ? manualUploadedEvidence : undefined);
    const description = manualDescription.trim();
    if (!analysis || !record || !description) {
      setError("Choose a document and enter the description that should appear in the index.");
      return;
    }
    const witnessInitials = "EX";
    const provisionalNumber = Math.max(0, ...candidates.map((candidate) => candidate.provisionalNumber)) + 1;
    const acknowledgedAt = new Date().toISOString();
    const manualCandidate: ExhibitCandidate = {
      id: `manual-${Date.now()}-${record.sha256.slice(0, 12)}`,
      mark: `${witnessInitials} ${provisionalNumber}`,
      provisionalNumber,
      description,
      date: manualDate.trim() || "Date not stated",
      paragraph: 0,
      citation: "",
      exhibitInitials: witnessInitials,
      exhibitSequence: 1,
      citationResolution: "none",
      discoverySignals: ["Manually added by reviewer"],
      evidenceId: record.id,
      confidence: 100,
      rationale: "Reviewer intentionally added an exhibit that is not cited in the witness statement",
      included: true,
      confirmed: true,
      confirmationMethod: "individual",
      confirmedAt: acknowledgedAt,
      witnessInitials,
      witnessKey: "general-exhibits::EX",
      sequenceOrder: Math.max(0, ...candidates.map((candidate) => candidate.sequenceOrder ?? candidate.provisionalNumber)) + 1_000,
      repeatDecision: candidates.some((candidate) => candidate.included && candidate.evidenceId && analysis.evidence.find((item) => item.id === candidate.evidenceId)?.sha256 === record.sha256) ? "separate" : undefined,
      manualAddition: true,
      manualAddedAt: acknowledgedAt,
      manualWarningAcknowledgedAt: acknowledgedAt,
    };
    if (manualUploadedEvidence?.id === record.id) {
      setAnalysis((current) => current ? { ...current, evidence: [...current.evidence, record], unreferenced: [...current.unreferenced, record] } : current);
      setEvidenceFiles((current) => [...current, record.file]);
    }
    setCandidates((current) => [...current, manualCandidate]);
    const origin = manualAddOrigin;
    closeManualAdd();
    setOrderPreview(null);
    setBuild(null);
    setReviewActionStatus(`Added ${description} as an exhibit. It is not cited in the statement.`);
    if (origin === "finalise") {
      setView("build");
      focusAfterArrangementChange(`[data-exhibit-id="${CSS.escape(`candidate-${manualCandidate.id}`)}"]`);
      return;
    }
    setShowPendingOnly(false);
    setShowDuplicatesOnly(false);
    setView("review");
    requestAnimationFrame(() => requestAnimationFrame(() => {
      const focusTarget = document.querySelector<HTMLElement>(confirmFocusSelector(manualCandidate.id))
        ?? document.querySelector<HTMLElement>(reviewCardSelector(manualCandidate.id));
      focusTarget?.focus({ preventScroll: true });
    }));
  }

  async function addEmailChildAsExhibit(parentCandidateId: string, child: EmailAttachmentChild) {
    if (!analysis) return;
    keepEmailReviewOpen(parentCandidateId);
    const existing = candidates.find((candidate) => candidate.parentEmailProvenance?.childIdentity === child.identity);
    if (existing) {
      setShowPendingOnly(false);
      setShowDuplicatesOnly(false);
      setView("review");
      requestAnimationFrame(() => requestAnimationFrame(() => {
        const focusTarget = document.querySelector<HTMLElement>(emailChildDescriptionSelector(existing.id))
          ?? firstActionableControl(document.querySelector<HTMLElement>(emailAttachmentsSelector(parentCandidateId)))
          ?? document.querySelector<HTMLElement>(reviewCardSelector(parentCandidateId));
        focusTarget?.focus({ preventScroll: true });
      }));
      return;
    }
    const parent = analysis.evidence.find((record) => record.emailAttachments?.some((item) => item.identity === child.identity));
    if (!parent) {
      setError("The parent email for this attachment is no longer available.");
      return;
    }
    setBusy("manual");
    setError(null);
    try {
      const [record] = await analyseEvidenceFiles([child.file], analysis.evidence.length);
      record.derivedFromEmail = { parentSha256: parent.sha256, childIdentity: child.identity };
      const acknowledgedAt = new Date().toISOString();
      const provisionalNumber = Math.max(0, ...candidates.map((candidate) => candidate.provisionalNumber)) + 1;
      const parentCandidate = candidates.find((candidate) => candidate.id === parentCandidateId);
      const parentSequence = parentCandidate?.sequenceOrder ?? parentCandidate?.provisionalNumber ?? 0;
      const ordinalPart = Number.parseFloat(child.ordinal) || 1;
      const description = child.name.replace(/\.[^.]+$/, "").replace(/[_-]+/g, " ");
      const manualCandidate: ExhibitCandidate = {
        id: `manual-${Date.now()}-${child.sha256.slice(0, 12)}`,
        mark: `EX ${provisionalNumber}`,
        provisionalNumber,
        description,
        date: "Date not stated",
        paragraph: 0,
        citation: "",
        exhibitInitials: "EX",
        exhibitSequence: 1,
        citationResolution: "none",
        discoverySignals: ["Manually added by reviewer", "Email attachment"],
        evidenceId: record.id,
        confidence: 100,
        rationale: "Reviewer added an email attachment as its own exhibit",
        included: true,
        confirmed: true,
        confirmationMethod: "individual",
        confirmedAt: acknowledgedAt,
        witnessInitials: "EX",
        witnessKey: "general-exhibits::EX",
        sequenceOrder: parentSequence + ordinalPart / 1_000,
        repeatDecision: candidates.some((candidate) => candidate.included && candidate.evidenceId && analysis.evidence.find((item) => item.id === candidate.evidenceId)?.sha256 === record.sha256) ? "separate" : undefined,
        manualAddition: true,
        manualAddedAt: acknowledgedAt,
        manualWarningAcknowledgedAt: acknowledgedAt,
        parentEmailProvenance: {
          parentName: parent.name,
          parentSha256: parent.sha256,
          childIdentity: child.identity,
          childSha256: child.sha256,
        },
      };
      setAnalysis((current) => current ? { ...current, evidence: [...current.evidence, record] } : current);
      setCandidates((current) => {
        const next = current.map((candidate) => candidate.id === parentCandidateId
          ? { ...candidate, emailAttachmentDispositions: { ...candidate.emailAttachmentDispositions, [child.identity]: "add-as-exhibit" as EmailChildDisposition } }
          : candidate);
        return [...next, manualCandidate];
      });
      setOrderPreview(null);
      setBuild(null);
      setShowPendingOnly(false);
      setShowDuplicatesOnly(false);
      setView("review");
      setReviewActionStatus(`Added ${description} as its own exhibit. It is not cited in the statement.`);
      requestAnimationFrame(() => requestAnimationFrame(() => {
        const focusTarget = document.querySelector<HTMLElement>(emailChildDescriptionSelector(manualCandidate.id))
          ?? firstActionableControl(document.querySelector<HTMLElement>(emailAttachmentsSelector(parentCandidateId)))
          ?? document.querySelector<HTMLElement>(reviewCardSelector(parentCandidateId));
        focusTarget?.focus({ preventScroll: true });
      }));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The email attachment could not be added as an exhibit.");
    } finally {
      setBusy(null);
    }
  }

  function setEmailChildDisposition(candidateId: string, child: EmailAttachmentChild, disposition: EmailChildDisposition) {
    const candidate = candidates.find((item) => item.id === candidateId);
    const record = candidate?.evidenceId ? analysis?.evidence.find((item) => item.id === candidate.evidenceId) : undefined;
    const group = candidate ? exhibitGroupLookupByKind.byCandidateId.get(candidate.id) : undefined;
    const repeatNeedsDecision = Boolean(group?.decisionPending || group?.selectionConflict || group?.collisionMembers.slice(1).some((repeat) => !repeat.confirmed));
    const nextDispositions = { ...candidate?.emailAttachmentDispositions, [child.identity]: disposition };
    const remaining = record?.emailAttachments?.length
      ? unresolvedEmailAttachments(record.emailAttachments, nextDispositions)
      : [];
    if (candidate?.confirmed && remaining.length === 0 && disposition !== "add-as-exhibit") {
      if (repeatNeedsDecision) keepCurrentReviewCard(candidateId);
      else beginConfirmViewportTransaction(candidateId);
      collapseConfirmedCard(candidateId);
    } else {
      keepEmailReviewOpen(candidateId);
    }
    if (disposition === "add-as-exhibit") {
      void addEmailChildAsExhibit(candidateId, child);
      return;
    }
    setCandidates((current) => current
      .filter((candidate) => candidate.parentEmailProvenance?.childIdentity !== child.identity)
      .map((candidate) => candidate.id === candidateId
        ? { ...candidate, emailAttachmentDispositions: { ...candidate.emailAttachmentDispositions, [child.identity]: disposition } }
        : candidate));
    setAnalysis((current) => current ? { ...current, evidence: current.evidence.filter((record) => record.derivedFromEmail?.childIdentity !== child.identity) } : current);
    setBuild(null);
  }

  function updateManualCandidate(candidateId: string, change: Pick<ExhibitCandidate, "description"> | Pick<ExhibitCandidate, "date"> | Pick<ExhibitCandidate, "pageStart"> | Pick<ExhibitCandidate, "pageEnd">) {
    setCandidates((current) => current.map((candidate) => candidate.id === candidateId ? { ...candidate, ...change, confirmed: true, confirmationMethod: "individual", confirmedAt: new Date().toISOString() } : candidate));
    setBuild(null);
  }

  function changeManualEvidence(candidateId: string, evidenceId: string) {
    setCandidates((current) => current.map((candidate) => candidate.id === candidateId ? { ...candidate, evidenceId, confirmed: true, confirmationMethod: "individual", confirmedAt: new Date().toISOString() } : candidate));
    setOrderPreview(null);
    setBuild(null);
  }

  function removeManualCandidate(candidateId: string) {
    const group = exhibitGroups.find((item) => item.id === candidateId) ?? exhibitGroups.find((item) => item.canonical.id === candidateId);
    const removed = candidates.find((candidate) => candidate.id === candidateId);
    const groupIndex = group ? finalOrder.indexOf(group.id) : -1;
    const focusExhibitId = groupIndex >= 0 ? finalOrder[groupIndex + 1] ?? finalOrder[groupIndex - 1] : undefined;
    setCandidates((current) => current
      .filter((candidate) => candidate.id !== candidateId)
      .map((candidate) => {
        const childIdentity = removed?.parentEmailProvenance?.childIdentity;
        if (!childIdentity || candidate.emailAttachmentDispositions?.[childIdentity] !== "add-as-exhibit") return candidate;
        const nextDispositions = { ...candidate.emailAttachmentDispositions };
        delete nextDispositions[childIdentity];
        return { ...candidate, emailAttachmentDispositions: nextDispositions };
      }));
    if (removed?.parentEmailProvenance) {
      setAnalysis((current) => current ? { ...current, evidence: current.evidence.filter((record) => record.derivedFromEmail?.childIdentity !== removed.parentEmailProvenance?.childIdentity) } : current);
    }
    if (group) setArrangement((current) => reconcileExhibitArrangement(current, rawExhibitGroups.filter((item) => item.id !== group.id)));
    setOrderPreview(null);
    setBuild(null);
    announceArrangement(`Removed manually added exhibit ${group?.canonical.description ?? candidateId}.`);
    focusAfterArrangementChange(focusExhibitId ? exhibitActionSelector(focusExhibitId, hasIndexHeadings ? "move-section" : "later") : "#new-index-heading", focusExhibitId ? exhibitActionSelector(focusExhibitId, "earlier") : undefined);
  }

  function updateCandidate(
    candidateId: string,
    change: Partial<ExhibitCandidate>,
  ) {
    const beforeCandidate = candidates.find((candidate) => candidate.id === candidateId);
    const sourceBefore = beforeCandidate?.evidenceId ? analysis?.evidence.find((record) => record.id === beforeCandidate.evidenceId) : undefined;
    if ("repeatDecision" in change && !("confirmed" in change) && !("evidenceId" in change)) {
      if (!beforeCandidate || beforeCandidate.repeatDecision === change.repeatDecision) return;
    }
    setCandidates((current) => applyCandidateReviewChange(current, candidateId, change, new Map((analysis?.evidence ?? []).map((record) => [record.id, record]))));
    if (change.included === false && beforeCandidate?.included) {
      setResolutions((current) => [
        ...current.filter((resolution) => resolution.candidateId !== candidateId),
        {
          blockerId: `candidate-exclusion-${candidateId}`,
          action: "exclude-candidate",
          candidateId,
          sourceId: sourceBefore?.id,
          sourceSha256: sourceBefore?.sha256,
          fileName: sourceBefore?.name,
          approvedAt: new Date().toISOString(),
          note: "Reviewer excluded this cited item from the exhibit bundle.",
        },
      ]);
    } else if (change.included === true) {
      setResolutions((current) => current.filter((resolution) => resolution.candidateId !== candidateId));
    }
    setBuild(null);
  }

  function confirmRepeatDecision(repeat: ExhibitCandidate, confirmed: boolean, cardId: string) {
    if (confirmed) keepCurrentReviewCard(cardId);
    updateCandidate(repeat.id, confirmed
      ? { confirmed: true, confirmationMethod: "individual", confirmedAt: new Date().toISOString() }
      : { confirmed: false, confirmationMethod: undefined, confirmedAt: undefined });
    setReviewActionStatus(confirmed
      ? `Repeat decision for paragraph ${repeat.paragraph} saved.`
      : `Repeat decision for paragraph ${repeat.paragraph} reopened.`);
  }

  function saveResolution(resolution: BuildResolution) {
    setResolutions((current) => [
      ...current.filter((existing) => !(existing.blockerId === resolution.blockerId && existing.sourceSha256 === resolution.sourceSha256 && existing.candidateId === resolution.candidateId)),
      resolution,
    ]);
    setBuild(null);
  }

  function approveOcrException(check: PreflightCheck) {
    if (!check.sourceId || !check.sourceSha256 || !isOcrCheck(check)) return;
    if (!visuallyReviewedSourceHashes.has(check.sourceSha256)) {
      setError("Open and visually review the original PDF before approving inclusion without OCR.");
      return;
    }
    saveResolution({
      blockerId: check.id,
      checkCode: check.code,
      profileId: profile.id,
      action: "proceed-without-ocr",
      sourceId: check.sourceId,
      sourceSha256: check.sourceSha256,
      fileName: check.fileName,
      approvedAt: new Date().toISOString(),
      note: "Reviewer visually reviewed the source and approved inclusion without a tool-generated OCR layer.",
      visualReviewConfirmed: true,
    });
  }

  function previewOriginalPdf(record: EvidenceRecord, origin: HTMLElement) {
    if (!record || record.extension !== "pdf") {
      setError("Only an original PDF can be opened in this visual-review step.");
      return;
    }
    ocrPreviewOrigin.current = origin;
    setOcrSourcePreview({ sourceId: record.id, sourceSha256: record.sha256, name: record.name, file: record.file, extractedText: record.text });
  }

  function previewOcrSource(check: PreflightCheck, origin: HTMLElement) {
    if (!check.sourceId || !check.sourceSha256 || !analysis) return;
    const record = evidenceById.get(check.sourceId);
    if (record) {
      previewOriginalPdf(record, origin);
      return;
    }
    const child = analysis.evidence.flatMap((item) => item.emailAttachments ?? []).find((item) => item.identity === check.sourceId && item.sha256 === check.sourceSha256);
    if (!child || child.extension !== "pdf") {
      setError("The original PDF is no longer available. Return to Sources and select it again.");
      return;
    }
    ocrPreviewOrigin.current = origin;
    setOcrSourcePreview({ sourceId: child.identity, sourceSha256: child.sha256, name: child.name, file: child.file, extractedText: "" });
  }

  function clearResolution(check: PreflightCheck) {
    setResolutions((current) => current.filter((resolution) => !(resolution.blockerId === check.id && (!resolution.sourceSha256 || resolution.sourceSha256 === check.sourceSha256))));
    setBuild(null);
  }

  function undoResolution(resolution: BuildResolution) {
    if (resolution.action === "proceed-without-ocr") {
      const check = rawPreflight.find((item) => item.id === resolution.blockerId);
      if (check) {
        clearResolution(check);
        return;
      }
    }
    setResolutions((current) => current.filter((item) => item !== resolution));
    setBuild(null);
  }

  function excludeSource(check: PreflightCheck) {
    if (!check.sourceId || !check.sourceSha256) return;
    const affectedCount = candidates.filter((candidate) => candidate.included && candidate.evidenceId === check.sourceId).length;
    const workbook = check.code === "workbook.fidelity_failed";
    if (affectedCount > 1 && !window.confirm(`${workbook ? "Leave this Excel file out of the bundle" : "Leave this file out of the bundle"}? Every exhibit that uses it is left out too. The witness statement is not edited.`)) return;
    setCandidates((current) => current.map((candidate) => candidate.evidenceId === check.sourceId ? { ...candidate, included: false, confirmed: false } : candidate));
    saveResolution({
      blockerId: check.id,
      action: "exclude-source",
      sourceId: check.sourceId,
      sourceSha256: check.sourceSha256,
      fileName: check.fileName,
      approvedAt: new Date().toISOString(),
      note: "Reviewer excluded this source and all current citations using it from the exhibit bundle.",
    });
  }

  function useBuiltInTemplates() {
    const slots = activeCustomTemplates.map((template) => template.slot);
    if (!slots.length) return;
    saveResolution({
      blockerId: "template-approval",
      action: "use-built-in-template",
      templateSlots: slots,
      templateHashes: Object.fromEntries(templates.filter((template) => slots.includes(template.slot)).map((template) => [template.slot, template.sha256])),
      approvedAt: new Date().toISOString(),
      note: "Reviewer selected the built-in layout instead of the custom template requiring review.",
    });
  }

  function applyArrangement(next: BundleArrangement) {
    setOrderHistory((history) => [...history.slice(-19), cloneBundleArrangement(arrangement)]);
    setArrangement(next);
    setOrderPreview(null);
    setBuild(null);
  }

  function announceArrangement(message: string) {
    setArrangementStatus("");
    window.setTimeout(() => setArrangementStatus(message), 0);
  }

  function preserveReviewViewport(candidateId?: string | null) {
    viewportProbe.current?.cancel();
    const cards = [...document.querySelectorAll<HTMLElement>(".exhibit-card-list .exhibit-review-card")];
    const anchoredId = candidateId ?? firstVisibleReviewCardId(cards, window.innerHeight);
    const current = anchoredId
      ? document.querySelector<HTMLElement>(reviewCardSelector(anchoredId))
      : cards[0] ?? null;
    const anchor = captureViewportAnchor(current, window.scrollY);
    const focusId = anchoredId ?? current?.getAttribute("data-candidate-id");
    viewportProbe.current = probeUntilStable(
      () => {
        const card = focusId ? document.querySelector<HTMLElement>(reviewCardSelector(focusId)) : null;
        return card ?? document.querySelector<HTMLElement>(".exhibit-card-list .exhibit-review-card") ?? reviewHeading.current;
      },
      (callback) => requestAnimationFrame(callback),
      (stable) => {
        restoreViewportAnchor(stable, anchor, (top) => window.scrollTo({ top, behavior: "auto" }));
        const cardId = stable.getAttribute?.("data-candidate-id") ?? focusId;
        const panel = cardId ? document.querySelector<HTMLElement>(emailAttachmentsSelector(cardId)) : null;
        const confirm = cardId ? document.querySelector<HTMLElement>(confirmFocusSelector(cardId)) : null;
        const focusTarget = firstActionableControl(panel)
          ?? firstActionableControl(confirm)
          ?? firstActionableControl(stable)
          ?? reviewHeading.current;
        focusTarget?.focus({ preventScroll: true });
      },
    );
  }

  function restoreDocumentPickerSummary(from: EventTarget | Node | null) {
    const picker = from instanceof Element ? from.closest(".document-picker") : null;
    const summary = picker?.querySelector("summary");
    if (summary instanceof HTMLElement) summary.focus({ preventScroll: true });
  }

  function focusConfirmControl(candidateId: string) {
    requestAnimationFrame(() => requestAnimationFrame(() => {
      document.querySelector<HTMLElement>(confirmFocusSelector(candidateId))?.focus({ preventScroll: true });
    }));
  }

  function beginConfirmViewportTransaction(confirmedId: string, nextPendingId?: string | null) {
    viewportProbe.current?.cancel();
    setOpenDocumentPickerId(null);
    const cards = [...document.querySelectorAll<HTMLElement>(".exhibit-card-list .exhibit-review-card")];
    confirmViewport.current = {
      id: ++confirmViewportSeq.current,
      scrollY: window.scrollY,
      confirmedId,
      nextPendingId: nextPendingId === undefined ? nextPendingConfirmCardId(cards, confirmedId) : nextPendingId,
    };
    document.querySelector<HTMLElement>(reviewCardListSelector())?.focus({ preventScroll: true });
  }

  function collapseConfirmedCard(candidateId: string) {
    setExpandedConfirmedCards((current) => {
      if (!current.has(candidateId)) return current;
      const next = new Set(current);
      next.delete(candidateId);
      return next;
    });
    setOpenEmailAttachmentsId((current) => current === candidateId ? null : current);
  }

  function preserveNextReviewCard(candidateId: string) {
    beginConfirmViewportTransaction(candidateId);
    collapseConfirmedCard(candidateId);
  }

  function keepCurrentReviewCard(candidateId: string) {
    beginConfirmViewportTransaction(candidateId, candidateId);
  }

  function keepEmailReviewOpen(candidateId: string) {
    setOpenDocumentPickerId(null);
    setExpandedConfirmedCards((current) => new Set(current).add(candidateId));
    setOpenEmailAttachmentsId(candidateId);
  }

  function focusAfterArrangementChange(primarySelector: string, fallbackSelector?: string) {
    requestAnimationFrame(() => requestAnimationFrame(() => {
      const primary = document.querySelector<HTMLElement>(primarySelector);
      const fallback = fallbackSelector ? document.querySelector<HTMLElement>(fallbackSelector) : null;
      const target = primary && !(primary instanceof HTMLButtonElement && primary.disabled) ? primary : fallback;
      target?.focus();
    }));
  }

  function exhibitActionSelector(exhibitId: string, action: string) {
    return `[data-exhibit-id="${CSS.escape(exhibitId)}"] [data-action="${action}"]`;
  }

  function sectionHeadingDomId(sectionId: string) {
    return `arrangement-heading-${encodeURIComponent(sectionId)}`;
  }

  function exhibitLocation(exhibitId: string) {
    for (let nodeIndex = 0; nodeIndex < arrangement.nodes.length; nodeIndex += 1) {
      const node = arrangement.nodes[nodeIndex];
      if (node.type === "exhibit" && node.exhibitId === exhibitId) return { sectionId: null, index: nodeIndex, length: arrangement.nodes.length } as const;
      if (node.type === "section") {
        const index = node.exhibits.findIndex((exhibit) => exhibit.exhibitId === exhibitId);
        if (index >= 0) return { sectionId: node.id, index, length: node.exhibits.length } as const;
      }
    }
    return null;
  }

  function moveCandidate(exhibitId: string, direction: -1 | 1) {
    const location = exhibitContainerLocation(arrangement, exhibitId);
    if (!location) return;
    const target = location.index + direction;
    if (target < 0 || target >= location.length) return;
    applyArrangement(moveArrangementExhibitInContainer(arrangement, exhibitId, target));
    const description = exhibitGroupLookupByKind.byGroupId.get(exhibitId)?.canonical.description ?? "Exhibit";
    announceArrangement(`Moved ${description} ${direction < 0 ? "earlier" : "later"}.`);
    focusAfterArrangementChange(exhibitActionSelector(exhibitId, direction < 0 ? "earlier" : "later"), hasIndexHeadings ? exhibitActionSelector(exhibitId, "move-section") : "#new-index-heading");
  }

  function moveCandidateBefore(exhibitId: string, targetId: string) {
    const source = exhibitLocation(exhibitId);
    const target = exhibitLocation(targetId);
    if (!source || !target || exhibitId === targetId) return;
    const targetIndex = source.sectionId === target.sectionId && source.index < target.index ? target.index - 1 : target.index;
    applyArrangement(moveArrangementExhibit(arrangement, exhibitId, { sectionId: target.sectionId, index: targetIndex }));
    announceArrangement(`Moved ${exhibitGroupLookupByKind.byGroupId.get(exhibitId)?.canonical.description ?? "exhibit"} to its new position.`);
  }

  function moveCandidateToEdge(exhibitId: string, edge: "top" | "bottom") {
    const location = exhibitContainerLocation(arrangement, exhibitId);
    if (!location) return;
    applyArrangement(moveArrangementExhibitInContainer(arrangement, exhibitId, edge === "top" ? 0 : location.length - 1));
    const description = exhibitGroupLookupByKind.byGroupId.get(exhibitId)?.canonical.description ?? "Exhibit";
    const place = location.sectionId ? "its current heading" : "this group of exhibits with no heading";
    announceArrangement(`Moved ${description} to the ${edge} of ${place}.`);
    focusAfterArrangementChange(exhibitActionSelector(exhibitId, edge), hasIndexHeadings ? exhibitActionSelector(exhibitId, "move-section") : exhibitActionSelector(exhibitId, "later"));
  }

  function moveCandidateToSection(exhibitId: string, sectionId: string | null) {
    const location = exhibitLocation(exhibitId);
    if (!location || location.sectionId === sectionId) return;
    const targetLength = sectionId === null
      ? arrangement.nodes.length - (location.sectionId === null ? 1 : 0)
      : arrangement.nodes.find((node): node is ArrangementSectionNode => node.type === "section" && node.id === sectionId)?.exhibits.length ?? 0;
    applyArrangement(moveArrangementExhibit(arrangement, exhibitId, { sectionId, index: targetLength }));
    const targetName = sectionId === null ? "No heading" : arrangement.nodes.find((node): node is ArrangementSectionNode => node.type === "section" && node.id === sectionId)?.heading ?? "the selected heading";
    announceArrangement(`Moved ${exhibitGroupLookupByKind.byGroupId.get(exhibitId)?.canonical.description ?? "exhibit"} to ${targetName}.`);
    focusAfterArrangementChange(hasIndexHeadings ? exhibitActionSelector(exhibitId, "move-section") : exhibitActionSelector(exhibitId, "later"));
  }

  function addSection() {
    const heading = newSectionHeading.trim();
    if (!heading) return;
    const sectionId = `section-${crypto.randomUUID()}`;
    applyArrangement(addArrangementSection(arrangement, { id: sectionId, heading, index: arrangement.nodes.length }));
    setNewSectionHeading("");
    announceArrangement(`Added index heading ${heading}.`);
    focusAfterArrangementChange(`#${CSS.escape(sectionHeadingDomId(sectionId))}-summary`);
  }

  function renameSection(sectionId: string) {
    const section = arrangement.nodes.find((node): node is ArrangementSectionNode => node.type === "section" && node.id === sectionId);
    if (!section) return;
    const heading = window.prompt("Rename this index heading", section.heading)?.trim();
    if (!heading || heading === section.heading) return;
    if (heading.length > 512) {
      setError("Index headings must be 512 characters or fewer.");
      return;
    }
    applyArrangement(renameArrangementSection(arrangement, sectionId, heading));
    announceArrangement(`Renamed index heading to ${heading}.`);
  }

  function deleteSection(sectionId: string) {
    const section = arrangement.nodes.find((node): node is ArrangementSectionNode => node.type === "section" && node.id === sectionId);
    if (!section || !window.confirm(`Delete the heading “${section.heading}”?\n\nIts ${section.exhibits.length} exhibit${section.exhibits.length === 1 ? "" : "s"} will be kept in the same place, still printing in the main index list.`)) return;
    const sections = arrangement.nodes.filter((node): node is ArrangementSectionNode => node.type === "section");
    const sectionIndex = sections.findIndex((node) => node.id === sectionId);
    const focusSection = sections[sectionIndex + 1] ?? sections[sectionIndex - 1];
    applyArrangement(deleteArrangementSectionKeepItems(arrangement, sectionId));
    announceArrangement(`Deleted heading ${section.heading}. Its ${section.exhibits.length} exhibit${section.exhibits.length === 1 ? " is" : "s are"} now under no heading and still print.`);
    focusAfterArrangementChange(focusSection ? `#${CSS.escape(sectionHeadingDomId(focusSection.id))}-summary` : "#new-index-heading");
  }

  function moveSection(sectionId: string, direction: -1 | 1) {
    const index = arrangement.nodes.findIndex((node) => node.type === "section" && node.id === sectionId);
    if (index < 0) return;
    const target = index + direction;
    if (target < 0 || target >= arrangement.nodes.length) return;
    applyArrangement(moveArrangementSection(arrangement, sectionId, target));
    const heading = arrangement.nodes.find((node): node is ArrangementSectionNode => node.type === "section" && node.id === sectionId)?.heading ?? "Index heading";
    announceArrangement(`Moved ${heading} ${direction < 0 ? "earlier" : "later"}.`);
    focusAfterArrangementChange(`#${CSS.escape(sectionHeadingDomId(sectionId))}-summary`);
  }

  function dropSectionBefore(sectionId: string, beforeNodeIndex: number) {
    applyArrangement(moveArrangementSectionBefore(arrangement, sectionId, beforeNodeIndex));
    const heading = arrangement.nodes.find((node): node is ArrangementSectionNode => node.type === "section" && node.id === sectionId)?.heading ?? "Index heading";
    announceArrangement(`Moved ${heading}.`);
    focusAfterArrangementChange(`#${CSS.escape(sectionHeadingDomId(sectionId))}-summary`);
  }

  function toggleSectionCollapsed(sectionId: string) {
    setCollapsedSectionIds((current) => {
      const next = new Set(current);
      if (next.has(sectionId)) next.delete(sectionId);
      else next.add(sectionId);
      return next;
    });
  }

  function previewOrderSort() {
    const byId = new Map(exhibitGroups.map((group) => [group.id, group]));
    const compare = (leftId: string, rightId: string) => {
      const left = byId.get(leftId);
      const right = byId.get(rightId);
      if (!left || !right) return 0;
      if (left.canonical.manualAddition !== right.canonical.manualAddition) return left.canonical.manualAddition ? 1 : -1;
      if (orderSort === "statement") return left.canonical.paragraph - right.canonical.paragraph;
      if (orderSort === "date") return left.canonical.date.localeCompare(right.canonical.date) || left.canonical.paragraph - right.canonical.paragraph;
      if (orderSort === "filename") return left.evidence.name.localeCompare(right.evidence.name) || left.canonical.paragraph - right.canonical.paragraph;
      return left.canonical.description.localeCompare(right.canonical.description) || left.canonical.paragraph - right.canonical.paragraph;
    };
    const labels = { statement: "statement order", date: "document date", filename: "source filename", description: "index description" } as const;
    setOrderPreview({ arrangement: sortBundleArrangementWithinSections(arrangement, compare), label: labels[orderSort] });
    focusAfterArrangementChange('[data-testid="order-preview-banner"]');
  }

  function applyOrderPreview() {
    if (!orderPreview) return;
    const label = orderPreview.label;
    applyArrangement(orderPreview.arrangement);
    announceArrangement(`Applied the proposed ${label}.`);
    focusAfterArrangementChange('[data-testid="preview-order-button"]');
  }

  function cancelOrderPreview() {
    setOrderPreview(null);
    announceArrangement("Cancelled the automatic order preview. The current arrangement is unchanged.");
    focusAfterArrangementChange('[data-testid="preview-order-button"]');
  }

  function undoOrderChange() {
    const previous = orderHistory.at(-1);
    if (!previous) return;
    setOrderHistory((history) => history.slice(0, -1));
    setArrangement(reconcileExhibitArrangement(previous, rawExhibitGroups));
    setOrderPreview(null);
    setBuild(null);
    announceArrangement("Undid the latest arrangement change.");
    focusAfterArrangementChange(orderHistory.length > 1 ? '[data-testid="undo-order-button"]' : '[data-testid="preview-order-button"]');
  }

  function confirmMatchedExhibits() {
    if (!bulkConfirmableCount) return;
    setBulkConfirmationAcknowledged(false);
    setBulkConfirmationOpen(true);
  }

  function confirmReviewedMatches() {
    if (!bulkConfirmationAcknowledged || !bulkConfirmableCount) return;
    const cards = [...document.querySelectorAll<HTMLElement>(".exhibit-card-list .exhibit-review-card")];
    preserveReviewViewport(firstVisibleReviewCardId(cards, window.innerHeight));
    const confirmedAt = new Date().toISOString();
    const confirmableIds = new Set(bulkConfirmableCandidates.map((candidate) => candidate.id));
    setCandidates((current) => current.map((candidate) => {
      return confirmableIds.has(candidate.id) ? { ...candidate, confirmed: true, confirmationMethod: "bulk", confirmedAt } : candidate;
    }));
    setBulkConfirmationOpen(false);
    setBulkConfirmationAcknowledged(false);
    setExpandedConfirmedCards(new Set());
    setBuild(null);
  }

  function buildReportFileBase() {
    return projectName.replace(/[^a-z0-9]+/gi, "_") || "Exhibit_Builder";
  }

  function buildReportPayload() {
    if (!build) return null;
    return createBuildReportPayload({
      projectName,
      build,
      candidates,
      analysis,
      preflight,
      resolutions,
    });
  }

  function downloadReadableBuildReport() {
    const payload = buildReportPayload();
    if (!payload) return;
    void downloadBytes(new TextEncoder().encode(formatBuildReportText(payload)), `${buildReportFileBase()}_Build_Report.txt`, "text/plain;charset=utf-8");
  }

  function downloadTechnicalBuildReport() {
    const payload = buildReportPayload();
    if (!payload) return;
    downloadJson(payload, `${buildReportFileBase()}_Build_Report.json`);
  }

  async function downloadVolumeZip() {
    if (!build?.volumeZipBytes || !build.volumeZipFileName) return;
    await downloadBytes(build.volumeZipBytes, build.volumeZipFileName, "application/zip");
  }

  function updateSheetSelection(evidenceId: string, sheetName: string, included: boolean, childIdentity?: string) {
    setAnalysis((current) => current ? {
      ...current,
      evidence: current.evidence.map((record) => {
        if (record.id !== evidenceId) return record;
        if (childIdentity) {
          return {
            ...record,
            emailAttachments: record.emailAttachments?.map((child) => child.identity !== childIdentity ? child : {
              ...child,
              sheetSelections: (child.sheetSelections ?? []).map((sheet) => sheet.name === sheetName ? { ...sheet, included } : sheet),
            }),
          };
        }
        return { ...record, sheetSelections: (record.sheetSelections ?? []).map((sheet) => sheet.name === sheetName ? { ...sheet, included } : sheet) };
      }),
    } : current);
    setBuild(null);
  }

  function handleStatement(event: ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files ?? []);
    const file = files[0] ?? null;
    setStatementFile(file);
    setStatements(file ? [makeStatement(file, 0)] : []);
    setBuild(null);
  }

  function handleEvidence(event: ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files ?? []);
    setEvidenceFiles(files);
    setBuild(null);
  }

  async function generateBundle() {
    if (!analysis) return;
    buildCancelRequested.current = false;
    setBusy("build");
    setError(null);
    setBuildProgress({ stage: "Preparing the build", detail: "Checking the confirmed exhibit bundle", startedAt: Date.now() });
    try {
      // Yield once so the progress card can paint before heavier local work
      // begins. Subsequent updates come from the real compositor stages.
      await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
      let result = await buildBundle(analysis, candidates, {
        pagination,
        templates,
        resolutions,
        layout: { ...layout, countOptionalPagesInReferences: countsOptionalPagesInReferences(pagination, layout) },
        arrangement,
        canonicalOrder: finalOrder,
        pageSizeChoices,
        workbookExporter: window.bundleBuilderDesktop?.exportWorkbook
          ? async (file, sheets) => window.bundleBuilderDesktop!.exportWorkbook(file.name, new Uint8Array(await file.arrayBuffer()), sheets)
          : undefined,
        isCancelled: () => buildCancelRequested.current,
        statements,
        onProgress: (stage, detail) => setBuildProgress((current) => ({ stage, detail, startedAt: current?.startedAt ?? Date.now() })),
      });
      if (buildCancelRequested.current) throw new Error("Bundle build cancelled. No finished bundle was replaced.");
      const currentSnapshot = await createSubstantiveBuildSnapshot({
        analysis,
        candidates,
        arrangement,
        templates,
        layout,
        pagination,
        pageSizeChoices,
        resolutions,
        volumePlan: result.buildPlan?.volumes.map((volume) => ({ number: volume.number, itemIds: volume.items.map((item) => item.id), totalPages: volume.totalPages, oversize: volume.oversize })) ?? null,
        statementSuggestions: buildStatementUpdateSuggestions(result.records).map((suggestion) => suggestion.line),
      });
      if (buildCancelRequested.current) throw new Error("Bundle build cancelled. No finished bundle was replaced.");
      const comparison = compareSubstantiveBuilds(lastBuildSnapshot, currentSnapshot);
      if (buildCancelRequested.current) throw new Error("Bundle build cancelled. No finished bundle was replaced.");
      result = await finalizeBuildAudit(result, currentSnapshot.fingerprint, comparison, () => buildCancelRequested.current);
      if (buildCancelRequested.current) throw new Error("Bundle build cancelled. No finished bundle was replaced.");
      setBuild(result);
      setLastBuildSnapshot(currentSnapshot);
      setRebuildComparison(comparison);
      setReorderReturn(null);
      setView("build");
    } catch (caught) {
      setError(
        caught instanceof Error ? caught.message : "Bundle build failed.",
      );
    } finally {
      buildCancelRequested.current = false;
      setBusy(null);
      setBuildProgress(null);
    }
  }

  async function resetWorkspace() {
    await dismissRecoveryOffer();
    const desktop = window.bundleBuilderDesktop;
    const activeRecoveryId = recoveryId.current;
    recoveryId.current = null;
    recoveryRevision.current = 0;
    recoverySourceDescriptors.current = [];
    if (desktop && activeRecoveryId) {
      try {
        await desktop.discardRecovery(activeRecoveryId);
        const begun = await desktop.beginRecovery();
        recoveryId.current = begun.recoveryId;
        recoveryRevision.current = begun.revision;
      } catch {
        // A fresh workspace must remain usable if journal cleanup is unavailable.
      }
    }
    setAnalysis(null);
    setCandidates([]);
    setBuild(null);
    setStatementFile(null);
    setEvidenceFiles([]);
    setStatements([]);
    setTemplates([]);
    setResolutions([]);
    setArrangement(bundleArrangementFromLegacyOrder([]));
    setOrderHistory([]);
    setOrderPreview(null);
    setNewSectionHeading("");
    setArrangementStatus("");
    setManualAddOpen(false);
    setManualEvidenceId("");
    setManualDescription("");
    setManualDate("Date not stated");
    setManualUploadedEvidence(null);
    setPageSizeChoices({});
    setStatementDrafts({});
    setReorderReturn(null);
    setLastBuildSnapshot(null);
    setRebuildComparison(null);
    setTemplatePreview(null);
    setOcrSourcePreview(null);
    setVisuallyReviewedSourceHashes(new Set());
    setTemplateDiscrepancyConfirmation(null);
    setOpenDocumentPickerId(null);
    setError(null);
    setView("review");
    skipGuidedSampleTour();
    if (statementInput.current) statementInput.current.value = "";
    if (evidenceInput.current) evidenceInput.current.value = "";
  }

  function updateStatementDraft(id: string, change: Partial<{ witnessName: string; witnessInitials: string }>) {
    const statement = statements.find((item) => item.id === id);
    if (!statement) return;
    setStatementDrafts((current) => ({
      ...current,
      [id]: { ...(current[id] ?? { witnessName: statement.witnessName, witnessInitials: statement.witnessInitials }), ...change },
    }));
  }

  function cancelStatementDraft(id: string) {
    setStatementDrafts((current) => {
      const next = { ...current };
      delete next[id];
      return next;
    });
  }

  function explicitPlaceholderInitials(id: string) {
    return [...new Set(candidates.filter((candidate) => candidate.statementId === id && candidate.citationToken).flatMap((candidate) => {
      const token = candidate.citationToken?.replace(/^[\[(]|[\])]$/g, "").trim() ?? "";
      if (/^exhib(?:it)?(?:\s+x+)?$/i.test(token)) return [];
      const match = token.match(/^([A-Z]{1,6})[\s-]*(?:\d+|x+)/i);
      return match ? [match[1].toUpperCase()] : [];
    }))];
  }

  function applyStatementDraft(id: string, confirmedConflict = false) {
    if (!analysis) return;
    const statement = statements.find((item) => item.id === id);
    const draft = statementDrafts[id];
    if (!statement || !draft) return;
    const proposed = draft.witnessInitials.trim().toUpperCase() || "EX";
    const existing = explicitPlaceholderInitials(id).filter((initials) => initials !== proposed);
    if (existing.length && !confirmedConflict) {
      setInitialsConfirmation({ statementId: id, existing, proposed });
      return;
    }
    const updatedStatement = { ...statement, witnessName: statement.witnessName, witnessInitials: proposed };
    const updated = applyWitnessDetails(analysis, candidates, updatedStatement);
    setStatements((current) => current.map((item) => item.id === id ? updatedStatement : item));
    setAnalysis(updated.analysis);
    setCandidates(updated.candidates);
    cancelStatementDraft(id);
    setBuild(null);
    setError(null);
    setInitialsConfirmation(null);
  }

  function commitPagination(change: Partial<PageNumberSettings>, warn = true) {
    const next = commitPaginationChange(pagination, paginationDraft, change);
    const paginationToUse = lockPagination(next.pagination);
    const draftToUse = lockPagination(next.draft);
    if (warn && numberingDiffersFromPdfOrder(paginationToUse)) {
      setPaginationConfirmation(paginationToUse);
      setPaginationPendingChange(change);
      setPaginationDraft(draftToUse);
      return;
    }
    setPagination(paginationToUse);
    setPaginationDraft(lockPagination(paginationDraftAfterWarningAccept(paginationToUse, draftToUse, change)));
    setLayout((current) => current.countOptionalPagesInReferences ? current : { ...current, countOptionalPagesInReferences: true });
  }

  function acceptPaginationChange() {
    if (!paginationConfirmation) return;
    const change = paginationPendingChange ?? {};
    setPagination(lockPagination(paginationConfirmation));
    setPaginationDraft((current) => lockPagination(paginationDraftAfterWarningAccept(paginationConfirmation, current, change)));
    setLayout((current) => current.countOptionalPagesInReferences ? current : { ...current, countOptionalPagesInReferences: true });
    setPaginationConfirmation(null);
    setPaginationPendingChange(null);
  }

  function cancelPaginationChange() {
    setPaginationDraft((current) => paginationDraftAfterWarningCancel(pagination, current));
    setPaginationConfirmation(null);
    setPaginationPendingChange(null);
  }

  function chooseVolumeNumbering(choice: PageNumberSettings["volumeNumbering"]) {
    if (choice === "restart" && pagination.volumeNumbering !== "restart") {
      setVolumeNumberingConfirmation(true);
      return;
    }
    const next = commitPaginationChange(pagination, paginationDraft, { volumeNumbering: choice });
    setPagination(lockPagination(next.pagination));
    setPaginationDraft(lockPagination(next.draft));
  }

  function acceptVolumeNumberingRestart() {
    const next = commitPaginationChange(pagination, paginationDraft, { volumeNumbering: "restart" });
    setPagination(lockPagination(next.pagination));
    setPaginationDraft(lockPagination(next.draft));
    setVolumeNumberingConfirmation(false);
  }

  function beginOrderChange() {
    if (!build) return;
    setOrderChangeConfirmation(true);
  }

  function confirmOrderChange() {
    if (!build || !retainedBuildInputs) return;
    setReorderReturn(captureRetainedBundle({
      build,
      arrangement: cloneBundleArrangement(arrangement),
      candidates: candidates.map((candidate) => ({ ...candidate })),
      pageSizeChoices: { ...pageSizeChoices },
      inputs: retainedBuildInputs,
    }));
    setBuild(null);
    setOrderHistory([]);
    setOrderPreview(null);
    setOrderChangeConfirmation(false);
  }

  function cancelOrderChange() {
    setOrderChangeConfirmation(false);
  }

  function keepCurrentBundle() {
    if (!reorderReturn || !retainedBuildInputs) return;
    const retained = dropStaleRetainedBuild(reorderReturn, retainedBuildInputs, { readyToBuild: retainReadyToBuild });
    if (!retained) return;
    const restored = restoredBundleFromRetain(retained, {
      candidates,
      pageSizeChoices,
      exhibitIds: rawExhibitGroups.map((group) => group.id),
    });
    setArrangement(restored.arrangement);
    setCandidates(restored.candidates);
    setPageSizeChoices(restored.pageSizeChoices);
    setBuild(restored.build);
    setReorderReturn(null);
    setOrderHistory([]);
    setOrderPreview(null);
  }

  useEffect(() => { setBuild(null); }, [pagination, templates, layout, templateDiscrepancyConfirmation]);
  useEffect(() => {
    if (!retainedBuildInputs) return;
    setReorderReturn((current) => dropStaleRetainedBuild(current, retainedBuildInputs, { readyToBuild: retainReadyToBuild }));
  }, [retainedBuildInputs, retainReadyToBuild]);

  async function chooseTemplate(slot: TemplateSlot, event: ChangeEvent<HTMLInputElement>) {
    const input = event.target;
    const file = input.files?.[0];
    if (!file) return;
    const extension = file.name.split(".").pop()?.toLowerCase() ?? "";
    if (!["pdf", "docx"].includes(extension)) {
      input.value = "";
      setError("Templates must be PDF or Word files (.pdf or .docx).");
      return;
    }
    setBusy("template");
    setError(null);
    try {
      const pdfFile = extension === "pdf" ? file : await convertWordTemplate(file);
      if (slot === "index") {
        const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
        const task = pdfjs.getDocument({ data: new Uint8Array(await pdfFile.arrayBuffer()), isEvalSupported: false, useWorkerFetch: false, verbosity: 0 });
        const document = await task.promise;
        const pages = document.numPages;
        await document.destroy();
        if (pages !== 1) {
          throw new Error("The index template must have exactly one page; it is repeated as needed for a long index.");
        }
      }
      const digest = await crypto.subtle.digest("SHA-256", await file.arrayBuffer());
      const sha256 = Array.from(new Uint8Array(digest)).map((part) => part.toString(16).padStart(2, "0")).join("");
      const pdfDigest = await crypto.subtle.digest("SHA-256", await pdfFile.arrayBuffer());
      const pdfSha256 = Array.from(new Uint8Array(pdfDigest)).map((part) => part.toString(16).padStart(2, "0")).join("");
      const matterReview = await reviewTemplateMatterPdf(pdfFile, file.name);
      setTemplates((current) => [...current.filter((template) => template.slot !== slot), { slot, file, sha256, sourceFormat: extension as "pdf" | "docx" | "doc", pdfFile, pdfSha256, reviewState: { matterReview } }]);
      setTemplateDiscrepancyConfirmation(null);
      setResolutions((current) => current.filter((resolution) => !(resolution.blockerId === "template-approval" && resolution.templateSlots?.includes(slot))));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The template could not be converted locally.");
    } finally {
      input.value = "";
      setBusy(null);
    }
  }

  function confirmTemplateReview(slot: TemplateSlot, kind: "appearanceConfirmation" | "matterConfirmation" | "placeholderConfirmation") {
    const reviewedAt = new Date().toISOString();
    const allowAmend = !(slot === "cover" && !coverWritesMatterText(layout));
    setTemplates((current) => current.map((template) => {
      if (template.slot !== slot || !template.pdfSha256) return template;
      const parsed = kind === "matterConfirmation" && allowAmend && matterDraft ? parseMatterDraft(matterDraft) : null;
      const confirmation = kind === "matterConfirmation"
        ? { pdfSha256: template.pdfSha256, confirmedAt: reviewedAt, ...(parsed ? { ...parsed.values, patches: parsed.patches } : {}) }
        : { pdfSha256: template.pdfSha256, confirmedAt: reviewedAt };
      return { ...template, reviewState: { ...template.reviewState, [kind]: confirmation } };
    }));
    setBuild(null);
  }

  function updateMatterOccurrence(findingId: string, value: string) {
    if (!matterDraft) return;
    const next: MatterDraft = { occurrences: matterDraft.occurrences.map((item) => item.findingId === findingId ? { ...item, value } : item) };
    setMatterDraft(next);
    const parsed = parseMatterDraft(next);
    setTemplates((templates) => templates.map((template) => {
      if (template.slot !== templatePreview?.slot) return template;
      const confirmed = template.reviewState?.matterConfirmation;
      if (!confirmed || confirmed.pdfSha256 !== template.pdfSha256) return template;
      const confirmedValues = matterValuesFromConfirmation(confirmed);
      if (confirmedValues && matterValuesEqual(parsed.values, confirmedValues)) return template;
      const { matterConfirmation: _removed, ...reviewState } = template.reviewState ?? {};
      return { ...template, reviewState };
    }));
  }

  async function previewTemplate(template: TemplateFile, origin?: HTMLElement | null) {
    templatePreviewOrigin.current = origin ?? (document.activeElement instanceof HTMLElement ? document.activeElement : null);
    setBusy("template");
    setError(null);
    setTemplatePreviewLoaded(false);
    try {
      const pdfFile = template.pdfFile ?? await convertWordTemplate(template.file);
      const pdfSha256 = template.pdfSha256 ?? await fileSha256(pdfFile);
      const matterReview = template.reviewState?.matterReview?.pdfSha256 === pdfSha256
        ? template.reviewState.matterReview
        : await reviewTemplateMatterPdf(pdfFile, template.file.name);
      if (!template.pdfFile || template.pdfSha256 !== pdfSha256 || template.reviewState?.matterReview?.pdfSha256 !== pdfSha256) {
        setTemplates((current) => current.map((item) => item.slot === template.slot ? { ...item, pdfFile, pdfSha256, reviewState: { matterReview } } : item));
        setTemplateDiscrepancyConfirmation(null);
      }
      setTemplatePreview({ slot: template.slot, name: template.file.name, file: pdfFile });
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The converted template could not be previewed.");
    } finally {
      setBusy(null);
    }
  }

  async function saveProject() {
    if (!analysis) return;
    try {
      const statementSnapshots = await verifiedStatementSnapshots(analysis, statements);
      const statementHashes = new Map(analysis.statementSources.map((source) => [source.statementId, source.sha256]));
      const statementSources = statementSnapshots.map((statement) => ({
        id: statement.id,
        role: "statement" as const,
        name: statement.file.name,
        sha256: statementHashes.get(statement.id)!,
        file: statement.file,
      }));
      const templateReviews: StoredTemplateReview[] = templates.flatMap((template) => template.pdfSha256 && template.reviewState ? [{
        slot: template.slot,
        sourceId: `template-${template.slot}`,
        renderedSourceId: template.sourceFormat === "pdf" ? undefined : `template-rendered-${template.slot}`,
        sourceFormat: template.sourceFormat ?? "pdf",
        sourceSha256: template.sha256,
        pdfSha256: template.pdfSha256,
        reviewState: template.reviewState,
      }] : []);
      const bytes = await createProjectArchive({
        schemaVersion: 8, name: projectName, createdAt: analysis.generatedAt, updatedAt: new Date().toISOString(), profileId: profile.id, pagination, layout, resolutions, arrangement, lastBuildSnapshot, pageSizeChoices, templateReviews, templateDiscrepancyConfirmation: templateDiscrepancyPending ? undefined : templateDiscrepancyConfirmation ?? undefined,
        witnessSettings: Object.fromEntries(statements.map((statement) => [statement.id, { initials: statement.witnessInitials, nextNumber: 1 }])),
        candidates, sheetSelections: analysis.evidence.flatMap((record) => [
          ...(record.extension !== "xlsx" || record.derivedFromEmail ? [] : (record.sheetSelections ?? []).flatMap((selection) => { const sheet = record.workbook?.sheets.find((item) => item.name === selection.name); return sheet ? [{ sourceSha256: record.sha256, sheetId: sheet.id, sheetPath: sheet.path, sheetName: sheet.name, included: selection.included, range: sheet.renderPlan.range, renderPlanHash: sheet.renderPlan.planHash }] : []; })),
          ...(record.emailAttachments ?? []).flatMap((child) => child.extension !== "xlsx" ? [] : (child.sheetSelections ?? []).flatMap((selection) => { const sheet = child.workbook?.sheets.find((item) => item.name === selection.name); return sheet ? [{ sourceSha256: child.sha256, sheetId: sheet.id, sheetPath: sheet.path, sheetName: sheet.name, included: selection.included, range: sheet.renderPlan.range, renderPlanHash: sheet.renderPlan.planHash }] : []; })),
        ]), analysis: { statements: statements.map(({ id, witnessName, witnessInitials, file }) => ({ id, witnessName, witnessInitials, name: file.name })), templates: templates.map((template) => template.slot) },
      }, [
        ...statementSources,
        ...analysis.evidence.filter((record) => !record.derivedFromEmail).map((record) => ({ id: record.id, role: "evidence" as const, name: record.name, sha256: record.sha256, file: record.file })),
        ...templates.map((template) => ({ id: `template-${template.slot}`, role: "template" as const, name: template.file.name, sha256: template.sha256, file: template.file })),
        ...templates.flatMap((template) => template.sourceFormat !== "pdf" && template.pdfFile && template.pdfSha256 ? [{ id: `template-rendered-${template.slot}`, role: "template-rendered" as const, name: `${template.file.name.replace(/\.[^.]+$/, "")}.rendered.pdf`, sha256: template.pdfSha256, file: template.pdfFile }] : []),
      ]);
      const saved = await downloadBytes(bytes, `${projectName.replace(/[^a-z0-9]+/gi, "_") || "bundle"}.bundle-project`, "application/zip");
      if (saved?.saved && saved.filePath && window.bundleBuilderDesktop && recoveryId.current) {
        const digest = await crypto.subtle.digest("SHA-256", new Uint8Array(bytes).buffer);
        const sha256 = Array.from(new Uint8Array(digest)).map((part) => part.toString(16).padStart(2, "0")).join("");
        const cleaned = await window.bundleBuilderDesktop.markRecoveryClean(recoveryId.current, recoveryRevision.current, { path: saved.filePath, sha256 });
        if (cleaned.revision) recoveryRevision.current = cleaned.revision;
      }
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Exhibit project could not be saved.");
    }
  }

  async function applyOpenedProject(opened: Awaited<ReturnType<typeof openProjectArchive>>, _legacyTemplateApprovals?: Map<string, boolean>) {
      const savedAnalysis = opened.snapshot.analysis as { statements?: Array<{ id: string; witnessName: string; witnessInitials: string; name: string }>; templates?: TemplateSlot[] };
      const restoredStatements: BundleStatementInput[] = (savedAnalysis.statements ?? []).flatMap((saved) => {
        const source = opened.sources.find((item) => item.id === saved.id);
        return source ? [{ id: saved.id, witnessName: saved.witnessName, witnessInitials: saved.witnessInitials, file: source.file }] : [];
      });
      const restoredEvidence = opened.sources.filter((source) => source.role === "evidence").map((source) => source.file);
      const restoredTemplates = restoreProjectTemplates(opened.snapshot, opened.sources);
      if (!restoredStatements.length || !restoredEvidence.length) throw new Error("The project does not contain its statements and evidence.");
      const result = await analyseBundleStatements(restoredStatements, restoredEvidence);
      const savedCandidates = opened.snapshot.candidates as ExhibitCandidate[];
      const savedSelections = opened.snapshot.sheetSelections ?? [];
      result.evidence = result.evidence.map((record) => {
        const next = record.extension !== "xlsx" ? record : { ...record, sheetSelections: (record.sheetSelections ?? []).map((selection) => { const sheet = record.workbook?.sheets.find((item) => item.name === selection.name); const saved = sheet ? savedSelections.find((item) => item.sourceSha256 === record.sha256 && item.sheetId === sheet.id && item.sheetPath === sheet.path) : undefined; return saved && saved.renderPlanHash === sheet?.renderPlan.planHash ? { ...selection, included: saved.included, range: saved.range } : selection; }) };
        if (!next.emailAttachments?.length) return next;
        return {
          ...next,
          emailAttachments: next.emailAttachments.map((child) => {
            if (child.extension !== "xlsx" || !child.workbook) return child;
            return {
              ...child,
              sheetSelections: (child.sheetSelections ?? []).map((selection) => {
                const sheet = child.workbook?.sheets.find((item) => item.name === selection.name);
                const saved = sheet ? savedSelections.find((item) => item.sourceSha256 === child.sha256 && item.sheetId === sheet.id && item.sheetPath === sheet.path) : undefined;
                return saved && saved.renderPlanHash === sheet?.renderPlan.planHash ? { ...selection, included: saved.included, range: saved.range } : selection;
              }),
            };
          }),
        };
      });
      const restoreSavedCandidate = (saved: ExhibitCandidate) => {
        if (saved.parentEmailProvenance) return { ...saved } as ExhibitCandidate;
        const savedSource = saved.evidenceId ? opened.sources.find((source) => source.id === saved.evidenceId && source.role === "evidence") : undefined;
        const currentEvidence = savedSource ? result.evidence.find((record) => record.sha256 === savedSource.sha256) : undefined;
        return { ...saved, evidenceId: currentEvidence?.id ?? null, confirmed: Boolean(currentEvidence && saved.confirmed) } as ExhibitCandidate;
      };
      const restoredCitedCandidates = result.candidates.map((candidate) => {
        const saved = savedCandidates.find((item) => item.id === candidate.id);
        if (!saved) return candidate;
        const restored = restoreSavedCandidate(saved);
        return restoreCitedCandidateDecision(candidate, saved, restored.evidenceId);
      });
      const restoredManualCandidates = savedCandidates.filter((candidate) => candidate.manualAddition && !restoredCitedCandidates.some((current) => current.id === candidate.id)).map(restoreSavedCandidate);
      const hydrated = await attachDerivedEmailEvidence(result, [...restoredCitedCandidates, ...restoredManualCandidates.filter((candidate) => candidate.parentEmailProvenance || candidate.evidenceId)]);
      setAnalysis(hydrated.analysis);
      setCandidates(hydrated.candidates);
      setStatements(restoredStatements.map((statement) => statement.id === result.statementId && result.witnessInitials ? { ...statement, witnessInitials: result.witnessInitials } : statement));
      setStatementFile(restoredStatements[0].file);
      setEvidenceFiles(restoredEvidence);
      setTemplates(restoredTemplates);
      setTemplateDiscrepancyConfirmation(opened.snapshot.templateDiscrepancyConfirmation ?? null);
      setResolutions(opened.snapshot.resolutions ?? []);
      setPagination(lockPagination(opened.snapshot.pagination ?? {}));
      setPaginationDraft(lockPagination(opened.snapshot.pagination ?? {}));
      setLayout({ ...DEFAULT_BUNDLE_LAYOUT, ...(opened.snapshot.layout ?? {}) });
      setPageSizeChoices(opened.snapshot.pageSizeChoices ?? {});
      setArrangement(opened.snapshot.arrangement ?? bundleArrangementFromLegacyOrder(opened.snapshot.finalOrder));
      setOrderHistory([]);
      setOrderPreview(null);
      setNewSectionHeading("");
      setArrangementStatus("");
      setLastBuildSnapshot((opened.snapshot.lastBuildSnapshot as SubstantiveBuildSnapshot | undefined) ?? null);
      setRebuildComparison(null);
      setReorderReturn(null);
      setProjectName(opened.snapshot.name);
      setBuild(null);
      setView("review");
  }

  async function openProject(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    await dismissRecoveryOffer();
    setBusy("analyse");
    setError(null);
    setBuild(null);
    try {
      const opened = await openProjectArchive(file);
      await applyOpenedProject(opened);
      const sourcePath = window.bundleBuilderDesktop?.sourcePath(file);
      if (sourcePath) recoverySourceDescriptors.current = [{ id: "saved-project-archive", role: "project", name: file.name, path: sourcePath, sha256: await fileSha256(file), size: file.size }];
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Exhibit project could not be opened.");
    } finally {
      setBusy(null);
      event.target.value = "";
    }
  }

  async function dismissRecoveryOffer() {
    const offer = recoveryOffer;
    setRecoveryOffer(null);
    if (!offer || !window.bundleBuilderDesktop) return;
    try {
      await window.bundleBuilderDesktop.discardRecovery(offer.recoveryId);
      const begun = await window.bundleBuilderDesktop.beginRecovery();
      recoveryId.current = begun.recoveryId;
      recoveryRevision.current = begun.revision;
      recoverySourceDescriptors.current = [];
      setRecoveryIssues([]);
      setRecoveryDataStored(false);
    } catch {
      // Starting a new analysis or project must not wait on recovery-journal cleanup.
    }
  }

  async function discardRecovery() {
    await dismissRecoveryOffer();
  }

  async function openRecoveryDataDialog() {
    const desktop = window.bundleBuilderDesktop;
    if (!desktop) return;
    setRecoveryDeleteAcknowledged(false);
    setRecoveryDataDialogOpen(true);
    try {
      const status = await desktop.recoveryStatus();
      setRecoveryDataStored(status.stored);
    } catch {
      setError("Exhibit Builder could not check the local recovery data.");
    }
  }

  async function clearLocalRecoveryData() {
    const desktop = window.bundleBuilderDesktop;
    if (!desktop || !recoveryDeleteAcknowledged) return;
    try {
      await desktop.clearRecoveryData();
      const begun = await desktop.beginRecovery();
      recoveryId.current = begun.recoveryId;
      recoveryRevision.current = begun.revision;
      recoverySourceDescriptors.current = [];
      setRecoveryOffer(null);
      setRecoveryIssues([]);
      setRecoveryDataStored(false);
      setRecoveryDataDialogOpen(false);
      setRecoveryDeleteAcknowledged(false);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The local recovery data could not be deleted.");
    }
  }

  async function restoreRecovery() {
    const desktop = window.bundleBuilderDesktop;
    if (!desktop || !recoveryOffer) return;
    setBusy("analyse");
    setError(null);
    recoveryRestoring.current = true;
    const issues: string[] = [];
    try {
      const journal = await desktop.loadRecovery(recoveryOffer.recoveryId);
      const payload = journal.payload as RecoveryProjectPayload;
      recoveryId.current = journal.recoveryId;
      recoveryRevision.current = journal.revision;
      recoverySourceDescriptors.current = payload.sources;
      const recovered = new Map<string, Awaited<ReturnType<typeof desktop.readRecoverySource>>>();
      for (const source of payload.sources) {
        try { recovered.set(source.id, await desktop.readRecoverySource(journal.recoveryId, source.id)); }
        catch (caught) { issues.push(caught instanceof Error ? caught.message : `${source.name} could not be restored.`); }
      }
      const projectSource = [...recovered.values()].find((source) => source.role === "project");
      if (projectSource) {
        const opened = await openProjectArchive(new File([projectSource.bytes as BlobPart], projectSource.name, { type: "application/zip" }));
        const recoveredProjectSources = [...recovered.values()]
          .filter((source): source is typeof source & { role: "statement" | "evidence" | "template" | "template-rendered" } => source.role !== "project")
          .map((source) => ({ id: source.id, role: source.role, name: source.name, sha256: source.sha256, file: new File([source.bytes as BlobPart], source.name) }));
        const merged = mergeRecoveryProjectDeltas(opened, payload, recoveredProjectSources);
        issues.push(...merged.issues);
        await applyOpenedProject(merged.opened, merged.templateApprovals);
      } else {
        const recoveredStatements: BundleStatementInput[] = (payload.statements ?? []).flatMap((saved) => {
          const source = recovered.get(saved.sourceId);
          return source ? [{ id: saved.id, witnessName: saved.witnessName, witnessInitials: saved.witnessInitials, file: new File([source.bytes as BlobPart], source.name) }] : [];
        });
        const evidenceSources = [...recovered.values()].filter((source) => source.role === "evidence");
        const recoveredEvidence = evidenceSources.map((source) => new File([source.bytes as BlobPart], source.name));
        if (!recoveredStatements.length || !recoveredEvidence.length) throw new Error("Recovery cannot continue because a witness statement or all evidence sources are unavailable.");
        const result = await analyseBundleStatements(recoveredStatements, recoveredEvidence);
        const decisions = payload.candidates ?? [];
        const currentStatementHashes = new Map(result.statementSources.map((source) => [source.statementId, source.sha256]));
        const recoveryStatementMatches = (saved: (typeof decisions)[number]) => {
          if (saved.manualAddition) return true;
          const descriptor = (payload.statements ?? []).find((statement) => statement.id === saved.statementId);
          return Boolean(saved.statementSha256
            && descriptor?.sourceSha256
            && saved.statementSha256 === descriptor.sourceSha256
            && currentStatementHashes.get(saved.statementId ?? "") === descriptor.sourceSha256);
        };
        const restoreRecoveryDecision = (saved: (typeof decisions)[number]) => {
          const statementMatches = recoveryStatementMatches(saved);
          if (saved.parentEmailProvenance) {
            const { id: _id, sourceSha256: _sourceSha256, statementSha256: _statementSha256, ...decision } = saved;
            return { ...decision, id: saved.id, confirmed: Boolean(decision.confirmed && statementMatches), repeatDecision: statementMatches ? decision.repeatDecision : undefined } as ExhibitCandidate;
          }
          const evidence = saved.sourceSha256 ? result.evidence.find((record) => record.sha256 === saved.sourceSha256) : undefined;
          const { id: _id, sourceSha256: _sourceSha256, statementSha256: _statementSha256, ...decision } = saved;
          return { ...decision, id: saved.id, evidenceId: evidence?.id ?? null, confirmed: Boolean(evidence && decision.confirmed && statementMatches), repeatDecision: statementMatches ? decision.repeatDecision : undefined } as ExhibitCandidate;
        };
        const restoredCandidates = result.candidates.map((candidate) => {
          const saved = decisions.find((decision) => decision.id === candidate.id);
          if (!saved) return candidate;
          const restored = restoreRecoveryDecision(saved);
          return restoreCitedCandidateDecision(candidate, saved as ExhibitCandidate, restored.evidenceId, recoveryStatementMatches(saved));
        });
        restoredCandidates.push(...decisions.filter((decision) => decision.manualAddition && !restoredCandidates.some((candidate) => candidate.id === decision.id)).map(restoreRecoveryDecision));
        const hydrated = await attachDerivedEmailEvidence(result, restoredCandidates);
        const recoveredTemplates: TemplateFile[] = (payload.templates ?? []).flatMap((saved) => {
          const source = recovered.get(saved.sourceId);
          if (!source) return [];
          const metadata = payload.templateReviews?.find((review) => review.slot === saved.slot && review.sourceId === saved.sourceId && review.sourceSha256 === source.sha256);
          const rendered = metadata?.renderedSourceId ? recovered.get(metadata.renderedSourceId) : saved.sourceFormat === "pdf" ? source : undefined;
          const exactReview = metadata && rendered?.sha256 === metadata.pdfSha256 && metadata.reviewState.matterReview?.pdfSha256 === metadata.pdfSha256 ? metadata.reviewState : undefined;
          return [{ slot: saved.slot, file: new File([source.bytes as BlobPart], source.name), sha256: source.sha256, sourceFormat: saved.sourceFormat, pdfFile: rendered ? new File([rendered.bytes as BlobPart], rendered.name, { type: "application/pdf" }) : undefined, pdfSha256: rendered?.sha256, reviewState: exactReview }];
        });
        const validHashes = new Set(hydrated.analysis.evidence.map((record) => record.sha256));
        setAnalysis(hydrated.analysis);
        setCandidates(hydrated.candidates);
        setStatements(recoveredStatements.map((statement) => statement.id === result.statementId && result.witnessInitials ? { ...statement, witnessInitials: result.witnessInitials } : statement));
        setStatementFile(recoveredStatements[0].file);
        setEvidenceFiles(recoveredEvidence);
        setTemplates(recoveredTemplates);
        setTemplateDiscrepancyConfirmation(payload.templateDiscrepancyConfirmation ?? null);
        setResolutions((payload.resolutions ?? []).filter((resolution) => !resolution.sourceSha256 || validHashes.has(resolution.sourceSha256)));
        setPagination(lockPagination(payload.pagination ?? {}));
        setPaginationDraft(lockPagination(payload.pagination ?? {}));
        setLayout({ ...DEFAULT_BUNDLE_LAYOUT, ...(payload.layout ?? {}) });
        setPageSizeChoices(payload.pageSizeChoices ?? {});
        setArrangement(payload.arrangement ?? bundleArrangementFromLegacyOrder(payload.finalOrder));
        setOrderHistory([]);
        setOrderPreview(null);
        setNewSectionHeading("");
        setArrangementStatus("");
        setProjectName(payload.project?.name ?? "Recovered exhibit project");
        setBuild(null);
        setLastBuildSnapshot(null);
        setRebuildComparison(null);
        setReorderReturn(null);
        setView("review");
      }
      setRecoveryOffer(null);
      setRecoveryIssues(issues);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The recovery journal could not be restored.");
      setRecoveryIssues(issues);
    } finally {
      recoveryRestoring.current = false;
      setBusy(null);
    }
  }

  function templateControl(slot: TemplateSlot, label: string, emptyLabel: string) {
    const template = templates.find((item) => item.slot === slot);
    const isWord = Boolean(template && template.sourceFormat !== "pdf");
    const exactHash = template?.pdfSha256;
    const appearanceConfirmed = !isWord || Boolean(exactHash && template?.reviewState?.appearanceConfirmation?.pdfSha256 === exactHash);
    const matterConfirmed = Boolean(exactHash && template?.reviewState?.matterConfirmation?.pdfSha256 === exactHash);
    const placeholdersConfirmed = !template?.reviewState?.matterReview?.placeholders.length || Boolean(exactHash && template.reviewState.placeholderConfirmation?.pdfSha256 === exactHash);
    const inputId = `template-file-${slot}`;
    return <div className="template-control">
      <label className="template-picker" htmlFor={inputId}><span>{label}</span>{template ? <strong className="template-selected-name">{template.file.name}</strong> : <small>{emptyLabel}</small>}</label>
      <input key={`${slot}-${template?.sha256 ?? "standard"}`} id={inputId} hidden={Boolean(template)} type="file" accept=".pdf,.docx" onChange={(event) => void chooseTemplate(slot, event)} />
      {template ? <div className="template-review-control">
        <div className="template-review-actions"><button className="secondary-button compact-action" type="button" onClick={() => document.getElementById(inputId)?.click()} disabled={busy !== null}>Change</button><button className="secondary-button compact-action" type="button" onClick={(event) => void previewTemplate(template, event.currentTarget)} disabled={busy !== null}>{isWord ? "Review converted PDF" : "Review this PDF"}</button><button className="quiet-text-button" type="button" onClick={() => { setTemplates((current) => current.filter((item) => item.slot !== slot)); setTemplateDiscrepancyConfirmation(null); setBuild(null); if (slot === "cover") setLayout((current) => ({ ...current, coverInsertion: "fit-a4", exactCoverPageNumber: false, exactCoverVolumeLabel: false })); }}>Use standard design</button></div>
        <ul className="template-review-status" aria-label={`Review status for ${template.file.name}`}>
          {isWord ? <li className={appearanceConfirmed ? "confirmed" : "pending"}>{appearanceConfirmed ? "Converted appearance confirmed" : "Converted appearance needs confirmation"}</li> : null}
          <li className={matterConfirmed ? "confirmed" : "pending"}>{slot === "cover" && !coverWritesMatterText(layout) ? (matterConfirmed ? "Cover as supplied confirmed" : "This cover needs confirmation as shown") : (matterConfirmed ? "Matter details and party names confirmed" : "Matter details and party names need confirmation")}</li>
          {template.reviewState?.matterReview?.placeholders.length ? <li className={placeholdersConfirmed ? "confirmed" : "pending"}>{placeholdersConfirmed ? "Visible placeholders confirmed" : "Visible placeholders need confirmation"}</li> : null}
        </ul>
        {slot === "cover" ? <fieldset className="cover-insertion">
          <legend>How should this cover be used?</legend>
          <label><input type="radio" name="cover-insertion" checked={layout.coverInsertion !== "exact"} onChange={() => setLayout((current) => ({ ...current, coverInsertion: "fit-a4" }))} /><span><strong>Finish from this template</strong><small>Recommended default. The layout stays. You can correct a misread name or case number, and those corrections are printed on the finished cover. The cover is fitted to A4 and receives the ordinary bundle page number. A split bundle also adds a volume label.</small></span></label>
          <label><input type="radio" name="cover-insertion" checked={layout.coverInsertion === "exact"} onChange={() => setLayout((current) => ({ ...current, coverInsertion: "exact", exactCoverPageNumber: false, exactCoverVolumeLabel: false }))} /><span><strong>Use this cover as supplied</strong><small>The page is used as shown. Names are not rewritten. It is still fitted to A4. A page number or volume label is added only if you tick that option.</small></span></label>
          {layout.coverInsertion === "exact" ? <div className="cover-insertion-extras">
            <p>Marks on a supplied cover stay off unless you choose them here.</p>
            <label className="checkbox-field"><input type="checkbox" checked={layout.exactCoverPageNumber} onChange={(event) => setLayout((current) => ({ ...current, exactCoverPageNumber: event.target.checked }))} /><span>Print a page number on this cover</span></label>
            <label className="checkbox-field"><input type="checkbox" checked={layout.exactCoverVolumeLabel} onChange={(event) => setLayout((current) => ({ ...current, exactCoverVolumeLabel: event.target.checked }))} /><span>Add a volume label if the bundle is split</span></label>
            {layout.volumePageLimit > 0 && !layout.exactCoverVolumeLabel ? <p className="settings-note">The bundle may be split, but this cover will not be marked Volume 1 of 3 unless you tick the option above.</p> : null}
          </div> : null}
        </fieldset> : null}
      </div> : slot === "cover" ? <div className="built-in-matter">
        <p>These details appear on the standard cover and index. They do not change the witness statement or any exhibit.</p>
        <label>Matter or case numbers<textarea value={(layout.builtInMatter?.matterNumbers ?? []).join("\n")} onChange={(event) => setLayout((current) => ({ ...current, builtInMatter: { ...(current.builtInMatter ?? { matterNumbers: [], partyNames: [], forums: [], matterTitles: [] }), matterNumbers: event.target.value.split("\n") } }))} rows={2} /></label>
        <label>Party names<textarea value={(layout.builtInMatter?.partyNames ?? []).join("\n")} onChange={(event) => setLayout((current) => ({ ...current, builtInMatter: { ...(current.builtInMatter ?? { matterNumbers: [], partyNames: [], forums: [], matterTitles: [] }), partyNames: event.target.value.split("\n") } }))} rows={3} /></label>
        <label>Forum or tribunal<textarea value={(layout.builtInMatter?.forums ?? []).join("\n")} onChange={(event) => setLayout((current) => ({ ...current, builtInMatter: { ...(current.builtInMatter ?? { matterNumbers: [], partyNames: [], forums: [], matterTitles: [] }), forums: event.target.value.split("\n") } }))} rows={2} /></label>
        <label>Matter title<textarea value={(layout.builtInMatter?.matterTitles ?? []).join("\n")} onChange={(event) => setLayout((current) => ({ ...current, builtInMatter: { ...(current.builtInMatter ?? { matterNumbers: [], partyNames: [], forums: [], matterTitles: [] }), matterTitles: event.target.value.split("\n") } }))} rows={2} /></label>
      </div> : null}
    </div>;
  }

  function renderManualExhibitPanel() {
    if (!manualAddOpen) return null;
    return <section className="manual-exhibit-panel" aria-labelledby="manual-exhibit-title">
      <div><h3 ref={manualPanelHeading} tabIndex={-1} id="manual-exhibit-title">Add an exhibit</h3><p>Choose an unused supplied document, or another local file.</p></div>
      <div className="manual-exhibit-fields">
        <label>Document<select value={manualEvidenceId} onChange={(event) => {
          const value = event.target.value;
          if (value === "__choose_local__") {
            event.target.value = manualEvidenceId;
            manualEvidenceInput.current?.click();
            return;
          }
          prepareManualEvidence(value);
        }}><option value="">Choose a document</option>{manualUploadedEvidence ? <option value={manualUploadedEvidence.id}>{manualUploadedEvidence.name} (new local file)</option> : null}{unreferencedEvidence.map((record) => <option key={record.id} value={record.id}>{record.name}</option>)}<option value="__choose_local__">Choose another local file</option></select></label>
        <input ref={manualEvidenceInput} className="visually-hidden" type="file" accept=".pdf,.docx,.eml,.txt,.xlsx" onChange={(event) => void uploadManualEvidence(event)} />
        <label>Index description<input value={manualDescription} onChange={(event) => setManualDescription(event.target.value)} placeholder="Description shown in the bundle index" /></label>
        <label>Document date<input value={manualDate} onChange={(event) => setManualDate(event.target.value)} placeholder="For example, 8 August 2026" /></label>
      </div>
      <p className="manual-exhibit-note">This exhibit is not cited in the statement. No statement reference is invented. It is listed at the end of the statement reference suggestions as an uncited exhibit.</p>
      <div className="download-row"><button className="primary-button" type="button" onClick={addManualExhibit} disabled={!manualEvidenceId || !manualDescription.trim() || busy !== null}>Add exhibit</button><button className="secondary-button" type="button" onClick={closeManualAdd}>Cancel</button></div>
    </section>;
  }

  function renderArrangementExhibit(exhibitId: string, sectionId: string | null, index: number, containerLength: number) {
    const group = exhibitGroupLookupByKind.byGroupId.get(exhibitId);
    if (!group) return null;
    const orderNumber = displayedOrderNumbers.get(exhibitId) ?? 0;
    const glued = Boolean(group.canonical.parentEmailProvenance);
    return <li key={exhibitId} data-exhibit-id={exhibitId} className={`finalise-order-item ${draggingExhibitId === exhibitId ? "dragging" : ""}`} draggable={!orderPreview && !draggingSectionId && !glued} onDragStart={() => { if (glued || draggingSectionId) return; setDraggingExhibitId(exhibitId); }} onDragEnd={() => setDraggingExhibitId(null)} onDragOver={(event) => { if (!orderPreview && !draggingSectionId && !glued) event.preventDefault(); }} onDrop={() => { if (glued || draggingSectionId) return; if (draggingExhibitId && !orderPreview) moveCandidateBefore(draggingExhibitId, exhibitId); setDraggingExhibitId(null); }}>
      <span className="drag-handle" aria-hidden="true">{glued ? "" : "::"}</span>
      <strong aria-label={`Exhibit position ${orderNumber}`}>{orderNumber}</strong>
      <div className="finalise-order-copy">
        {group.canonical.manualAddition ? <span className="manual-exhibit-badge">Added manually - not cited in statement</span> : null}
        {group.canonical.manualAddition && !orderPreview ? <div className="manual-row-editor"><label>Description<input value={group.canonical.description} onChange={(event) => updateManualCandidate(group.canonical.id, { description: event.target.value })} /></label><label>Date<input value={group.canonical.date} onChange={(event) => updateManualCandidate(group.canonical.id, { date: event.target.value })} /></label><label>Document<select value={group.canonical.evidenceId ?? ""} disabled={glued || Boolean(group.evidence.derivedFromEmail)} onChange={(event) => changeManualEvidence(group.canonical.id, event.target.value)}><option value={group.evidence.id}>{group.evidence.name}</option>{unreferencedEvidence.filter((record) => record.id !== group.evidence.id).map((record) => <option key={record.id} value={record.id}>{record.name}</option>)}</select></label></div> : <><b>{group.canonical.description}</b><small>{group.evidence.name}{group.references.length > 1 ? ` - cited at ${group.references.length} places` : ""}</small></>}
      </div>
      <div className="finalise-order-actions">
        {glued ? null : <button data-action="top" aria-label={`Move ${group.canonical.description} to the top`} type="button" disabled={Boolean(orderPreview) || index === 0} onClick={() => moveCandidateToEdge(exhibitId, "top")}>Top</button>}
        <button data-action="earlier" aria-label={`Move ${group.canonical.description} earlier`} type="button" disabled={Boolean(orderPreview) || index === 0} onClick={() => moveCandidate(exhibitId, -1)}>Earlier</button>
        <button data-action="later" aria-label={`Move ${group.canonical.description} later`} type="button" disabled={Boolean(orderPreview) || index === containerLength - 1} onClick={() => moveCandidate(exhibitId, 1)}>Later</button>
        {glued ? null : <button data-action="bottom" aria-label={`Move ${group.canonical.description} to the bottom`} type="button" disabled={Boolean(orderPreview) || index === containerLength - 1} onClick={() => moveCandidateToEdge(exhibitId, "bottom")}>Bottom</button>}
        {!glued && !orderPreview && hasIndexHeadings ? <label className="move-section-control"><span>Index heading</span><select data-action="move-section" aria-label={`Index heading for ${group.canonical.description}`} value={sectionId ?? ""} onChange={(event) => moveCandidateToSection(exhibitId, event.target.value || null)}><option value="">No heading</option>{arrangement.nodes.filter((node): node is ArrangementSectionNode => node.type === "section").map((section) => <option key={section.id} value={section.id}>{section.heading}</option>)}</select></label> : null}
        {group.canonical.manualAddition && !orderPreview ? <button className="remove-manual-exhibit" type="button" onClick={() => removeManualCandidate(group.canonical.id)}>Remove</button> : null}
      </div>
    </li>;
  }

  function renderArrangementNodes() {
    const items: Array<ReturnType<typeof renderArrangementExhibit>> = [];
    let run: ArrangementExhibitNode[] = [];
    const showUnheadedGroups = displayedArrangement.nodes.some((node) => node.type === "section");
    const flushRun = (key: string) => {
      if (!run.length) return;
      if (!showUnheadedGroups) {
        run.forEach((exhibit) => {
          const unheaded = exhibitContainerLocation(displayedArrangement, exhibit.exhibitId);
          items.push(renderArrangementExhibit(exhibit.exhibitId, null, unheaded?.index ?? 0, unheaded?.length ?? 1));
        });
      } else {
        const firstExhibitId = run[0].exhibitId;
        const beforeNodeIndex = displayedArrangement.nodes.findIndex((item) => item.type === "exhibit" && item.exhibitId === firstExhibitId);
        items.push(<li className={`arrangement-section${draggingSectionId && headingDropTargetKey === key ? " heading-drop-target" : ""}`} key={key} onDragOver={(event) => { if (!orderPreview && draggingSectionId) { event.preventDefault(); setHeadingDropTargetKey((current) => current === key ? current : key); } }} onDragLeave={(event) => { const next = event.relatedTarget; if (next instanceof Node && event.currentTarget.contains(next)) return; setHeadingDropTargetKey((current) => current === key ? null : current); }} onDrop={(event) => { event.preventDefault(); if (draggingSectionId && !orderPreview && beforeNodeIndex >= 0) dropSectionBefore(draggingSectionId, beforeNodeIndex); setDraggingSectionId(null); setHeadingDropTargetKey(null); }}>
          <h4 className="unheaded-run-label"><strong>No heading</strong><small>These exhibits still print in the main index list. Earlier and Later reorder this group only.</small></h4>
          <ol className={`finalise-order-list section-exhibit-list${draggingSectionId ? " heading-drop-passthrough" : ""}`}>{run.map((exhibit, exhibitIndex) => renderArrangementExhibit(exhibit.exhibitId, null, exhibitIndex, run.length))}</ol>
        </li>);
      }
      run = [];
    };
    displayedArrangement.nodes.forEach((node, nodeIndex) => {
      if (node.type === "section") {
        flushRun(`unheaded-before-${node.id}`);
        items.push(<li className={`arrangement-section ${draggingSectionId === node.id ? "dragging" : ""}${collapsedSectionIds.has(node.id) ? "" : " is-expanded"}${draggingSectionId && headingDropTargetKey === node.id ? " heading-drop-target" : ""}`} data-section-id={node.id} key={node.id} onDragOver={(event) => { if (!orderPreview && draggingSectionId && draggingSectionId !== node.id) { event.preventDefault(); setHeadingDropTargetKey((current) => current === node.id ? current : node.id); } }} onDragLeave={(event) => { const next = event.relatedTarget; if (next instanceof Node && event.currentTarget.contains(next)) return; setHeadingDropTargetKey((current) => current === node.id ? null : current); }} onDrop={(event) => { event.preventDefault(); if (draggingSectionId && draggingSectionId !== node.id && !orderPreview) dropSectionBefore(draggingSectionId, nodeIndex); setDraggingSectionId(null); setHeadingDropTargetKey(null); }}>
          <h4 id={sectionHeadingDomId(node.id)} className="visually-hidden">{node.heading}, {node.exhibits.length} exhibit{node.exhibits.length === 1 ? "" : "s"}</h4>
          <div className="section-heading-row">
            <span className="drag-handle" draggable={!orderPreview} data-action="drag-section" aria-hidden="true" onClick={(event) => { event.preventDefault(); event.stopPropagation(); }} onDragStart={(event) => { event.stopPropagation(); event.dataTransfer.effectAllowed = "move"; event.dataTransfer.setData("text/plain", node.id); setDraggingSectionId(node.id); }} onDragEnd={() => { setDraggingSectionId(null); setHeadingDropTargetKey(null); }}>::</span>
            <button type="button" id={`${sectionHeadingDomId(node.id)}-summary`} className="section-toggle" aria-labelledby={sectionHeadingDomId(node.id)} aria-expanded={!collapsedSectionIds.has(node.id)} onClick={() => toggleSectionCollapsed(node.id)}><span aria-hidden="true">{node.heading}</span><small aria-hidden="true">{node.exhibits.length} exhibit{node.exhibits.length === 1 ? "" : "s"}</small></button>
            {!orderPreview ? <div className="section-actions" aria-label={`Actions for ${node.heading}`}><button type="button" onClick={() => renameSection(node.id)}>Rename heading</button><button type="button" disabled={nodeIndex === 0} onClick={() => moveSection(node.id, -1)}>Move heading earlier</button><button type="button" disabled={nodeIndex === arrangement.nodes.length - 1} onClick={() => moveSection(node.id, 1)}>Move heading later</button><button className="remove-section" type="button" onClick={() => deleteSection(node.id)}>Delete heading</button></div> : null}
          </div>
          {collapsedSectionIds.has(node.id) ? null : node.exhibits.length ? <ol className={`finalise-order-list section-exhibit-list${draggingSectionId ? " heading-drop-passthrough" : ""}`}>{node.exhibits.map((exhibit, exhibitIndex) => renderArrangementExhibit(exhibit.exhibitId, node.id, exhibitIndex, node.exhibits.length))}</ol> : <p className="empty-section-note"><strong>This heading is empty.</strong> It is saved with the project but will not print in the index or appear as a PDF bookmark until it contains an exhibit.</p>}
        </li>);
        return;
      }
      run.push(node);
    });
    flushRun("unheaded-tail");
    return items;
  }

  function focusBlockerTarget(selector: string) {
    viewportProbe.current?.cancel();
    viewportProbe.current = probeSelectorUntilFound(
      (sel) => document.querySelector<HTMLElement>(sel),
      selector,
      (callback) => requestAnimationFrame(callback),
      (target) => {
        if (target instanceof HTMLDetailsElement) target.open = true;
        else {
          const details = target.closest("details");
          if (details) details.open = true;
        }
        const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
        const focusTarget = firstActionableControl(target) ?? target;
        focusTarget.scrollIntoView({ behavior: reducedMotion ? "auto" : "smooth", block: "center" });
        focusTarget.focus({ preventScroll: true });
      },
    );
  }

  function navigateToBlocker(blocker: BuildBlocker) {
    if (blocker.target === "sheets") {
      setView("sheets");
      focusBlockerTarget(blocker.sourceId ? `[data-source-id="${CSS.escape(blocker.sourceId)}"], [data-testid="sheet-readiness"]` : "[data-testid=\"sheet-readiness\"]");
      return;
    }
    if (blocker.target === "finalise") {
      setView("build");
      focusBlockerTarget("[data-testid=\"order-preview-banner\"]");
      return;
    }
    setView("review");
    if (blocker.kind === "approval") setShowPendingOnly(true);
    if (blocker.candidateId) {
      const emailAttachment = isEmailAttachmentBlocker(blocker);
      const targetCandidateId = emailAttachment
        ? exhibitGroupLookupByKind.byCandidateId.get(blocker.candidateId)?.canonical.id ?? blocker.candidateId
        : blocker.candidateId;
      setShowDuplicatesOnly(false);
      if (emailAttachment) setShowPendingOnly(false);
      else setShowPendingOnly(true);
      setExpandedConfirmedCards((current) => new Set(current).add(targetCandidateId));
      if (emailAttachment) setOpenEmailAttachmentsId(targetCandidateId);
      focusBlockerTarget(emailAttachment
        ? emailAttachmentsSelector(targetCandidateId)
        : reviewCardSelector(targetCandidateId));
      return;
    }
    if (blocker.target === "templates") {
      focusBlockerTarget(templateDiscrepancyPending ? "#template-discrepancy-title" : "#optional-settings");
      return;
    }
    focusBlockerTarget(`[data-blocker-id="${CSS.escape(blocker.id)}"]`);
  }

  return (
    <main className="app-shell">
      <div className="visually-hidden" role="status" aria-live="polite" aria-atomic="true">{analysisProgressAnnouncement}</div>
      <div className="visually-hidden" role="status" aria-live="polite" aria-atomic="true">{reviewActionStatus}</div>
      <header className="topbar">
        <a className="brand" href="#" aria-label="Exhibit Builder">
          <span className="brand-mark" aria-hidden="true">
            EB
          </span>
          <span>
            <strong>Exhibit Builder</strong>
            <small>Offline exhibit bundles</small>
          </span>
        </a>
        <div className="privacy-pill">
          <span className="privacy-dot" aria-hidden="true" />
          Offline desktop processing
        </div>
        <input ref={projectInput} className="visually-hidden" type="file" accept=".bundle-project,.zip" onChange={openProject} />
        <button className="text-button" type="button" onClick={() => projectInput.current?.click()}>
          Open exhibit project
        </button>
        {!analysis && window.bundleBuilderDesktop ? <button className="text-button" type="button" onClick={() => void openRecoveryDataDialog()}>Local recovery data</button> : null}
        {!analysis && guidedSamplePreferenceReady && guidedSampleHidden ? <button className="text-button" type="button" onClick={() => void setGuidedSampleHidden(false)}>Show guided sample</button> : null}
        {analysis && (
          <>
            <button className="text-button" type="button" data-tour="save" onClick={() => { if (tourActive) setTourSaved(true); void saveProject(); }}>Save exhibit project</button>
          <button className="text-button" type="button" onClick={() => { if (!analysis || window.confirm("Start a new exhibit project? This will discard the current statement, exhibit matches, approvals and build settings.")) void resetWorkspace(); }}>New exhibit project</button>
          </>
        )}
      </header>

      {recoveryOffer ? <div className="confirmation-backdrop" role="presentation"><section className="confirmation-dialog" role="dialog" aria-modal="true" aria-labelledby="recovery-title"><p className="eyebrow">Crash recovery</p><h2 id="recovery-title">Restore {recoveryOffer.projectName}?</h2><p>Exhibit Builder found an unfinished local recovery journal. Sources will be reread and SHA-256 checked. Confirmations and exceptions tied to missing or changed files will not be restored.</p><div className="download-row"><button className="primary-button" type="button" onClick={() => void restoreRecovery()} disabled={busy !== null}>Restore</button><button className="secondary-button" type="button" onClick={() => void discardRecovery()} disabled={busy !== null}>Discard recovery</button></div></section></div> : null}
      {recoveryIssues.length ? <details className="recovery-issues" open><summary>Automatic recovery needs attention</summary><ul>{recoveryIssues.map((issue) => <li key={issue}>{issue}</li>)}</ul></details> : null}

      {!analysis ? (
        <section className="welcome">
          <div className="welcome-copy">
            <p className="eyebrow">From placeholder to checked bundle</p>
            <h1>Build an exhibit bundle without the document chase.</h1>
            <p className="welcome-lede">
              Read one final witness statement with exhibit placeholders, reconcile every cited exhibit, and create one indexed PDF with a clear audit trail.
            </p>
            {guidedSamplePreferenceReady && !guidedSampleHidden ? <div className="guided-sample-panel">
              {guidedSampleAvailability !== "unavailable" ? <>
                <div className="hero-actions">
                  <button
                    className="primary-button"
                    type="button"
                    onClick={() => {
                      if (tourActive) setTourOpenedFolder(true);
                      void openGuidedSampleFolder();
                    }}
                    disabled={busy !== null || guidedSampleAvailability !== "available"}
                    data-testid="guided-sample-button"
                    data-tour="open-folder"
                    aria-describedby="guided-sample-note"
                  >
                    {guidedSampleAvailability === "checking"
                      ? "Checking guided sample"
                      : "Open the guided sample folder"}
                  </button>
                  <button
                    className="secondary-button"
                    type="button"
                    onClick={startGuidedSampleTour}
                    disabled={busy !== null || guidedSampleAvailability !== "available" || tourActive}
                    data-testid="guided-sample-tour-button"
                  >
                    Run guided sample
                  </button>
                </div>
                <p className="hero-note" id="guided-sample-note">
                  Open the guided sample folder and read the witness statement.
                </p>
              </> : <div className="guided-sample-unavailable" role="status"><div><strong>The optional guided sample is unavailable.</strong><span>You can still build a bundle from your own files.</span></div><button className="secondary-button compact-action" type="button" onClick={() => void retryGuidedSampleAvailability()}>Check again</button></div>}
              <button className="quiet-text-button" type="button" onClick={() => void setGuidedSampleHidden(true)}>Hide the guided sample from this screen</button>
            </div> : null}
          </div>

          <div className="workflow-card" aria-label="Bundle workflow">
            <div className="workflow-card-head">
              <span>Guided workflow</span>
              <span className="status-ready">Ready</span>
            </div>
            <ol className="workflow-list">
              <li>
                <span>01</span>
                <div>
                  <strong>Read the statement</strong>
                  <p>Pick up every exhibit placeholder and keep the surrounding paragraph.</p>
                </div>
              </li>
              <li>
                <span>02</span>
                <div>
                  <strong>Reconcile the evidence</strong>
                  <p>Match files and hold ambiguity for human review.</p>
                </div>
              </li>
              <li>
                <span>03</span>
                <div>
                  <strong>Build and validate</strong>
                  <p>Create the index, pagination, links and report.</p>
                </div>
              </li>
            </ol>
          </div>

          <div className="upload-panel">
            <div className="section-heading">
              <div>
                <p className="eyebrow">Or use your own files</p>
                <h2>Start a private local analysis</h2>
              </div>
              <span className="supported">DOCX · PDF · EML · XLSX</span>
            </div>
            <div className="upload-grid">
              <button
                className={`drop-card ${statementFile ? "has-file" : ""}`}
                type="button"
                data-tour="choose-statement"
                onClick={() => statementInput.current?.click()}
              >
                <span className="file-glyph" aria-hidden="true">
                </span>
                <span>
                  <strong>
                    {statementFile
                      ? `${statements.length || 1} witness statement${(statements.length || 1) === 1 ? "" : "s"} selected`
                      : "Choose witness statement"}
                  </strong>
                  <small>One DOCX witness statement</small>
                </span>
              </button>
              <input
                ref={statementInput}
                className="visually-hidden"
                type="file"
                accept=".docx"
                tabIndex={-1}
                onChange={handleStatement}
                aria-label="Choose witness statement"
              />

              <button
                className={`drop-card ${evidenceFiles.length ? "has-file" : ""}`}
                type="button"
                data-tour="choose-evidence"
                onClick={() => evidenceInput.current?.click()}
              >
                <span className="file-glyph" aria-hidden="true">
                  +
                </span>
                <span>
                  <strong>
                    {evidenceFiles.length
                      ? `${evidenceFiles.length} evidence files selected`
                      : "Choose evidence files"}
                  </strong>
                  <small>PDF, DOCX, EML and XLSX files</small>
                </span>
              </button>
              <input
                ref={evidenceInput}
                className="visually-hidden"
                type="file"
                accept=".pdf,.docx,.eml,.xlsx"
                multiple
                tabIndex={-1}
                onChange={handleEvidence}
                aria-label="Choose evidence files"
              />
            </div>
            <div className="upload-footer">
              <div className="upload-notes">
              <p>
                Files remain on this device. Exhibit Builder does not upload or retain
                document content.
              </p>
              <p>
                Automatic matching and OCR are English-only in this release. Source PDF
                page artwork is retained, but multilingual matching, OCR and generated
                typesetting are not supported; check every non-English output manually.
              </p>
              </div>
              {analysisProgress ? <div className="analysis-progress"><strong>{analysisProgress.stage}</strong><span>{analysisProgress.detail}</span><button className="secondary-button compact-action" type="button" onClick={() => { analysisCancelRequested.current = true; setAnalysisProgress((current) => current ? { ...current, detail: "Stopping safely after the current document." } : current); }}>Stop analysis</button></div> : null}
              <button
                className="primary-button"
                type="button"
                data-tour="analyse"
                disabled={!statementFile || !evidenceFiles.length || busy !== null}
                onClick={() =>
                  statementFile &&
                  runAnalysis(statementFile, evidenceFiles)
                }
              >
                {busy === "analyse" ? "Analysing" : "Analyse files"}
              </button>
            </div>
          </div>
        </section>
      ) : (
        <section className="workspace">
          <aside className="workspace-nav">
            <p className="nav-label">Exhibit project</p>
            <h2>{analysis.caseTitle}</h2>
            <div className="matter-meta">
              <span>Statement</span>
              <strong>{statements.length} statement{statements.length === 1 ? "" : "s"}</strong>
            </div>
            <nav aria-label="Bundle workflow">
              <button
                className={view === "sources" ? "active" : ""}
                type="button"
                onClick={() => setView("sources")}
                aria-current={view === "sources" ? "step" : undefined}
              >
                <span>01</span>
                <span>
                  <strong>Sources</strong><small>{evidenceFiles.length} evidence files</small>
                </span>
              </button>
              <button
                className={view === "review" || view === "reconcile" ? "active" : ""}
                type="button"
                onClick={() => setView("review")}
                aria-current={view === "review" || view === "reconcile" ? "step" : undefined}
              >
                <span>02</span>
                <span>
                  <strong>Review</strong><small>{needsDecisionCount} decisions</small>
                </span>
              </button>
              <button
                className={view === "sheets" ? "active" : ""}
                type="button"
                onClick={() => {
                  if (readyToLeaveReview && includedWorkbookCount) setView("sheets");
                }}
                disabled={busy !== null}
                aria-disabled={!readyToLeaveReview || !includedWorkbookCount}
                aria-describedby="sheets-nav-status"
                aria-current={view === "sheets" ? "step" : undefined}
              >
                <span>03</span>
                <span>
                  <strong>Sheets</strong><small id="sheets-nav-status">{!readyToLeaveReview ? "Complete exhibit review first" : workbookSourceCount ? `${includedWorkbookCount} of ${workbookSourceCount} confirmed` : "No workbook files"}</small>
                </span>
              </button>
              <button className={view === "build" ? "active" : ""} type="button" onClick={() => setView("build")} aria-current={view === "build" ? "step" : undefined}><span>04</span><span><strong>Finalise</strong><small>{build ? "Bundle ready" : "Set order"}</small></span></button>
            </nav>
            <div className="local-note">
              <span className="privacy-dot" aria-hidden="true" />
              <div>
                <strong>Private by design</strong>
                <p>All analysis and PDF assembly run inside this desktop app.</p>
              </div>
            </div>
          </aside>

          <div className="workspace-main">
            {view === "sources" && (
              <>
                <div className="workspace-header"><div><p className="eyebrow">Stage 1 of 4</p><h1 tabIndex={-1}>Sources</h1><p>Statements are read-only. Evidence stays local and nothing is included until reviewed.</p></div><button className="primary-button" onClick={() => setView("review")}>Continue to exhibit review</button></div>
                <section className="validation-card"><div><p className="eyebrow">Evidence files</p><h2>{analysis.evidence.filter((record) => !record.derivedFromEmail).length} supplied file{analysis.evidence.filter((record) => !record.derivedFromEmail).length === 1 ? "" : "s"}</h2><p>These files are available for matching. Cited items enter the bundle after you confirm them. Unused files can be added as uncited exhibits.</p></div><div className="check-list">{analysis.evidence.filter((record) => !record.derivedFromEmail).map((record) => <div key={record.id}><span className="check-pass">{record.extension.toUpperCase()}</span><p><strong>{record.name}</strong><small>{record.extension === "xlsx" ? `${record.workbook?.sheets.length ?? 0} worksheet tab${(record.workbook?.sheets.length ?? 0) === 1 ? "" : "s"}; choose which tabs are printed at the Sheets stage` : record.pageCount ? `${record.pageCount} page${record.pageCount === 1 ? "" : "s"}` : "Text-based document"}</small></p></div>)}</div></section>
              </>
            )}
            {view === "sheets" && (
              <>
                <div className="workspace-header"><div><p className="eyebrow">Stage 3 of 4</p><h1 tabIndex={-1}>Choose which Excel sheets to include</h1><p>An Excel workbook can contain several worksheet tabs. Tick only the tabs that belong in this exhibit. Unticked tabs stay in the source workbook but are omitted from the bundle.</p></div><div className="workspace-header-actions"><button className="secondary-button" type="button" onClick={() => setView("review")}>Back to exhibit review</button><button className="primary-button" type="button" data-tour="continue-finalise" onClick={() => { if (readyToBuild) setView("build"); else if (buildBlockerList[0]) navigateToBlocker(buildBlockerList[0]); }} disabled={busy !== null} aria-disabled={!readyToBuild} aria-describedby="sheet-continue-status">Continue to finalise</button><small id="sheet-continue-status" role="status" aria-live="polite" aria-atomic="true">{readyToBuild ? "Ready to continue." : `${buildBlockerList.length} requirement${buildBlockerList.length === 1 ? "" : "s"} remaining.`}</small></div></div>
                {!readyToBuild && busy === null && <section className="build-readiness sheet-readiness" data-testid="sheet-readiness">
                  <div className="build-readiness-heading"><div><p className="eyebrow">Workbook readiness</p><h2>Complete sheet selection</h2></div><span className="readiness-count needs-attention">{buildBlockerList.length} item{buildBlockerList.length === 1 ? "" : "s"} to resolve</span></div>
                  <p className="readiness-clear">The exhibit approval is complete. Select at least one readable worksheet for each workbook, or return to exhibit review for any other outstanding requirement.</p>
                  <div className="readiness-list">{buildBlockerList.map((blocker) => <div key={blocker.id} className={`readiness-item ${blocker.kind}`}><span aria-hidden="true">!</span><div className="readiness-copy"><p className="readiness-label"><strong>{blocker.label}</strong></p>{blocker.fileName ? <p className="readiness-file">{blocker.fileName}</p> : null}<p className="readiness-detail">{blocker.detail}</p></div></div>)}</div>
                </section>}
                {confirmedWorkbookExhibits.length === 0 ? <section className="validation-card"><h2>No confirmed workbook exhibits are ready</h2><p>Return to exhibit review and confirm a workbook before choosing its sheets.</p></section> : confirmedWorkbookExhibits.map((record) => <section className="sheet-workbook" key={record.key} data-source-id={record.evidenceId}><h2>{record.name}</h2><p className="sheet-note">For every ticked worksheet, the cell range shown below is the part Microsoft Excel will print into A4 bundle pages. Saved formula results are used without recalculation, and the source workbook is never changed.</p>{record.workbook?.sheets.map((sheet) => { const selection = record.sheetSelections?.find((item) => item.name === sheet.name); const plan = sheet.renderPlan; const tile = plan.tiles[0]; const preview = sheet.cells.filter((cell) => cell.row >= tile.top && cell.row <= tile.bottom && cell.col >= tile.left && cell.col <= tile.right); return <article className="sheet-card" key={sheet.name}><div><label><input type="checkbox" checked={selection?.included ?? false} onChange={(event) => updateSheetSelection(record.evidenceId, sheet.name, event.target.checked, record.childIdentity)} /> Include <strong>{sheet.name}</strong> in the bundle</label><span className={sheet.state === "visible" ? "success-badge" : "held-badge"}>{sheet.state}</span><p>Selected cell range: {plan.range} · {plan.orientation} A4 · about {plan.predictedPageCount} page{plan.predictedPageCount === 1 ? "" : "s"} · {plan.scalePercent}% print scale</p>{workbookPlanCheckCopy(sheet.name, plan.warnings).map((copy) => <small key={copy.idSuffix} className="warning">{copy.label}. {copy.detail}</small>)}</div><div className="sheet-preview" role="img" aria-label={`Thumbnail of the first planned A4 page for ${sheet.name}`}><div aria-hidden="true">{preview.map((cell) => <span key={`${cell.row}-${cell.col}`} style={{ gridColumn: cell.col - tile.left + 1, gridRow: cell.row - tile.top + 1 }}>{cell.value.slice(0, 40)}</span>)}</div></div></article>; })}</section>)}
              </>
            )}
            {view === "review" && (
              <>
                <div className="workspace-header">
                  <div>
                    <p className="eyebrow">Exhibit review</p>
                    <h1 ref={reviewHeading} tabIndex={-1}>Check the proposed exhibits</h1>
                    <p>
                      Confirm that the document selected for each statement reference is correct.
                      Nothing enters the final PDF until you approve it.
                    </p>
                  </div>
                </div>

                <details id="optional-settings" tabIndex={-1} className="statement-safety-note" aria-label="Optional project and bundle settings">
                  <summary>Optional settings <small>Project name, page numbering, templates and exhibit initials</small></summary>
                  <div className="advanced-settings">
                    <p className="advanced-settings-note"><strong>Most projects can leave these settings unchanged.</strong> Witness statements and source documents always remain read-only.</p>
                    <section className="advanced-settings-section">
                      <h3>Exhibit project name</h3>
                      <div className="advanced-settings-grid"><label className="wide"><span className="visually-hidden">Exhibit project name</span><input value={projectName} onChange={(event) => setProjectName(event.target.value)} /><small>Used for the saved project, crash recovery and report filenames. It is not printed on the bundle cover and does not affect page numbering. Most projects can leave this unchanged.</small></label></div>
                    </section>
                    <section className="advanced-settings-section">
                      <h3>Page numbering</h3>
                      <p className="advanced-settings-note">These page numbers appear on the finished PDF and are also used in the index and suggested witness-statement references.</p>
                      <p className="advanced-settings-note"><strong>One sequence across the whole bundle.</strong> {suppliedCoverOmitsPageNumber ? "This cover stays unnumbered unless you tick “Print a page number on this cover”. The index continues the sequence, and exhibit pages follow without restarting. If the bundle is split, the sequence continues into the next volume." : "The cover is page 1, the index continues the sequence, and exhibit pages follow without restarting. If the bundle is split, the sequence continues into the next volume."} The finished PDF is the only page-number source: stamps, the index and suggested statement references always use those printed labels.</p>
                      <div className="advanced-settings-grid continuous-numbering-controls">
                        <label>Prefix on every page<input data-testid="page-number-prefix" value={paginationDraft.prefix} onChange={(event) => setPaginationDraft((current) => updatePaginationDraft(current, { prefix: event.target.value }))} placeholder="For example, AH-" /><small>{suppliedCoverOmitsPageNumber ? "Optional. Applied to numbered pages. This cover stays unnumbered unless you tick “Print a page number on this cover”." : "Optional. Applied to cover, index and exhibit pages."}</small></label>
                        <label>Suffix on every page<input data-testid="page-number-suffix" value={paginationDraft.suffix} onChange={(event) => setPaginationDraft((current) => updatePaginationDraft(current, { suffix: event.target.value }))} placeholder="Optional" /><small>Placed directly after every page number.</small></label>
                        <label>Minimum number of digits<input data-testid="page-number-padding" type="number" min="0" max="12" value={paginationDraft.padding} onChange={(event) => setPaginationDraft((current) => updatePaginationDraft(current, { padding: Math.max(0, Math.min(12, Number(event.target.value) || 0)) }))} /><small>0 gives AH-1. Four digits gives AH-0001.</small></label>
                      </div><div className="numbering-preview" aria-live="polite"><strong>Example across the finished bundle</strong><span>{(() => { const first = Number.isFinite(paginationDraft.startAt) ? Math.max(1, Math.floor(paginationDraft.startAt)) : 1; const label = (page: number) => `${paginationDraft.prefix}${paginationDraft.padding ? String(page).padStart(paginationDraft.padding, "0") : String(page)}${paginationDraft.suffix}`; return `${label(first)}, ${label(first + 1)}, ${label(first + 2)}…`; })()}</span><small>{suppliedCoverOmitsPageNumber ? "Index and exhibit pages use this continuous sequence. This cover stays unnumbered unless you tick “Print a page number on this cover”." : "Cover, index and exhibit pages use this one continuous sequence, including across separate volumes."}</small></div><div className="numbering-apply-row"><button data-testid="apply-page-numbering" className="secondary-button" type="button" disabled={!paginationDraftChanged} onClick={() => commitPagination({})}>Apply page-number changes</button><small>{paginationDraftChanged ? "Review the example, then apply these prefix, suffix or padding changes. They become part of the printed PDF label used by the index and suggestions." : "The displayed page-number settings are applied."}</small></div>
                      <h4>Page numbers shown in the PDF</h4>
                      <p className="advanced-settings-note">Choose where the page number is printed. This does not change filenames, exhibit descriptions or bookmarks.</p>
                      <div className="advanced-settings-grid compact">
                        <label>Page-number position<select value={paginationDraft.position} onChange={(event) => commitPagination({ position: event.target.value as PageNumberSettings["position"] }, false)}><option value="bottom-left">Bottom left</option><option value="bottom-centre">Bottom centre</option><option value="bottom-right">Bottom right</option><option value="top-left">Top left</option><option value="top-centre">Top centre</option><option value="top-right">Top right</option><option value="inside-bottom">Inside bottom</option><option value="outside-bottom">Outside bottom</option></select><small>Applies to every printed page number.</small></label>
                        <label>Page-number font size<input type="number" min="6" max="16" value={paginationDraft.fontSize} onChange={(event) => commitPagination({ fontSize: Math.max(6, Math.min(16, Number(event.target.value) || 8)) }, false)} /><small>Applies to every printed page number.</small></label>
                      </div>
                      {numberingDiffersFromPdfOrder(pagination) ? <div className="numbering-divergence" role="status"><strong>Prefix, suffix and padding are part of the printed PDF page label.</strong><span>{numberingDifferenceExample(pagination)}</span><small>The index and suggested statement references use that same printed label, not a second sequence.</small></div> : null}
                    </section>
                    <section className="advanced-settings-section template-settings-section">
                      <h3>Templates and optional pages</h3>
                      <p className="advanced-settings-note">PDF templates stay as the page background, fitted to A4 without cropping. Word templates are converted locally using a simplified renderer; preview and approve the converted PDF before building.</p>
                      <div className="template-subsection">
                        <div className="template-subsection-heading"><h4>Bundle cover and index</h4><p>These pages are always created. Choose a custom template or leave the standard design in place.</p></div>
                        <div className="template-grid primary-template-grid">
                          <article className="template-card"><strong>Bundle cover</strong><p>The first page of each finished PDF. Choose a template to keep its layout, or leave the standard design and enter the matter details here.</p>{templateControl("cover", "Choose a custom cover template", "Standard cover design selected")}</article>
                          <article className="template-card"><strong>Index of exhibits</strong><p>Exhibit rows are written into the columns that template already has. A Date column is filled only if that template already has one. If party names or a case number appear on that page, you can correct a misread and those corrections are printed too.</p>{templateControl("index", "Choose a fixed-layout index background", "Standard index design selected")}</article>
                        </div>
                      </div>
                      <div className="template-subsection">
                        <div className="template-subsection-heading"><h4>Optional pages</h4><p>Turn on a page type before choosing its custom template.</p></div>
                        <div className="optional-page-cards">
                          <article className={`optional-page-card ${layout.includeDividerPages ? "enabled" : ""}`}>
                            <label className="optional-page-toggle"><input type="checkbox" checked={layout.includeDividerPages} onChange={(event) => setLayout((current) => ({ ...current, includeDividerPages: event.target.checked, countOptionalPagesInReferences: event.target.checked || current.includeExhibitCoverPages ? current.countOptionalPagesInReferences : false }))} /><span><strong>Divider pages</strong><small>Add a divider before the exhibit documents. The standard divider is used unless you choose a custom template.</small></span></label>
                            {layout.includeDividerPages ? templateControl("divider", "Choose a custom divider template", "Standard divider design selected") : <p className="template-disabled-note">Turn on divider pages to choose a divider template.</p>}
                          </article>
                          <article className={`optional-page-card ${layout.includeExhibitCoverPages ? "enabled" : ""}`}>
                            <label className="optional-page-toggle"><input type="checkbox" checked={layout.includeExhibitCoverPages} onChange={(event) => setLayout((current) => ({ ...current, includeExhibitCoverPages: event.target.checked, countOptionalPagesInReferences: event.target.checked || current.includeDividerPages ? current.countOptionalPagesInReferences : false }))} /><span><strong>Exhibit-cover pages</strong><small>Add a separate cover immediately before each exhibit. The standard design is used unless you choose a custom template.</small></span></label>
                            {layout.includeExhibitCoverPages ? templateControl("exhibitCover", "Choose a custom exhibit-cover template", "Standard exhibit-cover design selected") : <p className="template-disabled-note">Turn on exhibit-cover pages to choose an exhibit-cover template.</p>}
                          </article>
                        </div>
                        <label className="reference-counting-option"><input type="checkbox" disabled checked={(layout.includeDividerPages || layout.includeExhibitCoverPages) && countsOptionalPagesInReferences(paginationDraft, layout)} onChange={() => undefined} /><span><strong>Include optional pages when calculating exhibit page references</strong><small>These pages count because the finished PDF uses one page sequence. That cannot be turned off.</small></span></label>
                      </div>
                      {templateMatterDiscrepancies.length ? <div className="template-discrepancy-review" role="group" aria-labelledby="template-discrepancy-title">
                        <strong id="template-discrepancy-title">Selected templates may show different matter details</strong>
                        <p>Exhibit Builder has not decided which wording is correct. Open the exact previews and compare the possible differences below.</p>
                        <ul>{templateMatterDiscrepancies.map((discrepancy) => <li key={discrepancy.field}><b>{discrepancy.message}</b>{discrepancy.evidence.map((evidence) => <span key={`${discrepancy.field}-${evidence.templateId}`}>{evidence.role}: {evidence.values.join(", ")}</span>)}</li>)}</ul>
                        <label><input type="checkbox" checked={!templateDiscrepancyPending} onChange={(event) => setTemplateDiscrepancyConfirmation(event.target.checked ? { fingerprint: templateDiscrepancyFingerprint, confirmedAt: new Date().toISOString() } : null)} /> I compared these possible differences against the visible templates and confirm the selected templates belong to this matter.</label>
                      </div> : null}
                    </section>
                    <section className="advanced-settings-section">
                      <h3>Volume splitting</h3>
                      <div className="advanced-settings-grid">
                        <label>Approximate maximum pages per PDF<input type="number" min="0" step="1" value={layout.volumePageLimit} onChange={(event) => setLayout((current) => ({ ...current, volumePageLimit: Math.max(0, Number(event.target.value) || 0) }))} /><small>Enter 0 to make one PDF. Otherwise the tool starts a new volume between exhibits and never splits an exhibit.</small></label>
                      </div>
                      {layout.volumePageLimit > 0 ? <>
                        <fieldset className="numbering-mode volume-numbering-mode"><legend>If the bundle is split, how should page numbers continue?</legend>
                          <label><input type="radio" name="volume-numbering" checked={paginationDraft.volumeNumbering === "continuous"} onChange={() => chooseVolumeNumbering("continuous")} /><span><strong>Continue numbering across all volumes — recommended</strong><small>Every page number remains unique. For example: Volume 1 uses AH-001 to AH-200, Volume 2 begins AH-201, and Volume 3 continues from there.</small></span></label>
                          <label><input type="radio" name="volume-numbering" checked={paginationDraft.volumeNumbering === "restart"} onChange={() => chooseVolumeNumbering("restart")} /><span><strong>Restart numbering in each volume</strong><small>Each PDF begins again at its first configured number. Page numbers will repeat, so every reference must also identify the volume.</small></span></label>
                        </fieldset>
                        {paginationDraft.volumeNumbering === "continuous" ? <div className="numbering-divergence" role="status"><strong>Later volumes will not match their local PDF page positions.</strong><span>For example, physical PDF page 1 of Volume 2 may be printed as AH-201.</span><small>This is intentional: the printed sequence remains unique across the complete bundle.</small></div> : <div className="numbering-divergence" role="alert"><strong>Page numbers will repeat in separate volumes.</strong><span>For example, both Volume 1 and Volume 2 may contain AH-001.</span><small>Indexes and suggested references will identify the volume wherever repeated numbering could otherwise be ambiguous.</small></div>}
                        <p className="advanced-settings-note volume-cover-note"><strong>Volume identification:</strong> {suppliedCoverOmitsVolumeLabel ? "This cover will not be marked Volume 1 of 3 unless you tick “Add a volume label if the bundle is split”." : "a split bundle is clearly marked “Volume 1 of 3”, “Volume 2 of 3” and so on. A single-volume custom cover receives no additional bundle label."}</p>
                      </> : <p className="advanced-settings-note">The finished output will be one PDF. No volume label will be added to a custom cover.</p>}
                    </section>
                    <section className="advanced-settings-section">
                      <h3>Statement-reference mark</h3>
                      <p className="advanced-settings-note">Choose the exhibit initials to use throughout this bundle. They create reference marks such as AH1. Changing them here updates the generated bundle, index and suggested references; it does not reread or alter the witness statement.</p>
                      <div className="witness-settings">
                    {statements.map((statement) => { const draft = statementDrafts[statement.id] ?? { witnessName: statement.witnessName, witnessInitials: statement.witnessInitials }; const changed = draft.witnessInitials !== statement.witnessInitials; const explicit = explicitPlaceholderInitials(statement.id); const conflicts = explicit.filter((initials) => initials !== (draft.witnessInitials.trim() || "EX").toUpperCase()); return (
                      <fieldset key={statement.id} data-testid="witness-setting">
                        <legend>{statement.file.name}</legend>
                        <div className="witness-setting-fields">
                          <label className="witness-initials"><span>Exhibit initials used throughout this bundle</span><input value={draft.witnessInitials} maxLength={6} placeholder="AH" onChange={(event) => updateStatementDraft(statement.id, { witnessInitials: event.target.value.toUpperCase().replace(/[^A-Z0-9]/g, "") })} /><small>Final suggestions will use {(draft.witnessInitials.trim() || "EX").toUpperCase()}1. The original witness statement is not changed.</small></label>
                        </div>
                        {explicit.length ? <p className={conflicts.length ? "initials-conflict" : "initials-match"}>The statement contains explicit placeholder initials: {explicit.join(", ")}.{conflicts.length ? " These do not match the proposed initials; applying the change requires confirmation." : " These match the proposed exhibit labels."}</p> : <p className="initials-match">No explicit initialled placeholder was found. Generic [Exhibit] placeholders will use the initials selected here.</p>}
                        <div className="witness-setting-actions"><button className="primary-button compact-action" type="button" disabled={!changed || !draft.witnessInitials.trim()} onClick={() => applyStatementDraft(statement.id)}>Apply exhibit initials</button><button className="secondary-button compact-action" type="button" disabled={!changed} onClick={() => cancelStatementDraft(statement.id)}>Cancel changes</button><small>Existing source analysis and exhibit approvals are preserved.</small></div>
                      </fieldset>
                    ); })}
                      </div>
                    </section>
                  </div>
                </details>

                <section className="review-progress-strip" aria-label="Exhibit review progress">
                  <div><strong>{citedCandidates.length}</strong><span>reference{citedCandidates.length === 1 ? "" : "s"}</span></div>
                  <span aria-hidden="true">|</span>
                  <div><strong>{matchedCitationCount}</strong><span>document{matchedCitationCount === 1 ? "" : "s"} matched</span></div>
                  <span aria-hidden="true">|</span>
                  <div className={needsDecisionCount ? "needs-attention" : "complete"}><strong>{confirmedCitationCount} of {citedCandidates.length}</strong><span>references confirmed{repeatExhibitNote}</span></div>
                  {addedExhibitCount ? <><span aria-hidden="true">|</span><div className="added-exhibits"><strong>{addedExhibitCount}</strong><span>added exhibit{addedExhibitCount === 1 ? "" : "s"}</span></div></> : null}
                  {unreferencedEvidence.length ? <><span aria-hidden="true">|</span><button className="progress-link" type="button" onClick={() => setView("reconcile")}>{unreferencedEvidence.length} unused file{unreferencedEvidence.length === 1 ? "" : "s"}</button></> : null}
                  <details className="review-safety-note"><summary>Witness statement remains unchanged</summary><p>The tool calculates exhibit-bundle page ranges from confirmed exhibits. Exhibit Builder never writes them back into the DOCX.</p></details>
                </section>

                {analysis.statementWarnings.length > 0 && (
                  <div className="statement-warning-list">
                    {analysis.statementWarnings.map((warning) => (
                      <p key={warning}>{warning}</p>
                    ))}
                  </div>
                )}

                <div className="manifest-card">
                  <div className="review-list-toolbar">
                    {referenceMarkConflict ? <div className="reference-format-note reference-format-conflict" role="alert">
                      <span>Reference marks need review</span>
                      <strong>{[...includedReferenceMarks].join(" and ")}</strong>
                      <small>Choose one statement-reference mark in Optional settings before continuing.</small>
                    </div> : <div className="reference-format-note" aria-live="polite">
                      <span>Statement reference format</span>
                      <strong>[{referenceBundleMark}/page]</strong>
                      <small>{referenceBundleMark} will be used throughout this bundle and in the final suggestions.{layout.volumePageLimit > 0 && pagination.volumeNumbering === "restart" ? " Because page numbers restart in each volume, suggestions will also identify the volume." : ""}</small>
                    </div>}
                    <div className="review-list-controls">
                      <div className="review-filters" aria-label="Filter exhibits"><span>Show:</span>
                        <button type="button" className={!showPendingOnly && !showDuplicatesOnly ? "active" : ""} aria-pressed={!showPendingOnly && !showDuplicatesOnly} onClick={() => { setShowPendingOnly(false); setShowDuplicatesOnly(false); }}>All</button>
                        <button type="button" className={showPendingOnly ? "active" : ""} aria-pressed={showPendingOnly} onClick={() => { setShowPendingOnly(true); setShowDuplicatesOnly(false); }}>Outstanding {pendingCandidateIds.size}</button>
                        <button type="button" className={showDuplicatesOnly ? "active" : ""} aria-pressed={showDuplicatesOnly} onClick={() => { setShowDuplicatesOnly(true); setShowPendingOnly(false); }} disabled={!possibleDuplicateEvidenceIds.size}>Possible duplicates {possibleDuplicateEvidenceIds.size}</button>
                      </div>
                      <div className="review-list-actions">
                        <button className="secondary-button compact-action add-exhibit-button" type="button" onClick={() => openManualAdd("review")}>Add exhibit</button>
                        {bulkConfirmableCount > 0 ? <button className="secondary-button compact-action bulk-confirm-action" type="button" data-tour="confirm-all" onClick={confirmMatchedExhibits}>Confirm all {bulkConfirmableCount} proposed match{bulkConfirmableCount === 1 ? "" : "es"}</button> : null}
                      </div>
                    </div>
                  </div>
                  {renderManualExhibitPanel()}
                  {/* Retired manual-number table retained only as a migration note.
                    <table>
                      <thead>
                        <tr>
                          <th aria-label="Include" />
                          <th>Provisional number</th>
                          <th>Suggested description</th>
                          <th>Statement source</th>
                          <th>Matched file</th>
                          <th>Confidence</th>
                          <th>Human confirmation</th>
                        </tr>
                      </thead>
                      <tbody>
                        {candidates.map((candidate) => (
                          <tr
                            key={candidate.id}
                            className={!candidate.included ? "excluded" : ""}
                          >
                            <td>
                              <input
                                type="checkbox"
                                checked={candidate.included}
                                onChange={(event) =>
                                  updateCandidate(candidate.id, {
                                    included: event.target.checked,
                                  })
                                }
                                aria-label={`Include ${candidate.mark}`}
                              />
                            </td>
                            <td>
                              <label className="provisional-number">
                                <span>{candidate.witnessInitials ?? candidate.mark.replace(/\s*\d+$/, "")}</span>
                                <input
                                  type="number"
                                  min="1"
                                  max="999"
                                  value={candidate.provisionalNumber}
                                  onChange={(event) => {
                                    const provisionalNumber = Number(
                                      event.target.value,
                                    );
                                    if (
                                      Number.isInteger(provisionalNumber) &&
                                      provisionalNumber > 0
                                    ) {
                                      updateCandidate(candidate.id, {
                                        provisionalNumber,
                                        mark: `${candidate.witnessInitials ?? candidate.mark.replace(/\s*\d+$/, "")} ${provisionalNumber}`,
                                      });
                                    }
                                  }}
                                  aria-label={`Provisional number for paragraph ${candidate.paragraph}`}
                                />
                              </label>
                              <small className="provisional-suggestion">
                                {candidate.mark} - paragraph {candidate.paragraph}
                              </small>
                            </td>
                            <td>
                              <input
                                className="description-input"
                                type="text"
                                value={candidate.description}
                                onChange={(event) =>
                                  updateCandidate(candidate.id, {
                                    description: event.target.value,
                                  })
                                }
                                aria-label={`Description for ${candidate.mark}`}
                              />
                              <small>{candidate.date}</small>
                            </td>
                            <td>
                              <button
                                className="paragraph-link"
                                type="button"
                                title={candidate.citation}
                              >
                                 {candidate.paragraph}
                              </button>
                              <small className="citation-preview">
                                {(candidate.citationCount ?? 1) > 1 ? `Reference ${candidate.citationOrdinal} of ${candidate.citationCount}: ` : ""}{candidate.citation}
                              </small>
                              <small className="discovery-signal">
                                {candidate.discoverySignals.join(" + ")}
                              </small>
                            </td>
                            <td>
                              <select
                                value={candidate.evidenceId ?? ""}
                                onChange={(event) =>
                                  updateCandidate(candidate.id, {
                                    evidenceId: event.target.value || null,
                                    confidence: event.target.value ? 100 : 0,
                                    rationale: event.target.value
                                      ? "Source selected by reviewer"
                                      : "No source file matched",
                                  })
                                }
                                aria-label={`Matched file for ${candidate.mark}`}
                              >
                                <option value="">Select a file</option>
                                {analysis.evidence.map((record) => (
                                  <option key={record.id} value={record.id}>
                                    {record.name}
                                  </option>
                                ))}
                              </select>
                              <small title={candidate.rationale}>
                                {candidate.rationale}
                              </small>
                              {candidate.alternativeEvidenceIds?.length ? (
                                <small>Alternatives: {candidate.alternativeEvidenceIds.map((id) => analysis.evidence.find((record) => record.id === id)?.name).filter(Boolean).join("; ")}</small>
                              ) : null}
                              {analysis.evidence.find((record) => record.id === candidate.evidenceId)?.extension === "pdf" && (
                                <small>
                                  Pages <input type="number" min="1" value={candidate.pageStart ?? 1} onChange={(event) => updateCandidate(candidate.id, { pageStart: Math.max(1, Number(event.target.value) || 1) })} aria-label={`First page for ${candidate.mark}`} />
                                  - <input type="number" min="1" value={candidate.pageEnd ?? ""} onChange={(event) => updateCandidate(candidate.id, { pageEnd: Number(event.target.value) || undefined })} aria-label={`Last page for ${candidate.mark}`} placeholder="last" />
                                </small>
                              )}
                            </td>
                            <td>
                              <span
                                className={`confidence confidence-${confidenceLabel(
                                  candidate.confidence,
                                ).toLowerCase()}`}
                              >
                                {confidenceLabel(candidate.confidence)}
                                <b>{candidate.confidence}%</b>
                              </span>
                            </td>
                            <td>
                              <label
                                className={`confirmation-control ${
                                  candidate.confirmed ? "is-confirmed" : ""
                                }`}
                              >
                                <input
                                  type="checkbox"
                                  checked={candidate.confirmed}
                                  disabled={
                                    !candidate.included || !candidate.evidenceId
                                  }
                                  onChange={(event) =>
                                    updateCandidate(candidate.id, {
                                      confirmed: event.target.checked,
                                    })
                                  }
                                  aria-label={`Confirm ${candidate.mark} at paragraph ${candidate.paragraph}`}
                                />
                                <span>
                                  {candidate.confirmed
                                    ? "Confirmed"
                                    : "Confirm row"}
                                </span>
                              </label>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  */}
                  <div className="exhibit-card-list" tabIndex={-1}>
                    {reviewCandidates
                      .filter((candidate) => !showPendingOnly || pendingCandidateIds.has(candidate.id))
                      .filter((candidate) => !showDuplicatesOnly || (candidate.evidenceId ? possibleDuplicateEvidenceIds.has(candidate.evidenceId) : false))
                      .map((candidate) => {
                        const group = exhibitGroupLookupByKind.byCandidateId.get(candidate.id);
                        const record = candidate.evidenceId ? evidenceById.get(candidate.evidenceId) : undefined;
                        const itemNumber = reviewItemNumberByCandidate.get(candidate.id);
                        const repeatNeedsDecision = Boolean(group?.decisionPending || group?.selectionConflict || group?.collisionMembers.slice(1).some((repeat) => !repeat.confirmed));
                        const attachmentsNeedDecision = Boolean(record?.emailAttachments?.length && candidate.included && unresolvedEmailAttachments(record.emailAttachments, candidate.emailAttachmentDispositions).length);
                        const compact = candidate.confirmed && !repeatNeedsDecision && !attachmentsNeedDecision && !expandedConfirmedCards.has(candidate.id) && openEmailAttachmentsId !== candidate.id;
                        const repeatFocusId = group?.selectionConflict
                          ? group.collisionMembers.slice(1).find((repeat) => repeat.repeatDecision === "same")?.id
                          : group?.collisionMembers.slice(1).find((repeat) => !repeat.confirmed || !repeat.repeatDecision || repeat.repeatDecision === "pending")?.id;
                        const repeatControls = group?.collisionMembers.slice(1).map((repeat) => <div className="repeat-reference" key={repeat.id}>
                          <span>{repeat.statementName ? `${repeat.statementName}, ` : ""}paragraph {repeat.paragraph}: {repeat.description} {repeat.repeatDecision === "same" ? `Same exhibit as this card (still cited as ${referenceBundleMark})` : ""}</span>
                          <div role="group" aria-label={`Treatment of the repeated reference at paragraph ${repeat.paragraph}`}>
                            <button type="button" data-confirm-focus={!group.selectionConflict && repeat.id === repeatFocusId && (!repeat.repeatDecision || repeat.repeatDecision === "pending") ? "" : undefined} aria-pressed={repeat.repeatDecision === "same"} className={repeat.repeatDecision === "same" ? "selected" : ""} onClick={() => { if (repeat.repeatDecision !== "same") updateCandidate(repeat.id, { repeatDecision: "same" }); }}>Same exhibit</button>
                            <button type="button" data-confirm-focus={group.selectionConflict && repeat.id === repeatFocusId ? "" : undefined} aria-pressed={repeat.repeatDecision === "separate"} className={repeat.repeatDecision === "separate" ? "selected" : ""} onClick={() => { if (repeat.repeatDecision !== "separate") updateCandidate(repeat.id, { repeatDecision: "separate" }); }}>Separate exhibit</button>
                            <button className={`${repeat.confirmed ? "secondary-button" : "primary-button"} repeat-confirm-button`} type="button" data-tour="repeat-decision" data-confirm-focus={!group.selectionConflict && repeat.id === repeatFocusId && Boolean(repeat.repeatDecision && repeat.repeatDecision !== "pending") ? "" : undefined} disabled={!repeat.repeatDecision || repeat.repeatDecision === "pending"} onClick={() => confirmRepeatDecision(repeat, !repeat.confirmed, candidate.id)}>{repeat.confirmed ? "Undo repeat decision" : "Confirm repeat decision"}</button>
                          </div>
                        </div>);
                        return (
                          <article className={`exhibit-review-card ${!candidate.included ? "excluded" : ""} ${compact ? "compact" : ""}`} key={candidate.id} data-candidate-id={candidate.id} data-included={candidate.included ? "true" : "false"} data-confirmed={candidate.confirmed ? "true" : "false"} data-confirmable={candidate.included && candidate.evidenceId ? "true" : "false"} tabIndex={-1}>
                            <header>
                              <div className="review-card-identity">
                                <strong>{candidate.description}</strong>
                                <div className="review-card-chips">
                                  {candidate.manualAddition ? <span>Not cited in the statement</span> : <span>Cited at paragraph {candidate.paragraph}{(candidate.citationCount ?? 1) > 1 ? `, reference ${candidate.citationOrdinal} of ${candidate.citationCount}` : ""}</span>}
                                  {candidate.parentEmailProvenance ? <span>From {candidate.parentEmailProvenance.parentName}</span> : null}
                                </div>
                              </div>
                              <div className="review-card-status">
                                {repeatNeedsDecision ? <span className="needs-review-status">Repeat decision needed</span> : attachmentsNeedDecision ? <span className="needs-review-status">Attachment decision needed</span> : candidate.confirmed ? <span className="confirmed-status">Confirmed</span> : null}
                                {candidate.confirmed
                                  ? isAutomaticLowConfidenceMatch(candidate)
                                    ? <span className="match-provenance">Low-confidence automatic suggestion</span>
                                    : isReviewerSelectedSource(candidate)
                                      ? <span className="match-provenance">Confirmed by you</span>
                                      : null
                                  : <span className={`confidence confidence-${confidenceLabel(candidate.confidence).toLowerCase()}`}>{matchStrengthLabel(candidate.confidence, candidate.rationale)}</span>}
                                <label className="include-control"><input type="checkbox" checked={candidate.included} onChange={(event) => updateCandidate(candidate.id, { included: event.target.checked })} /> <span>{candidate.included ? "Included" : "Excluded"}</span></label>
                                {compact ? <button className="secondary-button compact-action" type="button" data-confirm-focus onClick={() => { setExpandedConfirmedCards((current) => new Set(current).add(candidate.id)); focusConfirmControl(candidate.id); }}>View or change</button> : candidate.confirmed && !repeatNeedsDecision && !attachmentsNeedDecision ? <button className="secondary-button compact-action" type="button" data-confirm-focus onClick={() => { setOpenDocumentPickerId((current) => current === candidate.id ? null : current); collapseConfirmedCard(candidate.id); focusConfirmControl(candidate.id); }}>Minimise</button> : null}
                              </div>
                            </header>
                            {!compact ? <>
                            <div className="exhibit-card-grid">
                              <section className="primary-review-panel">
                                <span className="field-label">Statement reference</span>
                                <div className="citation-preview"><strong>Paragraph {candidate.paragraph}{(candidate.citationCount ?? 1) > 1 ? ` - reference ${candidate.citationOrdinal} of ${candidate.citationCount}` : ""}</strong><p>{candidate.citation}</p></div>
                              </section>
                              <section className="primary-review-panel">
                                <label className="field-label">Proposed document</label>
                                <details className="document-picker" open={openDocumentPickerId === candidate.id} onToggle={(event) => { if (event.currentTarget.open) setOpenDocumentPickerId(candidate.id); else setOpenDocumentPickerId((current) => current === candidate.id ? null : current); }}>
                                  <summary aria-label={`Choose the proposed document for review list #${itemNumber}`}><span>{record?.name ?? "Choose a document"}</span><b>{record ? "Change" : "Choose"}</b></summary>
                                  {openDocumentPickerId === candidate.id ? <div className="document-picker-options" role="group" aria-label={`Available documents for review list #${itemNumber}`}>
                                    <button type="button" aria-pressed={!candidate.evidenceId} className={!candidate.evidenceId ? "selected" : ""} onClick={(event) => { if (candidate.evidenceId) updateCandidate(candidate.id, { evidenceId: null, confidence: 0, rationale: "No source file matched", repeatDecision: "pending", confirmed: false, confirmationMethod: undefined, confirmedAt: undefined }); restoreDocumentPickerSummary(event.currentTarget); setOpenDocumentPickerId(null); }}>No document selected</button>
                                    {analysis.evidence.filter((item) => !item.derivedFromEmail).map((item) => {
                                      const duplicateLabel = (evidenceHashCounts.get(item.sha256) ?? 0) > 1 ? " (identical physical copy)" : "";
                                      const selected = candidate.evidenceId === item.id;
                                      return <button type="button" aria-pressed={selected} className={selected ? "selected" : ""} key={item.id} onClick={(event) => { if (!selected) updateCandidate(candidate.id, { evidenceId: item.id, confidence: 100, rationale: REVIEWER_SELECTED_RATIONALE, repeatDecision: "pending", confirmed: false, confirmationMethod: undefined, confirmedAt: undefined }); restoreDocumentPickerSummary(event.currentTarget); setOpenDocumentPickerId(null); }}>{item.name}{duplicateLabel}</button>;
                                    })}
                                  </div> : null}
                                </details>
                                {!candidate.evidenceId ? <section className="match-explanation"><strong>Document required</strong><p>Select the source document referred to in the statement, or exclude this item.</p></section> : null}
                                {!record?.emailAttachments?.length ? <button className={`${candidate.confirmed ? "secondary-button" : "primary-button"} confirm-document-button`} type="button" data-confirm-document data-confirm-action={candidate.confirmed ? "undo" : "confirm"} disabled={!candidate.included || !candidate.evidenceId} onClick={() => { const confirmed = !candidate.confirmed; if (confirmed) repeatNeedsDecision ? keepCurrentReviewCard(candidate.id) : preserveNextReviewCard(candidate.id); updateCandidate(candidate.id, confirmed ? { confirmed: true, confirmationMethod: "individual", confirmedAt: new Date().toISOString() } : { confirmed: false, confirmationMethod: undefined, confirmedAt: undefined }); }}>{candidate.confirmed ? "Undo confirmation" : "Confirm this document"}</button> : null}
                              </section>
                            </div>
                            <div className="review-index-fields">
                              <section>
                                <label className="field-label" htmlFor={`index-description-${candidate.id}`}>Index description</label>
                                <input id={`index-description-${candidate.id}`} className="description-input" type="text" value={candidate.description} onChange={(event) => updateCandidate(candidate.id, { description: event.target.value })} aria-label={`Index description for ${candidate.mark}`} />
                              </section>
                              <section>
                                <label className="field-label" htmlFor={`document-date-${candidate.id}`}>Document date</label>
                                <input id={`document-date-${candidate.id}`} className="description-input date-input" type="text" value={candidate.date === "Date not stated" ? "" : candidate.date} placeholder="Date not stated" onChange={(event) => updateCandidate(candidate.id, { date: event.target.value.trim() ? event.target.value : "Date not stated" })} aria-label={`Document date for ${candidate.mark}`} />
                              </section>
                            </div>
                            {record?.emailAttachments?.length ? <details className="email-attachments-panel" data-email-attachments tabIndex={-1} open={attachmentsNeedDecision || openEmailAttachmentsId === candidate.id} onToggle={(event) => { if (event.currentTarget.open) setOpenEmailAttachmentsId(candidate.id); else setOpenEmailAttachmentsId((current) => current === candidate.id ? null : current); }}>
                              <summary>Email attachments ({record.emailAttachments.length})</summary>
                              <p className="supporting-copy">Each attachment needs a decision before this email can be built. Nested attachments are listed separately and are not included unless you choose them.</p>
                              <ul className="email-attachment-list">
                                {record.emailAttachments.map((child) => {
                                  const disposition = candidate.emailAttachmentDispositions?.[child.identity];
                                  const childExhibit = candidates.find((item) => item.parentEmailProvenance?.childIdentity === child.identity);
                                  const childRecord = childExhibit?.evidenceId ? evidenceById.get(childExhibit.evidenceId) : undefined;
                                  return <li key={child.identity}>
                                    <div>
                                      <strong>{child.name}</strong>
                                      <small>{child.supported ? child.extension.toUpperCase() : "Unsupported type"}{child.nested ? " · nested" : ""}</small>
                                    </div>
                                    <div className="email-attachment-actions" role="group" aria-label={`Treatment of ${child.name}`}>
                                      {child.supported ? <>
                                        <button type="button" data-tour="print-with-email" aria-pressed={disposition === "print-with-email"} className={disposition === "print-with-email" ? "selected" : ""} onClick={() => setEmailChildDisposition(candidate.id, child, "print-with-email")}>Print with this email</button>
                                        <button type="button" aria-pressed={disposition === "add-as-exhibit"} className={disposition === "add-as-exhibit" ? "selected" : ""} onClick={() => setEmailChildDisposition(candidate.id, child, "add-as-exhibit")}>Add as its own exhibit</button>
                                      </> : null}
                                      <button type="button" aria-pressed={disposition === "leave-out"} className={disposition === "leave-out" ? "selected" : ""} onClick={() => setEmailChildDisposition(candidate.id, child, "leave-out")}>Leave out</button>
                                    </div>
                                    <p className="email-attachment-result">{emailChildDispositionResult(disposition)}</p>
                                    {disposition === "add-as-exhibit" && childExhibit ? <div className="email-attachment-child" role="group" aria-label={`Index details for ${child.name}`} data-email-child-exhibit={childExhibit.id}>
                                      <label>Index description for {child.name}<input value={childExhibit.description} onChange={(event) => updateManualCandidate(childExhibit.id, { description: event.target.value })} aria-label={`Index description for ${child.name}`} /></label>
                                      <label>Document date for {child.name}<input value={childExhibit.date === "Date not stated" ? "" : childExhibit.date} placeholder="Date not stated" onChange={(event) => updateManualCandidate(childExhibit.id, { date: event.target.value.trim() ? event.target.value : "Date not stated" })} aria-label={`Document date for ${child.name}`} /></label>
                                      {childRecord?.extension === "pdf" ? <label>First source page for {child.name}<input type="number" min="1" value={childExhibit.pageStart ?? 1} onChange={(event) => updateManualCandidate(childExhibit.id, { pageStart: Math.max(1, Number(event.target.value) || 1) })} aria-label={`First source page for ${child.name}`} /></label> : null}
                                      {childRecord?.extension === "pdf" ? <label>Last source page for {child.name}<input type="number" min="1" value={childExhibit.pageEnd ?? ""} placeholder="last" onChange={(event) => updateManualCandidate(childExhibit.id, { pageEnd: Number(event.target.value) || undefined })} aria-label={`Last source page for ${child.name}`} /></label> : null}
                                      {childRecord?.extension === "xlsx" ? <small>Select worksheets on the workbook sheets stage.</small> : null}
                                    </div> : null}
                                  </li>;
                                })}
                              </ul>
                            </details> : null}
                            {record?.emailAttachments?.length ? <button className={`${candidate.confirmed ? "secondary-button" : "primary-button"} confirm-document-button`} type="button" data-confirm-document data-confirm-action={candidate.confirmed ? "undo" : "confirm"} disabled={!candidate.included || !candidate.evidenceId} onClick={() => { const confirmed = !candidate.confirmed; if (confirmed && attachmentsNeedDecision) keepEmailReviewOpen(candidate.id); else if (confirmed && repeatNeedsDecision) keepCurrentReviewCard(candidate.id); else if (confirmed) preserveNextReviewCard(candidate.id); updateCandidate(candidate.id, confirmed ? { confirmed: true, confirmationMethod: "individual", confirmedAt: new Date().toISOString() } : { confirmed: false, confirmationMethod: undefined, confirmedAt: undefined }); }}>{candidate.confirmed ? "Undo confirmation" : "Confirm this document"}</button> : null}
                            {group?.collision && repeatNeedsDecision ? <section className="repeat-panel needs-decision"><strong>Confirm how the repeated reference should be treated</strong><p>The same source file is selected for more than one statement reference. Choose whether each later reference uses the same exhibit or creates a separate exhibit, then confirm the decision.</p>{repeatControls}</section> : null}
                            <details className="supporting-review-details">
                              <summary>Review supporting details</summary>
                              <div className="supporting-review-grid">
                                <section><span className="field-label">Why this document was selected</span><p className="supporting-copy">{candidate.rationale}</p><p className="supporting-copy">{isReviewerSelectedSource(candidate) ? (candidate.confirmed ? "You selected this source. Confirmation records that selection; it is not a finding that the document is legally correct." : "Selected by you; confirmation is still required.") : candidate.confirmed ? `Originally suggested automatically (comparison score ${candidate.confidence}/100)${candidate.confirmationMethod === "bulk" ? " and confirmed in bulk" : candidate.confirmationMethod === "individual" ? " and confirmed individually" : ""}. The score is a ranking aid, not a finding that the document is correct.` : <>Automated comparison score: {candidate.confidence}/100. This is a ranking aid, not the probability that the document is correct.{isAutomaticLowConfidenceMatch(candidate) ? " Check the statement wording against the selected document before confirming." : ""}</>}</p>{candidate.discoverySignals.length ? <p className="supporting-copy">Statement signals: {candidate.discoverySignals.join(", ")}.</p> : null}{candidate.alternativeEvidenceIds?.length ? <p className="supporting-copy">Alternatives: {candidate.alternativeEvidenceIds.map((id) => evidenceById.get(id)?.name).filter(Boolean).join("; ")}.</p> : null}</section>
                                {candidate.contextParagraphs?.length ? <section className="citation-context supporting-wide"><strong>Statement context</strong>{candidate.contextParagraphs.map((item) => <p key={`${item.position}-${item.paragraph}`}><b>{item.position === "previous" ? "Previous" : "Following"} paragraph {item.paragraph}</b>{item.text}</p>)}</section> : candidate.context && candidate.context !== candidate.citation ? <section className="citation-context supporting-wide"><strong>Statement context</strong><p>{candidate.context}</p></section> : null}
                                {record ? <details className="source-preview supporting-wide" open={showSourcePreview === candidate.id} onToggle={(event) => setShowSourcePreview((event.currentTarget as HTMLDetailsElement).open ? candidate.id : null)}><summary>Review source document</summary><p>{record.text.trim().slice(0, 1_800) || "No extracted text preview is available for this source. Review the original file before confirming."}</p>{record.extension === "pdf" ? <button className="secondary-button compact-action" type="button" onClick={(event) => previewOriginalPdf(record, event.currentTarget)}>Open original PDF</button> : null}</details> : null}
                                {record?.extension === "pdf" ? <details className="advanced-pages supporting-wide"><summary>Include selected pages only</summary><p>Leave this closed to include the whole document. Set a range only when part of the selected PDF is the exhibit.</p><label>First source page <input type="number" min="1" value={candidate.pageStart ?? 1} onChange={(event) => updateCandidate(candidate.id, { pageStart: Math.max(1, Number(event.target.value) || 1) })} /></label><label>Last source page <input type="number" min="1" value={candidate.pageEnd ?? ""} onChange={(event) => updateCandidate(candidate.id, { pageEnd: Number(event.target.value) || undefined })} placeholder="last" /></label></details> : null}
                                <details className="review-card-details supporting-wide"><summary>Optional audit details{(candidate.aliases?.length || candidate.reviewNote) ? " - added" : ""}</summary><div className="review-card-actions"><label>Alternative document names<input value={(candidate.aliases ?? []).join("; ")} placeholder="Separate several names with ;" onChange={(event) => updateCandidate(candidate.id, { aliases: event.target.value.split(";").map((value) => value.trim()).filter(Boolean) })} /><small>Recorded in the build report; does not rename the source file.</small></label><label>Reviewer note<input value={candidate.reviewNote ?? ""} placeholder="Optional note for the build report" onChange={(event) => updateCandidate(candidate.id, { reviewNote: event.target.value })} /><small>Stored locally in the project and build report.</small></label></div></details>
                              </div>
                            </details>
                            {group?.collision && !repeatNeedsDecision ? <details className="repeat-panel repeat-resolved"><summary id={`repeat-resolved-${group.id}`} tabIndex={-1}>Repeat citation resolved</summary><p>The repeat decision is complete. Open this section if it needs to be changed.</p>{repeatControls}</details> : null}
                            </> : null}
                          </article>
                        );
                      })}
                    {showPendingOnly && pendingCandidateIds.size === 0 && <div className="pending-empty"><strong>No outstanding approvals</strong><p>Every included exhibit has a confirmed document match and any repeat decisions have been resolved.</p><button type="button" className="secondary-button" onClick={() => setShowPendingOnly(false)}>Show all exhibits</button></div>}
                  </div>
                </div>

                {technicalBuildBlockers.length > 0 && busy === null && <details className="build-readiness" data-testid="build-readiness" id="build-readiness" open={technicalBuildBlockers.some((blocker) => !blocker.candidateId) || undefined}>
                  <summary className="build-readiness-heading">
                    <div><p className="eyebrow">Other requirements</p><h2>{technicalBuildBlockers.length} item{technicalBuildBlockers.length === 1 ? "" : "s"} to resolve before continuing</h2></div>
                  </summary>
                  <div className="readiness-list readiness-jump-list">{technicalBuildBlockers.map((blocker) => {
                    const check = rawPreflight.find((item) => blocker.id === `preflight-${item.id}`);
                    const ocrApproved = Boolean(check && resolutions.some((resolution) => resolution.blockerId === check.id && resolution.action === "proceed-without-ocr" && resolution.sourceSha256 === check.sourceSha256));
                    const emailAttachment = isEmailAttachmentBlocker(blocker);
                    return <div key={blocker.id} data-blocker-id={blocker.id} tabIndex={-1} className={`readiness-item ${blocker.kind}`}>
                      <span aria-hidden="true">!</span>
                      <div className="readiness-copy"><p className="readiness-label"><strong>{blocker.label}</strong></p>{blocker.fileName ? <p className="readiness-file">{blocker.fileName}</p> : null}<p className="readiness-detail">{blocker.detail}</p>
                        <div className="readiness-controls">
                          {check && isOcrCheck(check) && <div className="readiness-action-group"><button className="secondary-button compact-action" type="button" onClick={(event) => previewOcrSource(check, event.currentTarget)}>Open original PDF for visual review</button><label className="exception-toggle"><input type="checkbox" checked={ocrApproved} disabled={!ocrApproved && !visuallyReviewedSourceHashes.has(check.sourceSha256 ?? "")} onChange={(event) => event.target.checked ? approveOcrException(check) : clearResolution(check)} /><span>Proceed without OCR after visual review</span></label><p className="readiness-action-note">The original source will remain in the exhibit bundle, but it will not have a searchable OCR text layer. The approval is tied to this exact source file.</p></div>}
                          {blocker.kind === "template" && <div className="readiness-action-group"><button className="secondary-button compact-action" type="button" onClick={() => navigateToBlocker(blocker)}>{blocker.actionLabel}</button><button className="quiet-text-button" type="button" onClick={useBuiltInTemplates}>Discard custom templates and use the built-in layout</button></div>}
                          {blocker.candidateId ? <button className="secondary-button compact-action" type="button" data-tour={emailAttachment ? "attachments" : undefined} onClick={() => navigateToBlocker(blocker)}>{(() => {
                            if (emailAttachment) return "Open attachment choices on this email";
                            const unmatched = candidates.find((item) => item.id === blocker.candidateId);
                            if (unmatched && !unmatched.manualAddition && unmatched.paragraph > 0) return `Open paragraph ${unmatched.paragraph}`;
                            return blocker.actionLabel;
                          })()}</button> : null}
                          {check && check.sourceId && !emailAttachment ? <div className="readiness-action-group"><button className="secondary-button compact-action" type="button" onClick={() => excludeSource(check)}>{check.code === "workbook.fidelity_failed" ? "Leave this Excel file out of the bundle" : "Leave this file out of the bundle"}</button><p className="readiness-action-note">{check.code === "workbook.fidelity_failed" ? "This Excel file and every exhibit that uses it are left out of the bundle. The witness statement is not edited, and the file stays on your computer." : "This file and every exhibit that uses it are left out of the bundle. The witness statement is not edited, and the file stays on your computer."}</p></div> : null}
                        </div>
                      </div>
                    </div>;
                  })}</div>
                </details>}

                {resolutions.length > 0 && <section className="resolution-summary" aria-live="polite" data-testid="resolution-summary">
                  <div>
                    <h2>{resolutions.length} technical exception{resolutions.length === 1 ? "" : "s"} recorded</h2>
                    <p>These exceptions are retained for this project and included in the build report. OCR approvals can be undone here; other exceptions can be changed from the relevant review card.</p>
                  </div>
                  <ul>{resolutions.map((resolution) => <li key={`${resolution.blockerId}-${resolution.sourceSha256 ?? resolution.candidateId ?? "project"}`}><strong>{resolution.fileName ?? "Project setting"}</strong><span>{resolution.action === "proceed-without-ocr" ? "Included without OCR" : resolution.action === "exclude-source" ? "Source excluded with its citations" : resolution.action === "exclude-candidate" ? "Citation excluded" : "Built-in template selected"}</span>{resolution.action === "proceed-without-ocr" && <button className="text-button resolution-undo" type="button" onClick={() => undoResolution(resolution)}>Undo approval</button>}</li>)}</ul>
                </section>}

                <div className="review-sticky-bar" aria-label="Review navigation">
                  <button className="secondary-button" type="button" onClick={() => setView("sources")}>Back to sources</button>
                  <span id="review-continue-status"><strong>{confirmedCitationCount} of {citedCandidates.length}</strong> statement references confirmed{repeatExhibitNote}{pendingCandidateIds.size ? ` - ${pendingCandidateIds.size} approval${pendingCandidateIds.size === 1 ? "" : "s"} remaining` : ""}{reviewContinueReason ? ` — ${reviewContinueReason}` : ""}</span>
                  <button className="primary-button" type="button" onClick={() => { if (readyToLeaveReview) setView(includedWorkbookInMatter ? "sheets" : "build"); else if (nextReviewBlocker) navigateToBlocker(nextReviewBlocker); else setShowPendingOnly(true); }} disabled={busy !== null} data-testid="build-bundle" data-tour={readyToLeaveReview ? (includedWorkbookInMatter ? "continue-sheets" : "continue-finalise") : undefined} aria-describedby="review-continue-status">{readyToLeaveReview ? (includedWorkbookInMatter ? "Continue to workbook sheets" : "Continue to finalise") : "Review next requirement"}</button>
                </div>
              </>
            )}

            {view === "reconcile" && (
              <>
                <div className="workspace-header">
                  <div>
                    <p className="eyebrow">Exception review</p>
                    <h1 tabIndex={-1}>Reconcile the evidence inbox</h1>
                    <p>
                      Files without a confirmed citation stay outside the bundle.
                    </p>
                  </div>
                  <button
                    className="secondary-button"
                    type="button"
                    onClick={() => setView("review")}
                  >
                    Back to exhibit review
                  </button>
                </div>
                <div className="reconcile-grid">
                  <article className="exception-card">
                    <div className="exception-icon" aria-hidden="true">
                      !
                    </div>
                    <div>
                      <p className="eyebrow">Supplied but not cited</p>
                      <h2>{unreferencedEvidence.length} file{unreferencedEvidence.length === 1 ? "" : "s"} not included</h2>
                      <p>
                        The builder does not silently add spare documents to the
                        bundle.
                      </p>
                    </div>
                  </article>
                  {unreferencedEvidence.map((record) => (
                    <article className="file-review-card" key={record.id}>
                      <div className="file-type">
                        {record.extension.toUpperCase()}
                      </div>
                      <div>
                        <h3>{record.name}</h3>
                        <p>
                          {selectedEvidenceHashes.has(record.sha256)
                            ? "Duplicate physical copy excluded: identical content is already selected elsewhere."
                            : record.marker === "N/A"
                            ? "The document is explicitly marked N/A and describes itself as a superseded draft."
                            : "No confirmed statement citation was found."}
                        </p>
                        <span>SHA-256 {shortHash(record.sha256)}</span>
                        {eligibleUnusedAdd(record) ? <button className="secondary-button add-exhibit-button" type="button" onClick={() => openManualAdd("reconcile", record.id)}>Add as exhibit</button> : null}
                      </div>
                      <span className="held-badge">{selectedEvidenceHashes.has(record.sha256) ? "Duplicate copy excluded" : "Excluded"}</span>
                    </article>
                  ))}
                  {renderManualExhibitPanel()}
                  <article className="audit-card">
                    <p className="eyebrow">Reconciliation rule</p>
                    <h2>Ambiguity never becomes an automatic choice.</h2>
                    <p>
                      Dates, titles, filenames, document content and any
                      existing source labels contribute to a suggested match.
                      Close alternatives are always returned to the reviewer.
                    </p>
                  </article>
                </div>
              </>
            )}

            {view === "build" && (
              <>
                <div className="workspace-header">
                  <div>
                    <p className="eyebrow">Final bundle order</p>
                    <h1 tabIndex={-1}>{build ? "Your bundle is ready" : "Finalise the exhibit bundle"}</h1>
                    <p>
                      {build
                        ? "The finished PDF has been reopened and checked."
                        : "Review the final exhibit order, then build the indexed PDF."}
                    </p>
                  </div>
                  {!build && (
                    <div className="workspace-header-actions"><button
                      className="primary-button"
                      type="button"
                      data-tour="build"
                      onClick={() => { if (readyToBuild) void generateBundle(); else if (buildBlockerList[0]) navigateToBlocker(buildBlockerList[0]); }}
                      disabled={busy !== null}
                      aria-disabled={!readyToBuild}
                      aria-describedby={!readyToBuild ? "finalise-build-help" : undefined}
                    >
                      {busy === "build" ? "Building" : reorderReturn ? "Regenerate bundle with new arrangement" : "Build exhibit bundle"}
                    </button>{reorderReturn ? <button className="secondary-button" type="button" disabled={busy !== null} onClick={keepCurrentBundle}>{reorderReturn.build ? "Discard arrangement changes and keep current bundle" : "Discard arrangement changes"}</button> : null}
                    <div className="visually-hidden" role="status" aria-live="polite" aria-atomic="true">
                      {busy === "build" && buildProgress
                        ? `${buildProgress.stage}. ${buildProgress.detail ?? "This may take a little time for long or scanned documents. Keep Exhibit Builder open until the final validation is complete."}`
                        : ""}
                    </div>
                    {busy === "build" && buildProgress && (
                      <section className="build-progress" aria-label="Build progress">
                        <span className="build-progress-spinner" aria-hidden="true" />
                        <div>
                          <p className="eyebrow">Building locally</p>
                          <p className="build-progress-stage">{buildProgress.stage}</p>
                          <p>{buildProgress.detail ?? "This may take a little time for long or scanned documents. Keep Exhibit Builder open until the final validation is complete."}</p>
                          <button className="secondary-button compact-action" type="button" onClick={() => { buildCancelRequested.current = true; setBuildProgress((current) => current ? { ...current, detail: "Stopping safely after the current document or validation pass." } : current); }}>Stop build</button>
                        </div>
                      </section>
                    )}
                    </div>
                  )}
                  {build ? <button className="secondary-button" type="button" onClick={beginOrderChange}>Change exhibit order</button> : null}
                </div>

                {reorderReturn?.build ? <div className="regeneration-warning" role="status"><strong>The current bundle is retained while you edit a new arrangement.</strong><p>Regenerating will replace the current result in this project and update the PDF, index, bookmarks, links, visible page numbers and witness-statement reference suggestions. Files you already downloaded are unaffected. Uploaded sources and review decisions are retained.</p></div> : reorderReturn ? <div className="regeneration-warning" role="status"><strong>The previous PDF is no longer kept.</strong><p>Settings that affect the finished bundle changed after that PDF was built, or a build requirement is outstanding. You can still discard the arrangement edits. Regenerating creates a new PDF that matches the current settings.</p></div> : null}

                {!build && !readyToBuild && <section className="build-readiness finalise-readiness" data-testid="finalise-readiness" id="finalise-build-help">
                  <div className="build-readiness-heading"><div><p className="eyebrow">Build readiness</p><h2>Complete the remaining checks</h2></div><span className="readiness-count needs-attention">{buildBlockerList.length} item{buildBlockerList.length === 1 ? "" : "s"} to resolve</span></div>
                  <p className="readiness-clear">The bundle order can be reviewed now, but the PDF cannot be built until each requirement below is resolved.</p>
                  <div className="readiness-list">{buildBlockerList.map((blocker) => <div key={blocker.id} data-blocker-id={blocker.id} className={`readiness-item ${blocker.kind}`}><span aria-hidden="true">!</span><div className="readiness-copy"><p className="readiness-label"><strong>{blocker.label}</strong></p>{blocker.fileName ? <p className="readiness-file">{blocker.fileName}</p> : null}<p className="readiness-detail">{blocker.detail}</p><div className="readiness-controls"><button className="secondary-button compact-action" type="button" onClick={() => navigateToBlocker(blocker)}>{blocker.actionLabel}</button></div></div></div>)}</div>
                </section>}

                {!build && <section className="finalise-card" aria-label="Final exhibit order">
                  <div className="finalise-card-heading">
                    <div><p className="eyebrow">Included exhibits</p><h2>Choose the exhibit order</h2><p>{hasIndexHeadings ? "Earlier and Later reorder an exhibit inside its current group. Drag a heading to move that group. To place an exhibit under a heading, use Index heading or drag it onto a row in that heading." : "The list below is the order used in the PDF, index and statement-reference suggestions. Drag an exhibit, or use the position buttons, to make a small change."}</p></div>
                    <div className="finalise-card-summary"><span>{exhibitGroups.length} exhibit{exhibitGroups.length === 1 ? "" : "s"}</span><button className="secondary-button add-exhibit-button" type="button" onClick={() => openManualAdd("finalise")}>+ Add an exhibit</button></div>
                  </div>
                  {nonA4Exhibits.length ? <section className="page-size-review" aria-labelledby="page-size-review-title">
                    <div><p className="eyebrow">Page-size review</p><h3 id="page-size-review-title">Choose how to include non-A4 exhibits</h3><p>Exhibit Builder scales PDF content directly and preserves searchable text and vector quality where the source permits. It does not turn the page into a screenshot.</p></div>
                    <div className="page-size-list">{nonA4Exhibits.map(({ choiceKey, name, nonA4, marginCount, annotatedCount }) => { const selectedHandling = pageSizeChoices[choiceKey] ?? (annotatedCount ? "keep-original" : "convert-to-a4"); return <fieldset key={choiceKey}><legend>{name}</legend><p>{nonA4.length} non-A4 page{nonA4.length === 1 ? "" : "s"}.{marginCount ? ` Converting ${marginCount} page${marginCount === 1 ? "" : "s"} will add white margins so no content is cropped or distorted.` : " The page proportions match A4 closely."}{annotatedCount ? ` ${annotatedCount} page${annotatedCount === 1 ? " contains" : "s contain"} PDF annotations, so the original size must be retained to preserve them.` : ""}</p><label><input type="radio" name={`page-size-${choiceKey}`} disabled={Boolean(annotatedCount)} checked={selectedHandling === "convert-to-a4"} onChange={() => setPageSizeChoices((current) => ({ ...current, [choiceKey]: "convert-to-a4" }))} /> <span><strong>Convert to A4{marginCount ? " with margins" : ""} — recommended</strong><small>{annotatedCount ? "Unavailable because scaling could remove or misplace PDF annotations." : "Scale the complete page proportionately. Nothing is stretched or cropped."}</small></span></label><label><input type="radio" name={`page-size-${choiceKey}`} checked={selectedHandling === "keep-original"} onChange={() => setPageSizeChoices((current) => ({ ...current, [choiceKey]: "keep-original" }))} /> <span><strong>Keep the original page size</strong><small>The finished bundle will contain this page at its existing non-A4 dimensions and preserve its PDF annotations.</small></span></label></fieldset>; })}</div>
                  </section> : null}
                  {renderManualExhibitPanel()}
                  {!orderPreview ? <div className="order-toolbar" aria-label="Final order tools">
                    <label>Arrange exhibits by<select value={orderSort} onChange={(event) => setOrderSort(event.target.value as typeof orderSort)}><option value="statement">Statement order</option><option value="date">Document date</option><option value="filename">Source filename</option><option value="description">Index description</option></select></label>
                    <button className="secondary-button" data-testid="preview-order-button" type="button" onClick={previewOrderSort}>Preview this order</button>
                    {orderHistory.length ? <button className="secondary-button" data-testid="undo-order-button" type="button" onClick={undoOrderChange}>Undo order change</button> : null}
                  </div> : <div className="order-preview-banner" data-testid="order-preview-banner" tabIndex={-1} role="status"><div><strong>Preview only - current order unchanged</strong><p>Showing the proposed {orderPreview.label}. Section headings and section membership will not change.</p></div><div><button className="primary-button" type="button" onClick={applyOrderPreview}>Use this order</button><button className="secondary-button" type="button" onClick={cancelOrderPreview}>Cancel preview</button></div></div>}
                  {!orderPreview ? <div className="section-add-control">
                    <div><strong>Add an index heading</strong><small>{hasIndexHeadings ? "A heading prints in the index and as a PDF bookmark only after you place at least one exhibit under it. Exhibits left on “No heading” still print. A new heading starts at the end of the bundle; assigning an exhibit moves it to the end of that heading." : "A heading prints in the index and as a PDF bookmark only after you place at least one exhibit under it. After you add a heading, each row can be moved into it. The heading starts at the end of the bundle, and assigning an exhibit moves that exhibit to the end of the heading."}</small></div>
                    <label><span>Heading</span><input id="new-index-heading" maxLength={512} value={newSectionHeading} onChange={(event) => setNewSectionHeading(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter") { event.preventDefault(); addSection(); } }} placeholder="For example, Agreements" /></label>
                    <button className="secondary-button" type="button" disabled={!newSectionHeading.trim()} onClick={addSection}>Add heading</button>
                  </div> : null}
                  <div className="visually-hidden" role="status" aria-live="polite" aria-atomic="true">{arrangementStatus}</div>
                  {displayedArrangement.nodes.some((node) => node.type === "section") ? <label className="section-jump-control"><span>Jump to index heading</span><select defaultValue="" onChange={(event) => { const sectionId = event.target.value; if (sectionId) focusAfterArrangementChange(`#${CSS.escape(sectionHeadingDomId(sectionId))}-summary`); event.target.value = ""; }}><option value="">Choose a heading</option>{displayedArrangement.nodes.filter((node): node is ArrangementSectionNode => node.type === "section").map((section) => <option key={section.id} value={section.id}>{section.heading} ({section.exhibits.length})</option>)}</select></label> : null}
                  <h3 className="order-list-heading">{orderPreview ? "Proposed order" : "Current bundle order"}</h3>
                  <ul className="finalise-arrangement-list">
                    {renderArrangementNodes()}
                  </ul>
                </section>}

                {!build && <details className="finalise-preflight" aria-label="File and build checks">
                  <summary>Show file and build checks{preflightWarningCount ? ` (${preflightWarningCount} warning${preflightWarningCount === 1 ? "" : "s"})` : ""}</summary>
                  <div className="check-list">
                    {preflight.map((check) => (
                      <div key={check.id}>
                        <span className={check.severity === "pass" ? "check-pass" : "check-warning"} aria-hidden="true">{check.severity === "pass" ? "" : check.severity === "blocking" ? "!" : "!"}</span>
                        <p><strong>{check.label}</strong>{check.fileName ? <small className="readiness-file">{check.fileName}</small> : null}<small>{check.detail}</small></p>
                      </div>
                    ))}
                  </div>
                </details>}
                {build ? (
                  <div className="completed-bundle-flow">
                  <div className="build-layout final-output-layout">
                    <section className="bundle-ready-card">
                      <div className="pdf-preview">
                        <div className="pdf-paper" aria-hidden="true">
                          <span>EXHIBIT BUNDLE</span>
                          <strong>{analysis.caseTitle}</strong>
                          <i />
                          {build.records.slice(0, 5).map((record, index) => (
                            <p key={`${record.sourceHash}-${index}`}>
                              <b>{record.mark}</b>
                              <span>{record.description}</span>
                              <em>{record.startPage}</em>
                            </p>
                          ))}
                          <small>Cover and linked index included</small>
                        </div>
                      </div>
                      <div className="bundle-summary">
                        <span className="success-badge">{retainedWarningCount ? `Bundle built — ${retainedWarningCount} review warning${retainedWarningCount === 1 ? "" : "s"} recorded separately` : "All checks passed"}</span>
                        {retainedWarningCount ? <p className="warning-separation-note"><strong>The warnings are not included in the exhibit bundle PDF.</strong> They appear only in the separate build report and audit information—never as bundle pages, index entries, exhibit annotations or printed warning text.</p> : null}
                        <h2>{build.fileName}</h2>
                        <p data-testid="bundle-output-summary">
                          {build.volumes ? `${build.volumes.length} volumes  ${build.records.length} exhibits  ${build.pageCount} total pages` : `${build.records.length} exhibits  ${build.pageCount} pages`}  searchable index  bookmarks
                        </p>
                        <div className="output-metrics" aria-label="Output summary">
                          <span><b>{templates.filter((template) => usedTemplateSlots.has(template.slot)).length}</b> selected template{templates.filter((template) => usedTemplateSlots.has(template.slot)).length === 1 ? "" : "s"} used</span>
                          <span><b>{(layout.includeDividerPages || layout.includeExhibitCoverPages) && countsOptionalPagesInReferences(pagination, layout) ? "Yes" : "No"}</b> optional pages included in exhibit references</span>
                          <span><b>{includedWorkbookCount}</b> spreadsheet{includedWorkbookCount === 1 ? "" : "s"}</span>
                          <span><b>{unreferencedEvidence.length}</b> source file{unreferencedEvidence.length === 1 ? "" : "s"} not referenced or included</span>
                          <span><b>{resolutions.filter((resolution) => resolution.action === "proceed-without-ocr").length}</b> OCR exception{resolutions.filter((resolution) => resolution.action === "proceed-without-ocr").length === 1 ? "" : "s"}</span>
                        </div>
                        {rebuildComparison ? <div className={`rebuild-comparison ${rebuildComparison.changed ? "changed" : "unchanged"}`}><strong>{rebuildComparison.summary}</strong>{rebuildComparison.categories.length ? <ul>{rebuildComparison.categories.map((category) => <li key={category}>{category}</li>)}</ul> : null}</div> : null}
                        <div className="hash-line">
                          <span>Output SHA-256</span>
                          <code>{shortHash(build.sha256)}</code>
                        </div>
                        {build.volumes ? <div className="volume-hashes">{build.volumes.map((volume) => <p key={volume.number}><span>{volume.label}: {volume.pageCount} pages</span><code>{shortHash(volume.sha256)}</code></p>)}</div> : null}
                        <div className="download-row">
                          {build.volumes ? build.volumes.map((volume) => <button key={volume.number} className="secondary-button" type="button" onClick={() => downloadBytes(volume.bytes, volume.fileName, "application/pdf")}>Download {volume.label}</button>) : null}
                          {build.volumes ? <button className="primary-button" type="button" onClick={() => void downloadVolumeZip()}>Download all volumes (.zip)</button> : null}
                          {!build.volumes ? <button
                            className="primary-button"
                            type="button"
                            data-tour="download"
                            onClick={() => {
                              if (tourActive) setTourDownloaded(true);
                              void downloadBytes(
                                build.bytes,
                                build.fileName,
                                "application/pdf",
                              );
                            }}
                          >
                            Download bundle PDF
                          </button> : null}
                          <button
                            className="secondary-button"
                            type="button"
                            onClick={() =>
                              downloadJson(
                                build.manifest,
                                `${projectName.replace(/[^a-z0-9]+/gi, "_") || "Exhibit_Builder"}_Build_Manifest.json`,
                              )
                            }
                          >
                            Download audit manifest (.json)
                          </button>
                          <button className="primary-button" type="button" onClick={downloadReadableBuildReport}>
                            Download readable report (.txt)
                          </button>
                          <button className="secondary-button" type="button" onClick={downloadTechnicalBuildReport}>
                            Download technical report (.json)
                          </button>
                        </div>
                      </div>
                    </section>

                    <section className="statement-update-card">
                      <div>
                        <p className="eyebrow">Statement update suggestions</p>
                        <h2>Statement reference suggestions</h2>
                        <p>These are suggestions only. The original witness statement remains unchanged.</p>
                      </div>
                      {statementSuggestionSections.ready.length ? <>
                        <h3>Ready to copy</h3>
                        <pre>{statementSuggestionSections.ready.map((item) => item.line).join("\n")}</pre>
                      </> : null}
                      {statementSuggestionSections.verify.length ? <>
                        <h3>Verify the selected source before copying</h3>
                        <pre>{statementSuggestionSections.verify.map((item) => item.line).join("\n")}</pre>
                      </> : null}
                      {statementSuggestionSections.uncited.length ? <>
                        <h3>Uncited exhibits</h3>
                        <pre>{statementSuggestionSections.uncited.map((item) => item.line).join("\n")}</pre>
                      </> : null}
                      {!statementUpdateText ? <pre>No final exhibit page ranges are available.</pre> : null}
                      <div className="download-row">
                        <button className="secondary-button" type="button" onClick={() => void (async () => {
                          const result = await copyPlainText(statementUpdateText);
                          setCopyStatus(result.copied
                            ? { kind: "success", message: statementSuggestionSections.verify.length
                              ? "Copied all suggestions, including lines that need source verification before use."
                              : "Copied to the clipboard." }
                            : { kind: "failure", message: result.detail ?? "Copy failed. Download the .txt file instead." });
                        })()} disabled={!statementUpdateText}>Copy suggestions</button>
                        <button className="secondary-button" type="button" onClick={() => downloadBytes(new TextEncoder().encode(statementUpdateText), "Statement_Update_Suggestions.txt", "text/plain")} disabled={!statementUpdateText}>Download .txt</button>
                      </div>
                      <p className={`copy-feedback${copyStatus?.kind === "failure" ? " is-failure" : ""}`} role="status" aria-live="polite" aria-atomic="true">{copyStatus?.message ?? ""}</p>
                    </section>
                  </div>
                    <details className="completed-checks">
                      <summary>Review final checks{preflightWarningCount ? ` (${preflightWarningCount} warning${preflightWarningCount === 1 ? "" : "s"})` : ""}</summary>
                      <p>These checks were completed before the bundle was made. They do not prevent you from downloading it.</p>
                      <div className="completed-check-grid">
                        <section>
                          <h3>Preflight</h3>
                          <div className="check-list">
                            {preflight.map((check) => <div key={check.id}><span className={check.severity === "pass" ? "check-pass" : "check-warning"} aria-hidden="true">{check.severity === "pass" ? "" : "!"}</span><p><strong>{check.label}</strong>{check.fileName ? <small className="readiness-file">{check.fileName}</small> : null}<small>{check.detail}</small></p></div>)}
                          </div>
                        </section>
                        <section>
                          <h3>Release checks</h3>
                          <div className="check-list">
                            {build.checks.map((check, index) => <div key={`${check.label}-${index}`}><span className={check.status === "pass" ? "check-pass" : "check-warning"} aria-hidden="true">{check.status === "pass" ? "" : "!"}</span><p><strong>{check.label}</strong><small>{check.detail}</small></p></div>)}
                          </div>
                        </section>
                      </div>
                    </details>
                  </div>
                ) : null}
              </>
            )}
          </div>
        </section>
      )}

      {templatePreview && (
        <dialog ref={templatePreviewDialog} className="template-preview-dialog" aria-labelledby="template-preview-title" aria-describedby="template-preview-description">
          <div className="template-preview-heading">
            <div><p className="eyebrow">Template review</p><h2 id="template-preview-title">{templatePreview.name}</h2><p id="template-preview-description">{previewedTemplate?.slot === "cover" && !coverWritesMatterText(layout) ? "Check the visible page. This cover will be used as shown. Names are not rewritten." : "Check the visible page. If the app misread a name or case number, correct it here. Those corrections are printed on the finished bundle."}</p></div>
            <button ref={templatePreviewCloseButton} className="secondary-button" type="button" onClick={() => setTemplatePreview(null)}>Close preview</button>
          </div>
          <div className="template-preview-body">
            <OriginalPdfReview file={templatePreview.file} name={templatePreview.name} purpose="template-preview" onPageRendered={() => setTemplatePreviewLoaded(true)} />
            <aside className="template-matter-panel" aria-live="polite">
              <h3>Matter details to check</h3>
              {previewedTemplate?.slot === "cover" && !coverWritesMatterText(layout) ? <p>This page will be used as shown. There is nothing to amend on an as-supplied cover.</p> : <p>These names will be printed on the finished bundle. The source PDF stays as the background. If the printed page is still the wrong file, choose a different template.</p>}
              {previewedTemplate?.slot === "cover" && !coverWritesMatterText(layout) ? <p>{previewedTemplate.reviewState?.matterReview?.notice ?? "Read the visible cover carefully before confirming it."}</p> : previewedTemplate?.reviewState?.matterReview && matterDraft ? <>
                <div className="template-matter-fields">
                  {matterDraft.occurrences.map((occurrence) => (
                    <label key={occurrence.findingId}>{occurrence.kind === "matter-number" ? "Matter or case number" : occurrence.kind === "party-name" ? "Party name" : occurrence.kind === "forum" ? "Forum or tribunal" : "Matter title"}<input value={occurrence.value} onChange={(event) => updateMatterOccurrence(occurrence.findingId, event.target.value)} /><small>Read as: {occurrence.originalValue}</small></label>
                  ))}
                </div>
                {!matterDraft.occurrences.length ? <p>No matter details were read from this page. Confirm the visible page, or choose a different template.</p> : null}
                {previewedTemplate.reviewState.matterReview.placeholders.length ? <div className="template-placeholder-warning"><strong>Possible placeholders remain visible</strong><ul>{previewedTemplate.reviewState.matterReview.placeholders.map((finding) => <li key={`${finding.id}-${finding.pageNumbers.join("-")}`}>{finding.value} (page {finding.pageNumbers.join(", ")})</li>)}</ul></div> : null}
              </> : <p>{previewedTemplate?.reviewState?.matterReview?.notice ?? "Read the visible template carefully before confirming it."}</p>}
              {!templatePreviewLoaded ? <p className="template-preview-loading">Waiting for the PDF preview to open...</p> : null}
              <div className="template-confirmation-actions">
                {previewedTemplate && previewedTemplate.sourceFormat !== "pdf" ? <button type="button" className="secondary-button" disabled={!templatePreviewLoaded || previewedTemplate.reviewState?.appearanceConfirmation?.pdfSha256 === previewedTemplate.pdfSha256} onClick={() => confirmTemplateReview(previewedTemplate.slot, "appearanceConfirmation")}>Confirm converted appearance</button> : null}
                {previewedTemplate ? <button type="button" className="primary-button" disabled={!templatePreviewLoaded || previewedTemplate.reviewState?.matterConfirmation?.pdfSha256 === previewedTemplate.pdfSha256} onClick={() => confirmTemplateReview(previewedTemplate.slot, "matterConfirmation")}>{previewedTemplate.slot === "cover" && !coverWritesMatterText(layout) ? "Confirm this cover as shown" : "Confirm matter details and party names"}</button> : null}
                {previewedTemplate?.reviewState?.matterReview?.placeholders.length ? <button type="button" className="secondary-button" disabled={!templatePreviewLoaded || previewedTemplate.reviewState.placeholderConfirmation?.pdfSha256 === previewedTemplate.pdfSha256} onClick={() => confirmTemplateReview(previewedTemplate.slot, "placeholderConfirmation")}>Confirm these placeholders are intentional</button> : null}
              </div>
            </aside>
          </div>
        </dialog>
      )}

      {ocrSourcePreview && <dialog ref={ocrPreviewDialog} className="template-preview-dialog source-review-dialog" aria-labelledby="ocr-source-preview-title" aria-describedby="ocr-source-preview-description">
        <div className="template-preview-heading"><div><p className="eyebrow">Original source review</p><h2 id="ocr-source-preview-title">{ocrSourcePreview.name}</h2><p id="ocr-source-preview-description">Inspect the original PDF itself. This preview uses the tool's local PDF renderer and does not add OCR or change the source document.</p></div><button ref={ocrPreviewCloseButton} className="secondary-button" type="button" onClick={() => setOcrSourcePreview(null)}>Close preview</button></div>
        <OriginalPdfReview file={ocrSourcePreview.file} name={ocrSourcePreview.name} accessibleText={ocrSourcePreview.extractedText} onPageRendered={() => setVisuallyReviewedSourceHashes((current) => current.has(ocrSourcePreview.sourceSha256) ? current : new Set(current).add(ocrSourcePreview.sourceSha256))} />
      </dialog>}

      {orderChangeConfirmation ? <div className="confirmation-backdrop" role="presentation"><section className="confirmation-dialog" role="dialog" aria-modal="true" aria-labelledby="order-change-confirmation-title"><p className="eyebrow">Exhibit order</p><h2 id="order-change-confirmation-title">Change the exhibit order?</h2><p>Your uploaded documents and review decisions will be kept. If you save a changed order, the existing PDF, index, bookmarks, links and witness-statement page-reference suggestions must be regenerated.</p><div className="download-row"><button className="primary-button" type="button" onClick={confirmOrderChange}>Change exhibit order</button><button className="secondary-button" type="button" onClick={cancelOrderChange}>Cancel</button></div></section></div> : null}

      {paginationConfirmation ? <div className="confirmation-backdrop" role="presentation"><section className="confirmation-dialog" role="dialog" aria-modal="true" aria-labelledby="numbering-confirmation-title"><p className="eyebrow">Numbering warning</p><h2 id="numbering-confirmation-title">The printed page label will include a prefix, suffix or padding</h2><p className="confirmation-example">{numberingDifferenceExample(paginationConfirmation)}</p><p>That printed label is what appears on the finished PDF. The index and witness-statement suggestions use the same label.</p><div className="download-row"><button className="primary-button" type="button" onClick={acceptPaginationChange}>Use this numbering</button><button className="secondary-button" type="button" onClick={cancelPaginationChange}>Cancel</button></div></section></div> : null}

      {volumeNumberingConfirmation ? <div className="confirmation-backdrop" role="presentation"><section className="confirmation-dialog" role="dialog" aria-modal="true" aria-labelledby="volume-numbering-confirmation-title"><p className="eyebrow">Repeated page-number warning</p><h2 id="volume-numbering-confirmation-title">Page numbers will restart in every volume</h2><p className="confirmation-example">Volume 1: AH-001…AH-200<br />Volume 2: AH-001…AH-200</p><p>The same visible page number may appear in more than one PDF. Every index entry and legal reference must therefore be read together with its volume number.</p><p>Continuous numbering is safer when pages may be printed, extracted or circulated separately.</p><div className="download-row"><button className="primary-button" type="button" onClick={acceptVolumeNumberingRestart}>Restart in each volume</button><button className="secondary-button" type="button" onClick={() => setVolumeNumberingConfirmation(false)}>Keep continuous numbering</button></div></section></div> : null}

      {initialsConfirmation ? <div className="confirmation-backdrop" role="presentation"><section className="confirmation-dialog" role="dialog" aria-modal="true" aria-labelledby="initials-confirmation-title"><p className="eyebrow">Placeholder mismatch</p><h2 id="initials-confirmation-title">Use one mark throughout this bundle?</h2><p>The statement contains explicit placeholder initials <strong>{initialsConfirmation.existing.join(", ")}</strong>, but you selected <strong>{initialsConfirmation.proposed}</strong>.</p><p>The generated bundle and suggestions will use <strong>{initialsConfirmation.proposed}</strong> throughout. The source statement remains unchanged, so update any different placeholders separately if needed.</p><div className="download-row"><button className="primary-button" type="button" onClick={() => applyStatementDraft(initialsConfirmation.statementId, true)}>Use {initialsConfirmation.proposed} throughout this bundle</button><button className="secondary-button" type="button" onClick={() => setInitialsConfirmation(null)}>Cancel</button></div></section></div> : null}

      {bulkConfirmationOpen ? <div className="confirmation-backdrop" role="presentation"><section className="confirmation-dialog bulk-confirmation-dialog" role="dialog" aria-modal="true" aria-labelledby="bulk-confirmation-title"><p className="eyebrow">Bulk confirmation</p><h2 id="bulk-confirmation-title">Confirm only the matches you have reviewed</h2><p>The following {bulkConfirmableCount} proposed document match{bulkConfirmableCount === 1 ? " is" : "es are"} ready. Repeat conflicts, unmatched references and automatic matches below {AUTOMATIC_MATCH_REVIEW_THRESHOLD} are not included.</p><div className="bulk-confirmation-list">{bulkConfirmableCandidates.map((candidate) => <div key={candidate.id}><span><strong>{candidate.description}</strong><small>{analysis?.evidence.find((record) => record.id === candidate.evidenceId)?.name ?? "No document selected"}</small></span><b>{matchStrengthLabel(candidate.confidence, candidate.rationale)}</b></div>)}</div><label className="confirmation-acknowledgement"><input type="checkbox" checked={bulkConfirmationAcknowledged} onChange={(event) => setBulkConfirmationAcknowledged(event.target.checked)} /> I have reviewed the proposed document for each listed exhibit.</label><div className="download-row"><button className="primary-button" type="button" disabled={!bulkConfirmationAcknowledged} onClick={confirmReviewedMatches}>Confirm these reviewed matches</button><button className="secondary-button" type="button" onClick={() => { setBulkConfirmationOpen(false); setBulkConfirmationAcknowledged(false); }}>Cancel</button></div></section></div> : null}

      {recoveryDataDialogOpen ? <div className="confirmation-backdrop" role="presentation"><section className="confirmation-dialog" role="dialog" aria-modal="true" aria-labelledby="recovery-data-title"><p className="eyebrow">Local recovery data</p><h2 id="recovery-data-title">Delete automatic recovery data?</h2><p>{recoveryDataStored ? "This computer currently holds an automatic recovery journal for Exhibit Builder." : "No automatic recovery journal is currently stored on this computer."}</p><p>Deleting it removes only Exhibit Builder's local crash-recovery metadata. It does not delete source documents, saved exhibit projects, finished bundles or your guided-sample preference.</p>{recoveryDataStored ? <label className="confirmation-acknowledgement"><input type="checkbox" checked={recoveryDeleteAcknowledged} onChange={(event) => setRecoveryDeleteAcknowledged(event.target.checked)} /> I understand that an unfinished project cannot be restored from this recovery data after deletion.</label> : null}<div className="download-row">{recoveryDataStored ? <button className="primary-button destructive-action" type="button" disabled={!recoveryDeleteAcknowledged} onClick={() => void clearLocalRecoveryData()}>Delete recovery data</button> : null}<button className="secondary-button" type="button" onClick={() => { setRecoveryDataDialogOpen(false); setRecoveryDeleteAcknowledged(false); }}>Close</button></div></section></div> : null}

      {error && (
        <div className="error-toast" role="alert">
          <strong>Action needed</strong>
          <span>{error}</span>
          <button type="button" onClick={() => setError(null)}>
            Dismiss
          </button>
        </div>
      )}
      {tourActive ? <GuidedSampleTour stepId={tourStep} dialogOpen={Boolean(confirmationDialogKey)} onSkip={skipGuidedSampleTour} /> : null}
    </main>
  );
}
