import {
  lockPagination,
  type NonA4PageHandling,
  type TemplateDiscrepancyConfirmation,
} from "./bundle-types.ts";
import { reconcileBundleArrangement, type BundleArrangement } from "./bundle-arrangement.ts";
import {
  nonArrangementSubstantiveKey,
  substantiveBuildCanonical,
  type BuildFingerprintInput,
} from "./rebuild-comparison.ts";

export type RetainedBuildInputs = {
  substantiveKey: string;
  templateDiscrepancyConfirmation: TemplateDiscrepancyConfirmation | null;
};

export type RetainedBundle<TBuild, TArrangement, TCandidate> = {
  build: TBuild | null;
  arrangement: TArrangement;
  candidates: TCandidate[];
  pageSizeChoices: Record<string, NonA4PageHandling>;
  inputs: RetainedBuildInputs;
};

export type LiveReviewState<TCandidate> = {
  candidates: TCandidate[];
  pageSizeChoices: Record<string, NonA4PageHandling>;
  exhibitIds: readonly string[];
};

function stableInputs(inputs: RetainedBuildInputs): string {
  return JSON.stringify({
    substantiveKey: inputs.substantiveKey,
    templateDiscrepancyConfirmation: inputs.templateDiscrepancyConfirmation,
  });
}

/** Snapshot every non-arrangement input that a finished PDF was built from. */
export function retainedBuildInputsFrom(
  input: BuildFingerprintInput & { templateDiscrepancyConfirmation?: TemplateDiscrepancyConfirmation | null },
): RetainedBuildInputs {
  return {
    substantiveKey: nonArrangementSubstantiveKey(substantiveBuildCanonical({
      ...input,
      pagination: lockPagination(input.pagination),
    })),
    templateDiscrepancyConfirmation: input.templateDiscrepancyConfirmation ?? null,
  };
}

export function retainedBuildInputsMatch(left: RetainedBuildInputs, right: RetainedBuildInputs) {
  return stableInputs(left) === stableInputs(right);
}

/** Preview-only blockers disable generate without invalidating a captured PDF. */
export const TRANSIENT_RETAIN_READINESS_BLOCKERS = new Set(["order-preview-pending"]);

export function retainBuildReadiness(blockers: Array<{ id: string }>) {
  return blockers.every((blocker) => TRANSIENT_RETAIN_READINESS_BLOCKERS.has(blocker.id));
}

export function captureRetainedBundle<TBuild, TArrangement, TCandidate>(input: {
  build: TBuild;
  arrangement: TArrangement;
  candidates: TCandidate[];
  pageSizeChoices: Record<string, NonA4PageHandling>;
  inputs: RetainedBuildInputs;
}): RetainedBundle<TBuild, TArrangement, TCandidate> {
  return {
    build: input.build,
    arrangement: input.arrangement,
    candidates: input.candidates,
    pageSizeChoices: input.pageSizeChoices,
    inputs: input.inputs,
  };
}

/**
 * Keep arrangement restore, but drop the finished PDF when any non-arrangement
 * substantive input no longer matches, or when a non-transient readiness
 * blocker means the captured PDF cannot be regenerated.
 */
export function dropStaleRetainedBuild<TBuild, TArrangement, TCandidate>(
  retained: RetainedBundle<TBuild, TArrangement, TCandidate> | null,
  currentInputs: RetainedBuildInputs,
  options: { readyToBuild?: boolean } = {},
): RetainedBundle<TBuild, TArrangement, TCandidate> | null {
  if (!retained) return null;
  const readyToBuild = options.readyToBuild !== false;
  if (retained.build === null || (readyToBuild && retainedBuildInputsMatch(retained.inputs, currentInputs))) return retained;
  return { ...retained, build: null };
}

/**
 * Matching arrangement-only keep restores the captured PDF and review snapshot.
 * After the PDF is dropped, restore only the captured arrangement against live
 * exhibit IDs and keep the current candidate and page-size state.
 */
export function restoredBundleFromRetain<TBuild, TArrangement, TCandidate>(
  retained: RetainedBundle<TBuild, TArrangement, TCandidate>,
  live?: LiveReviewState<TCandidate>,
): {
  arrangement: TArrangement;
  candidates: TCandidate[];
  pageSizeChoices: Record<string, NonA4PageHandling>;
  build: TBuild | null;
} {
  if (retained.build) {
    return {
      arrangement: retained.arrangement,
      candidates: retained.candidates,
      pageSizeChoices: retained.pageSizeChoices,
      build: retained.build,
    };
  }
  const arrangement = live?.exhibitIds
    ? reconcileBundleArrangement(retained.arrangement as BundleArrangement, live.exhibitIds) as TArrangement
    : retained.arrangement;
  return {
    arrangement,
    candidates: live?.candidates ?? retained.candidates,
    pageSizeChoices: live?.pageSizeChoices ?? retained.pageSizeChoices,
    build: null,
  };
}
