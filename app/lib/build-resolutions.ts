import type {
  BuildResolution,
  PreflightCheck,
  TemplateFile,
  TemplateSlot,
} from "./bundle-types.ts";

export function isOcrCheck(check: Pick<PreflightCheck, "label">) {
  return check.label === "OCR failed" || check.label === "OCR unavailable";
}

export function resolutionMatchesCheck(
  resolution: BuildResolution,
  check: Pick<PreflightCheck, "id" | "code" | "policy" | "sourceId" | "sourceSha256">,
) {
  const sourceWideOcrMatch = resolution.action === "proceed-without-ocr" && resolution.checkCode?.startsWith("ocr.") && check.code?.startsWith("ocr.");
  if (resolution.blockerId !== check.id && !sourceWideOcrMatch) return false;
  if (sourceWideOcrMatch && (resolution.sourceId !== check.sourceId || resolution.sourceSha256 !== check.sourceSha256)) return false;
  if (resolution.checkCode && resolution.checkCode !== check.code) return false;
  // A source-bound resolution must never apply to a check which has lost its
  // source identity.  This makes replacement of a file invalidate the choice.
  if (resolution.sourceId && resolution.sourceId !== check.sourceId) return false;
  if (resolution.sourceSha256 && resolution.sourceSha256 !== check.sourceSha256) return false;
  return true;
}

export function findResolution(
  resolutions: BuildResolution[],
  check: Pick<PreflightCheck, "id" | "code" | "policy" | "sourceId" | "sourceSha256">,
) {
  return resolutions.find((resolution) => resolutionMatchesCheck(resolution, check));
}

/**
 * Applies only safe, explicit technical exceptions to preflight.  Legal
 * identity/numbering/repeat decisions remain blocking regardless of what is
 * stored in the project.
 */
export function applyBuildResolutions(
  checks: PreflightCheck[],
  resolutions: BuildResolution[] = [],
) {
  const resolved = checks.map((check) => {
    const resolution = findResolution(resolutions, check);
    if (!resolution) return check;
    if (resolution.action === "proceed-without-ocr" && isOcrCheck(check) && check.policy === "exception-eligible" && resolution.visualReviewConfirmed === true) {
      return {
        ...check,
        severity: "warning" as const,
        detail: `${check.detail} Approved exception: the original PDF will be included without a tool-generated OCR text layer.`,
      };
    }
    return check;
  });

  if (!resolved.some((check) => check.severity === "blocking") && !resolved.some((check) => check.id === "ready" || check.id === "ready-after-exceptions")) {
    resolved.push({
      id: resolutions.length ? "ready-after-exceptions" : "ready",
      severity: "pass",
      label: resolutions.length ? "Ready with approved exceptions" : "Ready to build",
      detail: resolutions.length
        ? "No unresolved blocking issue remains; one or more technical exceptions are recorded below."
        : "No blocking preflight issue was found.",
    });
  }
  return resolved;
}

export function sourceIsExcluded(
  resolutions: BuildResolution[],
  sourceId: string,
  sourceSha256: string,
) {
  return resolutions.some(
    (resolution) =>
      resolution.action === "exclude-source" &&
      resolution.sourceId === sourceId &&
      resolution.sourceSha256 === sourceSha256,
  );
}

export function candidateIsExcluded(
  resolutions: BuildResolution[],
  candidateId: string,
) {
  return resolutions.some(
    (resolution) =>
      resolution.action === "exclude-candidate" &&
      resolution.candidateId === candidateId,
  );
}

export function templateFallbackSlots(
  resolutions: BuildResolution[],
  templates?: Array<Pick<TemplateFile, "slot" | "sha256">>,
) {
  const slots = new Set<TemplateSlot>();
  for (const resolution of resolutions) {
      if (resolution.action !== "use-built-in-template") continue;
    for (const slot of resolution.templateSlots ?? []) {
      const expectedHash = resolution.templateHashes?.[slot];
      const current = templates?.find((template) => template.slot === slot);
      // A fallback decision is an approval for one exact custom template.  Do
      // not let legacy or hand-crafted resolutions silently replace a newly
      // selected template (or a missing template) without that binding.
      if (!expectedHash || !current || expectedHash !== current.sha256) continue;
      slots.add(slot);
    }
  }
  return slots;
}

export function shouldSkipOcr(
  sourceId: string,
  sourceSha256: string,
  resolutions: BuildResolution[],
) {
  return resolutions.some(
    (resolution) =>
      resolution.action === "proceed-without-ocr" &&
      resolution.sourceId === sourceId &&
      resolution.sourceSha256 === sourceSha256 &&
      (resolution.checkCode?.startsWith("ocr.") || resolution.blockerId === `ocr-${sourceId}`),
  );
}
