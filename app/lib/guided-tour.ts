export const TOUR_STEP_IDS = [
  "open-folder",
  "choose-statement",
  "choose-evidence",
  "analyse",
  "confirm-all",
  "repeat-decision",
  "attachments",
  "print-with-email",
  "continue-sheets",
  "continue-finalise",
  "build",
  "download",
  "save",
] as const;

export type TourStepId = (typeof TOUR_STEP_IDS)[number];

export type TourWorkspaceView = "home" | "review" | "sheets" | "build" | "other";

export type TourContext = {
  active: boolean;
  openedFolder: boolean;
  statementSelected: boolean;
  evidenceSelected: boolean;
  analysisReady: boolean;
  bulkConfirmableCount: number;
  repeatPending: boolean;
  attachmentPending: boolean;
  attachmentChoicesOpen: boolean;
  printWithEmailVisible: boolean;
  view: TourWorkspaceView;
  includedWorkbook: boolean;
  hasBuild: boolean;
  downloaded: boolean;
  saved: boolean;
};

export const TOUR_COPY: Record<TourStepId, { title: string; body: string }> = {
  "open-folder": {
    title: "Open the sample folder",
    body: "Open the guided sample folder in File Explorer. Read the witness statement, then come back here.",
  },
  "choose-statement": {
    title: "Choose the witness statement",
    body: "Click Choose witness statement and select 01_GUIDED_SAMPLE_Witness_Statement.docx from Witness statement.",
  },
  "choose-evidence": {
    title: "Choose the evidence files",
    body: "Click Choose evidence files and select the files in the Exhibits folder.",
  },
  analyse: {
    title: "Analyse the files",
    body: "Click Analyse files. The sample matches are mapped so you can practise confirmation.",
  },
  "confirm-all": {
    title: "Confirm the proposed matches",
    body: "Click Confirm all, then confirm the reviewed matches in the dialog.",
  },
  "repeat-decision": {
    title: "Confirm the repeated reference",
    body: "Click Confirm repeat decision. Same exhibit is already selected.",
  },
  attachments: {
    title: "Open attachment choices",
    body: "Click Open attachment choices on this email.",
  },
  "print-with-email": {
    title: "Choose a disposition",
    body: "Click Print with this email for the sample attachment.",
  },
  "continue-sheets": {
    title: "Continue to workbook sheets",
    body: "Click Continue to workbook sheets.",
  },
  "continue-finalise": {
    title: "Continue to finalise",
    body: "Click Continue to finalise.",
  },
  build: {
    title: "Build the bundle",
    body: "Click Build exhibit bundle.",
  },
  download: {
    title: "Download the PDF",
    body: "Click Download bundle PDF.",
  },
  save: {
    title: "Save the project",
    body: "Click Save exhibit project.",
  },
};

function homeStep(ctx: TourContext): TourStepId | null {
  if (ctx.analysisReady) return null;
  if (!ctx.openedFolder && !ctx.statementSelected) return "open-folder";
  if (!ctx.statementSelected) return "choose-statement";
  if (!ctx.evidenceSelected) return "choose-evidence";
  return "analyse";
}

function reviewStep(ctx: TourContext): TourStepId | null {
  if (!ctx.analysisReady) return null;
  if (ctx.bulkConfirmableCount > 0) return "confirm-all";
  if (ctx.repeatPending) return "repeat-decision";
  if (ctx.attachmentPending) {
    if (ctx.printWithEmailVisible || ctx.attachmentChoicesOpen) return "print-with-email";
    return "attachments";
  }
  if (ctx.view === "review") return ctx.includedWorkbook ? "continue-sheets" : "continue-finalise";
  return null;
}

function finishStep(ctx: TourContext): TourStepId | null {
  if (!ctx.analysisReady) return null;
  if (ctx.view === "sheets") return "continue-finalise";
  if (ctx.view !== "build" && ctx.view !== "other") return null;
  if (!ctx.hasBuild) return "build";
  if (!ctx.downloaded) return "download";
  if (!ctx.saved) return "save";
  return null;
}

export function resolveTourStep(ctx: TourContext): TourStepId | null {
  if (!ctx.active) return null;
  return homeStep(ctx) ?? reviewStep(ctx) ?? finishStep(ctx);
}

export function tourWorkspaceView(
  analysisReady: boolean,
  view: "sources" | "sheets" | "review" | "reconcile" | "build",
): TourWorkspaceView {
  if (!analysisReady) return "home";
  if (view === "review") return "review";
  if (view === "sheets") return "sheets";
  if (view === "build") return "build";
  return "other";
}
