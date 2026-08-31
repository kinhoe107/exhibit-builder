export {};

declare global {
  interface Window {
    bundleBuilderDesktop?: {
      saveFile: (
        bytes: Uint8Array,
        fileName: string,
        mediaType: string,
        options?: { protectedSourcePaths?: string[]; allowedOverwritePath?: string },
      ) => Promise<{ saved: boolean; filePath?: string }>;
      copyText: (text: string) => Promise<{ copied: boolean }>;
      convertTemplate: (
        html: string,
        sourceName: string,
      ) => Promise<Uint8Array>;
      exportWorkbook: (
        fileName: string,
        bytes: Uint8Array,
        sheets: Array<{ name: string; range: string; orientation: "portrait" | "landscape"; scalePercent: number; columnBreaks: number[]; rowBreaks: number[]; titleRows: string; titleColumns: string; pageOrder: "downThenOver"; margins: { left: number; right: number; top: number; bottom: number; header: number; footer: number }; printableWidthPoints: number; printableHeightPoints: number; geometryChecks: Array<{ axis: "horizontal" | "vertical"; ranges: string[] }>; expectedPageCount: number }>,
      ) => Promise<Array<{ name: string; range: string; orientation?: "portrait" | "landscape"; bytes: Uint8Array }>>;
      sourcePath: (file: File) => string;
      readPreferences: () => Promise<{ hideGuidedSample: boolean }>;
      setGuidedSampleHidden: (hidden: boolean) => Promise<{ hideGuidedSample: boolean }>;
      openGuidedSampleFolder: () => Promise<{ opened: boolean }>;
      recoveryStatus: () => Promise<{ available: boolean; stored: boolean; corrupt?: boolean; issue?: string; recoveryId?: string; revision?: number; projectName?: string }>;
      beginRecovery: () => Promise<{ recoveryId: string; revision: number }>;
      writeRecovery: (recoveryId: string, revision: number, payload: unknown) => Promise<{ saved: boolean; revision: number }>;
      loadRecovery: (recoveryId: string) => Promise<{ recoveryId: string; revision: number; dirty: boolean; payload: any; savedArchive?: { path: string; sha256: string } | null }>;
      readRecoverySource: (recoveryId: string, sourceId: string) => Promise<{ id: string; role: "statement" | "evidence" | "template" | "project"; name: string; sha256: string; bytes: Uint8Array }>;
      discardRecovery: (recoveryId: string) => Promise<{ discarded: boolean }>;
      clearRecoveryData: () => Promise<{ cleared: boolean }>;
      markRecoveryClean: (recoveryId: string, revision: number, savedArchive?: { path: string; sha256: string } | null) => Promise<{ cleaned: boolean; revision?: number }>;
    };
  }
}
