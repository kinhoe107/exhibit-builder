const { contextBridge, ipcRenderer, webUtils } = require("electron");

contextBridge.exposeInMainWorld("bundleBuilderDesktop", {
  saveFile(bytes, fileName, mediaType) {
    return ipcRenderer.invoke("bundle-builder:save-file", {
      bytes,
      fileName,
      mediaType,
    });
  },
  copyText(text) {
    return ipcRenderer.invoke("bundle-builder:clipboard-write", text);
  },
  convertTemplate(html, sourceName) {
    return ipcRenderer.invoke("bundle-builder:convert-template", {
      html,
      sourceName,
    });
  },
  exportWorkbook(fileName, bytes, sheets) {
    return ipcRenderer.invoke("bundle-builder:export-workbook", { fileName, bytes, sheets });
  },
  sourcePath(file) {
    try {
      return webUtils.getPathForFile(file) || "";
    } catch {
      return "";
    }
  },
  readPreferences() {
    return ipcRenderer.invoke("bundle-builder:preferences-read");
  },
  setGuidedSampleHidden(hidden) {
    return ipcRenderer.invoke("bundle-builder:guided-sample-hidden", hidden);
  },
  recoveryStatus() {
    return ipcRenderer.invoke("bundle-builder:recovery-status");
  },
  beginRecovery() {
    return ipcRenderer.invoke("bundle-builder:recovery-begin");
  },
  writeRecovery(recoveryId, revision, payload) {
    return ipcRenderer.invoke("bundle-builder:recovery-write", { recoveryId, revision, payload });
  },
  loadRecovery(recoveryId) {
    return ipcRenderer.invoke("bundle-builder:recovery-load", { recoveryId });
  },
  readRecoverySource(recoveryId, sourceId) {
    return ipcRenderer.invoke("bundle-builder:recovery-read-source", { recoveryId, sourceId });
  },
  discardRecovery(recoveryId) {
    return ipcRenderer.invoke("bundle-builder:recovery-discard", { recoveryId });
  },
  clearRecoveryData() {
    return ipcRenderer.invoke("bundle-builder:recovery-clear-all");
  },
  markRecoveryClean(recoveryId, revision, savedArchive) {
    return ipcRenderer.invoke("bundle-builder:recovery-mark-clean", { recoveryId, revision, savedArchive });
  },
});
