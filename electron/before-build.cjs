// The desktop application ships a fully bundled renderer and uses only
// Electron/Node built-ins from the main process.  Tell electron-builder not
// to run a package-manager dependency collector for the packaged artifact.
// This keeps packaging reproducible with pnpm while leaving the runtime
// artifact intentionally free of external node_modules.
module.exports = async function beforeBuild() {
  return false;
};
