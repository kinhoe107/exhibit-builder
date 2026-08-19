"use strict";

function installDiagnosticStreamErrorHandlers(streams, { record, fatal }) {
  let fatalHandled = false;
  const handler = (error) => {
    if (error?.code === "EPIPE") return;
    try {
      record(error);
    } catch {
      // The fallback recorder must not replace the original stream failure.
    }
    if (fatalHandled) return;
    fatalHandled = true;
    fatal(error);
  };
  for (const stream of streams) stream?.on?.("error", handler);
  return handler;
}

module.exports = { installDiagnosticStreamErrorHandlers };
