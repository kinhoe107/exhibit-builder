import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import test from "node:test";

import diagnosticStream from "../electron/diagnostic-stream.cjs";

test("only EPIPE is ignored; another stream failure is recorded and fatal", () => {
  const stream = new EventEmitter();
  const recorded = [];
  const fatal = [];
  diagnosticStream.installDiagnosticStreamErrorHandlers([stream], {
    record: (error) => recorded.push(error.code),
    fatal: (error) => fatal.push(error.code),
  });
  stream.emit("error", Object.assign(new Error("closed pipe"), { code: "EPIPE" }));
  assert.deepEqual(recorded, []);
  assert.deepEqual(fatal, []);
  stream.emit("error", Object.assign(new Error("bad descriptor"), { code: "EBADF" }));
  assert.deepEqual(recorded, ["EBADF"]);
  assert.deepEqual(fatal, ["EBADF"]);
  stream.emit("error", Object.assign(new Error("input/output failure"), { code: "EIO" }));
  assert.deepEqual(recorded, ["EBADF", "EIO"], "every non-EPIPE error remains durably observable");
  assert.deepEqual(fatal, ["EBADF"], "one controlled fatal path is enough");
});
