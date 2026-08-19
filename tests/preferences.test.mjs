import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import preferences from "../electron/preferences.cjs";

const { PreferenceStore, validatePreferences } = preferences;

test("guided-sample visibility is validated and persists outside the disposable browser profile", async () => {
  const root = await mkdtemp(path.join(tmpdir(), "exhibit-preferences-test-"));
  try {
    const store = new PreferenceStore(root);
    assert.deepEqual(await store.read(), { hideGuidedSample: false });
    assert.deepEqual(await store.write({ hideGuidedSample: true }), { hideGuidedSample: true });
    assert.deepEqual(await store.read(), { hideGuidedSample: true });
    assert.deepEqual(JSON.parse(await readFile(path.join(root, "preferences.json"), "utf8")), { hideGuidedSample: true });

    await writeFile(path.join(root, "preferences.json"), "not json", "utf8");
    assert.deepEqual(await store.read(), { hideGuidedSample: false });
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("unknown or malformed preference fields cannot alter the supported setting", () => {
  assert.deepEqual(validatePreferences(null), { hideGuidedSample: false });
  assert.deepEqual(validatePreferences({ hideGuidedSample: "yes", unrelated: true }), { hideGuidedSample: false });
});
