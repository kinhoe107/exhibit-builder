const { mkdir, readFile, rename, writeFile } = require("node:fs/promises");
const path = require("node:path");

const DEFAULT_PREFERENCES = Object.freeze({ hideGuidedSample: false });

function validatePreferences(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return { ...DEFAULT_PREFERENCES };
  return {
    hideGuidedSample: typeof value.hideGuidedSample === "boolean" ? value.hideGuidedSample : false,
  };
}

class PreferenceStore {
  constructor(root) {
    this.root = root;
    this.filePath = path.join(root, "preferences.json");
    this.writeQueue = Promise.resolve();
  }

  async read() {
    try {
      return validatePreferences(JSON.parse(await readFile(this.filePath, "utf8")));
    } catch {
      return { ...DEFAULT_PREFERENCES };
    }
  }

  write(change) {
    const pending = this.writeQueue.then(async () => {
      const current = await this.read();
      const next = validatePreferences({ ...current, ...change });
      await mkdir(this.root, { recursive: true });
      const temporary = `${this.filePath}.tmp`;
      await writeFile(temporary, `${JSON.stringify(next, null, 2)}\n`, "utf8");
      await rename(temporary, this.filePath);
      return next;
    });
    this.writeQueue = pending.catch(() => {});
    return pending;
  }
}

module.exports = { DEFAULT_PREFERENCES, PreferenceStore, validatePreferences };
