import assert from "node:assert/strict";
import { readFile, stat } from "node:fs/promises";
import test from "node:test";

test("ships a multi-size Windows icon and wires it into the desktop shell", async () => {
  const [ico, png, favicon, main, installer, html, packageJson] = await Promise.all([
    readFile(new URL("../build/icon.ico", import.meta.url)),
    readFile(new URL("../build/icon.png", import.meta.url)),
    readFile(new URL("../public/favicon.ico", import.meta.url)),
    readFile(new URL("../electron/main.cjs", import.meta.url), "utf8"),
    readFile(new URL("../build/standalone-installer.nsi", import.meta.url), "utf8"),
    readFile(new URL("../index.html", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
  ]);
  const manifest = JSON.parse(packageJson);

  assert.equal(ico.readUInt16LE(0), 0);
  assert.equal(ico.readUInt16LE(2), 1, "the file must be an ICO, not a CUR");
  assert.equal(ico.readUInt16LE(4), 7, "the ICO must include 16, 24, 32, 48, 64, 128 and 256 pixel frames");
  assert.ok(ico.byteLength > 4000, "a real multi-size ICO is larger than a stub directory");
  assert.ok(ico.includes(Buffer.from([0x89, 0x50, 0x4e, 0x47])), "each frame is stored as PNG");

  assert.equal(png[0], 0x89);
  assert.equal(png.toString("ascii", 1, 4), "PNG");
  assert.ok((await stat(new URL("../build/icon.png", import.meta.url))).size > 1000);

  assert.equal(favicon.readUInt16LE(2), 1);
  assert.ok(favicon.readUInt16LE(4) >= 2);

  assert.equal(manifest.build.win.icon, "build/icon.ico");
  assert.equal(manifest.build.win.signExecutable, false);
  assert.equal(manifest.build.win.signAndEditExecutable, undefined);
  assert.ok(manifest.build.extraFiles.some((entry) => entry.from === "build/icon.ico" && entry.to === "icon.ico"));
  assert.ok(manifest.build.extraFiles.some((entry) => entry.from === "public/guided-sample" && entry.to === "Guided Sample"));
  assert.match(main, /function resolveAppIcon\(\)/);
  assert.match(main, /icon: resolveAppIcon\(\)/);
  assert.match(main, /path\.join\(__dirname, "\.\.", "build", "icon\.ico"\)/);
  assert.match(main, /path\.join\(path\.dirname\(process\.execPath\), "icon\.ico"\)/);
  assert.match(installer, /!define MUI_ICON "\$\{SOURCE_DIR\}\\icon\.ico"/);
  assert.match(installer, /!define MUI_UNICON "\$\{SOURCE_DIR\}\\icon\.ico"/);
  assert.match(installer, /DisplayIcon" "\$INSTDIR\\Exhibit Builder\.exe"/);
  assert.match(html, /rel="icon" href="\/favicon\.ico"/);
});
