const { app, BrowserWindow } = require("electron");
const { mkdirSync, writeFileSync } = require("node:fs");
const path = require("node:path");

const outDir = path.join(__dirname, "..", "public", "guided-sample", "readme-images");
mkdirSync(outDir, { recursive: true });

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

app.whenReady().then(async () => {
  const win = new BrowserWindow({
    width: 1280,
    height: 900,
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  });
  await win.loadURL("http://localhost:5173");
  const page = win.webContents;

  async function waitFor(expression, timeout = 60_000) {
    const end = Date.now() + timeout;
    while (Date.now() < end) {
      if (await page.executeJavaScript(`Boolean(${expression})`)) return;
      await delay(120);
    }
    throw new Error(`Timed out waiting for ${expression}`);
  }

  async function shot(name, selector) {
    if (selector) {
      const rect = await page.executeJavaScript(`(() => {
        const node = document.querySelector(${JSON.stringify(selector)});
        if (!node) return null;
        node.scrollIntoView({ block: "center" });
        const box = node.getBoundingClientRect();
        return { x: Math.max(0, box.x - 12), y: Math.max(0, box.y - 12), width: box.width + 24, height: box.height + 24 };
      })()`);
      if (!rect || rect.width < 8 || rect.height < 8) throw new Error(`Missing ${selector} for ${name}`);
      const image = await page.capturePage(rect);
      writeFileSync(path.join(outDir, name), image.toPNG());
      return;
    }
    const image = await page.capturePage();
    writeFileSync(path.join(outDir, name), image.toPNG());
  }

  await waitFor(`document.querySelector('[data-testid="guided-sample-button"]') && document.querySelector('[data-testid="guided-sample-tour-button"]')`);
  await waitFor(`document.querySelector('[data-testid="guided-sample-button"]')?.textContent.includes("Open the guided sample folder")`);
  await page.executeJavaScript(`document.querySelector(".confirmation-backdrop button.secondary-button")?.click()`);
  await delay(400);
  await shot("home.png");
  await shot("choose-files.png", ".upload-panel");

  await page.executeJavaScript(`window.dispatchEvent(new Event("exhibit-builder:analyse-guided-sample"))`);
  await waitFor(`document.querySelector(".exhibit-review-card") && document.body.innerText.includes("Check the proposed exhibits")`, 120_000);
  const reviewError = await page.executeJavaScript(`document.querySelector(".error-toast span")?.textContent || ""`);
  if (reviewError) throw new Error(reviewError);
  await delay(600);
  await shot("review-confirm.png", ".exhibit-review-card");

  await page.executeJavaScript(`document.querySelector(".progress-link")?.click()`);
  await waitFor(`document.querySelector(".file-review-card")`);
  await delay(300);
  await shot("unused-file.png", ".file-review-card");

  await page.executeJavaScript(`[...document.querySelectorAll(".workspace-nav button")].find((button) => button.textContent.includes("Finalise"))?.click()`);
  await waitFor(`document.body.innerText.includes("Build exhibit bundle")`);
  await delay(300);
  await shot("finalise-build.png", ".workspace-header");
  await shot("save-download.png", "header");

  app.exit(0);
}).catch((error) => {
  console.error(error);
  app.exit(1);
});
