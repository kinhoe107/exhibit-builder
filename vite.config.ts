import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import { rm } from "node:fs/promises";
import { resolve } from "node:path";

const generatedGuidedOutputs = ["Exhibit_Bundle.pdf", "Guided_sample_exhibit_bundle_Build_Manifest.json"];

export default defineConfig({
  base: "./",
  plugins: [react(), {
    name: "exclude-generated-guided-outputs",
    async closeBundle() {
      await Promise.all(generatedGuidedOutputs.map((name) => rm(resolve("dist", "guided-sample", name), { force: true })));
    },
  }],
  build: {
    target: "es2022",
    chunkSizeWarningLimit: 1100,
  },
});
