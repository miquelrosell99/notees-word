import { mkdirSync, copyFileSync } from "node:fs";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const src = require.resolve("@microsoft/office-js/package.json");
const { dirname, join } = await import("node:path");
const distDir = join(dirname(src), "dist");
mkdirSync("dist/lib", { recursive: true });
copyFileSync(join(distDir, "office.js"), "dist/lib/office.js");
console.log("vendored office.js");
