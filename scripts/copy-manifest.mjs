import { readFileSync, writeFileSync, mkdirSync } from "node:fs";

const host = process.env.NOTEES_WORD_HOST ?? "https://localhost:3000";
const manifest = readFileSync("manifest.xml", "utf8").replaceAll("${HOST}", host);
if (manifest.includes("${HOST}")) {
  throw new Error("unresolved ${HOST} placeholder in manifest.xml");
}
mkdirSync("dist", { recursive: true });
writeFileSync("dist/manifest.xml", manifest);
console.log(`manifest written with host ${host}`);
