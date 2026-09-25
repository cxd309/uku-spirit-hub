// join every src file into dist/SpiritHub-<version>.js
// Consts.js first so shared constants exist before other files use them
// then every other src file alphabetically
// dist is emptied first so it only ever holds the current build

import { copyFileSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import vm from "node:vm";

const SRC = "src";
const DIST = "dist";
const FIRST = "Consts.js";

const consts = readFileSync(join(SRC, FIRST), "utf8");
const version = /HUB_VERSION\s*=\s*"([^"]+)"/.exec(consts)?.[1];
if (!version) throw new Error(`HUB_VERSION not found in ${SRC}/${FIRST}`);

const files = [FIRST, ...readdirSync(SRC).filter((f) => f.endsWith(".js") && f !== FIRST).sort()];

const header = [
  `// SpiritHub ${version}`,
  `// built ${new Date().toISOString().slice(0, 10)}`,
].join("\n");
const body = files
  .map((f) => `// ---- ${f} ----\n${readFileSync(join(SRC, f), "utf8").trim()}\n`)
  .join("\n");
const bundle = `${header}\n\n${body}`;

// top-level code runs on load in apps script too
// google services are only used inside functions, so an empty context is enough
vm.runInNewContext(bundle, {}, { filename: `SpiritHub-${version}.js` });

rmSync(DIST, { recursive: true, force: true });
mkdirSync(DIST);
writeFileSync(join(DIST, `SpiritHub-${version}.js`), bundle);
// clasp refuses to push a folder without the project manifest
copyFileSync(join(SRC, "appsscript.json"), join(DIST, "appsscript.json"));
console.log(`built ${DIST}/SpiritHub-${version}.js from ${files.length} files`);
