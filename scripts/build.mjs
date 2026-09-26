// join every src file into dist/SpiritHub-<version>.js
// Consts.ts first so shared constants exist before other files use them
// then every other src file alphabetically
// types are stripped, not compiled: the code stays exactly as written
// comments are removed, development happens in src/ where they live
// dist is emptied first so it only ever holds the current build

import { execFileSync } from "node:child_process";
import { copyFileSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import vm from "node:vm";
import tsBlankSpace from "ts-blank-space";
import ts from "typescript";

const SRC = "src";
const DIST = "dist";
const FIRST = "Consts.ts";

const consts = readFileSync(join(SRC, FIRST), "utf8");
const version = /HUB_VERSION\s*=\s*"([^"]+)"/.exec(consts)?.[1];
if (!version) throw new Error(`HUB_VERSION not found in ${SRC}/${FIRST}`);

const files = [FIRST, ...readdirSync(SRC).filter((f) => f.endsWith(".ts") && f !== FIRST).sort()];

/**
 * a src file as plain javascript
 * types are replaced with spaces so every line stays where it was
 *
 * @param {string} file  file name in src
 * @returns {string} the javascript
 */
function stripTypes(file) {
  /** @type {string[]} */
  const errors = [];
  const js = tsBlankSpace(readFileSync(join(SRC, file), "utf8"), (node) => errors.push(node.getText()));
  // tsconfig's erasableSyntaxOnly should already have caught these in `just typecheck`
  if (errors.length > 0) throw new Error(`${file}: cannot strip types from: ${errors.join(", ")}`);
  return js;
}

/**
 * javascript with every comment blanked out
 * comments are found from the parsed code, so text inside strings and regexes is never touched
 * the build's own file markers are added afterwards, so they are kept
 *
 * @param {string} js    the javascript
 * @param {string} file  file name, for the parser
 * @returns {string} the javascript without comments
 */
function stripComments(js, file) {
  const source = ts.createSourceFile(file, js, ts.ScriptTarget.ES2022, true, ts.ScriptKind.JS);
  /** @type {Map<number, number>} */
  const ranges = new Map();
  /** @param {ts.Node} node */
  const visit = (node) => {
    for (const r of ts.getLeadingCommentRanges(js, node.pos) ?? []) ranges.set(r.pos, r.end);
    for (const r of ts.getTrailingCommentRanges(js, node.end) ?? []) ranges.set(r.pos, r.end);
    ts.forEachChild(node, visit);
  };
  visit(source);
  // blank from the end backwards so earlier positions stay valid, newlines kept so dprint can tidy up
  return [...ranges]
    .sort(([a], [b]) => b - a)
    .reduce(
      (text, [pos, end]) => text.slice(0, pos) + text.slice(pos, end).replace(/[^\n]/g, " ") + text.slice(end),
      js,
    );
}

const header = [
  `// SpiritHub ${version}`,
  `// built ${new Date().toISOString().slice(0, 10)}`,
].join("\n");
const body = files
  .map((f) => `// ---- ${f} ----\n${stripComments(stripTypes(f), f).trim()}\n`)
  .join("\n");

// stripping leaves spaces where the types and comments were, and blank lines where whole lines were removed
// dprint tidies them, using the same settings as the source
const bundle = execFileSync("npx", ["dprint", "fmt", "--stdin", "bundle.js"], {
  input: `${header}\n\n${body}`,
  encoding: "utf8",
});

// top-level code runs on load in apps script too
// google services are only used inside functions, so an empty context is enough
vm.runInNewContext(bundle, {}, { filename: `SpiritHub-${version}.js` });

rmSync(DIST, { recursive: true, force: true });
mkdirSync(DIST);
writeFileSync(join(DIST, `SpiritHub-${version}.js`), bundle);
// clasp refuses to push a folder without the project manifest
copyFileSync(join(SRC, "appsscript.json"), join(DIST, "appsscript.json"));
console.log(`built ${DIST}/SpiritHub-${version}.js from ${files.length} files`);
