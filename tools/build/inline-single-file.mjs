/**
 * Folds the build into one HTML file that opens from disk.
 *
 * A browser refuses ES module imports and `fetch` over `file://`, so a build
 * left as index.html plus separate .js and .css would load and immediately fail
 * when someone double-clicks it. Inlining every asset removes the requests, and
 * with the death-probability grid already embedded by embed-mortality.mjs there
 * is nothing left for the page to ask for: no server, no network, no CDN.
 *
 * Written here rather than pulled in as a plugin so the one thing the whole
 * delivery promise rests on is forty lines you can read.
 */
import { readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const dist = resolve(here, "../../apps/web/dist");
const htmlPath = join(dist, "index.html");

let html = readFileSync(htmlPath, "utf8");

/** Everything the page would otherwise request, by the tag that requests it. */
const script = /<script[^>]*src="([^"]+)"[^>]*><\/script>/g;
const stylesheet = /<link[^>]*rel="stylesheet"[^>]*href="([^"]+)"[^>]*>/g;

const read = (href) => readFileSync(join(dist, href.replace(/^\.?\//, "")), "utf8");

// A classic script, not `type="module"` -- see vite.config.ts's own comment
// on why. A module script defers itself; a classic one runs the moment the
// parser reaches it, which is a problem since Vite places it in `<head>`,
// before `<body>` -- `#app` included -- exists to find. `defer` cannot fix
// this either: the attribute is defined to do nothing on a script with no
// `src`, inline or not. So this moves the tag to just before `</body>`
// instead, the ordinary fix for exactly this, rather than leaving it where
// Vite put it.
let inlineScript = "";
html = html.replace(script, (_, href) => {
  const code = read(href);
  // `</script>` inside the bundle would close the tag early.
  inlineScript = `<script>${code.replace(/<\/script>/gi, "<\\/script>")}</script>`;
  return "";
});
html = html.replace(stylesheet, (_, href) => `<style>${read(href)}</style>`);
if (inlineScript === "") throw new Error("no <script src> tag found to inline");
if (!html.includes("</body>")) throw new Error("no </body> to place the inlined script before");
// A replacer *function*, not a template string: `String.replace`'s string form
// treats `$&`, `$$`, `` $` `` etc. in the replacement as special patterns, and a
// minified bundle saying `$` for some identifier right next to a `&` (bitwise
// AND, as common as the variable name itself in output this size) spells `$&`
// by pure chance -- which is exactly what happened here once code elsewhere in
// this change shifted the minifier's own name allocation onto it, silently
// splicing the literal matched text ("</body>") into the middle of the script
// instead of the script itself. A function's return value is inserted as-is,
// with no such reinterpretation, whatever the bundle happens to contain.
html = html.replace("</body>", () => `${inlineScript}</body>`);

/**
 * What a leftover `src`/`href` is allowed to be.
 *
 * The check is for subresources -- things the page fetches while rendering,
 * which is what fails silently over `file://`. Somewhere to *navigate* is not
 * one of those: an `http` link, a `mailto:` handed to the mail client, a `#`
 * anchor and a `data:` URI all work from a file on a stick. Anything else is a
 * request, and has to be inlined.
 */
const navigates = (v) =>
  v.startsWith("data:") ||
  v.startsWith("http") ||
  v.startsWith("mailto:") ||
  v.startsWith("tel:") ||
  v.startsWith("#");

const left = [...html.matchAll(/(?:src|href)="([^"]+)"/g)]
  .map((m) => m[1])
  .filter((v) => !navigates(v));
if (left.length > 0) {
  throw new Error(
    `the page still requests ${left.join(", ")}, which would fail over file://. ` +
      `Inline it or embed it in the bundle.`,
  );
}

const out = join(dist, "typfallsmodellen.html");
writeFileSync(out, html);

// The separate assets are what the single file replaces; leaving them invites
// shipping the broken pair by mistake.
for (const entry of readdirSync(dist)) {
  if (entry !== "typfallsmodellen.html") rmSync(join(dist, entry), { recursive: true, force: true });
}

const kb = Math.round(Buffer.byteLength(html) / 1024);
console.log(`inline-single-file: apps/web/dist/typfallsmodellen.html (${kb} kB, no external requests)`);
