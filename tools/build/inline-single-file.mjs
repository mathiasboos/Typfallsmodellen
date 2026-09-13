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

html = html.replace(script, (_, href) => {
  const code = read(href);
  // `</script>` inside the bundle would close the tag early.
  return `<script type="module">${code.replace(/<\/script>/gi, "<\\/script>")}</script>`;
});
html = html.replace(stylesheet, (_, href) => `<style>${read(href)}</style>`);

const left = [...html.matchAll(/(?:src|href)="([^"]+)"/g)]
  .map((m) => m[1])
  .filter((v) => !v.startsWith("data:") && !v.startsWith("http") && !v.startsWith("#"));
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
