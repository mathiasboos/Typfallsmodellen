import { defineConfig } from "vite";

/**
 * The build has to produce a file that opens from disk.
 *
 * Relative asset paths and an inlining pass after the build are what make that
 * work: a browser refuses module imports and fetch over file://, so anything
 * left as a separate request would fail silently when the page is opened by
 * double-clicking it. See tools/build/inline-single-file.mjs.
 *
 * `format: "iife"` -- a plain, self-invoking classic script instead of
 * Rollup's default `"es"` -- for the same reason, one step further: Mobile
 * Safari can fail to run a `<script type="module">` at all when the page is
 * opened from Files/Mail/Messages over `file://`, even with every import
 * already inlined and nothing left to fetch -- reported after exactly that
 * (a blank/black screen on iPhone). `iife` needs no `type="module"` on the
 * tag (`inline-single-file.mjs` writes a classic one), and needs a single
 * chunk to work at all -- already true here via `inlineDynamicImports`.
 */
export default defineConfig({
  base: "./",
  build: {
    target: "es2022",
    // One chunk, so there is one thing to inline.
    rollupOptions: { output: { inlineDynamicImports: true, format: "iife" } },
    assetsInlineLimit: 0,
    cssCodeSplit: false,
    reportCompressedSize: false,
  },
});
