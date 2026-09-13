import { defineConfig } from "vite";

/**
 * The build has to produce a file that opens from disk.
 *
 * Relative asset paths and an inlining pass after the build are what make that
 * work: a browser refuses module imports and fetch over file://, so anything
 * left as a separate request would fail silently when the page is opened by
 * double-clicking it. See tools/build/inline-single-file.mjs.
 */
export default defineConfig({
  base: "./",
  build: {
    target: "es2022",
    // One chunk, so there is one thing to inline.
    rollupOptions: { output: { inlineDynamicImports: true } },
    assetsInlineLimit: 0,
    cssCodeSplit: false,
    reportCompressedSize: false,
  },
});
