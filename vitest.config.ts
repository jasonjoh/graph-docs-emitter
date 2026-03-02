import { defineConfig } from "vitest/config";
import alloyPlugin from "@alloy-js/rollup-plugin";

export default defineConfig({
  test: {
    include: ["test/**/*.test.ts", "test/**/*.test.tsx"],
    exclude: ["test/**/*.d.ts"],
  },
  esbuild: {
    jsx: "preserve",
    sourcemap: "both",
  },
  plugins: [alloyPlugin()],
});
