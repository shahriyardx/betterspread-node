import { defineConfig } from "tsup"

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  dts: true,
  clean: true,
  splitting: false,
  platform: "node",
  target: "esnext",
  external: [
    /@googleapis\//,
    "zod",
  ],
  esbuildOptions(options) {
    options.packages = "external"
  },
})
