import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['esm'],
  target: 'es2022',
  outDir: 'dist',
  sourcemap: true,
  clean: true,
  dts: true,
  minify: false,
  // Single output file — avoids hashed chunk files (e.g. XmlHelper-*.js) going missing after rebuilds
  splitting: false,
});
