import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteSingleFile } from 'vite-plugin-singlefile';

// Set BASE_PATH at build time to deploy under a sub-path (e.g. GitHub Pages).
// Example:  BASE_PATH=/AIDEN/ npm run build
// Set SINGLE_FILE=1 to inline everything into a single index.html that can be
// opened directly from disk (no web server required). Example:
//   SINGLE_FILE=1 npm run build  -> produces dist/index.html
const base = process.env.BASE_PATH || '/';
const singleFile = process.env.SINGLE_FILE === '1';

export default defineConfig({
  base: singleFile ? './' : base,
  plugins: [react(), ...(singleFile ? [viteSingleFile()] : [])],
});
