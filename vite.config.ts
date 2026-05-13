import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Set BASE_PATH at build time to deploy under a sub-path (e.g. GitHub Pages).
// Example: BASE_PATH=/AIDEN/ npm run build
const base = process.env.BASE_PATH || '/';

export default defineConfig({
  base,
  plugins: [react()],
});
