import { defineConfig } from 'vite';

// Use relative asset paths so dist/index.html can be opened directly via file://
export default defineConfig({
  base: './',
});