import react from '@vitejs/plugin-react';
import { copyFileSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import dts from 'vite-plugin-dts';
import { defineConfig } from 'vite';

const projectRoot = fileURLToPath(new URL('.', import.meta.url));

export default defineConfig({
  plugins: [
    react(),
    dts({
      insertTypesEntry: true,
      tsconfigPath: './tsconfig.build.json',
    }),
    {
      name: 'copy-pdf-worker',
      closeBundle() {
        const source = resolve(projectRoot, 'node_modules/pdfjs-dist/build/pdf.worker.min.mjs');
        const target = resolve(projectRoot, 'dist/pdf.worker.min.mjs');

        mkdirSync(dirname(target), { recursive: true });
        copyFileSync(source, target);
      },
    },
  ],
  build: {
    lib: {
      entry: 'src/index.ts',
      name: 'LocalDocumentViewer',
      fileName: 'index',
    },
    rollupOptions: {
      external: ['react', 'react-dom', 'react/jsx-runtime'],
      output: {
        globals: {
          react: 'React',
          'react-dom': 'ReactDOM',
          'react/jsx-runtime': 'ReactJSXRuntime',
        },
      },
    },
  },
});
