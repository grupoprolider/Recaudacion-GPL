import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';
import { exec } from 'child_process';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');

function autoSyncExcelPlugin() {
  let lastSignature = '';
  let syncing = false;

  const computeSignature = () => {
    try {
      const isExcel = (f) =>
        (f.toLowerCase().endsWith('.xlsx') || f.toLowerCase().endsWith('.xls')) && !f.startsWith('~$');
      const files = fs.readdirSync(rootDir).filter(isExcel).sort();
      return files
        .map(f => {
          const st = fs.statSync(path.join(rootDir, f));
          return `${f}:${st.size}:${Math.floor(st.mtimeMs)}`;
        })
        .join('|');
    } catch (_e) {
      return '';
    }
  };

  return {
    name: 'auto-sync-excel-github',
    configureServer() {
      lastSignature = computeSignature();
      setInterval(() => {
        if (syncing) return;
        const currentSig = computeSignature();
        if (currentSig && currentSig !== lastSignature) {
          lastSignature = currentSig;
          syncing = true;
          console.log('[Auto-Sync] Cambio detectado en planillas Excel. Sincronizando y subiendo a GitHub...');
          exec(
            'node sync-data.js && cd .. && git add . && git commit -m "Auto-actualizacion de planillas Excel" && git pull --rebase -X ours origin main && git push origin main',
            { cwd: __dirname },
            (err, stdout, stderr) => {
              if (stdout) console.log(stdout);
              if (stderr) console.log(stderr);
              if (err) {
                console.warn('[Auto-Sync] Aviso al sincronizar:', err.message);
              } else {
                console.log('[Auto-Sync] Planillas subidas a GitHub exitosamente.');
              }
              syncing = false;
            }
          );
        }
      }, 4000);
    }
  };
}

export default defineConfig({
  base: './',
  plugins: [react(), autoSyncExcelPlugin()],
});
