import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const rootDir = path.resolve(__dirname, '..');
const publicDir = path.resolve(__dirname, 'public');
const dataDir = path.resolve(publicDir, 'data');

if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir, { recursive: true });
}

// 1. Copy LOGO-gpl.png if present in root
const logoSrc = path.join(rootDir, 'LOGO-gpl.png');
const logoDest = path.join(publicDir, 'logo-gpl.png');
if (fs.existsSync(logoSrc)) {
  fs.copyFileSync(logoSrc, logoDest);
}

// 2. Copy all .xlsx files from rootDir to public/data/
const rootFiles = fs.readdirSync(rootDir);
rootFiles.forEach(file => {
  if (file.toLowerCase().endsWith('.xlsx') && !file.startsWith('~$')) {
    fs.copyFileSync(path.join(rootDir, file), path.join(dataDir, file));
  }
});

// 3. Generate manifest.json of all .xlsx files in public/data/
const dataFiles = fs.readdirSync(dataDir).filter(
  f => f.toLowerCase().endsWith('.xlsx') && !f.startsWith('~$')
);

// Put TABLA ACUMULADOS first so historical base loads first, then RECA files
dataFiles.sort((a, b) => {
  const aIsAcum = a.toUpperCase().includes('ACUMULADO') ? 0 : 1;
  const bIsAcum = b.toUpperCase().includes('ACUMULADO') ? 0 : 1;
  if (aIsAcum !== bIsAcum) return aIsAcum - bIsAcum;
  return a.localeCompare(b);
});

fs.writeFileSync(
  path.join(dataDir, 'manifest.json'),
  JSON.stringify({ updatedAt: new Date().toISOString(), files: dataFiles }, null, 2)
);

console.log('Datos sincronizados para web/GitHub:', dataFiles);
