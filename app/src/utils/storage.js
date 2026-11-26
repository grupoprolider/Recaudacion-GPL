import localforage from 'localforage';
import { parseExcelBuffer } from './parser';

const DATA_KEY = 'recaudacion_data_v9';
const MANIFEST_KEY = 'recaudacion_manifest_ts';

const GITHUB_API_CONTENTS = 'https://api.github.com/repos/grupoprolider/Recaudacion-GPL/contents/';
const GITHUB_RAW_BASE = 'https://raw.githubusercontent.com/grupoprolider/Recaudacion-GPL/main/';

const isLocalhost = () =>
  typeof window !== 'undefined' &&
  (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1');

const isExcelFile = (f) =>
  typeof f === 'string' &&
  (f.toLowerCase().endsWith('.xlsx') || f.toLowerCase().endsWith('.xls')) &&
  !f.startsWith('~$');

const sortExcelFiles = (files) => {
  return [...files].sort((a, b) => {
    const aIsAcum = a.toUpperCase().includes('ACUMULADO') ? 0 : 1;
    const bIsAcum = b.toUpperCase().includes('ACUMULADO') ? 0 : 1;
    if (aIsAcum !== bIsAcum) return aIsAcum - bIsAcum;
    return a.localeCompare(b);
  });
};

export const mergeParsedData = (existingData, parsedData) => {
  const newData = [...existingData];

  parsedData.forEach(d => {
    const platform = d.subPlatform ? d.subPlatform : 'GLOBAL';
    const newItem = { ...d, platform };

    const existingIndex = newData.findIndex(item =>
      item.year === newItem.year &&
      item.month === newItem.month &&
      item.platform === newItem.platform
    );

    if (existingIndex >= 0) {
      const prev = newData[existingIndex];
      newData[existingIndex] = {
        ...newItem,
        daily: (newItem.daily && newItem.daily.length > 0) ? newItem.daily : (prev.daily || [])
      };
    } else {
      newData.push(newItem);
    }
  });

  return newData;
};

export const saveData = async (parsedData, _filename) => {
  const existingData = (await localforage.getItem(DATA_KEY)) || [];
  const newData = mergeParsedData(existingData, parsedData);
  await localforage.setItem(DATA_KEY, newData);
  return newData;
};

// Obtiene la lista de archivos Excel y su huella de versión en tiempo real
// (directo de la rama main de GitHub para no depender de que GitHub Pages demore o falle al publicar)
const fetchRemoteCatalog = async () => {
  const bust = Date.now();

  // 1) En localhost priorizamos los archivos locales de la carpeta, pero también consultamos manifest local
  if (isLocalhost()) {
    try {
      const res = await fetch(`./data/manifest.json?t=${bust}`, { cache: 'no-store' });
      if (res.ok) {
        const man = await res.json();
        if (Array.isArray(man.files) && man.files.length > 0) {
          return {
            fingerprint: `local:${man.updatedAt || ''}:${man.files.join(',')}`,
            files: sortExcelFiles(man.files).map(name => ({
              name,
              url: `./data/${encodeURIComponent(name)}?t=${bust}`
            }))
          };
        }
      }
    } catch (_e) {
      // Fallback below
    }
  }

  // 2) Consultar la API de GitHub en vivo (detecta al instante cualquier cambio o archivo nuevo en la rama main)
  try {
    const apiRes = await fetch(`${GITHUB_API_CONTENTS}?t=${bust}`, { cache: 'no-store' });
    if (apiRes.ok) {
      const items = await apiRes.json();
      if (Array.isArray(items)) {
        const excelItems = items.filter(it => it && it.type === 'file' && isExcelFile(it.name));
        if (excelItems.length > 0) {
          const sorted = sortExcelFiles(excelItems.map(it => it.name));
          const shaMap = new Map(excelItems.map(it => [it.name, it.sha || `${it.size}`]));
          const fingerprint = sorted.map(name => `${name}:${shaMap.get(name)}`).join('|');
          return {
            fingerprint: `gh-api:${fingerprint}`,
            files: sorted.map(name => ({
              name,
              url: `${GITHUB_RAW_BASE}${encodeURIComponent(name)}?sha=${encodeURIComponent(shaMap.get(name) || bust)}`
            }))
          };
        }
      }
    }
  } catch (_e) {
    // Si la API de GitHub alcanza límite de rate, usamos raw manifest.json directo de main
  }

  // 3) Fallback a manifest.json directo de la rama main en raw.githubusercontent.com
  try {
    const rawManRes = await fetch(`${GITHUB_RAW_BASE}app/public/data/manifest.json?t=${bust}`, { cache: 'no-store' });
    if (rawManRes.ok) {
      const man = await rawManRes.json();
      if (Array.isArray(man.files) && man.files.length > 0) {
        const sorted = sortExcelFiles(man.files);
        const fp = `gh-raw:${man.updatedAt || bust}:${sorted.join(',')}`;
        return {
          fingerprint: fp,
          files: sorted.map(name => ({
            name,
            url: `${GITHUB_RAW_BASE}${encodeURIComponent(name)}?v=${encodeURIComponent(man.updatedAt || bust)}`
          }))
        };
      }
    }
  } catch (_e) {
    // Fallback a ./data/manifest.json relativo
  }

  // 4) Último fallback: ./data/manifest.json relativo del sitio
  let defaultFiles = [
    'TABLA ACUMULADOS.xlsx',
    'RECA.01.xlsx',
    'RECA.02.xlsx',
    'RECA.03.xlsx',
    'RECA.04.xlsx',
    'RECA.05.xlsx',
    'RECA.06.xlsx',
    'RECA.07.xlsx',
    'RECA.08.xlsx',
    'RECA.09.xlsx',
    'RECA.10.xlsx'
  ];
  let ts = String(bust);
  try {
    const manRes = await fetch(`./data/manifest.json?t=${bust}`, { cache: 'no-store' });
    if (manRes.ok) {
      const man = await manRes.json();
      if (Array.isArray(man.files) && man.files.length > 0) {
        defaultFiles = sortExcelFiles(man.files);
      }
      if (man.updatedAt) ts = man.updatedAt;
    }
  } catch (_e) {
    // ignore
  }

  return {
    fingerprint: `static:${ts}:${defaultFiles.join(',')}`,
    files: defaultFiles.map(name => ({
      name,
      url: `./data/${encodeURIComponent(name)}?v=${encodeURIComponent(ts)}`
    }))
  };
};

export const seedDefaultFiles = async (preserveUserUploads = false) => {
  const catalog = await fetchRemoteCatalog();
  let combined = preserveUserUploads ? ((await localforage.getItem(DATA_KEY)) || []) : [];

  for (const fileEntry of catalog.files) {
    try {
      let res = await fetch(fileEntry.url, { cache: 'no-store' });
      // Fallback a ruta local ./data/ si raw.githubusercontent fallara por red
      if (!res.ok) {
        res = await fetch(`./data/${encodeURIComponent(fileEntry.name)}?t=${Date.now()}`, { cache: 'no-store' });
      }
      if (res.ok) {
        const buf = await res.arrayBuffer();
        const parsed = parseExcelBuffer(buf);
        combined = mergeParsedData(combined, parsed);
      }
    } catch (e) {
      console.warn('No se pudo precargar archivo base:', fileEntry.name, e);
    }
  }

  if (combined.length > 0) {
    await localforage.setItem(DATA_KEY, combined);
    if (catalog.fingerprint) {
      await localforage.setItem(MANIFEST_KEY, catalog.fingerprint);
    }
  }
  return combined;
};

export const loadData = async () => {
  const stored = await localforage.getItem(DATA_KEY);
  const storedFp = await localforage.getItem(MANIFEST_KEY);

  try {
    const catalog = await fetchRemoteCatalog();
    if (catalog.fingerprint && catalog.fingerprint !== storedFp) {
      // Si cambió cualquier planilla en GitHub o en la carpeta local, recargamos desde cero las planillas actualizadas
      const seeded = await seedDefaultFiles(false);
      if (seeded.length > 0) return seeded;
    }
  } catch (_e) {
    // Ignore network check error
  }

  const hasPlatforms = stored && stored.some(d => d.platform !== 'GLOBAL');
  const hasDaily = stored && stored.some(d => Array.isArray(d.daily) && d.daily.length > 0);
  if (stored && stored.length > 0 && hasPlatforms && hasDaily) {
    return stored;
  }
  return await seedDefaultFiles(false);
};

export const clearData = async () => {
  await localforage.setItem(DATA_KEY, []);
};
