import localforage from 'localforage';
import { parseExcelBuffer } from './parser';

const DATA_KEY = 'recaudacion_data_v6';
const MANIFEST_KEY = 'recaudacion_manifest_ts';

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
      newData[existingIndex] = newItem;
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

export const seedDefaultFiles = async (preserveUserUploads = true) => {
  let defaultFiles = [
    'TABLA ACUMULADOS.xlsx',
    'RECA.04.xlsx',
    'RECA.05.xlsx'
  ];
  let manifestTs = null;

  try {
    const manRes = await fetch('./data/manifest.json?t=' + Date.now());
    if (manRes.ok) {
      const manifest = await manRes.json();
      if (Array.isArray(manifest.files) && manifest.files.length > 0) {
        defaultFiles = manifest.files;
      }
      manifestTs = manifest.updatedAt || null;
    }
  } catch (_e) {
    // Fallback to default file list
  }

  let combined = preserveUserUploads ? ((await localforage.getItem(DATA_KEY)) || []) : [];
  for (const filename of defaultFiles) {
    try {
      const url = `./data/${encodeURIComponent(filename)}`;
      const res = await fetch(url);
      if (res.ok) {
        const buf = await res.arrayBuffer();
        const parsed = parseExcelBuffer(buf);
        combined = mergeParsedData(combined, parsed);
      }
    } catch (e) {
      console.warn('No se pudo precargar archivo base:', filename, e);
    }
  }

  if (combined.length > 0) {
    await localforage.setItem(DATA_KEY, combined);
    if (manifestTs) {
      await localforage.setItem(MANIFEST_KEY, manifestTs);
    }
  }
  return combined;
};

export const loadData = async () => {
  const stored = await localforage.getItem(DATA_KEY);
  const storedTs = await localforage.getItem(MANIFEST_KEY);

  try {
    const manRes = await fetch('./data/manifest.json?t=' + Date.now());
    if (manRes.ok) {
      const manifest = await manRes.json();
      if (manifest.updatedAt && manifest.updatedAt !== storedTs) {
        const seeded = await seedDefaultFiles(true);
        if (seeded.length > 0) return seeded;
      }
    }
  } catch (_e) {
    // Ignore network check error
  }

  const hasPlatforms = stored && stored.some(d => d.platform !== 'GLOBAL');
  if (stored && stored.length > 0 && hasPlatforms) {
    return stored;
  }
  return await seedDefaultFiles(true);
};

export const clearData = async () => {
  await localforage.setItem(DATA_KEY, []);
};
