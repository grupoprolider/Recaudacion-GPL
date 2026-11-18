import * as XLSX from 'xlsx';

// Helper to convert Excel serial date or string date to JS Date
const parsePossibleDate = (cell) => {
  if (typeof cell === 'number' && cell > 38000 && cell < 60000) {
    const d = new Date(Math.round((cell - 25569) * 86400 * 1000));
    return new Date(d.getTime() + d.getTimezoneOffset() * 60000);
  }
  if (cell instanceof Date && !isNaN(cell.getTime())) {
    return cell;
  }
  if (typeof cell === 'string') {
    const clean = cell.trim();
    const m = clean.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
    if (m) {
      const day = parseInt(m[1], 10);
      const month = parseInt(m[2], 10) - 1;
      let year = parseInt(m[3], 10);
      if (year < 100) year += 2000;
      return new Date(year, month, day);
    }
  }
  return null;
};

const monthNames = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
];

// Find the most frequent (year, month) among date cells in a 2D array
const detectSheetYearMonth = (aoa) => {
  const counts = new Map();
  for (let i = 0; i < Math.min(32, aoa.length); i++) {
    const row = aoa[i];
    if (!row) continue;
    for (let c = 0; c <= 1; c++) {
      const d = parsePossibleDate(row[c]);
      if (d) {
        const y = d.getFullYear();
        const m = d.getMonth() + 1;
        if (y >= 2018 && y <= 2035) {
          const key = `${y}-${m}`;
          counts.set(key, (counts.get(key) || 0) + 1);
        }
      }
    }
  }
  if (counts.size === 0) return null;
  let bestKey = null;
  let bestCount = -1;
  for (const [k, cnt] of counts.entries()) {
    if (cnt > bestCount) {
      bestCount = cnt;
      bestKey = k;
    }
  }
  const [year, month] = bestKey.split('-').map(Number);
  return { year, month, monthName: monthNames[month - 1] };
};

export const parseExcelBuffer = (arrayBuffer) => {
  const data = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(data, { type: 'array' });
  const results = [];

  // 1. Intentar procesar como TABLA ACUMULADOS (Hojas AÑO YYYY)
  let isAcumulados = false;
  workbook.SheetNames.forEach(sheetName => {
    if (!sheetName.toUpperCase().includes('AÑO')) return;
    isAcumulados = true;

    const yearMatch = sheetName.match(/\d{4}/);
    if (!yearMatch) return;
    const year = parseInt(yearMatch[0], 10);

    const worksheet = workbook.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, blankrows: true });
    if (aoa.length < 2) return;

    let headerRowIndex = -1;
    for (let i = 0; i < Math.min(10, aoa.length); i++) {
      if (aoa[i] && aoa[i].some(cell => typeof cell === 'string' && cell.toUpperCase().includes('MONTO COBRADO'))) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) return;

    const headerRow = aoa[headerRowIndex];
    const monthBlocks = [];

    headerRow.forEach((cell, index) => {
      if (typeof cell === 'string' && cell.toUpperCase().includes('MONTO COBRADO')) {
        const monthIndex = monthBlocks.length + 1;
        if (monthIndex <= 12) {
          monthBlocks.push({
            monthIndex,
            monthName: monthNames[monthIndex - 1],
            colFecha: index - 3,
            colCobrados: index - 1,
            colMonto: index,
          });
        }
      }
    });

    monthBlocks.forEach(block => {
      let totalMonto = 0;
      let totalClientes = 0;

      // En TABLA ACUMULADOS:
      // - En 2024..2026 el total está en la Fila 25 (índice 24).
      // - En 2020..2023 algunos meses tienen 23 días hábiles (hasta Fila 25) y el total está en la Fila 26 (índice 25).
      // Tomamos entre Fila 25 (idx 24) y Fila 26 (idx 25) aquella que tenga el total del mes (mayor monto, siempre que no sea una fila de reintegros menor).
      const r25 = aoa[24];
      const r26 = aoa[25];
      const m25 = r25 && typeof r25[block.colMonto] === 'number' ? r25[block.colMonto] : 0;
      const m26 = r26 && typeof r26[block.colMonto] === 'number' ? r26[block.colMonto] : 0;
      const c25 = r25 && typeof r25[block.colCobrados] === 'number' ? r25[block.colCobrados] : 0;
      const c26 = r26 && typeof r26[block.colCobrados] === 'number' ? r26[block.colCobrados] : 0;

      let chosenRowIndex = -1;
      if (m25 > 0 && m26 > m25 * 2) {
        // Fila 25 era el día hábil 23 y Fila 26 es el total mensual
        totalMonto = m26;
        totalClientes = c26 > 0 ? Math.round(c26) : 0;
        chosenRowIndex = 25;
      } else if (m25 > 0) {
        totalMonto = m25;
        totalClientes = c25 > 0 ? Math.round(c25) : 0;
        chosenRowIndex = 24;
      } else if (m26 > 0) {
        totalMonto = m26;
        totalClientes = c26 > 0 ? Math.round(c26) : 0;
        chosenRowIndex = 25;
      }

      if (totalMonto > 0 && totalClientes === 0) {
        let sumClientes = 0;
        const limitRow = chosenRowIndex > 0 ? chosenRowIndex : 24;
        for (let r = headerRowIndex + 1; r < limitRow; r++) {
          if (aoa[r] && typeof aoa[r][block.colCobrados] === 'number') {
            sumClientes += aoa[r][block.colCobrados];
          }
        }
        totalClientes = Math.round(sumClientes);
      }

      if (totalMonto > 0) {
        results.push({
          year,
          month: block.monthIndex,
          monthName: block.monthName,
          clientes: totalClientes,
          monto: totalMonto,
          acumulado: 0,
          difMontos: 0,
          difClientes: 0,
          subPlatform: null
        });
      }
    });
  });

  // 2. Si no es tabla de acumulados, procesar como RECA (todas las solapas de plataformas)
  if (!isAcumulados) {
    let workbookYearMonth = null;
    for (const sn of workbook.SheetNames) {
      const aoa = XLSX.utils.sheet_to_json(workbook.Sheets[sn], { header: 1, defval: null });
      const ym = detectSheetYearMonth(aoa);
      if (ym) {
        workbookYearMonth = ym;
        if (sn.toUpperCase() === 'GLOBAL') break;
      }
    }

    workbook.SheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

      const ym = detectSheetYearMonth(aoa) || workbookYearMonth;
      if (!ym) return;

      const { year, month, monthName } = ym;

      let headerRowIndex = -1;
      let colCobrados = 3;
      let colMonto = 7;
      let colTotalBruto = -1;

      for (let i = 0; i < Math.min(10, aoa.length); i++) {
        if (aoa[i] && aoa[i].some(cell => typeof cell === 'string' && cell.toUpperCase().includes('MONTO COBRADO'))) {
          headerRowIndex = i;
          const foundCob = aoa[i].findIndex(c => {
            if (typeof c !== 'string') return false;
            const u = c.toUpperCase().trim();
            return u === 'COBRADOS' || u === 'CANT. COB.' || u === 'CANT. COB';
          });
          if (foundCob !== -1) colCobrados = foundCob;

          const foundMonto = aoa[i].findIndex(c => typeof c === 'string' && c.toUpperCase().includes('MONTO COBRADO'));
          if (foundMonto !== -1) colMonto = foundMonto;

          const foundBruto = aoa[i].findIndex(c => {
            if (typeof c !== 'string') return false;
            const u = c.toUpperCase().trim();
            return u === 'TOTAL COBRADO' || u === 'TOTAL' || u === 'SUMA COBRADOS';
          });
          if (foundBruto !== -1) colTotalBruto = foundBruto;
          break;
        }
      }

      if (headerRowIndex === -1) return;

      let totalMonto = 0;
      let totalClientes = 0;

      let totalRow = aoa.find((row, idx) =>
        idx >= 22 &&
        row &&
        row.some(cell => {
          if (typeof cell !== 'string') return false;
          const clean = cell.toUpperCase().replace(/\s+/g, '');
          return clean === 'TOTAL$' || clean === 'TOTAL';
        })
      );

      if (totalRow) {
        totalClientes = Math.round(totalRow[colCobrados] || 0);
        const isGlobalSheet = sheetName.toUpperCase().trim() === 'GLOBAL';
        if (isGlobalSheet && colTotalBruto !== -1 && typeof totalRow[colTotalBruto] === 'number' && totalRow[colTotalBruto] > 0) {
          totalMonto = totalRow[colTotalBruto];
        } else if (typeof totalRow[colMonto] === 'number' && totalRow[colMonto] > 0) {
          totalMonto = totalRow[colMonto];
        } else if (colTotalBruto !== -1 && typeof totalRow[colTotalBruto] === 'number' && totalRow[colTotalBruto] > 0) {
          totalMonto = totalRow[colTotalBruto];
        }
      }

      if (totalMonto === 0) {
        for (let i = 24; i <= Math.min(28, aoa.length - 1); i++) {
          const row = aoa[i];
          if (!row) continue;
          const mVal = typeof row[colMonto] === 'number' && row[colMonto] > 0
            ? row[colMonto]
            : (colTotalBruto !== -1 && typeof row[colTotalBruto] === 'number' ? row[colTotalBruto] : 0);
          if (mVal > 0 && (typeof row[colCobrados] === 'number' && row[colCobrados] > 0)) {
            totalClientes = Math.round(row[colCobrados]);
            totalMonto = mVal;
            break;
          }
        }
      }

      if (totalMonto === 0) {
        let sumM = 0;
        let sumC = 0;
        for (let r = headerRowIndex + 1; r < Math.min(25, aoa.length); r++) {
          const row = aoa[r];
          if (!row || !parsePossibleDate(row[0])) continue;
          if (typeof row[colMonto] === 'number' && row[colMonto] > 0) sumM += row[colMonto];
          if (typeof row[colCobrados] === 'number' && row[colCobrados] > 0) sumC += row[colCobrados];
        }
        if (sumM > 0) {
          totalMonto = sumM;
          totalClientes = Math.round(sumC);
        }
      }

      if (totalMonto > 0) {
        const cleanSheet = sheetName.trim();
        const isGlobal = cleanSheet.toUpperCase() === 'GLOBAL';
        results.push({
          year,
          month,
          monthName,
          clientes: totalClientes,
          monto: totalMonto,
          acumulado: 0,
          difMontos: 0,
          difClientes: 0,
          subPlatform: isGlobal ? null : cleanSheet
        });
      }
    });
  }

  return results;
};

export const parseExcelFile = async (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        resolve(parseExcelBuffer(e.target.result));
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};
