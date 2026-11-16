import * as XLSX from 'xlsx';

// Helper to convert Excel date to JS date
const excelToDate = (excelDate) => {
  return new Date(Math.round((excelDate - 25569) * 86400 * 1000));
};

const monthNames = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
];

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
    const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
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

      // En TABLA ACUMULADOS, la fila 25 (índice 24) contiene siempre el total histórico del mes.
      // Las filas inferiores (26-29) suelen tener reintegros, diferencias o anotaciones al margen.
      const row25 = aoa[24];
      if (row25 && typeof row25[block.colMonto] === 'number' && row25[block.colMonto] > 0) {
        totalMonto = row25[block.colMonto];
        if (typeof row25[block.colCobrados] === 'number' && row25[block.colCobrados] > 0) {
          totalClientes = Math.round(row25[block.colCobrados]);
        } else {
          // Si en algún mes histórico (ej. algunos meses de 2021) no se sumó la celda de cobrados en fila 25,
          // sumamos los cobrados diarios entre el encabezado y la fila 24.
          let sumClientes = 0;
          for (let r = headerRowIndex + 1; r < 24; r++) {
            if (aoa[r] && typeof aoa[r][block.colCobrados] === 'number') {
              sumClientes += aoa[r][block.colCobrados];
            }
          }
          totalClientes = Math.round(sumClientes);
        }
      } else {
        // Fallback si alguna hoja tiene el total en fila 26 (índice 25) con palabra TOTAL
        for (let i = 23; i <= Math.min(28, aoa.length - 1); i++) {
          const row = aoa[i];
          if (!row) continue;
          const monto = row[block.colMonto];
          const clientes = row[block.colCobrados];
          if (typeof monto === 'number' && monto > 0 && typeof clientes === 'number' && clientes > 0) {
            totalMonto = monto;
            totalClientes = Math.round(clientes);
            break;
          }
        }
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
    workbook.SheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

      // Buscar la fecha base para saber mes y año
      let recordDate = null;
      for (let i = 0; i < Math.min(30, aoa.length); i++) {
        const fechaCell = aoa[i] && aoa[i][0];
        if (typeof fechaCell === 'number' && fechaCell > 40000) {
          recordDate = excelToDate(fechaCell);
          recordDate = new Date(recordDate.getTime() + recordDate.getTimezoneOffset() * 60000);
          break;
        }
      }

      if (recordDate) {
        const year = recordDate.getFullYear();
        const month = recordDate.getMonth() + 1;
        const monthName = monthNames[recordDate.getMonth()];

        let colCobrados = 3;
        let colMonto = 7;
        let colTotalBruto = -1;

        for (let i = 0; i < Math.min(10, aoa.length); i++) {
          if (aoa[i] && aoa[i].some(cell => typeof cell === 'string' && cell.toUpperCase().includes('MONTO COBRADO'))) {
            const foundCob = aoa[i].findIndex(c => {
              if (typeof c !== 'string') return false;
              const u = c.toUpperCase().trim();
              return u === 'COBRADOS' || u === 'CANT. COB.' || u === 'CANT. COB';
            });
            if (foundCob !== -1) colCobrados = foundCob;

            const foundMonto = aoa[i].findIndex(c => typeof c === 'string' && c.toUpperCase().includes('MONTO COBRADO'));
            if (foundMonto !== -1) colMonto = foundMonto;

            const foundBruto = aoa[i].findIndex(c => typeof c === 'string' && c.toUpperCase().trim() === 'TOTAL COBRADO');
            if (foundBruto !== -1) colTotalBruto = foundBruto;
            break;
          }
        }

        let totalMonto = 0;
        let totalClientes = 0;

        // 1) Buscar fila marcada con "TOTAL$" (presente en todas las solapas de RECA)
        const totalRow = aoa.find(row =>
          row && row.some(cell => typeof cell === 'string' && cell.toUpperCase().replace(/\s+/g, '') === 'TOTAL$')
        );

        if (totalRow && typeof totalRow[colMonto] === 'number') {
          totalClientes = Math.round(totalRow[colCobrados] || 0);
          // En la hoja GLOBAL, TOTAL COBRADO coincide con el importe bruto pasado a TABLA ACUMULADOS
          if (sheetName.toUpperCase() === 'GLOBAL' && colTotalBruto !== -1 && typeof totalRow[colTotalBruto] === 'number' && totalRow[colTotalBruto] > 0) {
            totalMonto = totalRow[colTotalBruto];
          } else {
            totalMonto = totalRow[colMonto] || 0;
          }
        } else {
          // Fallback a filas 25-28 (índices 24-27)
          for (let i = 24; i <= Math.min(28, aoa.length - 1); i++) {
            if (aoa[i] && typeof aoa[i][colMonto] === 'number' && aoa[i][colMonto] > 0) {
              totalClientes = Math.round(aoa[i][colCobrados] || 0);
              totalMonto = aoa[i][colMonto] || 0;
              break;
            }
          }
        }

        if (totalMonto > 0) {
          results.push({
            year,
            month,
            monthName,
            clientes: totalClientes,
            monto: totalMonto,
            acumulado: 0,
            difMontos: 0,
            difClientes: 0,
            subPlatform: sheetName.toUpperCase() === 'GLOBAL' ? null : sheetName.trim()
          });
        }
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
