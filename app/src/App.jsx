import React, { useState, useEffect, useCallback, useRef } from 'react';
import {
  UploadCloud, TrendingUp, TrendingDown, Trash2, Download, FileSpreadsheet, RefreshCw, Upload, X
} from 'lucide-react';
import { toPng } from 'html-to-image';
import * as XLSX from 'xlsx';
import { parseExcelFile } from './utils/parser';
import { saveData, loadData, clearData, seedDefaultFiles } from './utils/storage';
import {
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, BarChart, Bar, LabelList, Legend
} from 'recharts';

const renderTopBarLabel = (fill, fontSize) => (props) => {
  const { x, y, width, value } = props;
  if (value === undefined || value === null || value === '') return null;
  return (
    <text
      x={Number(x) + Number(width) / 2}
      y={Number(y) - 8}
      fill={fill}
      textAnchor="middle"
      dominantBaseline="Auto"
      fontSize={fontSize}
      fontWeight="700"
      fontFamily="Inter, Segoe UI, Arial, sans-serif"
      style={{ fontSize: `${fontSize}px`, fontWeight: 700, fill }}
    >
      {value}
    </text>
  );
};

const renderRightBarLabel = (fill, fontSize) => (props) => {
  const { x, y, width, height, value } = props;
  if (value === undefined || value === null || value === '') return null;
  return (
    <text
      x={Number(x) + Number(width) + 8}
      y={Number(y) + Number(height) / 2 + 1}
      fill={fill}
      textAnchor="start"
      dominantBaseline="middle"
      fontSize={fontSize}
      fontWeight="700"
      fontFamily="Inter, Segoe UI, Arial, sans-serif"
      style={{ fontSize: `${fontSize}px`, fontWeight: 700, fill }}
    >
      {value}
    </text>
  );
};

function App() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [showUploader, setShowUploader] = useState(false);
  const [uploadNotice, setUploadNotice] = useState(null);
  const [selectedPlatform, setSelectedPlatform] = useState('GLOBAL');
  const [selectedPeriod, setSelectedPeriod] = useState('LAST_12');
  const [selectedPlatformMonth, setSelectedPlatformMonth] = useState('');
  const [selectedDailyMonth, setSelectedDailyMonth] = useState('');
  const [selectedCutoffDay, setSelectedCutoffDay] = useState(0); // 0 = último día hábil cargado
  const [dailyMetric, setDailyMetric] = useState('MONTO');
  const [compareDailyPrev, setCompareDailyPrev] = useState(true);
  const [isMobile, setIsMobile] = useState(() => typeof window !== 'undefined' && window.innerWidth <= 768);
  const dashboardRef = useRef(null);

  useEffect(() => {
    loadData().then(setData);
    const onResize = () => setIsMobile(window.innerWidth <= 768);
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, []);

  const processFiles = async (files) => {
    setLoading(true);
    setUploadNotice(null);
    try {
      let newData = [...data];
      let totalParsed = 0;
      const detectedPeriods = new Set();
      let latestRecaKey = null;

      for (const file of files) {
        const parsed = await parseExcelFile(file);
        totalParsed += parsed.length;
        parsed.forEach(p => {
          detectedPeriods.add(`${p.monthName} ${p.year}`);
          if (p.subPlatform) {
            latestRecaKey = `${p.year}-${p.month}`;
          }
        });
        newData = await saveData(parsed, file.name);
      }

      if (totalParsed === 0) {
        setUploadNotice({
          type: 'error',
          text: 'No se encontraron filas de recaudación válidas en el archivo seleccionado. Verifica que tenga las columnas COBRADOS y MONTO COBRADO.'
        });
      } else {
        setData(newData);
        if (latestRecaKey) {
          setSelectedPlatformMonth(latestRecaKey);
          setSelectedDailyMonth(latestRecaKey);
          setSelectedCutoffDay(0);
        }
        const periodsList = Array.from(detectedPeriods).slice(0, 5).join(', ');
        setUploadNotice({
          type: 'success',
          text: `¡Planilla actualizada! Se procesaron ${totalParsed} registros (${periodsList}).`
        });
        setShowUploader(false);
      }
    } catch (error) {
      console.error('Error parsing files', error);
      setUploadNotice({
        type: 'error',
        text: 'Hubo un error procesando el archivo. Asegúrate de que sea un archivo Excel (.xlsx) válido.'
      });
    } finally {
      setLoading(false);
      setDragOver(false);
    }
  };

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFiles(e.dataTransfer.files);
    }
  }, [data]);

  const handleDragOver = (e) => {
    e.preventDefault();
    setDragOver(true);
  };

  const handleDragLeave = () => setDragOver(false);

  const handleClear = async () => {
    if (confirm('¿Estás seguro de que quieres borrar los datos cargados en este navegador?')) {
      await clearData();
      setData([]);
      setShowUploader(true);
      setUploadNotice(null);
    }
  };

  const handleReloadDefaults = async () => {
    setLoading(true);
    try {
      const seeded = await seedDefaultFiles(false);
      setData(seeded);
      setShowUploader(false);
      setUploadNotice({
        type: 'success',
        text: 'Se recargaron las planillas actualizadas desde el repositorio.'
      });
    } finally {
      setLoading(false);
    }
  };

  // Extract unique platforms
  const uniquePlatforms = [
    'GLOBAL',
    ...new Set(data.map(d => d.platform).filter(p => p !== 'GLOBAL' && p !== 'AGGREGATED'))
  ];

  // Filter by Platform
  let platformFilteredData = [];
  if (selectedPlatform === 'GLOBAL') {
    platformFilteredData = data
      .filter(d => d.platform === 'GLOBAL')
      .map(d => ({ ...d }));

    if (platformFilteredData.length === 0 && data.length > 0) {
      const map = new Map();
      data.forEach(d => {
        const key = `${d.year}-${d.month}`;
        if (!map.has(key)) {
          map.set(key, { ...d, platform: 'AGGREGATED' });
        } else {
          const existing = map.get(key);
          existing.monto += d.monto;
          existing.clientes += d.clientes;
        }
      });
      platformFilteredData = Array.from(map.values());
    }
  } else {
    platformFilteredData = data
      .filter(d => d.platform === selectedPlatform)
      .map(d => ({ ...d }));
  }

  platformFilteredData.sort((a, b) => {
    if (a.year !== b.year) return a.year - b.year;
    return a.month - b.month;
  });

  // Calculate dynamic Month-over-Month differences before time filtering
  platformFilteredData.forEach((item, index) => {
    if (index > 0) {
      const prev = platformFilteredData[index - 1];
      const isConsecutive =
        (item.year === prev.year && item.month === prev.month + 1) ||
        (item.year === prev.year + 1 && item.month === 1 && prev.month === 12);
      if (isConsecutive) {
        item.difMontos = item.monto - prev.monto;
        item.difClientes = item.clientes - prev.clientes;
        item.pctMontos = prev.monto > 0 ? ((item.monto - prev.monto) / prev.monto) * 100 : 0;
        item.pctClientes = prev.clientes > 0 ? ((item.clientes - prev.clientes) / prev.clientes) * 100 : 0;
      } else {
        item.difMontos = 0;
        item.difClientes = 0;
        item.pctMontos = 0;
        item.pctClientes = 0;
      }
    } else {
      item.difMontos = 0;
      item.difClientes = 0;
      item.pctMontos = 0;
      item.pctClientes = 0;
    }
  });

  // Filter by Period
  let displayData = [...platformFilteredData];
  if (selectedPeriod === 'LAST_12' && displayData.length > 12) {
    displayData = displayData.slice(displayData.length - 12);
  } else if (selectedPeriod === 'LAST_6' && displayData.length > 6) {
    displayData = displayData.slice(displayData.length - 6);
  } else if (selectedPeriod !== 'ALL' && selectedPeriod !== 'LAST_12' && selectedPeriod !== 'LAST_6') {
    displayData = displayData.filter(d => d.year.toString() === selectedPeriod);
  }

  const latest = displayData.length > 0 ? displayData[displayData.length - 1] : null;

  // Formatters
  const formatMoney = (value) =>
    new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS', maximumFractionDigits: 0 }).format(value);
  const formatCompactMoney = (value) => {
    if (!value) return '$0';
    const abs = Math.abs(value);
    const sign = value < 0 ? '-' : '';
    if (abs >= 1000000) {
      return `${sign}$${(abs / 1000000).toLocaleString('es-AR', { minimumFractionDigits: isMobile ? 0 : 1, maximumFractionDigits: isMobile ? 0 : 1 })}M`;
    }
    return `${sign}$${Math.round(abs / 1000).toLocaleString('es-AR')}K`;
  };
  const formatNumber = (value) => new Intl.NumberFormat('es-AR').format(Math.round(value || 0));

  // Prepare chart data
  const chartData = displayData.map(d => ({
    name: `${d.monthName.substring(0, 3)} ${String(d.year).slice(-2)}`,
    fullName: `${d.monthName.split(' ')[0]} ${d.year}`,
    Recaudacion: d.monto,
    RecaudacionLabel: formatCompactMoney(d.monto),
    Clientes: Math.round(d.clientes),
    ClientesLabel: formatNumber(d.clientes)
  }));

  // Available months for Daily Business Day Curve (filtered by selectedPlatform)
  const recordsWithDaily = [...platformFilteredData]
    .filter(d => Array.isArray(d.daily) && d.daily.length > 0)
    .reverse();

  const activeDailyRecord =
    (selectedDailyMonth && recordsWithDaily.find(d => `${d.year}-${d.month}` === selectedDailyMonth)) ||
    (latest && Array.isArray(latest.daily) && latest.daily.length > 0 ? latest : recordsWithDaily[0]) ||
    null;

  const activeDailyMonthKey = activeDailyRecord ? `${activeDailyRecord.year}-${activeDailyRecord.month}` : '';

  // Find consecutive previous month record with daily data for same-business-day comparison
  const prevDailyRecord = activeDailyRecord
    ? platformFilteredData.find(d => {
        const isPrev =
          (activeDailyRecord.month > 1 && d.year === activeDailyRecord.year && d.month === activeDailyRecord.month - 1) ||
          (activeDailyRecord.month === 1 && d.year === activeDailyRecord.year - 1 && d.month === 12);
        return isPrev && Array.isArray(d.daily) && d.daily.length > 0;
      }) || null
    : null;

  // Map daily entries by diaHabil (1..23)
  const currDailyMap = new Map();
  let maxCurrDiaHabil = 0;
  if (activeDailyRecord && Array.isArray(activeDailyRecord.daily)) {
    activeDailyRecord.daily.forEach((d, idx) => {
      const dh = d.diaHabil || idx + 1;
      currDailyMap.set(dh, d);
      if ((d.monto > 0 || d.clientes > 0) && dh > maxCurrDiaHabil) {
        maxCurrDiaHabil = dh;
      }
    });
  }

  const prevDailyMap = new Map();
  let maxPrevDiaHabil = 0;
  if (prevDailyRecord && Array.isArray(prevDailyRecord.daily)) {
    prevDailyRecord.daily.forEach((d, idx) => {
      const dh = d.diaHabil || idx + 1;
      prevDailyMap.set(dh, d);
      if ((d.monto > 0 || d.clientes > 0) && dh > maxPrevDiaHabil) {
        maxPrevDiaHabil = dh;
      }
    });
  }

  // Cutoff business day for same-day comparison (defaults to last loaded business day of activeDailyRecord)
  const effectiveCutoffDay =
    selectedCutoffDay > 0 && selectedCutoffDay <= maxCurrDiaHabil
      ? selectedCutoffDay
      : maxCurrDiaHabil;

  // Compute cumulative totals from Día Hábil 1 up to effectiveCutoffDay for both current and previous month
  let currAcumMontoAtCutoff = 0;
  let currAcumClientesAtCutoff = 0;
  let prevAcumMontoAtCutoff = 0;
  let prevAcumClientesAtCutoff = 0;

  for (let dh = 1; dh <= effectiveCutoffDay; dh++) {
    const cItem = currDailyMap.get(dh);
    if (cItem) {
      currAcumMontoAtCutoff += cItem.monto || 0;
      currAcumClientesAtCutoff += cItem.clientes || 0;
    }
    const pItem = prevDailyMap.get(dh);
    if (pItem) {
      prevAcumMontoAtCutoff += pItem.monto || 0;
      prevAcumClientesAtCutoff += pItem.clientes || 0;
    }
  }

  const currCutoffItem = currDailyMap.get(effectiveCutoffDay) || null;
  const prevCutoffItem = prevDailyMap.get(effectiveCutoffDay) || null;

  const difSameDayClientes = currAcumClientesAtCutoff - prevAcumClientesAtCutoff;
  const pctSameDayClientes =
    prevAcumClientesAtCutoff > 0 ? (difSameDayClientes / prevAcumClientesAtCutoff) * 100 : 0;

  const difSameDayMonto = currAcumMontoAtCutoff - prevAcumMontoAtCutoff;
  const pctSameDayMonto =
    prevAcumMontoAtCutoff > 0 ? (difSameDayMonto / prevAcumMontoAtCutoff) * 100 : 0;

  const difSingleDayClientes = (currCutoffItem?.clientes || 0) - (prevCutoffItem?.clientes || 0);
  const difSingleDayMonto = (currCutoffItem?.monto || 0) - (prevCutoffItem?.monto || 0);

  const dailyTotalMonto = activeDailyRecord
    ? activeDailyRecord.daily.reduce((acc, item) => acc + (item.monto || 0), 0) || activeDailyRecord.monto
    : 0;

  // Build Daily Chart Data (1..maxCurrDiaHabil) including previous month comparison
  let runCurrM = 0;
  let runCurrC = 0;
  let runPrevM = 0;
  let runPrevC = 0;

  const dailyChartData = [];
  for (let dh = 1; dh <= maxCurrDiaHabil; dh++) {
    const cItem = currDailyMap.get(dh) || { diaHabil: dh, fecha: '', clientes: 0, monto: 0 };
    const pItem = prevDailyMap.get(dh) || { diaHabil: dh, fecha: '', clientes: 0, monto: 0 };

    runCurrM += cItem.monto || 0;
    runCurrC += cItem.clientes || 0;
    runPrevM += pItem.monto || 0;
    runPrevC += pItem.clientes || 0;

    const pctAcum = dailyTotalMonto > 0 ? (runCurrM / dailyTotalMonto) * 100 : 0;
    const valCurr = dailyMetric === 'MONTO' ? cItem.monto : cItem.clientes;
    const valPrev = dailyMetric === 'MONTO' ? pItem.monto : pItem.clientes;

    dailyChartData.push({
      name: isMobile ? `D${dh}` : `Día ${dh}`,
      fullName: `Día Hábil ${dh}${cItem.fecha ? ` (${cItem.fecha})` : ''}`,
      diaHabil: dh,
      fecha: cItem.fecha || '',
      fechaPrev: pItem.fecha || '',
      Monto: cItem.monto || 0,
      Clientes: cItem.clientes || 0,
      MontoPrev: pItem.monto || 0,
      ClientesPrev: pItem.clientes || 0,
      AcumMonto: runCurrM,
      AcumClientes: runCurrC,
      AcumMontoPrev: runPrevM,
      AcumClientesPrev: runPrevC,
      DifAcumMonto: runCurrM - runPrevM,
      DifAcumClientes: runCurrC - runPrevC,
      Valor: valCurr,
      ValorPrev: valPrev,
      Label: dailyMetric === 'MONTO' ? formatCompactMoney(cItem.monto) : formatNumber(cItem.clientes),
      LabelPrev: dailyMetric === 'MONTO' ? formatCompactMoney(pItem.monto) : formatNumber(pItem.clientes),
      pctAcum: Math.min(100, pctAcum)
    });
  }

  // Daily summary stats (Peak day, First 10 business days %, Daily average)
  const peakDay =
    dailyChartData.length > 0
      ? dailyChartData.reduce((best, cur) => (cur.Monto > best.Monto ? cur : best), dailyChartData[0])
      : null;

  const first10DaysMonto = dailyChartData
    .filter(d => d.diaHabil <= 10)
    .reduce((acc, d) => acc + d.Monto, 0);
  const first10DaysPct = dailyTotalMonto > 0 ? (first10DaysMonto / dailyTotalMonto) * 100 : 0;
  const avgDailyMonto = dailyChartData.length > 0 ? dailyTotalMonto / dailyChartData.length : 0;
  const avgDailyClientes =
    dailyChartData.length > 0
      ? dailyChartData.reduce((acc, d) => acc + d.Clientes, 0) / dailyChartData.length
      : 0;

  // Available months that have individual platform breakdown (from RECA files)
  const platformRecords = data.filter(d => d.platform !== 'GLOBAL' && d.platform !== 'AGGREGATED');
  const availablePlatformMonths = Array.from(
    new Map(
      platformRecords.map(d => [
        `${d.year}-${String(d.month).padStart(2, '0')}`,
        { key: `${d.year}-${d.month}`, label: `${d.monthName.split(' ')[0]} ${d.year}`, year: d.year, month: d.month }
      ])
    ).values()
  ).sort((a, b) => (a.year !== b.year ? b.year - a.year : b.month - a.month));

  const activePlatformMonthKey =
    selectedPlatformMonth && availablePlatformMonths.some(m => m.key === selectedPlatformMonth)
      ? selectedPlatformMonth
      : availablePlatformMonths.length > 0
      ? availablePlatformMonths[0].key
      : null;

  const activePlatformMonthLabel =
    availablePlatformMonths.find(m => m.key === activePlatformMonthKey)?.label || '';

  const platformData = activePlatformMonthKey
    ? platformRecords
        .filter(d => `${d.year}-${d.month}` === activePlatformMonthKey)
        .sort((a, b) => b.monto - a.monto)
        .map(d => ({
          name: d.platform,
          Recaudacion: d.monto,
          Clientes: d.clientes,
          barLabel: isMobile
            ? `${formatCompactMoney(d.monto)} (${formatNumber(d.clientes)})`
            : `${formatCompactMoney(d.monto)} (${formatNumber(d.clientes)} cl.)`
        }))
    : [];

  // Download Dashboard as PNG Image
  const handleDownloadImage = async () => {
    if (!dashboardRef.current) return;
    setExporting(true);
    try {
      const dataUrl = await toPng(dashboardRef.current, {
        backgroundColor: '#f3f6f8',
        pixelRatio: 2,
        cacheBust: true,
        filter: (node) => {
          if (node?.classList?.contains('recharts-tooltip-wrapper')) return false;
          if (node?.classList?.contains('recharts-bar-cursor')) return false;
          return true;
        },
        style: {
          padding: isMobile ? '12px' : '24px',
          borderRadius: '16px'
        }
      });
      const link = document.createElement('a');
      const periodLabel = selectedPeriod === 'LAST_12' ? 'Ultimos_12_Meses' : selectedPeriod === 'LAST_6' ? 'Ultimos_6_Meses' : selectedPeriod === 'ALL' ? 'Historico' : selectedPeriod;
      link.download = `Informe_GPL_${selectedPlatform}_${periodLabel}.png`;
      link.href = dataUrl;
      link.click();
    } catch (err) {
      console.error('Error exportando imagen:', err);
      alert('No se pudo generar la imagen. Intenta nuevamente.');
    } finally {
      setExporting(false);
    }
  };

  // Download filtered data as Excel
  const handleDownloadExcel = () => {
    const wb = XLSX.utils.book_new();

    const evolutionRows = displayData.map(d => ({
      Año: d.year,
      Mes: d.monthName.split(' ')[0],
      Plataforma: d.platform,
      'Clientes Cobrados': Math.round(d.clientes),
      'Monto Cobrado ($)': Math.round(d.monto * 100) / 100,
      'Dif. Monto vs Mes Ant. ($)': Math.round((d.difMontos || 0) * 100) / 100,
      'Dif. Clientes vs Mes Ant.': Math.round(d.difClientes || 0)
    }));
    const wsEvol = XLSX.utils.json_to_sheet(evolutionRows);
    XLSX.utils.book_append_sheet(wb, wsEvol, 'Evolución');

    if (dailyChartData.length > 0 && activeDailyRecord) {
      const dailyRows = dailyChartData.map(d => ({
        Período: `${activeDailyRecord.monthName} ${activeDailyRecord.year}`,
        Plataforma: selectedPlatform,
        'Día Hábil': d.diaHabil,
        Fecha: d.fecha,
        'Clientes Día': d.Clientes,
        'Monto Día ($)': Math.round(d.Monto * 100) / 100,
        'Acum. Clientes Mes Actual': d.AcumClientes,
        'Acum. Clientes Mes Ant. (mismo día)': d.AcumClientesPrev,
        'Dif. Acum. Clientes vs Mes Ant.': d.DifAcumClientes,
        'Acum. Monto Mes Actual ($)': Math.round(d.AcumMonto * 100) / 100,
        'Acum. Monto Mes Ant. ($)': Math.round(d.AcumMontoPrev * 100) / 100,
        'Dif. Acum. Monto vs Mes Ant. ($)': Math.round(d.DifAcumMonto * 100) / 100
      }));
      const wsDaily = XLSX.utils.json_to_sheet(dailyRows);
      XLSX.utils.book_append_sheet(wb, wsDaily, 'Comparativo Día Hábil');
    }

    if (platformData.length > 0) {
      const platRows = platformData.map(p => ({
        Período: activePlatformMonthLabel,
        Plataforma: p.name,
        'Clientes Cobrados': Math.round(p.Clientes),
        'Monto Cobrado ($)': Math.round(p.Recaudacion * 100) / 100
      }));
      const wsPlat = XLSX.utils.json_to_sheet(platRows);
      XLSX.utils.book_append_sheet(wb, wsPlat, 'Por Plataforma');
    }

    const periodLabel = selectedPeriod === 'LAST_12' ? 'Ultimos_12_Meses' : selectedPeriod === 'LAST_6' ? 'Ultimos_6_Meses' : selectedPeriod === 'ALL' ? 'Historico' : selectedPeriod;
    XLSX.writeFile(wb, `Informe_GPL_${selectedPlatform}_${periodLabel}.xlsx`);
  };

  const showDetailedLabels = chartData.length <= 24;
  const periodText =
    selectedPeriod === 'LAST_12'
      ? 'Últimos 12 meses'
      : selectedPeriod === 'LAST_6'
      ? 'Últimos 6 meses'
      : selectedPeriod === 'ALL'
      ? 'Histórico completo'
      : `Año ${selectedPeriod}`;

  return (
    <div className="animate-fade-in">
      {/* Barra de controles superior */}
      <div className="toolbar-container">
        <div className="toolbar-filters">
          {data.length > 0 && (
            <>
              <div className="filter-group">
                <span className="filter-label">Plataforma:</span>
                <select
                  className="select-input"
                  value={selectedPlatform}
                  onChange={(e) => {
                    setSelectedPlatform(e.target.value);
                    setSelectedCutoffDay(0);
                  }}
                >
                  {uniquePlatforms.map(p => (
                    <option key={p} value={p}>
                      {p === 'GLOBAL' ? 'GLOBAL (Consolidado)' : `Plataforma ${p}`}
                    </option>
                  ))}
                </select>
              </div>

              <div className="filter-group">
                <span className="filter-label">Período:</span>
                <select
                  className="select-input"
                  value={selectedPeriod}
                  onChange={(e) => setSelectedPeriod(e.target.value)}
                >
                  <option value="LAST_6">Últimos 6 meses</option>
                  <option value="LAST_12">Últimos 12 meses</option>
                  <option value="ALL">Todo el histórico</option>
                  {[...new Set(platformFilteredData.map(d => d.year))].reverse().map(year => (
                    <option key={year} value={year.toString()}>Año {year}</option>
                  ))}
                </select>
              </div>
            </>
          )}
        </div>

        <div className="toolbar-actions">
          <button
            className="btn btn-secondary"
            onClick={() => setShowUploader(prev => !prev)}
            title="Subir o actualizar archivos RECA / TABLA ACUMULADOS"
          >
            <Upload size={16} /> {showUploader ? 'Ocultar Carga' : 'Actualizar Planillas'}
          </button>

          {data.length > 0 && (
            <>
              <button className="btn" onClick={handleDownloadImage} disabled={exporting}>
                <Download size={16} /> {exporting ? 'Generando...' : 'Descargar PNG'}
              </button>
              <button className="btn btn-teal" onClick={handleDownloadExcel}>
                <FileSpreadsheet size={16} /> Exportar Excel
              </button>
              <button className="btn btn-secondary" onClick={handleReloadDefaults} title="Recargar planillas desde GitHub">
                <RefreshCw size={16} /> {isMobile ? 'Base' : ''}
              </button>
              <button className="btn btn-secondary" onClick={handleClear} title="Limpiar datos del navegador">
                <Trash2 size={16} /> {isMobile ? 'Limpiar' : ''}
              </button>
            </>
          )}
        </div>
      </div>

      {uploadNotice && (
        <div className={`notice-banner ${uploadNotice.type === 'error' ? 'error' : ''}`}>
          <span>{uploadNotice.text}</span>
          <button
            onClick={() => setUploadNotice(null)}
            style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: 'inherit', display: 'flex' }}
          >
            <X size={18} />
          </button>
        </div>
      )}

      {/* Zona desplegable para cargar nuevos archivos Excel */}
      {(showUploader || data.length === 0) && (
        <div
          className={`upload-zone ${dragOver ? 'drag-over' : ''}`}
          onDrop={handleDrop}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onClick={() => document.getElementById('file-upload').click()}
        >
          <input
            id="file-upload"
            type="file"
            multiple
            accept=".xlsx,.xls"
            style={{ display: 'none' }}
            onChange={(e) => {
              if (e.target.files && e.target.files.length > 0) {
                processFiles(e.target.files);
                e.target.value = '';
              }
            }}
          />
          <UploadCloud className="icon" />
          <p>
            {loading
              ? 'Procesando planillas...'
              : 'Arrastra aquí los archivos RECA o TABLA ACUMULADOS (.xlsx), o toca para seleccionarlos'}
          </p>
        </div>
      )}

      {/* Área capturable del Informe Mensual Corporativo */}
      {latest && (
        <div ref={dashboardRef} id="dashboard-capture">
          {/* Encabezado Institucional Grupo ProLíder */}
          <div className="report-header">
            <div className="report-header-bar" />
            <div className="report-header-content">
              <div className="brand-block">
                <img src="./logo-gpl.png" alt="Grupo ProLíder" className="brand-logo" />
                <div className="brand-divider" />
                <div className="brand-titles">
                  <h1>INFORME DE RECAUDACIÓN Y COBRANZAS</h1>
                  <p>Grupo ProLíder • www.grupoprolider.com.ar</p>
                </div>
              </div>

              <div className="report-meta">
                <span className="meta-badge">
                  {selectedPlatform === 'GLOBAL' ? 'CONSOLIDADO GLOBAL' : `PLATAFORMA: ${selectedPlatform}`}
                </span>
                <span className="meta-sub">
                  Corte: <strong>{latest.monthName.split(' ')[0]} {latest.year}</strong> • Vista: <strong>{periodText}</strong>
                </span>
              </div>
            </div>
          </div>

          {/* Tarjetas de Indicadores Clave (KPIs) */}
          <div className="dashboard-grid">
            <div className="glass-panel kpi-card">
              <div className="kpi-title">Recaudación Último Mes ({latest.monthName.split(' ')[0]} {latest.year})</div>
              <div className="kpi-value">{formatMoney(latest.monto)}</div>
              <div className={`kpi-change ${latest.difMontos > 0 ? 'positive' : latest.difMontos < 0 ? 'negative' : 'neutral'}`}>
                {latest.difMontos > 0 ? <TrendingUp className="icon" /> : <TrendingDown className="icon" />}
                {formatMoney(Math.abs(latest.difMontos))} ({latest.pctMontos >= 0 ? '+' : ''}{latest.pctMontos.toFixed(1)}%) vs cierre mes anterior
              </div>
            </div>

            <div className="glass-panel kpi-card accent-cyan">
              <div className="kpi-title">Clientes Cobrados ({latest.monthName.split(' ')[0]} {latest.year})</div>
              <div className="kpi-value">{formatNumber(latest.clientes)}</div>
              <div className={`kpi-change ${latest.difClientes > 0 ? 'positive' : latest.difClientes < 0 ? 'negative' : 'neutral'}`}>
                {latest.difClientes > 0 ? <TrendingUp className="icon" /> : <TrendingDown className="icon" />}
                {formatNumber(Math.abs(latest.difClientes))} ({latest.pctClientes >= 0 ? '+' : ''}{latest.pctClientes.toFixed(1)}%) vs cierre mes anterior
              </div>
            </div>

            <div className="glass-panel kpi-card accent-magenta">
              <div className="kpi-title">Ticket Promedio por Cliente</div>
              <div className="kpi-value">
                {latest.clientes > 0 ? formatMoney(latest.monto / latest.clientes) : '$0'}
              </div>
              <div className="kpi-change neutral">
                Período evaluado: {latest.monthName.split(' ')[0]} {latest.year}
              </div>
            </div>
          </div>

          {/* NUEVO INFORME: Comparativo de Avance al Mismo Día Hábil vs. Mes Anterior */}
          {activeDailyRecord && prevDailyRecord && effectiveCutoffDay > 0 && (
            <div className="dashboard-grid">
              <div
                className="glass-panel"
                style={{
                  gridColumn: '1 / -1',
                  borderLeft: `6px solid ${difSameDayClientes >= 0 ? '#0f8a5f' : '#c82333'}`
                }}
              >
                <div
                  style={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    flexWrap: 'wrap',
                    gap: '0.75rem',
                    marginBottom: '0.9rem'
                  }}
                >
                  <div>
                    <h2 className="section-title" style={{ marginBottom: '0.2rem' }}>
                      <span className="section-dot" />
                      Comparativo al Mismo Día Hábil: {activeDailyRecord.monthName} vs. {prevDailyRecord.monthName} {prevDailyRecord.year}
                    </h2>
                    <p style={{ fontSize: '0.88rem', color: '#4e6772', fontWeight: 500 }}>
                      Comparación exacta del acumulado desde el <strong>Día Hábil 1 al Día Hábil {effectiveCutoffDay}</strong>{' '}
                      {currCutoffItem?.fecha ? `(corte ${currCutoffItem.fecha})` : ''} contra igual día hábil del mes anterior{' '}
                      {prevCutoffItem?.fecha ? `(${prevCutoffItem.fecha})` : ''}.
                    </p>
                  </div>

                  <div style={{ display: 'flex', gap: '0.65rem', flexWrap: 'wrap', alignItems: 'center' }}>
                    <div className="filter-group">
                      <span className="filter-label">Corte Día Hábil:</span>
                      <select
                        className="select-input"
                        value={effectiveCutoffDay}
                        onChange={(e) => setSelectedCutoffDay(Number(e.target.value))}
                      >
                        {Array.from({ length: maxCurrDiaHabil }, (_, i) => i + 1).map(dh => (
                          <option key={dh} value={dh}>
                            Hasta Día Hábil {dh}{dh === maxCurrDiaHabil ? ' (Último cargado)' : ''}
                          </option>
                        ))}
                      </select>
                    </div>

                    {recordsWithDaily.length > 1 && (
                      <div className="filter-group">
                        <span className="filter-label">Mes:</span>
                        <select
                          className="select-input"
                          value={activeDailyMonthKey}
                          onChange={(e) => {
                            setSelectedDailyMonth(e.target.value);
                            setSelectedCutoffDay(0);
                          }}
                        >
                          {recordsWithDaily.map(r => (
                            <option key={`${r.year}-${r.month}`} value={`${r.year}-${r.month}`}>
                              {r.monthName} {r.year}
                            </option>
                          ))}
                        </select>
                      </div>
                    )}
                  </div>
                </div>

                {/* Frase ejecutiva de diagnóstico directo */}
                <div
                  style={{
                    background: difSameDayClientes >= 0 && difSameDayMonto >= 0 ? '#ecfdf5' : '#fff7ed',
                    border: `1px solid ${difSameDayClientes >= 0 && difSameDayMonto >= 0 ? '#6ee7b7' : '#fdba74'}`,
                    color: '#093947',
                    padding: '0.85rem 1.1rem',
                    borderRadius: '10px',
                    marginBottom: '1rem',
                    fontSize: '0.95rem',
                    fontWeight: 600
                  }}
                >
                  Al <strong>Día Hábil {effectiveCutoffDay}</strong> de <strong>{activeDailyRecord.monthName}</strong>, estamos{' '}
                  <strong style={{ color: difSameDayClientes >= 0 ? '#0f8a5f' : '#c82333' }}>
                    {difSameDayClientes >= 0 ? 'ARRIBA' : 'ABAJO'} en {formatNumber(Math.abs(difSameDayClientes))} clientes ({pctSameDayClientes >= 0 ? '+' : ''}{pctSameDayClientes.toFixed(1)}%)
                  </strong>{' '}
                  y{' '}
                  <strong style={{ color: difSameDayMonto >= 0 ? '#0f8a5f' : '#c82333' }}>
                    {difSameDayMonto >= 0 ? 'ARRIBA' : 'ABAJO'} en {formatMoney(Math.abs(difSameDayMonto))} ({pctSameDayMonto >= 0 ? '+' : ''}{pctSameDayMonto.toFixed(1)}%)
                  </strong>{' '}
                  respecto al mismo día hábil de <strong>{prevDailyRecord.monthName}</strong>.
                </div>

                {/* 3 Tarjetas de Comparación al Día Hábil N */}
                <div
                  style={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(auto-fit, minmax(230px, 1fr))',
                    gap: '0.9rem'
                  }}
                >
                  <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '10px', padding: '0.9rem 1rem' }}>
                    <div style={{ fontSize: '0.75rem', fontWeight: 700, color: '#4e6772', textTransform: 'uppercase' }}>
                      Clientes Acumulados (Día 1 al {effectiveCutoffDay})
                    </div>
                    <div style={{ fontSize: '1.35rem', fontWeight: 800, color: difSameDayClientes >= 0 ? '#0f8a5f' : '#c82333', margin: '0.25rem 0' }}>
                      {difSameDayClientes >= 0 ? '+' : '-'}{formatNumber(Math.abs(difSameDayClientes))} clientes ({pctSameDayClientes >= 0 ? '+' : ''}{pctSameDayClientes.toFixed(1)}%)
                    </div>
                    <div style={{ fontSize: '0.82rem', color: '#4e6772', fontWeight: 600 }}>
                      {activeDailyRecord.monthName.substring(0, 3)}: <strong>{formatNumber(currAcumClientesAtCutoff)}</strong> vs.{' '}
                      {prevDailyRecord.monthName.substring(0, 3)}: <strong>{formatNumber(prevAcumClientesAtCutoff)}</strong>
                    </div>
                  </div>

                  <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '10px', padding: '0.9rem 1rem' }}>
                    <div style={{ fontSize: '0.75rem', fontWeight: 700, color: '#4e6772', textTransform: 'uppercase' }}>
                      Recaudación Acumulada (Día 1 al {effectiveCutoffDay})
                    </div>
                    <div style={{ fontSize: '1.35rem', fontWeight: 800, color: difSameDayMonto >= 0 ? '#0f8a5f' : '#c82333', margin: '0.25rem 0' }}>
                      {difSameDayMonto >= 0 ? '+' : '-'}{formatMoney(Math.abs(difSameDayMonto))} ({pctSameDayMonto >= 0 ? '+' : ''}{pctSameDayMonto.toFixed(1)}%)
                    </div>
                    <div style={{ fontSize: '0.82rem', color: '#4e6772', fontWeight: 600 }}>
                      {activeDailyRecord.monthName.substring(0, 3)}: <strong>{formatCompactMoney(currAcumMontoAtCutoff)}</strong> vs.{' '}
                      {prevDailyRecord.monthName.substring(0, 3)}: <strong>{formatCompactMoney(prevAcumMontoAtCutoff)}</strong>
                    </div>
                  </div>

                  <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '10px', padding: '0.9rem 1rem' }}>
                    <div style={{ fontSize: '0.75rem', fontWeight: 700, color: '#4e6772', textTransform: 'uppercase' }}>
                      Solo el Día Hábil {effectiveCutoffDay} (Puntual)
                    </div>
                    <div style={{ fontSize: '1.1rem', fontWeight: 800, color: '#093947', margin: '0.25rem 0' }}>
                      {formatNumber(currCutoffItem?.clientes || 0)} cl. • {formatCompactMoney(currCutoffItem?.monto || 0)}
                    </div>
                    <div style={{ fontSize: '0.82rem', color: difSingleDayClientes >= 0 ? '#0f8a5f' : '#c82333', fontWeight: 600 }}>
                      Vs. Día {effectiveCutoffDay} {prevDailyRecord.monthName.substring(0, 3)} ({formatNumber(prevCutoffItem?.clientes || 0)} cl.):{' '}
                      {difSingleDayClientes >= 0 ? '+' : ''}{formatNumber(difSingleDayClientes)} cl. ({difSingleDayMonto >= 0 ? '+' : ''}{formatCompactMoney(difSingleDayMonto)})
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Gráfico 1: Evolución de Recaudación ($) */}
          <div className="dashboard-grid">
            <div className="glass-panel" style={{ gridColumn: '1 / -1' }}>
              <h2 className="section-title">
                <span className="section-dot teal" />
                Evolución Mensual de Recaudación ($)
              </h2>
              <div className="chart-container" style={{ height: isMobile ? '290px' : '360px' }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart
                    data={chartData}
                    margin={{ top: 28, right: isMobile ? 8 : 20, left: isMobile ? -12 : 10, bottom: isMobile ? 20 : 10 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                    <XAxis
                      dataKey="name"
                      stroke="#4e6772"
                      angle={isMobile ? -35 : 0}
                      textAnchor={isMobile ? 'end' : 'middle'}
                      height={isMobile ? 48 : 30}
                      interval={0}
                      tick={{ fill: '#093947', fontSize: isMobile ? 10 : 12, fontWeight: 600 }}
                    />
                    <YAxis
                      stroke="#4e6772"
                      tick={{ fill: '#4e6772', fontSize: isMobile ? 10 : 12 }}
                      tickFormatter={(value) => `$${(value / 1000000).toFixed(0)}M`}
                    />
                    <Tooltip
                      isAnimationActive={false}
                      allowEscapeViewBox={{ x: false, y: false }}
                      contentStyle={{ backgroundColor: '#093947', border: 'none', borderRadius: '8px', color: '#fff' }}
                      itemStyle={{ color: '#fff' }}
                      labelStyle={{ color: '#21C0AC', fontWeight: 700 }}
                      formatter={(value) => [formatMoney(value), 'Recaudación']}
                      labelFormatter={(label, payload) => payload?.[0]?.payload?.fullName || label}
                    />
                    <Bar dataKey="Recaudacion" fill="#123C49" radius={[5, 5, 0, 0]} isAnimationActive={false}>
                      {showDetailedLabels && (
                        <LabelList
                          dataKey="RecaudacionLabel"
                          position="top"
                          isAnimationActive={false}
                          content={renderTopBarLabel('#093947', isMobile ? 9 : 11)}
                        />
                      )}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          {/* Gráfico 2: Evolución de Clientes Cobrados */}
          <div className="dashboard-grid">
            <div className="glass-panel" style={{ gridColumn: '1 / -1' }}>
              <h2 className="section-title">
                <span className="section-dot cyan" />
                Evolución Mensual de Clientes Cobrados
              </h2>
              <div className="chart-container" style={{ height: isMobile ? '280px' : '330px' }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart
                    data={chartData}
                    margin={{ top: 28, right: isMobile ? 8 : 20, left: isMobile ? -12 : 10, bottom: isMobile ? 20 : 10 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                    <XAxis
                      dataKey="name"
                      stroke="#4e6772"
                      angle={isMobile ? -35 : 0}
                      textAnchor={isMobile ? 'end' : 'middle'}
                      height={isMobile ? 48 : 30}
                      interval={0}
                      tick={{ fill: '#093947', fontSize: isMobile ? 10 : 12, fontWeight: 600 }}
                    />
                    <YAxis
                      stroke="#4e6772"
                      tick={{ fill: '#4e6772', fontSize: isMobile ? 10 : 12 }}
                      tickFormatter={(value) => formatNumber(value)}
                    />
                    <Tooltip
                      isAnimationActive={false}
                      allowEscapeViewBox={{ x: false, y: false }}
                      contentStyle={{ backgroundColor: '#093947', border: 'none', borderRadius: '8px', color: '#fff' }}
                      itemStyle={{ color: '#fff' }}
                      labelStyle={{ color: '#21C0AC', fontWeight: 700 }}
                      formatter={(value) => [formatNumber(value), 'Clientes Cobrados']}
                      labelFormatter={(label, payload) => payload?.[0]?.payload?.fullName || label}
                    />
                    <Bar dataKey="Clientes" fill="#159A8A" radius={[5, 5, 0, 0]} isAnimationActive={false}>
                      {showDetailedLabels && (
                        <LabelList
                          dataKey="ClientesLabel"
                          position="top"
                          isAnimationActive={false}
                          content={renderTopBarLabel('#093947', isMobile ? 9 : 11)}
                        />
                      )}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          {/* Gráfico 3: Curva de Cobranza por Día Hábil (con Comparativo opcional vs. Mes Anterior) */}
          {dailyChartData.length > 0 && activeDailyRecord && (
            <div className="dashboard-grid">
              <div className="glass-panel" style={{ gridColumn: '1 / -1' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.85rem', flexWrap: 'wrap', gap: '0.75rem' }}>
                  <h2 className="section-title" style={{ marginBottom: 0 }}>
                    <span className="section-dot" />
                    Curva de Cobranza por Día Hábil (Día 1 al {maxCurrDiaHabil}) — {activeDailyRecord.monthName} {activeDailyRecord.year}
                  </h2>

                  <div style={{ display: 'flex', gap: '0.75rem', flexWrap: 'wrap', alignItems: 'center' }}>
                    {prevDailyRecord && (
                      <div className="filter-group">
                        <span className="filter-label">Comparar:</span>
                        <select
                          className="select-input"
                          value={compareDailyPrev ? 'YES' : 'NO'}
                          onChange={(e) => setCompareDailyPrev(e.target.value === 'YES')}
                        >
                          <option value="YES">Vs. {prevDailyRecord.monthName}</option>
                          <option value="NO">Solo {activeDailyRecord.monthName}</option>
                        </select>
                      </div>
                    )}

                    <div className="filter-group">
                      <span className="filter-label">Métrica:</span>
                      <select
                        className="select-input"
                        value={dailyMetric}
                        onChange={(e) => setDailyMetric(e.target.value)}
                      >
                        <option value="MONTO">Recaudación ($)</option>
                        <option value="CLIENTES">Clientes Cobrados</option>
                      </select>
                    </div>

                    {recordsWithDaily.length > 1 && (
                      <div className="filter-group">
                        <span className="filter-label">Mes:</span>
                        <select
                          className="select-input"
                          value={activeDailyMonthKey}
                          onChange={(e) => {
                            setSelectedDailyMonth(e.target.value);
                            setSelectedCutoffDay(0);
                          }}
                        >
                          {recordsWithDaily.map(r => (
                            <option key={`${r.year}-${r.month}`} value={`${r.year}-${r.month}`}>
                              {r.monthName} {r.year}
                            </option>
                          ))}
                        </select>
                      </div>
                    )}
                  </div>
                </div>

                {/* Resumen ejecutivo del ritmo diario del mes */}
                <div
                  style={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(auto-fit, minmax(210px, 1fr))',
                    gap: '0.75rem',
                    marginBottom: '1rem',
                    background: '#f8fafc',
                    padding: '0.75rem 1rem',
                    borderRadius: '10px',
                    border: '1px solid #e2e8f0'
                  }}
                >
                  {peakDay && (
                    <div>
                      <div style={{ fontSize: '0.74rem', fontWeight: 700, color: '#4e6772', textTransform: 'uppercase' }}>
                        Pico Máximo de Cobro
                      </div>
                      <div style={{ fontSize: '0.95rem', fontWeight: 800, color: '#7B1C55' }}>
                        Día Hábil {peakDay.diaHabil} {peakDay.fecha ? `(${peakDay.fecha})` : ''} — {formatCompactMoney(peakDay.Monto)} ({formatNumber(peakDay.Clientes)} cl.)
                      </div>
                    </div>
                  )}

                  <div>
                    <div style={{ fontSize: '0.74rem', fontWeight: 700, color: '#4e6772', textTransform: 'uppercase' }}>
                      Primeros 10 Días Hábiles
                    </div>
                    <div style={{ fontSize: '0.95rem', fontWeight: 800, color: '#093947' }}>
                      {first10DaysPct.toFixed(1)}% del mes ({formatCompactMoney(first10DaysMonto)})
                    </div>
                  </div>

                  <div>
                    <div style={{ fontSize: '0.74rem', fontWeight: 700, color: '#4e6772', textTransform: 'uppercase' }}>
                      Promedio por Día Hábil
                    </div>
                    <div style={{ fontSize: '0.95rem', fontWeight: 800, color: '#159A8A' }}>
                      {formatCompactMoney(avgDailyMonto)} / día ({formatNumber(avgDailyClientes)} cl./día)
                    </div>
                  </div>
                </div>

                <div className="chart-container" style={{ height: isMobile ? '300px' : '360px' }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={dailyChartData}
                      margin={{ top: 28, right: isMobile ? 8 : 20, left: isMobile ? -12 : 10, bottom: isMobile ? 20 : 10 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                      <XAxis
                        dataKey="name"
                        stroke="#4e6772"
                        angle={isMobile ? -45 : 0}
                        textAnchor={isMobile ? 'end' : 'middle'}
                        height={isMobile ? 48 : 30}
                        interval={0}
                        tick={{ fill: '#093947', fontSize: isMobile ? 9 : 11, fontWeight: 600 }}
                      />
                      <YAxis
                        stroke="#4e6772"
                        tick={{ fill: '#4e6772', fontSize: isMobile ? 10 : 12 }}
                        tickFormatter={(value) =>
                          dailyMetric === 'MONTO' ? `$${(value / 1000000).toFixed(0)}M` : formatNumber(value)
                        }
                      />
                      <Tooltip
                        isAnimationActive={false}
                        allowEscapeViewBox={{ x: false, y: false }}
                        contentStyle={{ backgroundColor: '#093947', border: 'none', borderRadius: '8px', color: '#fff' }}
                        itemStyle={{ color: '#fff' }}
                        labelStyle={{ color: '#f78da7', fontWeight: 700 }}
                        formatter={(_value, name, props) => {
                          const p = props.payload;
                          if (name === 'MesAnterior') {
                            return [
                              `${formatMoney(p.MontoPrev)} • ${formatNumber(p.ClientesPrev)} cl. (Acum al Día ${p.diaHabil}: ${formatCompactMoney(p.AcumMontoPrev)} / ${formatNumber(p.AcumClientesPrev)} cl.)`,
                              `${prevDailyRecord?.monthName || 'Mes Anterior'}`
                            ];
                          }
                          const difCl = p.DifAcumClientes;
                          const difMo = p.DifAcumMonto;
                          return [
                            `${formatMoney(p.Monto)} • ${formatNumber(p.Clientes)} cl. (Acum al Día ${p.diaHabil}: ${difCl >= 0 ? '+' : ''}${formatNumber(difCl)} cl. / ${difMo >= 0 ? '+' : ''}${formatCompactMoney(difMo)} vs mes ant.)`,
                            `${activeDailyRecord.monthName}`
                          ];
                        }}
                        labelFormatter={(label, payload) => payload?.[0]?.payload?.fullName || label}
                      />
                      {compareDailyPrev && prevDailyRecord && (
                        <Legend
                          verticalAlign="top"
                          height={28}
                          formatter={(val) =>
                            val === 'MesAnterior'
                              ? `${prevDailyRecord.monthName} ${prevDailyRecord.year} (Mes Anterior)`
                              : `${activeDailyRecord.monthName} ${activeDailyRecord.year} (Actual)`
                          }
                        />
                      )}
                      {compareDailyPrev && prevDailyRecord && (
                        <Bar
                          dataKey="ValorPrev"
                          name="MesAnterior"
                          fill="#cbd5e1"
                          radius={[4, 4, 0, 0]}
                          isAnimationActive={false}
                        />
                      )}
                      <Bar
                        dataKey="Valor"
                        name="MesActual"
                        fill={dailyMetric === 'MONTO' ? '#7B1C55' : '#159A8A'}
                        radius={[4, 4, 0, 0]}
                        isAnimationActive={false}
                      >
                        <LabelList
                          dataKey="Label"
                          position="top"
                          isAnimationActive={false}
                          content={renderTopBarLabel('#7B1C55', isMobile ? 8 : 10)}
                        />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
          )}

          {/* Gráfico 4: Desglose por Plataforma */}
          {platformData.length > 0 && (
            <div className="dashboard-grid">
              <div className="glass-panel" style={{ gridColumn: '1 / -1' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem', flexWrap: 'wrap', gap: '0.75rem' }}>
                  <h2 className="section-title" style={{ marginBottom: 0 }}>
                    <span className="section-dot" />
                    Participación por Plataforma — {activePlatformMonthLabel}
                  </h2>
                  {availablePlatformMonths.length > 1 && (
                    <div className="filter-group">
                      <span className="filter-label">Mes RECA:</span>
                      <select
                        className="select-input"
                        value={activePlatformMonthKey || ''}
                        onChange={(e) => setSelectedPlatformMonth(e.target.value)}
                      >
                        {availablePlatformMonths.map(m => (
                          <option key={m.key} value={m.key}>{m.label}</option>
                        ))}
                      </select>
                    </div>
                  )}
                </div>

                <div className="chart-container" style={{ height: `${Math.max(270, platformData.length * (isMobile ? 42 : 46))}px` }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={platformData}
                      layout="vertical"
                      margin={{ top: 8, right: isMobile ? 115 : 175, left: isMobile ? 4 : 20, bottom: 8 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                      <XAxis
                        type="number"
                        stroke="#4e6772"
                        tick={{ fill: '#4e6772', fontSize: isMobile ? 10 : 12 }}
                        tickFormatter={(value) => `$${(value / 1000000).toFixed(0)}M`}
                      />
                      <YAxis
                        dataKey="name"
                        type="category"
                        stroke="#4e6772"
                        tick={{ fill: '#093947', fontSize: isMobile ? 10 : 12, fontWeight: 700 }}
                        width={isMobile ? 74 : 100}
                      />
                      <Tooltip
                        isAnimationActive={false}
                        allowEscapeViewBox={{ x: false, y: false }}
                        contentStyle={{ backgroundColor: '#093947', border: 'none', borderRadius: '8px', color: '#fff' }}
                        itemStyle={{ color: '#fff' }}
                        labelStyle={{ color: '#f78da7', fontWeight: 700 }}
                        formatter={(value, _name, props) => [
                          `${formatMoney(value)} (${formatNumber(props.payload.Clientes)} clientes)`,
                          'Recaudación'
                        ]}
                      />
                      <Bar dataKey="Recaudacion" fill="#7B1C55" radius={[0, 6, 6, 0]} isAnimationActive={false}>
                        <LabelList
                          dataKey="barLabel"
                          position="right"
                          isAnimationActive={false}
                          content={renderRightBarLabel('#7B1C55', isMobile ? 10 : 12)}
                        />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
          )}

          <div className="report-footer">
            <span>Grupo ProLíder (GPL) • Sistema de Estadísticas de Cobranza</span>
            <span>www.grupoprolider.com.ar</span>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
