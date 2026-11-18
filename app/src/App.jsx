import React, { useState, useEffect, useCallback, useRef } from 'react';
import {
  UploadCloud, TrendingUp, TrendingDown, Trash2, Download, FileSpreadsheet, RefreshCw, Upload, X
} from 'lucide-react';
import { toPng } from 'html-to-image';
import * as XLSX from 'xlsx';
import { parseExcelFile } from './utils/parser';
import { saveData, loadData, clearData, seedDefaultFiles } from './utils/storage';
import {
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, BarChart, Bar, LabelList
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
  const [dailyMetric, setDailyMetric] = useState('MONTO');
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
        text: 'Se recargaron las planillas base publicadas en el servidor.'
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
    if (value >= 1000000) {
      return `$${(value / 1000000).toLocaleString('es-AR', { minimumFractionDigits: isMobile ? 0 : 1, maximumFractionDigits: isMobile ? 0 : 1 })}M`;
    }
    return `$${Math.round(value / 1000).toLocaleString('es-AR')}K`;
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

  const dailyTotalMonto = activeDailyRecord
    ? activeDailyRecord.daily.reduce((acc, item) => acc + (item.monto || 0), 0) || activeDailyRecord.monto
    : 0;

  const dailyChartData = activeDailyRecord
    ? activeDailyRecord.daily.map((d, idx) => {
        const pctAcum = dailyTotalMonto > 0 ? ((d.acumulado || 0) / dailyTotalMonto) * 100 : 0;
        return {
          name: isMobile ? `D${d.diaHabil || idx + 1}` : `Día ${d.diaHabil || idx + 1}`,
          fullName: `Día Hábil ${d.diaHabil || idx + 1}${d.fecha ? ` (${d.fecha})` : ''}`,
          diaHabil: d.diaHabil || idx + 1,
          fecha: d.fecha || '',
          Monto: d.monto,
          Clientes: d.clientes,
          Valor: dailyMetric === 'MONTO' ? d.monto : d.clientes,
          Label: dailyMetric === 'MONTO' ? formatCompactMoney(d.monto) : formatNumber(d.clientes),
          pctAcum: Math.min(100, pctAcum)
        };
      })
    : [];

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
        'Clientes Cobrados': d.Clientes,
        'Monto Cobrado ($)': Math.round(d.Monto * 100) / 100,
        '% Acumulado del Mes': `${d.pctAcum.toFixed(1)}%`
      }));
      const wsDaily = XLSX.utils.json_to_sheet(dailyRows);
      XLSX.utils.book_append_sheet(wb, wsDaily, 'Por Día Hábil');
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
                  onChange={(e) => setSelectedPlatform(e.target.value)}
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
              <button className="btn btn-secondary" onClick={handleReloadDefaults} title="Recargar planillas originales">
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
                {formatMoney(Math.abs(latest.difMontos))} ({latest.pctMontos >= 0 ? '+' : ''}{latest.pctMontos.toFixed(1)}%) vs mes anterior
              </div>
            </div>

            <div className="glass-panel kpi-card accent-cyan">
              <div className="kpi-title">Clientes Cobrados ({latest.monthName.split(' ')[0]} {latest.year})</div>
              <div className="kpi-value">{formatNumber(latest.clientes)}</div>
              <div className={`kpi-change ${latest.difClientes > 0 ? 'positive' : latest.difClientes < 0 ? 'negative' : 'neutral'}`}>
                {latest.difClientes > 0 ? <TrendingUp className="icon" /> : <TrendingDown className="icon" />}
                {formatNumber(Math.abs(latest.difClientes))} ({latest.pctClientes >= 0 ? '+' : ''}{latest.pctClientes.toFixed(1)}%) vs mes anterior
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

          {/* Gráfico Nuevo (#5): Curva de Cobranza por Día Hábil (Día 1 al 22) */}
          {dailyChartData.length > 0 && activeDailyRecord && (
            <div className="dashboard-grid">
              <div className="glass-panel" style={{ gridColumn: '1 / -1' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.85rem', flexWrap: 'wrap', gap: '0.75rem' }}>
                  <h2 className="section-title" style={{ marginBottom: 0 }}>
                    <span className="section-dot" />
                    Curva de Cobranza por Día Hábil (Día 1 al {dailyChartData.length}) — {activeDailyRecord.monthName} {activeDailyRecord.year}
                  </h2>

                  <div style={{ display: 'flex', gap: '0.75rem', flexWrap: 'wrap', alignItems: 'center' }}>
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
                          onChange={(e) => setSelectedDailyMonth(e.target.value)}
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

                <div className="chart-container" style={{ height: isMobile ? '290px' : '350px' }}>
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
                        contentStyle={{ backgroundColor: '#093947', border: 'none', borderRadius: '8px', color: '#fff' }}
                        itemStyle={{ color: '#fff' }}
                        labelStyle={{ color: '#f78da7', fontWeight: 700 }}
                        formatter={(_value, _name, props) => {
                          const p = props.payload;
                          return [
                            `${formatMoney(p.Monto)} • ${formatNumber(p.Clientes)} clientes (Acum: ${p.pctAcum.toFixed(1)}%)`,
                            'Cobranza del Día'
                          ];
                        }}
                        labelFormatter={(label, payload) => payload?.[0]?.payload?.fullName || label}
                      />
                      <Bar
                        dataKey="Valor"
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
