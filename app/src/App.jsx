import React, { useState, useEffect, useCallback, useRef } from 'react';
import {
  UploadCloud, TrendingUp, TrendingDown, Trash2, Download, FileSpreadsheet, RefreshCw, Upload
} from 'lucide-react';
import { toPng } from 'html-to-image';
import * as XLSX from 'xlsx';
import { parseExcelFile } from './utils/parser';
import { saveData, loadData, clearData, seedDefaultFiles } from './utils/storage';
import {
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, BarChart, Bar, LabelList
} from 'recharts';

function App() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [showUploader, setShowUploader] = useState(false);
  const [selectedPlatform, setSelectedPlatform] = useState('GLOBAL');
  const [selectedPeriod, setSelectedPeriod] = useState('LAST_12');
  const [selectedPlatformMonth, setSelectedPlatformMonth] = useState('');
  const dashboardRef = useRef(null);

  useEffect(() => {
    loadData().then(setData);
  }, []);

  const processFiles = async (files) => {
    setLoading(true);
    try {
      let newData = [...data];
      for (const file of files) {
        const parsed = await parseExcelFile(file);
        newData = await saveData(parsed, file.name);
      }
      setData(newData);
      setShowUploader(false);
    } catch (error) {
      console.error('Error parsing files', error);
      alert('Hubo un error procesando el archivo. Asegúrate de que sea el formato correcto.');
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
    }
  };

  const handleReloadDefaults = async () => {
    setLoading(true);
    try {
      const seeded = await seedDefaultFiles();
      setData(seeded);
      setShowUploader(false);
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
      return `$${(value / 1000000).toLocaleString('es-AR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}M`;
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
          barLabel: `${formatCompactMoney(d.monto)} (${formatNumber(d.clientes)} cl.)`
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
        style: {
          padding: '24px',
          borderRadius: '16px'
        }
      });
      const link = document.createElement('a');
      const periodLabel = selectedPeriod === 'LAST_12' ? 'Ultimos_12_Meses' : selectedPeriod === 'ALL' ? 'Historico' : selectedPeriod;
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

    const periodLabel = selectedPeriod === 'LAST_12' ? 'Ultimos_12_Meses' : selectedPeriod === 'ALL' ? 'Historico' : selectedPeriod;
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
        <div style={{ display: 'flex', gap: '1rem', flexWrap: 'wrap', alignItems: 'center' }}>
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

        <div style={{ display: 'flex', gap: '0.6rem', flexWrap: 'wrap' }}>
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
                <Download size={16} /> {exporting ? 'Generando PNG...' : 'Descargar Informe (PNG)'}
              </button>
              <button className="btn btn-teal" onClick={handleDownloadExcel}>
                <FileSpreadsheet size={16} /> Exportar Excel
              </button>
              <button className="btn btn-secondary" onClick={handleReloadDefaults} title="Recargar planillas originales">
                <RefreshCw size={16} />
              </button>
              <button className="btn btn-secondary" onClick={handleClear} title="Limpiar datos del navegador">
                <Trash2 size={16} />
              </button>
            </>
          )}
        </div>
      </div>

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
            onChange={(e) => e.target.files && processFiles(e.target.files)}
          />
          <UploadCloud className="icon" />
          <p>
            {loading
              ? 'Procesando planillas...'
              : 'Arrastra aquí los archivos RECA o TABLA ACUMULADOS (.xlsx), o haz clic para seleccionarlos'}
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
              <div className="chart-container" style={{ height: '360px' }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={chartData} margin={{ top: 28, right: 20, left: 10, bottom: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                    <XAxis dataKey="name" stroke="#4e6772" tick={{ fill: '#093947', fontSize: 12, fontWeight: 600 }} />
                    <YAxis
                      stroke="#4e6772"
                      tick={{ fill: '#4e6772', fontSize: 12 }}
                      tickFormatter={(value) => `$${(value / 1000000).toFixed(0)}M`}
                    />
                    <Tooltip
                      contentStyle={{ backgroundColor: '#093947', border: 'none', borderRadius: '8px', color: '#fff' }}
                      itemStyle={{ color: '#fff' }}
                      labelStyle={{ color: '#21C0AC', fontWeight: 700 }}
                      formatter={(value) => [formatMoney(value), 'Recaudación']}
                      labelFormatter={(label, payload) => payload?.[0]?.payload?.fullName || label}
                    />
                    <Bar dataKey="Recaudacion" fill="#123C49" radius={[6, 6, 0, 0]}>
                      {showDetailedLabels && (
                        <LabelList
                          dataKey="RecaudacionLabel"
                          position="top"
                          fill="#093947"
                          fontSize={11}
                          fontWeight={700}
                          offset={8}
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
              <div className="chart-container" style={{ height: '330px' }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={chartData} margin={{ top: 28, right: 20, left: 10, bottom: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                    <XAxis dataKey="name" stroke="#4e6772" tick={{ fill: '#093947', fontSize: 12, fontWeight: 600 }} />
                    <YAxis
                      stroke="#4e6772"
                      tick={{ fill: '#4e6772', fontSize: 12 }}
                      tickFormatter={(value) => formatNumber(value)}
                    />
                    <Tooltip
                      contentStyle={{ backgroundColor: '#093947', border: 'none', borderRadius: '8px', color: '#fff' }}
                      itemStyle={{ color: '#fff' }}
                      labelStyle={{ color: '#21C0AC', fontWeight: 700 }}
                      formatter={(value) => [formatNumber(value), 'Clientes Cobrados']}
                      labelFormatter={(label, payload) => payload?.[0]?.payload?.fullName || label}
                    />
                    <Bar dataKey="Clientes" fill="#159A8A" radius={[6, 6, 0, 0]}>
                      {showDetailedLabels && (
                        <LabelList
                          dataKey="ClientesLabel"
                          position="top"
                          fill="#093947"
                          fontSize={11}
                          fontWeight={700}
                          offset={8}
                        />
                      )}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          {/* Gráfico 3: Desglose por Plataforma */}
          {platformData.length > 0 && (
            <div className="dashboard-grid">
              <div className="glass-panel" style={{ gridColumn: '1 / -1' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem', flexWrap: 'wrap', gap: '0.75rem' }}>
                  <h2 className="section-title" style={{ marginBottom: 0 }}>
                    <span className="section-dot" />
                    Participación por Plataforma de Cobranza — {activePlatformMonthLabel}
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

                <div className="chart-container" style={{ height: `${Math.max(290, platformData.length * 46)}px` }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={platformData} layout="vertical" margin={{ top: 10, right: 175, left: 20, bottom: 10 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                      <XAxis
                        type="number"
                        stroke="#4e6772"
                        tick={{ fill: '#4e6772', fontSize: 12 }}
                        tickFormatter={(value) => `$${(value / 1000000).toFixed(0)}M`}
                      />
                      <YAxis
                        dataKey="name"
                        type="category"
                        stroke="#4e6772"
                        tick={{ fill: '#093947', fontSize: 12, fontWeight: 700 }}
                        width={100}
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
                      <Bar dataKey="Recaudacion" fill="#7B1C55" radius={[0, 6, 6, 0]}>
                        <LabelList
                          dataKey="barLabel"
                          position="right"
                          fill="#7B1C55"
                          fontSize={12}
                          fontWeight={700}
                          offset={10}
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
