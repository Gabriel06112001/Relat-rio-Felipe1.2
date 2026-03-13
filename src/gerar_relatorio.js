const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ─── Leitura da planilha ─────────────────────────────────────────
const wb = XLSX.readFile(path.join(__dirname, '..', 'data', 'CAP_Semana_12.xlsx'));
const ws = wb.Sheets['Planilha1'];
const rawData = XLSX.utils.sheet_to_json(ws, { defval: '' });

// Colunas relevantes (sem __EMPTY*)
const COLUNAS = [
  'id', 'data_inclusao', 'data_vencimento', 'data_pagamento',
  'descricao', 'nf', 'categoria', 'sub_categoria', 'valor',
  'verdadeiro ou falso', 'origem', 'tipo_pagamento', 'status',
  'responsavel', 'banco', 'observação', 'Recorrencia'
];

// Converte serial Excel para data legível
function excelDateToStr(serial) {
  if (!serial || isNaN(serial)) return '';
  const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
  return date.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
}

// Campos de data
const DATE_FIELDS = ['data_inclusao', 'data_vencimento', 'data_pagamento'];

// Normalizar dados
const dados = rawData.map(row => {
  const obj = {};
  COLUNAS.forEach(col => {
    let val = row[col];
    if (DATE_FIELDS.includes(col) && typeof val === 'number') {
      val = excelDateToStr(val);
    } else if (col === 'verdadeiro ou falso') {
      val = val === true || val === 'TRUE' || val === 'true' ? 'Sim' : (val === false || val === 'FALSE' || val === 'false' ? 'Não' : '');
    } else if (col === 'valor') {
      val = typeof val === 'number' ? val : parseFloat(val) || 0;
    }
    obj[col] = val === undefined || val === null ? '' : val;
  });
  return obj;
});

// ─── Estatísticas gerais ──────────────────────────────────────────
const totalGeral = dados.reduce((s, r) => s + (r['valor'] || 0), 0);
const totalRegistros = dados.length;

// Por categoria
const porCategoria = {};
dados.forEach(r => {
  const cat = r['categoria'] || 'Sem categoria';
  porCategoria[cat] = (porCategoria[cat] || 0) + (r['valor'] || 0);
});

// Por status
const porStatus = {};
dados.forEach(r => {
  const s = r['status'] || 'Sem status';
  porStatus[s] = (porStatus[s] || 0) + 1;
});

// Por tipo_pagamento
const porTipo = {};
dados.forEach(r => {
  const t = r['tipo_pagamento'] || 'N/A';
  porTipo[t] = (porTipo[t] || 0) + (r['valor'] || 0);
});

// Por responsavel
const porResp = {};
dados.forEach(r => {
  const resp = r['responsavel'] || 'N/A';
  porResp[resp] = (porResp[resp] || 0) + (r['valor'] || 0);
});

// Por sub_categoria (top 8)
const porSubCat = {};
dados.forEach(r => {
  const sc = r['sub_categoria'] || 'N/A';
  porSubCat[sc] = (porSubCat[sc] || 0) + (r['valor'] || 0);
});
const topSubCat = Object.entries(porSubCat)
  .sort((a, b) => b[1] - a[1])
  .slice(0, 8);

// ─── Atrasados ───────────────────────────────────────────────────
// Considera hoje como a data de geração do relatório
const HOJE_STR = new Date().toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' });
function strToDate(str) {
  if (!str || typeof str !== 'string') return null;
  const [d, m, y] = str.split('/').map(Number);
  if (!d || !m || !y) return null;
  return new Date(y, m - 1, d);
}
const hoje = strToDate(HOJE_STR);

const atrasados = dados.filter(r => {
  const dataVenc = strToDate(r['data_vencimento']);
  if (!dataVenc) return false;
  const status = (r['status'] || '').toUpperCase();
  return dataVenc <= hoje && !status.includes('PAGO');
});

const totalAtrasado = atrasados.reduce((s, r) => s + (r['valor'] || 0), 0);
const atrasadosJSON = JSON.stringify(atrasados);

// Por categoria dos atrasados
const atrasadosPorCat = {};
atrasados.forEach(r => {
  const cat = r['categoria'] || 'Sem categoria';
  atrasadosPorCat[cat] = (atrasadosPorCat[cat] || 0) + (r['valor'] || 0);
});

// ─── Serializar dados para o HTML ────────────────────────────────
const dadosJSON = JSON.stringify(dados);
const colunasJSON = JSON.stringify(COLUNAS);

// ─── Paleta de cores ─────────────────────────────────────────────
const CORES = [
  '#2563eb','#64748b','#059669','#d97706','#0891b2',
  '#7c3aed','#dc2626','#0f766e','#475569','#1d4ed8'
];

const fmt = v => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

// ─── Helper: botões de tipo de gráfico ───────────────────────────
const TYPE_LABELS = { bar:'Barras', line:'Linha', doughnut:'Rosca', pie:'Pizza', radar:'Radar', polarArea:'Polar' };
function typeBtns(chartId, types, def) {
  return `<div class="chart-type-btns">${types.map(t =>
    `<button class="chart-type-btn${t===def?' active':''}" onclick="switchType('${chartId}','${t}')">${TYPE_LABELS[t]}</button>`
  ).join('')}</div>`;
}

// ─── HTML ─────────────────────────────────────────────────────────
const html = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Relatório CAP – Semana 12</title>

  <!-- Fonts -->
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet" />

  <!-- Chart.js -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"><\/script>

  <!-- SheetJS (export Excel) -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"><\/script>

  <!-- ag-Grid Community (tabela interativa) -->
  <script src="https://cdn.jsdelivr.net/npm/ag-grid-community@31.3.2/dist/ag-grid-community.min.js"><\/script>
  <link  href="https://cdn.jsdelivr.net/npm/ag-grid-community@31.3.2/styles/ag-grid.min.css" rel="stylesheet" />
  <link  href="https://cdn.jsdelivr.net/npm/ag-grid-community@31.3.2/styles/ag-theme-alpine.min.css" rel="stylesheet" />

  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    :root {
      --primary: #2563eb;
      --primary-dark: #1e3a5f;
      --accent:  #059669;
      --bg:      #f4f6f9;
      --surface: #ffffff;
      --text:    #1e293b;
      --muted:   #64748b;
      --radius:  12px;
      --shadow:  0 2px 12px rgba(0,0,0,.07);
    }

    body {
      font-family: 'Inter', sans-serif;
      background: var(--bg);
      color: var(--text);
      min-height: 100vh;
    }

    /* ── Header ── */
    header {
      background: linear-gradient(135deg, var(--primary-dark) 0%, var(--primary) 100%);
      color: #fff;
      padding: 28px 40px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 12px;
    }
    header h1 { font-size: 1.6rem; font-weight: 700; letter-spacing: -.5px; }
    header p  { font-size: .9rem; opacity: .8; margin-top: 4px; }
    .badge {
      background: rgba(255,255,255,.2);
      border-radius: 20px;
      padding: 6px 16px;
      font-size: .8rem;
      font-weight: 600;
    }

    /* ── Layout ── */
    .container { max-width: 1400px; margin: 0 auto; padding: 28px 24px; }

    /* ── KPI Cards ── */
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
      gap: 16px;
      margin-bottom: 28px;
    }
    .kpi-card {
      background: var(--surface);
      border-radius: var(--radius);
      padding: 20px 24px;
      box-shadow: var(--shadow);
      border-left: 4px solid var(--primary);
      transition: transform .2s;
    }
    .kpi-card:hover { transform: translateY(-3px); }
    .kpi-card:nth-child(2) { border-left-color: #0891b2; }
    .kpi-card:nth-child(3) { border-left-color: #d97706; }
    .kpi-card:nth-child(4) { border-left-color: #059669; }
    .kpi-card .label { font-size: .75rem; font-weight: 600; text-transform: uppercase; color: var(--muted); letter-spacing: .5px; }
    .kpi-card .value { font-size: 1.5rem; font-weight: 700; margin-top: 6px; color: var(--text); }
    .kpi-card .sub   { font-size: .78rem; color: var(--muted); margin-top: 4px; }

    /* ── Charts ── */
    .charts-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(340px, 1fr));
      gap: 20px;
      margin-bottom: 28px;
    }
    .chart-card {
      background: var(--surface);
      border-radius: var(--radius);
      padding: 22px;
      box-shadow: var(--shadow);
      transition: opacity .35s ease, transform .35s ease;
    }
    .chart-card.chart-hidden {
      display: none !important;
    }
    .chart-card.chart-fading-in {
      animation: fadeInUp .4s ease both;
    }
    @keyframes fadeInUp {
      from { opacity: 0; transform: translateY(14px); }
      to   { opacity: 1; transform: translateY(0); }
    }
    .chart-card h3 {
      font-size: .88rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: .5px;
      color: var(--muted);
      flex: 1;
    }
    .chart-header {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 10px;
      margin-bottom: 14px;
      border-bottom: 1px solid #eee;
      padding-bottom: 10px;
      flex-wrap: wrap;
    }
    .chart-type-btns {
      display: flex;
      gap: 4px;
      flex-wrap: wrap;
      justify-content: flex-end;
      flex-shrink: 0;
    }
    .chart-type-btn {
      padding: 3px 9px;
      border: 1.5px solid #dde3f5;
      border-radius: 6px;
      background: #f4f6ff;
      color: var(--muted);
      font-size: .70rem;
      font-weight: 600;
      cursor: pointer;
      transition: all .15s;
      white-space: nowrap;
    }
    .chart-type-btn:hover { background: #e0e7ff; border-color: var(--primary); color: var(--primary); }
    .chart-type-btn.active { background: var(--primary); border-color: var(--primary); color: #fff; }
    .chart-wrap { position: relative; height: 260px; }
    .chart-wrap-featured { position: relative; height: 340px; }
    .chart-desc {
      font-size: .78rem;
      color: var(--muted);
      margin-top: 10px;
      margin-bottom: 0;
      line-height: 1.5;
      border-top: 1px dashed #e5e7f3;
      padding-top: 8px;
    }

    /* ── Card em destaque ── */
    .chart-card-featured {
      grid-column: 1 / -1;
      background: #1e293b;
      color: #fff;
      border-radius: var(--radius);
      padding: 28px 30px 22px;
      box-shadow: 0 4px 20px rgba(0,0,0,.16);
      transition: opacity .35s ease, transform .35s ease;
      position: relative;
    }
    .chart-card-featured .chart-header {
      border-bottom-color: rgba(255,255,255,.15);
    }
    .chart-card-featured h3 {
      color: rgba(255,255,255,.7);
      font-size: 1rem;
    }
    .chart-card-featured .chart-type-btn {
      background: rgba(255,255,255,.08);
      border-color: rgba(255,255,255,.18);
      color: rgba(255,255,255,.7);
    }
    .chart-card-featured .chart-type-btn:hover {
      background: rgba(255,255,255,.15);
      border-color: rgba(255,255,255,.5);
      color: #fff;
    }
    .chart-card-featured .chart-type-btn.active {
      background: #2563eb;
      border-color: #2563eb;
      color: #fff;
    }
    .chart-card-featured .chart-desc {
      color: rgba(255,255,255,.5);
      border-top-color: rgba(255,255,255,.1);
    }
    /* Mini-KPIs de dias */
    .dias-kpis {
      display: grid;
      grid-template-columns: repeat(5, 1fr);
      gap: 10px;
      margin-top: 18px;
      position: relative;
      z-index: 1;
    }
    .dia-kpi {
      background: rgba(255,255,255,.1);
      border-radius: 10px;
      padding: 12px 10px;
      text-align: center;
      border: 1.5px solid rgba(255,255,255,.12);
      transition: transform .2s;
    }
    .dia-kpi:hover { transform: translateY(-2px); }
    .dia-kpi.dia-maior {
      background: rgba(37,99,235,.25);
      border-color: #2563eb;
    }
    .dia-kpi .dia-nome  { font-size: .72rem; font-weight: 700; text-transform: uppercase; letter-spacing: .5px; color: rgba(255,255,255,.7); margin-bottom: 4px; }
    .dia-kpi .dia-valor { font-size: .92rem; font-weight: 700; color: #fff; line-height: 1.2; }
    .dia-kpi .dia-qtd   { font-size: .68rem; color: rgba(255,255,255,.55); margin-top: 3px; }

    /* ── Column Selector ── */
    .section-title {
      font-size: 1rem;
      font-weight: 700;
      color: var(--primary-dark);
      margin-bottom: 14px;
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .section-title::before {
      content: '';
      display: inline-block;
      width: 4px; height: 18px;
      background: var(--primary);
      border-radius: 2px;
    }

    .col-selector-card {
      background: var(--surface);
      border-radius: var(--radius);
      padding: 22px;
      box-shadow: var(--shadow);
      margin-bottom: 20px;
    }
    .col-selector-actions {
      display: flex;
      gap: 10px;
      margin-bottom: 14px;
      flex-wrap: wrap;
    }
    .btn {
      padding: 7px 18px;
      border: none;
      border-radius: 6px;
      font-size: .83rem;
      font-weight: 600;
      cursor: pointer;
      transition: opacity .2s, transform .15s;
    }
    .btn:hover { opacity: .85; transform: translateY(-1px); }
    .btn-primary   { background: var(--primary); color: #fff; }
    .btn-secondary { background: #e9ecef; color: var(--text); }
    .btn-danger    { background: var(--accent); color: #fff; }
    .btn-success   { background: #2a9d8f; color: #fff; }

    .col-checks {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
    }
    .col-check-item {
      display: flex;
      align-items: center;
      gap: 6px;
      background: #f4f6ff;
      border: 1.5px solid #dde3f5;
      border-radius: 8px;
      padding: 6px 14px;
      cursor: pointer;
      transition: background .15s, border-color .15s;
      user-select: none;
    }
    .col-check-item:hover { background: #e0e7ff; }
    .col-check-item.checked { background: #4361ee18; border-color: var(--primary); }
    .col-check-item input { accent-color: var(--primary); cursor: pointer; }
    .col-check-item span  { font-size: .82rem; font-weight: 500; }

    /* ── ag-Grid wrapper ── */
    .table-card {
      background: var(--surface);
      border-radius: var(--radius);
      padding: 22px;
      box-shadow: var(--shadow);
      margin-bottom: 28px;
    }
    .table-toolbar {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 14px;
      align-items: center;
      justify-content: space-between;
    }
    .search-box {
      padding: 8px 14px;
      border: 1.5px solid #dde3f5;
      border-radius: 8px;
      font-size: .85rem;
      width: 280px;
      outline: none;
      transition: border-color .2s;
    }
    .search-box:focus { border-color: var(--primary); }
    #grid { width: 100%; height: 500px; }

    /* ── Footer ── */
    footer {
      text-align: center;
      padding: 18px;
      font-size: .78rem;
      color: var(--muted);
    }

    /* ── Seção Atrasados ── */
    .atraso-section {
      background: #fff5f5;
      border: 2px solid #fca5a5;
      border-radius: var(--radius);
      padding: 22px 24px;
      margin-bottom: 28px;
    }
    .atraso-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 12px;
      margin-bottom: 18px;
    }
    .atraso-title {
      font-size: 1rem;
      font-weight: 700;
      color: #b91c1c;
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .atraso-kpis {
      display: flex;
      gap: 16px;
      flex-wrap: wrap;
      margin-bottom: 18px;
    }
    .atraso-kpi {
      background: #fee2e2;
      border-radius: 10px;
      padding: 12px 20px;
      min-width: 160px;
    }
    .atraso-kpi .ak-label {
      font-size: .72rem;
      font-weight: 700;
      text-transform: uppercase;
      color: #b91c1c;
      letter-spacing: .4px;
      margin-bottom: 4px;
    }
    .atraso-kpi .ak-value {
      font-size: 1.2rem;
      font-weight: 700;
      color: #7f1d1d;
    }
    .atraso-cat-bars {
      display: flex;
      flex-direction: column;
      gap: 7px;
      margin-bottom: 18px;
    }
    .atraso-bar-row {
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .atraso-bar-label { font-size: .8rem; width: 180px; flex-shrink: 0; color: #7f1d1d; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .atraso-bar-track { flex: 1; background: #fecaca; border-radius: 6px; height: 10px; }
    .atraso-bar-fill  { height: 10px; border-radius: 6px; background: #dc2626; transition: width .4s; }
    .atraso-bar-val   { font-size: .78rem; font-weight: 600; color: #7f1d1d; width: 130px; text-align: right; flex-shrink: 0; }
    .atraso-table-wrap { overflow-x: auto; }
    .atraso-table {
      width: 100%;
      border-collapse: collapse;
      font-size: .82rem;
    }
    .atraso-table th {
      background: #fecaca;
      color: #7f1d1d;
      font-weight: 700;
      padding: 8px 10px;
      text-align: left;
      white-space: nowrap;
    }
    .atraso-table td {
      padding: 7px 10px;
      border-bottom: 1px solid #fee2e2;
      color: #1e293b;
      max-width: 200px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .atraso-table tr:last-child td { border-bottom: none; }
    .atraso-table tr:hover td { background: #fff1f2; }
    .atraso-table .td-valor { font-weight: 700; text-align: right; color: #b91c1c; }
    .atraso-dias-badge {
      display: inline-block;
      background: #dc2626;
      color: #fff;
      border-radius: 12px;
      padding: 2px 9px;
      font-size: .72rem;
      font-weight: 700;
    }
    .kpi-card.kpi-danger {
      border-left-color: #dc2626;
    }
    .kpi-card.kpi-danger .value { color: #b91c1c; }

    /* ── Modal Detalhe Dia ── */
    .modal-overlay {
      display: none;
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,.55);
      z-index: 1000;
      align-items: flex-start;
      justify-content: center;
      padding: 40px 16px 24px;
      overflow-y: auto;
    }
    .modal-overlay.open { display: flex; }
    .modal-box {
      background: var(--surface);
      border-radius: 14px;
      box-shadow: 0 24px 60px rgba(0,0,0,.25);
      width: 100%;
      max-width: 860px;
      padding: 32px 28px 28px;
      position: relative;
      animation: modalIn .22s ease;
    }
    @keyframes modalIn { from { opacity:0; transform:translateY(-18px); } to { opacity:1; transform:none; } }
    .modal-close {
      position: absolute;
      top: 14px; right: 16px;
      background: none; border: none;
      font-size: 1.6rem; color: var(--muted);
      cursor: pointer; line-height: 1;
      transition: color .15s;
    }
    .modal-close:hover { color: var(--text); }
    .modal-title {
      font-size: 1.25rem;
      font-weight: 700;
      color: var(--text);
      margin: 0 0 4px;
    }
    .modal-dates {
      font-size: .82rem;
      color: var(--muted);
      margin-bottom: 20px;
    }
    .modal-kpis {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
      gap: 12px;
      margin-bottom: 24px;
    }
    .modal-kpi {
      background: #f1f5ff;
      border-radius: 10px;
      padding: 14px 16px;
    }
    .modal-kpi .mk-label { font-size: .72rem; color: var(--muted); font-weight: 600; text-transform: uppercase; margin-bottom: 4px; }
    .modal-kpi .mk-value { font-size: 1.05rem; font-weight: 700; color: var(--text); }
    .modal-section-title {
      font-size: .85rem;
      font-weight: 700;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: .5px;
      margin: 20px 0 10px;
    }
    .cat-bars { display: flex; flex-direction: column; gap: 8px; margin-bottom: 8px; }
    .cat-bar-row { display: flex; align-items: center; gap: 10px; }
    .cat-bar-label { font-size: .8rem; width: 180px; flex-shrink: 0; color: var(--text); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .cat-bar-track { flex: 1; background: #e8ecf5; border-radius: 6px; height: 10px; }
    .cat-bar-fill  { height: 10px; border-radius: 6px; background: var(--primary); transition: width .4s; }
    .cat-bar-val   { font-size: .78rem; font-weight: 600; color: var(--text); width: 110px; text-align: right; flex-shrink: 0; }
    .modal-table-wrap { overflow-x: auto; margin-top: 4px; }
    .modal-table {
      width: 100%;
      border-collapse: collapse;
      font-size: .8rem;
    }
    .modal-table th {
      background: #f1f5ff;
      color: var(--muted);
      font-weight: 700;
      padding: 8px 10px;
      text-align: left;
      white-space: nowrap;
      cursor: pointer;
      user-select: none;
    }
    .modal-table th:hover { background: #e2e8f8; }
    .modal-table th.sort-asc::after  { content: ' \u25B2'; font-size: .65rem; }
    .modal-table th.sort-desc::after { content: ' \u25BC'; font-size: .65rem; }
    .modal-table td {
      padding: 7px 10px;
      border-bottom: 1px solid #f0f0f5;
      color: var(--text);
      max-width: 180px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .modal-table tr:last-child td { border-bottom: none; }
    .modal-table tr:hover td { background: #f8faff; }
    .modal-table .td-valor { font-weight: 600; text-align: right; }
    .status-badge {
      display: inline-block;
      padding: 2px 8px;
      border-radius: 20px;
      font-size: .72rem;
      font-weight: 600;
    }
    .status-badge.pago     { background: #d1fae5; color: #065f46; }
    .status-badge.pendente { background: #fef3c7; color: #92400e; }
    .status-badge.outro    { background: #e2e8f8; color: #475569; }
    .dia-kpi { cursor: pointer; }

    /* ── Responsivo: Tablet (≤ 900px) ── */
    @media (max-width: 900px) {
      header { padding: 20px 20px; }
      header h1 { font-size: 1.3rem; }
      .container { padding: 20px 16px; }
      .charts-grid { grid-template-columns: 1fr; }
      .chart-wrap-featured { height: 260px; }
      .kpi-grid { grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); }
      .dias-kpis { grid-template-columns: repeat(3, 1fr); }
      .search-box { width: 100%; }
      .table-toolbar { flex-direction: column; align-items: stretch; }
      .table-toolbar .btn { width: 100%; text-align: center; }
      #grid { height: 420px; }
    }

    /* ── Responsivo: Mobile (≤ 600px) ── */
    @media (max-width: 600px) {
      header { padding: 16px 14px; flex-direction: column; align-items: flex-start; }
      header h1 { font-size: 1.1rem; }
      header p  { font-size: .8rem; }
      .badge { align-self: flex-start; }
      .container { padding: 14px 10px; }

      .kpi-grid { grid-template-columns: 1fr 1fr; gap: 10px; }
      .kpi-card { padding: 14px 14px; }
      .kpi-card .value { font-size: 1.15rem; }

      .charts-grid { gap: 14px; }
      .chart-card { padding: 16px 14px; }
      .chart-card-featured { padding: 18px 14px 16px; }
      .chart-wrap { height: 200px; }
      .chart-wrap-featured { height: 210px; }
      .chart-header { flex-direction: column; align-items: flex-start; gap: 8px; }
      .chart-type-btns { justify-content: flex-start; }

      .dias-kpis { grid-template-columns: repeat(3, 1fr); gap: 7px; margin-top: 12px; }
      .dia-kpi { padding: 8px 6px; }
      .dia-kpi .dia-nome  { font-size: .65rem; }
      .dia-kpi .dia-valor { font-size: .78rem; }
      .dia-kpi .dia-qtd   { font-size: .62rem; }

      .col-selector-card { padding: 14px 12px; }
      .col-selector-actions { flex-direction: column; }
      .col-selector-actions .btn { width: 100%; text-align: center; }
      .col-check-item { padding: 5px 10px; }

      .table-card { padding: 14px 10px; }
      #grid { height: 380px; }

      .modal-overlay { padding: 16px 8px 16px; }
      .modal-box { padding: 22px 14px 18px; }
      .modal-title { font-size: 1rem; }
      .modal-kpis { grid-template-columns: 1fr 1fr; gap: 8px; }
      .modal-kpi { padding: 10px 12px; }
      .cat-bar-label { width: 100px; font-size: .72rem; }
      .cat-bar-val   { width: 80px; font-size: .72rem; }
    }

    /* ── Responsivo: Mobile muito pequeno (≤ 380px) ── */
    @media (max-width: 380px) {
      .kpi-grid { grid-template-columns: 1fr; }
      .dias-kpis { grid-template-columns: repeat(2, 1fr); }
      .modal-kpis { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>

<header>
  <div>
    <h1>📊 Relatório CAP – Semana 12</h1>
  </div>
  <span class="badge">${totalRegistros} registros</span>
</header>

<div class="container">

  <!-- KPIs -->
  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="label">Total Geral</div>
      <div class="value" id="kpi-total">${fmt(totalGeral)}</div>
      <div class="sub">Soma de todos os valores</div>
    </div>
    <div class="kpi-card">
      <div class="label">Total Registros</div>
      <div class="value" id="kpi-count">${totalRegistros}</div>
      <div class="sub">Linhas na planilha</div>
    </div>
    <div class="kpi-card">
      <div class="label">Maior Despesa Variável</div>
      <div class="value">${fmt(porCategoria['DESPESA VARIAVEL'] || 0)}</div>
      <div class="sub">Categoria mais representativa</div>
    </div>
    <div class="kpi-card kpi-danger">
      <div class="label">⚠️ Em Atraso</div>
      <div class="value">${fmt(totalAtrasado)}</div>
      <div class="sub">${atrasados.length} lançamento${atrasados.length !== 1 ? 's' : ''} vencido${atrasados.length !== 1 ? 's' : ''}</div>
    </div>
  </div>

  <!-- Seção Atrasados -->
  ${atrasados.length > 0 ? `
  <div class="atraso-section">
    <div class="atraso-header">
      <div class="atraso-title">🔴 Lançamentos em Atraso</div>
      <span style="font-size:.8rem;color:#b91c1c;font-weight:600">${HOJE_STR}</span>
    </div>
    <div class="atraso-kpis">
      <div class="atraso-kpi"><div class="ak-label">Total em Atraso</div><div class="ak-value">${fmt(totalAtrasado)}</div></div>
      <div class="atraso-kpi"><div class="ak-label">Lançamentos</div><div class="ak-value">${atrasados.length}</div></div>
      <div class="atraso-kpi"><div class="ak-label">Categorias</div><div class="ak-value">${Object.keys(atrasadosPorCat).length}</div></div>
    </div>
    <div style="font-size:.8rem;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#b91c1c;margin-bottom:8px">Por Categoria</div>
    <div class="atraso-cat-bars">
      ${Object.entries(atrasadosPorCat).sort((a,b)=>b[1]-a[1]).map(([cat,val]) =>
        `<div class="atraso-bar-row">
          <span class="atraso-bar-label" title="${cat}">${cat}</span>
          <div class="atraso-bar-track"><div class="atraso-bar-fill" style="width:${Math.round(val/totalAtrasado*100)}%"></div></div>
          <span class="atraso-bar-val">${fmt(val)}</span>
        </div>`
      ).join('')}
    </div>
    <div style="font-size:.8rem;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#b91c1c;margin:14px 0 8px">Detalhamento</div>
    <div class="atraso-table-wrap">
      <table class="atraso-table">
        <thead><tr>
          <th>Vencimento</th>
          <th>Dias em Atraso</th>
          <th>Descrição</th>
          <th>Categoria</th>
          <th>Sub-Categoria</th>
          <th>Responsável</th>
          <th>Origem</th>
          <th>Tipo Pgto</th>
          <th style="text-align:right">Valor (R$)</th>
        </tr></thead>
        <tbody>
          ${atrasados.sort((a,b)=>{ const da=strToDate(a['data_vencimento']),db=strToDate(b['data_vencimento']); return da-db; }).map(r => {
            const venc = strToDate(r['data_vencimento']);
            const dias = venc ? Math.floor((hoje - venc)/(1000*60*60*24)) : 0;
            return `<tr>
              <td>${r['data_vencimento']||'—'}</td>
              <td><span class="atraso-dias-badge">${dias}d</span></td>
              <td title="${r['descricao']||''}">${r['descricao']||'—'}</td>
              <td>${r['categoria']||'—'}</td>
              <td>${r['sub_categoria']||'—'}</td>
              <td>${r['responsavel']||'—'}</td>
              <td>${r['origem']||'—'}</td>
              <td>${r['tipo_pagamento']||'—'}</td>
              <td class="td-valor">${fmt(r['valor']||0)}</td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>
  </div>` : ''}

  <!-- Charts -->
  <div class="charts-grid" id="chartsGrid">

    <!-- DESTAQUE: Dia da Semana -->
    <div class="chart-card chart-card-featured" id="card-chartDiaSemana">
      <div class="chart-header"><h3>📅 Valor Total a Pagar por Dia da Semana</h3>${typeBtns('chartDiaSemana',['bar','line','radar','polarArea'],'bar')}</div>
      <div class="chart-wrap-featured"><canvas id="chartDiaSemana"></canvas></div>
      <div class="dias-kpis" id="diasKpis">
        <div class="dia-kpi" id="diakpi-0" onclick="openDayDetail(0)" title="Ver detalhes"><div class="dia-nome">Segunda</div><div class="dia-valor" id="dvk-0">—</div><div class="dia-qtd" id="dqk-0"></div></div>
        <div class="dia-kpi" id="diakpi-1" onclick="openDayDetail(1)" title="Ver detalhes"><div class="dia-nome">Terça</div><div class="dia-valor" id="dvk-1">—</div><div class="dia-qtd" id="dqk-1"></div></div>
        <div class="dia-kpi" id="diakpi-2" onclick="openDayDetail(2)" title="Ver detalhes"><div class="dia-nome">Quarta</div><div class="dia-valor" id="dvk-2">—</div><div class="dia-qtd" id="dqk-2"></div></div>
        <div class="dia-kpi" id="diakpi-3" onclick="openDayDetail(3)" title="Ver detalhes"><div class="dia-nome">Quinta</div><div class="dia-valor" id="dvk-3">—</div><div class="dia-qtd" id="dqk-3"></div></div>
        <div class="dia-kpi" id="diakpi-4" onclick="openDayDetail(4)" title="Ver detalhes"><div class="dia-nome">Sexta</div><div class="dia-valor" id="dvk-4">—</div><div class="dia-qtd" id="dqk-4"></div></div>
      </div>
      <p class="chart-desc">Total acumulado (R$) por dia útil de vencimento. O dia com maior valor fica em destaque.</p>
    </div>
    <div class="chart-card" id="card-chartCategoria">
      <div class="chart-header"><h3>Valor por Categoria</h3>${typeBtns('chartCategoria',['bar','line','radar','polarArea'],'bar')}</div>
      <div class="chart-wrap"><canvas id="chartCategoria"></canvas></div>
      <p class="chart-desc">Soma total (R$) agrupada por categoria de despesa. Permite identificar quais tipos de gasto têm maior peso no período.</p>
    </div>
    <div class="chart-card" id="card-chartResp">
      <div class="chart-header"><h3>Valor por Responsável</h3>${typeBtns('chartResp',['bar','line','doughnut'],'bar')}</div>
      <div class="chart-wrap"><canvas id="chartResp"></canvas></div>
      <p class="chart-desc">Valor total comprometido por cada centro de responsabilidade ou unidade de negócio.</p>
    </div>
    <div class="chart-card" id="card-chartSubCat">
      <div class="chart-header"><h3>Top 8 Sub-Categorias</h3>${typeBtns('chartSubCat',['bar','line','radar'],'bar')}</div>
      <div class="chart-wrap"><canvas id="chartSubCat"></canvas></div>
      <p class="chart-desc">As oito sub-categorias de maior montante financeiro, facilitando a priorização de redução de custos.</p>
    </div>
    <div class="chart-card" id="card-chartDescricao">
      <div class="chart-header"><h3>Top 10 Fornecedores / Descrições</h3>${typeBtns('chartDescricao',['bar','line','radar'],'bar')}</div>
      <div class="chart-wrap"><canvas id="chartDescricao"></canvas></div>
      <p class="chart-desc">Os dez fornecedores ou descrições de lançamento com maior valor acumulado no período.</p>
    </div>
  </div>

  <!-- Column Selector -->
  <div class="col-selector-card">
    <div class="section-title">Escolha as Colunas da Tabela</div>
    <div class="col-selector-actions">
      <button class="btn btn-primary"   onclick="selectAll(true)">Selecionar Todas</button>
      <button class="btn btn-secondary" onclick="selectAll(false)">Limpar Seleção</button>
      <button class="btn btn-success"   onclick="aplicarColunas()">▶ Aplicar</button>
    </div>
    <div class="col-checks" id="colChecks"></div>
  </div>

  <!-- Table -->
  <div class="table-card">
    <div class="section-title">Tabela de Dados</div>
    <div class="table-toolbar">
      <input class="search-box" id="searchBox" type="text" placeholder="🔍  Filtrar qualquer coluna..." oninput="onSearch(this.value)" />
      <div style="display:flex;gap:8px;">
        <button class="btn btn-secondary" onclick="exportCSV()">⬇ CSV</button>
        <button class="btn btn-secondary" onclick="exportExcel()">⬇ Excel (.xlsx)</button>
      </div>
    </div>
    <div id="grid" class="ag-theme-alpine"></div>
  </div>
</div>

<footer>Relatório gerado com ag-Grid + Chart.js · Fonte: CAP_Semana_12.xlsx</footer>

<!-- Modal detalhe dia -->
<div class="modal-overlay" id="dayModal">
  <div class="modal-box" id="dayModalBox">
    <button class="modal-close" onclick="closeDayModal()" aria-label="Fechar">&times;</button>
    <div id="dayModalContent"></div>
  </div>
</div>

<script>
// ── Dados ─────────────────────────────────────────────────────────
const DADOS   = ${dadosJSON};
const COLUNAS = ${colunasJSON};

// ── ag-Grid ───────────────────────────────────────────────────────
let gridApi;
let activeColumns = [...COLUNAS];

function buildColDefs(cols) {
  return cols.map(col => {
    const def = {
      field: col,
      headerName: col.replace(/_/g,' ').replace(/\\b\\w/g, c => c.toUpperCase()),
      sortable: true,
      filter: true,
      resizable: true,
      minWidth: 110,
    };
    if (col === 'valor') {
      def.valueFormatter = p => p.value != null && p.value !== ''
        ? Number(p.value).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })
        : '';
      def.cellStyle = { textAlign: 'right', fontWeight: '600' };
    }
    return def;
  });
}

window.addEventListener('DOMContentLoaded', () => {
  // Renderiza checkboxes
  const container = document.getElementById('colChecks');
  COLUNAS.forEach(col => {
    const item = document.createElement('label');
    item.className = 'col-check-item checked';
    item.innerHTML = \`<input type="checkbox" value="\${col}" checked />
      <span>\${col.replace(/_/g,' ')}</span>\`;
    item.querySelector('input').addEventListener('change', () => {
      item.classList.toggle('checked', item.querySelector('input').checked);
    });
    container.appendChild(item);
  });

  // Init ag-Grid
  const gridDiv = document.getElementById('grid');
  const gridOptions = {
    columnDefs: buildColDefs(COLUNAS),
    rowData: DADOS,
    defaultColDef: {
      sortable: true,
      filter: true,
      resizable: true,
      floatingFilter: true,
    },
    pagination: true,
    paginationPageSize: 20,
    paginationPageSizeSelector: [10, 20, 50, 100],
    animateRows: true,
    rowSelection: 'multiple',
    suppressRowClickSelection: true,
  };

  gridApi = agGrid.createGrid(gridDiv, gridOptions);
});

function selectAll(check) {
  document.querySelectorAll('#colChecks input[type=checkbox]').forEach(cb => {
    cb.checked = check;
    cb.closest('.col-check-item').classList.toggle('checked', check);
  });
}

function aplicarColunas() {
  const selecionadas = [...document.querySelectorAll('#colChecks input:checked')].map(cb => cb.value);
  if (selecionadas.length === 0) { alert('Selecione ao menos uma coluna.'); return; }
  activeColumns = selecionadas;
  gridApi.setGridOption('columnDefs', buildColDefs(selecionadas));
  syncCharts(selecionadas);
}

function onSearch(val) {
  gridApi.setGridOption('quickFilterText', val);
}

function exportCSV() {
  gridApi.exportDataAsCsv({ fileName: 'CAP_Semana12_export.csv' });
}

function exportExcel() {
  // Monta os dados respeitando colunas ativas e filtros atuais da grid
  const cols = activeColumns;
  const headers = cols.map(c => c.replace(/_/g,' ').replace(/\b\w/g, ch => ch.toUpperCase()));

  const rows = [];
  gridApi.forEachNodeAfterFilterAndSort(node => {
    if (!node.data) return;
    rows.push(cols.map(col => {
      const v = node.data[col];
      if (col === 'valor') return typeof v === 'number' ? v : (parseFloat(v) || '');
      return v === null || v === undefined ? '' : v;
    }));
  });

  const wsData = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Largura automática das colunas
  const colWidths = headers.map((h, i) => {
    const maxLen = Math.max(h.length, ...rows.map(r => String(r[i] ?? '').length));
    return { wch: Math.min(maxLen + 2, 40) };
  });
  ws['!cols'] = colWidths;

  // Formata coluna 'valor' como moeda (número)
  const valorIdx = cols.indexOf('valor');
  if (valorIdx >= 0) {
    rows.forEach((_, ri) => {
      const cellRef = XLSX.utils.encode_cell({ r: ri + 1, c: valorIdx });
      if (ws[cellRef] && typeof ws[cellRef].v === 'number') {
        ws[cellRef].z = 'R$ #,##0.00';
      }
    });
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'CAP Semana 12');
  XLSX.writeFile(wb, 'CAP_Semana12_export.xlsx');
}

// ── Charts ────────────────────────────────────────────────────────
const CORES = ${JSON.stringify(CORES)};

// Tipo atual por gráfico
const currentChartTypes = {
  chartCategoria:  'bar',
  chartResp:       'bar',
  chartSubCat:     'bar',
  chartDescricao:  'bar',
  chartDiaSemana:  'bar',
};

// Colunas necessárias por gráfico
const CHART_DEPS = {
  chartCategoria:   ['categoria',     'valor'],
  chartResp:        ['responsavel',   'valor'],
  chartSubCat:      ['sub_categoria', 'valor'],
  chartDescricao:   ['descricao',     'valor'],
  chartDiaSemana:   ['data_vencimento','valor'],
};

const chartInstances = {};

// ── Helpers de agregação ──────────────────────────────────────────
function aggregateBy(groupField, valueField, topN) {
  const map = {};
  DADOS.forEach(r => {
    const key = r[groupField] || 'N/A';
    map[key] = (map[key] || 0) + (parseFloat(r[valueField]) || 0);
  });
  let entries = Object.entries(map).sort((a, b) => b[1] - a[1]);
  if (topN) entries = entries.slice(0, topN);
  return entries;
}

function aggregateByCount(groupField, topN) {
  const map = {};
  DADOS.forEach(r => {
    const key = r[groupField] || 'N/A';
    map[key] = (map[key] || 0) + 1;
  });
  let entries = Object.entries(map).sort((a, b) => b[1] - a[1]);
  if (topN) entries = entries.slice(0, topN);
  return entries;
}

function aggregateByMonth(dateField, valueField) {
  const map = {};
  DADOS.forEach(r => {
    const ds = r[dateField];
    if (!ds || typeof ds !== 'string') return;
    const parts = ds.split('/');
    if (parts.length < 3) return;
    const key = parts[1] + '/' + parts[2]; // MM/AAAA
    map[key] = (map[key] || 0) + (parseFloat(r[valueField]) || 0);
  });
  return Object.entries(map).sort((a, b) => {
    const [ma, ya] = a[0].split('/');
    const [mb, yb] = b[0].split('/');
    return (parseInt(ya) * 12 + parseInt(ma)) - (parseInt(yb) * 12 + parseInt(mb));
  });
}

function aggregateByWeekday(dateField, valueField) {
  const DIAS = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta'];
  // js getDay(): 0=Dom,1=Seg,2=Ter,3=Qua,4=Qui,5=Sex,6=Sab
  const jsToIdx = { 1: 0, 2: 1, 3: 2, 4: 3, 5: 4 };
  const totals = [0, 0, 0, 0, 0];
  const counts = [0, 0, 0, 0, 0];
  const dates  = [[], [], [], [], []];
  DADOS.forEach(r => {
    const ds = r[dateField];
    if (!ds || typeof ds !== 'string') return;
    const parts = ds.split('/');
    if (parts.length < 3) return;
    const d = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const y = parseInt(parts[2], 10);
    const date = new Date(y, m, d);
    const dow = date.getDay();
    // Avança para o próximo dia útil se cair no fim de semana
    let effDate = date;
    if (dow === 6) effDate = new Date(y, m, d + 2);      // Sábado → próxima Segunda
    else if (dow === 0) effDate = new Date(y, m, d + 1); // Domingo → próxima Segunda
    const effDow = effDate.getDay();
    if (effDow >= 1 && effDow <= 5) {
      const idx = jsToIdx[effDow];
      totals[idx] += parseFloat(r[valueField]) || 0;
      counts[idx] += 1;
      const effD = effDate.getDate();
      const effM = effDate.getMonth() + 1;
      const effY = effDate.getFullYear();
      const dateStr = String(effD).padStart(2,'0') + '/' + String(effM).padStart(2,'0') + '/' + effY;
      if (!dates[idx].includes(dateStr)) dates[idx].push(dateStr);
    }
  });
  dates.forEach(arr => arr.sort((a, b) => {
    const [da,ma,ya] = a.split('/').map(Number);
    const [db,mb,yb] = b.split('/').map(Number);
    return new Date(ya,ma-1,da) - new Date(yb,mb-1,db);
  }));
  return { labels: DIAS, totals, counts, dates };
}

// ── Build chart ───────────────────────────────────────────────────
function buildChart(id, type, labels, data, label, valueType = 'currency') {
  if (chartInstances[id]) { chartInstances[id].destroy(); delete chartInstances[id]; }
  const ctx = document.getElementById(id);
  if (!ctx) return;

  const isLine      = type === 'line';
  const isRadar     = type === 'radar';
  const isRound     = type === 'doughnut' || type === 'pie';
  const isPolar     = type === 'polarArea';
  const hasAxes     = type === 'bar' || isLine;

  let dataset = { label, data };

  if (isRadar) {
    dataset.backgroundColor = CORES[0] + '35';
    dataset.borderColor     = CORES[0];
    dataset.borderWidth     = 2;
    dataset.pointBackgroundColor = CORES[0];
    dataset.fill = true;
  } else if (isLine) {
    dataset.backgroundColor = CORES[0] + '25';
    dataset.borderColor     = CORES[0];
    dataset.borderWidth     = 2.5;
    dataset.fill            = true;
    dataset.tension         = 0.4;
    dataset.pointBackgroundColor = CORES[0];
    dataset.pointRadius     = 4;
    dataset.pointHoverRadius = 6;
  } else {
    // bar, doughnut, pie, polarArea
    dataset.backgroundColor = CORES.slice(0, Math.max(labels.length, 1)).map(c => c + (type === 'bar' ? 'cc' : '99'));
    dataset.borderColor     = CORES.slice(0, Math.max(labels.length, 1));
    dataset.borderWidth     = type === 'bar' ? 0 : 1.5;
    dataset.hoverOffset     = 6;
  }

  const fmtVal = v => {
    if (typeof v !== 'number') return String(v);
    return valueType === 'currency'
      ? v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })
      : v.toLocaleString('pt-BR');
  };

  const scalesCb = valueType === 'currency'
    ? v => 'R$ ' + (v >= 1e6 ? (v/1e6).toFixed(1)+'M' : v >= 1e3 ? (v/1e3).toFixed(0)+'K' : v)
    : v => v.toLocaleString('pt-BR');

  chartInstances[id] = new Chart(ctx.getContext('2d'), {
    type,
    data: { labels, datasets: [dataset] },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 400 },
      plugins: {
        legend: {
          position: isRound || isPolar ? 'right' : 'top',
          labels: { font: { size: 11 }, boxWidth: 14 }
        },
        tooltip: {
          callbacks: {
            label: c => {
              const v = c.parsed?.y ?? c.parsed?.r ?? c.raw ?? 0;
              return ' ' + fmtVal(typeof v === 'number' ? v : parseFloat(v) || 0);
            }
          }
        }
      },
      scales: hasAxes ? {
        x: { ticks: { font: { size: 11 }, maxRotation: 38 } },
        y: { ticks: { font: { size: 11 }, callback: scalesCb } }
      } : undefined
    }
  });
}

// ── Render por ID ─────────────────────────────────────────────────
function renderChart(id) {
  const type = currentChartTypes[id];
  if (id === 'chartCategoria') {
    const e = aggregateBy('categoria', 'valor');
    buildChart(id, type, e.map(x=>x[0]), e.map(x=>x[1]), 'Valor (R$)', 'currency');
  } else if (id === 'chartResp') {
    const e = aggregateBy('responsavel', 'valor');
    buildChart(id, type, e.map(x=>x[0]), e.map(x=>x[1]), 'Valor (R$)', 'currency');
  } else if (id === 'chartSubCat') {
    const e = aggregateBy('sub_categoria', 'valor', 8);
    buildChart(id, type, e.map(x=>x[0]), e.map(x=>x[1]), 'Valor (R$)', 'currency');
  } else if (id === 'chartDescricao') {
    const e = aggregateBy('descricao', 'valor', 10);
    // Abrevia labels longos
    buildChart(id, type, e.map(x => x[0].length > 28 ? x[0].slice(0,26)+'…' : x[0]), e.map(x=>x[1]), 'Valor (R$)', 'currency');
  } else if (id === 'chartDiaSemana') {
    const { labels, totals, counts, dates } = aggregateByWeekday('data_vencimento', 'valor');
    const labelsRich = labels.map((d, i) => dates[i].length > 0 ? d + ' (' + dates[i].join(', ') + ')' : d);
    // Usa cores especiais no featured: rosa para o maior, azul claro para os demais
    const maxVal = Math.max(...totals);
    const bgColors = totals.map(v => v === maxVal ? '#2563ebcc' : '#94a3b870');
    const bdColors = totals.map(v => v === maxVal ? '#2563eb'   : '#94a3b8');
    const type = currentChartTypes[id];
    if (chartInstances[id]) { chartInstances[id].destroy(); delete chartInstances[id]; }
    const ctx = document.getElementById(id);
    if (ctx) {
      const isRadar = type === 'radar';
      const isPolar = type === 'polarArea';
      const hasAxes = type === 'bar' || type === 'line';
      let ds = { label: 'Total a Pagar (R$)', data: totals };
      if (isRadar) {
        ds.backgroundColor = '#f7258535'; ds.borderColor = '#f72585';
        ds.borderWidth = 2; ds.pointBackgroundColor = '#f72585'; ds.fill = true;
      } else if (type === 'line') {
        ds.backgroundColor = '#f7258520'; ds.borderColor = '#f72585';
        ds.borderWidth = 3; ds.fill = true; ds.tension = 0.4;
        ds.pointBackgroundColor = '#f72585'; ds.pointRadius = 5; ds.pointHoverRadius = 7;
      } else {
        ds.backgroundColor = bgColors; ds.borderColor = bdColors; ds.borderWidth = isPolar ? 1.5 : 0;
      }
      chartInstances[id] = new Chart(ctx.getContext('2d'), {
        type,
        data: { labels: labelsRich, datasets: [ds] },
        options: {
          responsive: true, maintainAspectRatio: false,
          animation: { duration: 400 },
          onClick: (evt, elements) => { if (elements.length > 0) openDayDetail(elements[0].index); },
          plugins: {
            legend: { display: false },
            tooltip: { callbacks: { label: c => {
              const v = c.parsed?.y ?? c.parsed?.r ?? c.raw ?? 0;
              return ' ' + Number(v).toLocaleString('pt-BR', { style:'currency', currency:'BRL' });
            }}}
          },
          scales: hasAxes ? {
            x: { ticks: { color:'rgba(255,255,255,.7)', font:{ size:11, weight:'600' } }, grid: { color:'rgba(255,255,255,.08)' } },
            y: { ticks: { color:'rgba(255,255,255,.6)', font:{size:11},
              callback: v => 'R$ '+(v>=1e6?(v/1e6).toFixed(1)+'M':v>=1e3?(v/1e3).toFixed(0)+'K':v)
            }, grid: { color:'rgba(255,255,255,.08)' } }
          } : undefined
        }
      });
      ctx.style.cursor = 'pointer';
    }
    // Atualiza mini-KPIs
    const maxIdx = totals.indexOf(maxVal);
    totals.forEach((val, i) => {
      const el = document.getElementById('dvk-'+i);
      const eq = document.getElementById('dqk-'+i);
      const card = document.getElementById('diakpi-'+i);
      if (el) el.textContent = val.toLocaleString('pt-BR', { style:'currency', currency:'BRL', maximumFractionDigits:0 });
      if (eq) eq.textContent = dates[i].length > 0 ? dates[i].join(', ') : '';
      if (card) card.classList.toggle('dia-maior', i === maxIdx);
    });
  }
}

// ── Troca tipo do gráfico ao clicar no botão ──────────────────────
function switchType(chartId, newType) {
  currentChartTypes[chartId] = newType;
  // Atualiza botão ativo
  const card = document.getElementById('card-' + chartId);
  card.querySelectorAll('.chart-type-btn').forEach(btn => {
    btn.classList.toggle('active', btn.textContent === '${Object.fromEntries(Object.entries(TYPE_LABELS))}'[newType]);
  });
  // Solução mais limpa: comparar pelo onclick
  card.querySelectorAll('.chart-type-btn').forEach(btn => {
    const onclickVal = btn.getAttribute('onclick');
    btn.classList.toggle('active', onclickVal && onclickVal.includes(\`'\${newType}'\`));
  });
  renderChart(chartId);
}

// ── Sync visibilidade ─────────────────────────────────────────────
function syncCharts(selecionadas) {
  const sel = new Set(selecionadas);
  let algumVisivel = false;

  Object.entries(CHART_DEPS).forEach(([chartId, deps]) => {
    const card = document.getElementById('card-' + chartId);
    const deveAparecer = deps.every(d => sel.has(d));

    if (deveAparecer) {
      const eraOculto = card.classList.contains('chart-hidden');
      card.classList.remove('chart-hidden');
      if (eraOculto) {
        card.classList.add('chart-fading-in');
        card.addEventListener('animationend', () => card.classList.remove('chart-fading-in'), { once: true });
      }
      renderChart(chartId);
      algumVisivel = true;
    } else {
      if (chartInstances[chartId]) { chartInstances[chartId].destroy(); delete chartInstances[chartId]; }
      card.classList.add('chart-hidden');
    }
  });

  document.getElementById('chartsGrid').style.display = algumVisivel ? '' : 'none';
}

// ── Detalhe por Dia ──────────────────────────────────────────────
let _weekdayData = null;
function _ensureWeekdayData() {
  if (!_weekdayData) _weekdayData = aggregateByWeekday('data_vencimento', 'valor');
  return _weekdayData;
}

function openDayDetail(dayIdx) {
  const DAY_NAMES = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira'];
  const { totals, counts, dates } = _ensureWeekdayData();
  const dayDates = dates[dayIdx];
  const total    = totals[dayIdx];
  const qtd      = counts[dayIdx];

  // Filtrar registros do dia
  const rows = DADOS.filter(r => {
    const dv = r['data_vencimento'];
    return dv && dayDates.includes(dv);
  });

  if (rows.length === 0) return;

  // Agrupamento por categoria
  const porCat = {};
  rows.forEach(r => {
    const c = r['categoria'] || 'Sem categoria';
    porCat[c] = (porCat[c] || 0) + (parseFloat(r['valor']) || 0);
  });
  const catEntries = Object.entries(porCat).sort((a,b) => b[1]-a[1]);
  const maxCat = catEntries[0]?.[1] || 1;

  // Maior lançamento
  const maiorVal = Math.max(...rows.map(r => parseFloat(r['valor']) || 0));
  const maiorRow = rows.find(r => (parseFloat(r['valor'])||0) === maiorVal);

  const fmtC = v => Number(v).toLocaleString('pt-BR', { style:'currency', currency:'BRL', maximumFractionDigits:2 });

  // Tabela — estado de ordenação
  let sortCol = 'valor', sortDir = -1;
  const TABLE_COLS = [
    { key:'data_vencimento', label:'Vencimento' },
    { key:'descricao',       label:'Descrição' },
    { key:'categoria',       label:'Categoria' },
    { key:'sub_categoria',   label:'Sub-Categoria' },
    { key:'valor',           label:'Valor (R$)' },
    { key:'status',          label:'Status' },
    { key:'responsavel',     label:'Responsável' },
    { key:'tipo_pagamento',  label:'Tipo Pgto' },
    { key:'banco',           label:'Banco' },
  ];

  function buildTable(col, dir) {
    const sorted = [...rows].sort((a,b) => {
      let va = a[col], vb = b[col];
      if (col === 'valor') { va = parseFloat(va)||0; vb = parseFloat(vb)||0; }
      if (va < vb) return -dir; if (va > vb) return dir; return 0;
    });
    const thHTML = TABLE_COLS.map(c => {
      let cls = '';
      if (c.key === col) cls = dir === 1 ? ' sort-asc' : ' sort-desc';
      return \`<th class="\${cls}" onclick="_sortDayTable('\${c.key}')">\${c.label}</th>\`;
    }).join('');
    const tdRows = sorted.map(r => {
      const v = parseFloat(r['valor']) || 0;
      const st = (r['status'] || '').toLowerCase();
      const badgeCls = st.includes('pago') ? 'pago' : st.includes('pend') ? 'pendente' : 'outro';
      const tds = TABLE_COLS.map(c => {
        if (c.key === 'valor') return \`<td class="td-valor">\${fmtC(v)}</td>\`;
        if (c.key === 'status') return \`<td><span class="status-badge \${badgeCls}">\${r[c.key] || '—'}</span></td>\`;
        return \`<td title="\${r[c.key] || ''}">\${r[c.key] || '—'}</td>\`;
      }).join('');
      return \`<tr>\${tds}</tr>\`;
    }).join('');
    return \`<table class="modal-table"><thead><tr>\${thHTML}</tr></thead><tbody>\${tdRows}</tbody></table>\`;
  }

  window._sortDayTable = (col) => {
    if (sortCol === col) sortDir *= -1; else { sortCol = col; sortDir = -1; }
    document.querySelector('.modal-table-wrap').innerHTML = buildTable(sortCol, sortDir);
  };

  const catBarsHTML = catEntries.map(([cat, val]) => \`
    <div class="cat-bar-row">
      <span class="cat-bar-label" title="\${cat}">\${cat}</span>
      <div class="cat-bar-track"><div class="cat-bar-fill" style="width:\${Math.round(val/maxCat*100)}%"></div></div>
      <span class="cat-bar-val">\${fmtC(val)}</span>
    </div>\`).join('');

  const content = \`
    <h2 class="modal-title">\${DAY_NAMES[dayIdx]}</h2>
    <div class="modal-dates">Datas: \${dayDates.join(' • ')} &nbsp;&middot;&nbsp; \${qtd} lançamento\${qtd !== 1 ? 's' : ''}</div>
    <div class="modal-kpis">
      <div class="modal-kpi"><div class="mk-label">Total do Dia</div><div class="mk-value">\${fmtC(total)}</div></div>
      <div class="modal-kpi"><div class="mk-label">Lançamentos</div><div class="mk-value">\${qtd}</div></div>
      <div class="modal-kpi"><div class="mk-label">Categorias</div><div class="mk-value">\${catEntries.length}</div></div>
      <div class="modal-kpi"><div class="mk-label">Maior Lançamento</div><div class="mk-value">\${fmtC(maiorVal)}</div></div>
    </div>
    <div class="modal-section-title">Distribuição por Categoria</div>
    <div class="cat-bars">\${catBarsHTML}</div>
    <div class="modal-section-title" style="margin-top:24px">Todos os Lançamentos</div>
    <div class="modal-table-wrap">\${buildTable(sortCol, sortDir)}</div>
  \`;

  document.getElementById('dayModalContent').innerHTML = content;
  document.getElementById('dayModal').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeDayModal() {
  document.getElementById('dayModal').classList.remove('open');
  document.body.style.overflow = '';
}

window.addEventListener('DOMContentLoaded', () => {
  syncCharts(COLUNAS);
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeDayModal(); });
  document.getElementById('dayModal').addEventListener('click', e => {
    if (e.target === document.getElementById('dayModal')) closeDayModal();
  });
});
<\/script>
</body>
</html>`;

// ─── Salvar ───────────────────────────────────────────────────────
const OUTPUT_DIR = path.join(__dirname, '..', 'output');
fs.mkdirSync(OUTPUT_DIR, { recursive: true });
const OUTPUT_PATH = path.join(OUTPUT_DIR, 'index.html');
fs.writeFileSync(OUTPUT_PATH, html, 'utf8');
console.log('✅  Relatório gerado:', OUTPUT_PATH);
