// ============================================================
// DASHBOARD CRM — Apps Script
// Cole no Google Apps Script do Sheets e configure o trigger.
// ============================================================

// ── CONFIGURAÇÕES ──────────────────────────────────────────
const GITHUB_REPO  = 'ALINERAULINO2/crm-dashboard-cesar';
const GITHUB_FILE  = 'index.html';
const SHEET_GID    = 848653007; // aba do CRM

// Token GitHub fica em: Projeto > Configurações > Propriedades do script
// Chave: GITHUB_TOKEN   Valor: (o token)

// ── ENTRADA PRINCIPAL ──────────────────────────────────────
function atualizarDashboard() {
  const dados = lerDados();
  const html  = gerarHTML(dados);
  publicarGitHub(html);
}

// Trigger de edição — chame isso para instalar automaticamente
function instalarTrigger() {
  ScriptApp.newTrigger('atualizarDashboard')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  // Trigger de tempo: atualiza a cada hora mesmo sem edição
  ScriptApp.newTrigger('atualizarDashboard')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('✅ Triggers instalados com sucesso!');
}

// ── LEITURA DOS DADOS ──────────────────────────────────────
function lerDados() {
  const ss    = SpreadsheetApp.getActive();
  const abas  = ss.getSheets();
  let aba     = ss.getSheetByName('CRM') || abas[0];

  // Tenta encontrar a aba pelo GID
  for (const a of abas) {
    if (a.getSheetId() === SHEET_GID) { aba = a; break; }
  }

  const valores = aba.getDataRange().getValues();

  const leads = [];
  const HEADER_MARCADOR = 'ID_do_Atendimento';

  let cabecalho = null;
  for (let i = 0; i < valores.length; i++) {
    const linha = valores[i];
    if (linha[0] === HEADER_MARCADOR || String(linha[0]).includes('ID_do_Atend')) {
      cabecalho = linha;
      continue;
    }
    if (!cabecalho) continue;
    if (!linha[0] || linha[0] === '') continue;
    // Pula tabelas auxiliares (Etapa | Origem_Lead)
    if (String(linha[0]).trim() === 'Etapa' || String(linha[0]).trim() === 'Chave') continue;

    leads.push({
      id:           String(linha[0] || '').trim(),
      dataEntrada:  linha[1] ? new Date(linha[1]) : null,
      dataFim:      linha[2] ? new Date(linha[2]) : null,
      origem:       String(linha[3] || '').trim(),
      nome:         String(linha[4] || '').trim(),
      contato:      String(linha[5] || '').trim(),
      carro:        String(linha[6] || '').trim(),
      troca:        String(linha[7] || '').trim(),
      etapa:        String(linha[8] || '').trim().toUpperCase(),
      obs:          String(linha[9] || '').trim(),
      vendedor:     String(linha[10] || '').trim(),
      proximoPasso: String(linha[11] || '').trim(),
      proxContato:  linha[12] ? new Date(linha[12]) : null,
      testDrive:    String(linha[13] || '').trim(),
      tipoVeiculo:  String(linha[14] || '').trim(),
      dataVenda:    linha[15] ? new Date(linha[15]) : null,
    });
  }

  // Remove duplicatas pelo ID
  const vistos = new Set();
  const leadsUnicos = leads.filter(l => {
    if (!l.id || vistos.has(l.id)) return false;
    vistos.add(l.id);
    return true;
  });

  return processarDados(leadsUnicos);
}

function processarDados(leads) {
  const hoje = new Date();

  const ativos       = leads.filter(l => l.etapa === 'ATIVO');
  const comprou      = leads.filter(l => l.etapa === 'COMPROU AQUI');
  const concorrente  = leads.filter(l => l.etapa === 'COMPROU CONCORRENTE');
  const encerrados   = leads.filter(l => l.etapa === 'ENCERRADO DEFINITIVO');

  // Não respondem (obs contém "NAO RESPONDE")
  const naoRespondem = ativos.filter(l => l.obs.toUpperCase().includes('NAO RESPONDE') || l.obs.toUpperCase().includes('NÃO RESPONDE'));

  // Próximos contatos agrupados por data
  const proxContatos = {};
  for (const l of ativos) {
    if (!l.proxContato) continue;
    const key = formatarData(l.proxContato);
    if (!proxContatos[key]) proxContatos[key] = [];
    proxContatos[key].push(l);
  }
  const proxContatosOrdenados = Object.keys(proxContatos)
    .sort((a, b) => new Date(a.split('/').reverse().join('-')) - new Date(b.split('/').reverse().join('-')))
    .map(k => ({ data: k, leads: proxContatos[k] }));

  // Contatos urgentes (próximos 13 dias)
  const em13Dias = new Date(hoje);
  em13Dias.setDate(em13Dias.getDate() + 13);
  const contatosUrgentes = ativos.filter(l => l.proxContato && l.proxContato <= em13Dias).length;

  // Origem dos leads ativos
  const origemAtivos = {};
  for (const l of ativos) {
    const o = mapearOrigem(l.origem);
    origemAtivos[o] = (origemAtivos[o] || 0) + 1;
  }

  // Origem das vendas
  const origemVendas = {};
  for (const l of comprou) {
    const o = mapearOrigem(l.origem);
    origemVendas[o] = (origemVendas[o] || 0) + 1;
  }

  // Modelos pipeline ativo
  const modelosAtivos = {};
  for (const l of ativos) {
    const m = (l.carro || '').trim();
    if (!m) { modelosAtivos['Sem modelo'] = (modelosAtivos['Sem modelo'] || 0) + 1; continue; }
    const mNorm = normalizarModelo(m);
    modelosAtivos[mNorm] = (modelosAtivos[mNorm] || 0) + 1;
  }

  // Modelos vendidos
  const modelosVendidos = {};
  for (const l of comprou) {
    const m = normalizarModelo(l.carro || '');
    if (!m) continue;
    modelosVendidos[m] = (modelosVendidos[m] || 0) + 1;
  }

  // Vendas por mês
  const vendasMes = {};
  for (const l of comprou) {
    if (!l.dataVenda) continue;
    const mes = `${mesNome(l.dataVenda.getMonth())}/${l.dataVenda.getFullYear()}`;
    vendasMes[mes] = (vendasMes[mes] || 0) + 1;
  }

  return {
    totalAtivos: ativos.length,
    totalVendas: comprou.length,
    naoRespondem: naoRespondem.length,
    naoRespondemNomes: naoRespondem.map(l => l.nome || l.id).join(', '),
    totalConcorrente: concorrente.length,
    contatosUrgentes,
    origemAtivos,
    origemVendas,
    modelosAtivos,
    modelosVendidos,
    vendasMes,
    ativos,
    comprou,
    concorrente,
    proxContatosOrdenados,
    dataAtualizacao: formatarDataHora(hoje),
  };
}

function mapearOrigem(origem) {
  if (!origem) return 'Outros';
  const o = origem.toLowerCase();
  if (o.includes('meta') || o.includes('instagram') || o.includes('facebook')) return 'META';
  if (o.includes('site') || o.includes('whats site')) return 'Whats Site';
  if (o.includes('ligação') || o.includes('ligacao') || o.includes('ligaçao')) return 'Ligação';
  if (o.includes('carteira')) return 'Carteira';
  if (o.includes('balcão') || o.includes('balcao') || o.includes('loja')) return 'Balcão';
  if (o.includes('indica')) return 'Indicação';
  if (o.includes('olx')) return 'OLX';
  if (o.includes('webmotors')) return 'Webmotors';
  return 'Outros';
}

function normalizarModelo(m) {
  const s = m.toUpperCase();
  if (s.includes('NIVUS')) return 'Nivus';
  if (s.includes('TERA')) return 'Tera';
  if (s.includes('T-CROSS') || s.includes('TCROSS')) return 'T-Cross';
  if (s.includes('TAOS')) return 'Taos';
  if (s.includes('VIRTUS')) return 'Virtus';
  if (s.includes('POLO')) return 'Polo';
  if (s.includes('GOLF')) return 'Golf GTI';
  if (s.includes('TIGUAN')) return 'Tiguan';
  if (s.includes('SAVEIRO')) return 'Saveiro';
  if (s.includes('AMAROK')) return 'Amarok';
  if (s.includes('COROLLA')) return 'Corolla';
  if (s.includes('KWID')) return 'Kwid';
  if (s.includes('HB20') || s.includes('HB-20')) return 'HB-20';
  if (!m.trim()) return 'Sem modelo';
  return m.length > 20 ? m.substring(0, 20) + '...' : m;
}

function formatarData(d) {
  if (!d) return '';
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
}

function formatarDataHora(d) {
  return `${formatarData(d)} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
}

function mesNome(m) {
  return ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'][m];
}

// ── GERAÇÃO DO HTML ────────────────────────────────────────
function gerarHTML(d) {

  // Helpers para charts
  const origemAtivosLabels = JSON.stringify(Object.keys(d.origemAtivos));
  const origemAtivosData   = JSON.stringify(Object.values(d.origemAtivos));
  const origemVendasLabels = JSON.stringify(Object.keys(d.origemVendas));
  const origemVendasData   = JSON.stringify(Object.values(d.origemVendas));

  const topModelosAtivos = Object.entries(d.modelosAtivos).sort((a,b)=>b[1]-a[1]).slice(0,7);
  const topModelosVendidos = Object.entries(d.modelosVendidos).sort((a,b)=>b[1]-a[1]).slice(0,6);

  const vendasMesLabels = JSON.stringify(Object.keys(d.vendasMes));
  const vendasMesData   = JSON.stringify(Object.values(d.vendasMes));

  // Barras de modelos ativos
  const maxModelo = topModelosAtivos[0] ? topModelosAtivos[0][1] : 1;
  const coresBar  = ['#3b82f6','#14b8a6','#8b5cf6','#f97316','#ec4899','#22c55e','#64748b'];
  const barrasHtml = topModelosAtivos.map(([nome, qtd], i) => `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:9px">
      <div style="width:110px;font-size:12px;color:#e2e8f0">${nome}</div>
      <div style="flex:1;background:#2d3148;border-radius:4px;height:18px;overflow:hidden">
        <div style="width:${Math.round(qtd/maxModelo*100)}%;height:100%;background:${coresBar[i]||'#64748b'};display:flex;align-items:center;padding-left:8px;font-size:11px;font-weight:700;color:#fff">${qtd}</div>
      </div>
    </div>`).join('');

  // Tabela pipeline ativo
  const linhasAtivos = d.ativos.map(l => {
    const obs = (l.obs||'').toUpperCase();
    let statusCor = '#94a3b8', statusTxt = l.obs ? l.obs.substring(0,22) : 'ATIVO';
    if (obs.includes('NAO RESPONDE') || obs.includes('NÃO RESPONDE')) { statusCor='#f87171'; statusTxt='NÃO RESPONDE'; }
    else if (obs.includes('PENSANDO')) { statusCor='#facc15'; statusTxt='PENSANDO'; }
    else if (obs.includes('CONVERSANDO')) { statusCor='#4ade80'; statusTxt='CONVERSANDO'; }
    else if (obs.includes('VIRA AQUI')) { statusCor='#a78bfa'; statusTxt='VIRA AQUI'; }

    const origemBadge = badgeOrigem(l.origem);
    const proxData = l.proxContato ? formatarData(l.proxContato) : '—';

    return `<tr>
      <td style="color:#64748b;font-size:11px">${l.id}</td>
      <td>${l.nome || '—'}</td>
      <td style="color:#94a3b8">${l.carro || '—'}</td>
      <td>${origemBadge}</td>
      <td style="color:${statusCor};font-size:11px;font-weight:600">${statusTxt}</td>
      <td style="color:#94a3b8">${proxData}</td>
    </tr>`;
  }).join('');

  // Próximos contatos
  const proxContatosHtml = d.proxContatosOrdenados.slice(0,4).map((grupo, idx) => {
    const cores = ['#facc15','#fb923c','#60a5fa','#94a3b8'];
    const bgs   = ['rgba(234,179,8,0.2)','rgba(249,115,22,0.2)','rgba(59,130,246,0.2)','rgba(148,163,184,0.1)'];
    const pills = grupo.leads.slice(0,8).map(l =>
      `<div style="background:#2d3148;border-radius:6px;padding:4px 10px;font-size:11px;display:inline-flex;align-items:center;gap:6px;margin:2px">
        <span style="color:#e2e8f0">${(l.nome||'—').split(' ')[0]}</span>
        <span style="color:#64748b;font-size:10px">${l.carro||'—'}</span>
      </div>`).join('');
    const extra = grupo.leads.length > 8 ? `<div style="font-size:10px;color:#64748b;margin-top:4px">+${grupo.leads.length-8} mais</div>` : '';
    return `<div>
      <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;padding:3px 8px;border-radius:4px;display:inline-block;background:${bgs[idx]};color:${cores[idx]};margin-bottom:6px">${grupo.data} — ${grupo.leads.length} lead${grupo.leads.length>1?'s':''}</div>
      <div>${pills}${extra}</div>
    </div>`;
  }).join('');

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashboard CRM — Vendas</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"><\/script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',sans-serif;background:#0f1117;color:#e2e8f0}
.header{background:linear-gradient(135deg,#1e40af,#1e3a8a);padding:22px 32px;display:flex;align-items:center;justify-content:space-between;border-bottom:3px solid #3b82f6}
.header h1{font-size:20px;font-weight:700}
.header .sub{font-size:12px;color:#93c5fd;margin-top:4px}
.header .date{font-size:12px;color:#bfdbfe;background:rgba(255,255,255,.1);padding:5px 12px;border-radius:20px}
.wrap{padding:24px 32px;max-width:1400px;margin:0 auto}
.kpi-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:24px}
.kpi{background:#1e2030;border-radius:10px;padding:18px;border-left:4px solid #3b82f6}
.kpi.g{border-left-color:#22c55e}.kpi.r{border-left-color:#ef4444}.kpi.o{border-left-color:#f97316}.kpi.y{border-left-color:#eab308}
.kpi .lbl{font-size:10px;color:#94a3b8;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px}
.kpi .val{font-size:34px;font-weight:800;line-height:1}
.kpi .val.b{color:#60a5fa}.kpi.g .val{color:#4ade80}.kpi.r .val{color:#f87171}.kpi.o .val{color:#fb923c}.kpi.y .val{color:#facc15}
.kpi .hint{font-size:10px;color:#64748b;margin-top:4px}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px}
.g3{display:grid;grid-template-columns:1fr 2fr;gap:18px;margin-bottom:18px}
.card{background:#1e2030;border-radius:10px;padding:18px}
.ct{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#94a3b8;margin-bottom:14px;padding-bottom:8px;border-bottom:1px solid #2d3148;display:flex;align-items:center;gap:7px}
.dot{width:7px;height:7px;border-radius:50%;display:inline-block}
.db{background:#3b82f6}.dg{background:#22c55e}.dr{background:#ef4444}.do{background:#f97316}.dy{background:#eab308}.dp{background:#a855f7}
.funnel{display:flex;gap:4px;height:90px;align-items:stretch}
.fs{flex:1;border-radius:7px;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:8px}
.fs .fv{font-size:26px;font-weight:800}.fs .fl{font-size:9px;color:rgba(255,255,255,.7);text-align:center;margin-top:1px}
.fa{background:linear-gradient(135deg,#1e40af,#2563eb)}.fn{background:linear-gradient(135deg,#7c3aed,#8b5cf6)}
.fp{background:linear-gradient(135deg,#d97706,#f59e0b)}.fg{background:linear-gradient(135deg,#15803d,#22c55e)}.fc{background:linear-gradient(135deg,#b91c1c,#ef4444)}
.insight{background:rgba(59,130,246,.08);border-left:3px solid #3b82f6;border-radius:4px;padding:9px 12px;margin-top:10px;font-size:11px;color:#93c5fd;line-height:1.5}
.alerta{background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.3);border-radius:7px;padding:9px 12px;margin-bottom:8px;font-size:11px;color:#fca5a5}
.alerta strong{color:#f87171}
table{width:100%;border-collapse:collapse;font-size:12px}
th{text-align:left;padding:7px 9px;color:#64748b;font-size:10px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid #2d3148}
td{padding:7px 9px;border-bottom:1px solid #1a1d2e}
tr:last-child td{border-bottom:none}
tr:hover td{background:rgba(255,255,255,.02)}
.bm{display:inline-block;padding:2px 7px;border-radius:8px;font-size:10px;font-weight:600;text-transform:uppercase}
.bm-meta{background:#1e3a8a;color:#93c5fd}.bm-site{background:#065f46;color:#6ee7b7}.bm-lig{background:#713f12;color:#fde68a}
.bm-cart{background:#4c1d95;color:#c4b5fd}.bm-balc{background:#1f2937;color:#9ca3af}.bm-ind{background:#134e4a;color:#5eead4}
.bm-olx{background:#2d1b69;color:#ddd6fe}.bm-o{background:#1c1c2e;color:#64748b}
.ch{position:relative;height:170px}
</style>
</head>
<body>
<div class="header">
  <div>
    <h1>Dashboard CRM — Vendas</h1>
    <div class="sub">Vendedor: Cesar &nbsp;|&nbsp; Atualização automática via Google Sheets</div>
  </div>
  <div class="date">Atualizado: ${d.dataAtualizacao}</div>
</div>
<div class="wrap">

  <!-- KPIs -->
  <div class="kpi-grid">
    <div class="kpi"><div class="lbl">Pipeline ativo</div><div class="val b">${d.totalAtivos}</div><div class="hint">leads em andamento</div></div>
    <div class="kpi g"><div class="lbl">Vendas realizadas</div><div class="val">${d.totalVendas}</div><div class="hint">total registrado</div></div>
    <div class="kpi r"><div class="lbl">Não respondem</div><div class="val">${d.naoRespondem}</div><div class="hint">${d.totalAtivos > 0 ? Math.round(d.naoRespondem/d.totalAtivos*100) : 0}% do pipeline ativo</div></div>
    <div class="kpi o"><div class="lbl">Perdidos (concorrente)</div><div class="val">${d.totalConcorrente}</div><div class="hint">leads perdidos</div></div>
    <div class="kpi y"><div class="lbl">Contatos próximos 13d</div><div class="val">${d.contatosUrgentes}</div><div class="hint">requerem ação</div></div>
  </div>

  <!-- FUNIL + ORIGEM VENDAS -->
  <div class="g2">
    <div class="card">
      <div class="ct"><span class="dot db"></span>Funil de Vendas</div>
      <div class="funnel">
        <div class="fs fa"><div class="fv">${d.totalAtivos}</div><div class="fl">ATIVO</div></div>
        <div class="fs fn"><div class="fv">0</div><div class="fl">NEGOCIAÇÃO</div></div>
        <div class="fs fp"><div class="fv">0</div><div class="fl">PROPOSTA</div></div>
        <div class="fs fg"><div class="fv">${d.totalVendas}</div><div class="fl">COMPROU AQUI</div></div>
        <div class="fs fc"><div class="fv">${d.totalConcorrente}</div><div class="fl">CONCORRENTE</div></div>
      </div>
      <div class="insight">⚠️ Funil sem etapas intermediárias registradas. Atualizar "Etapa" no CRM ajuda a prever fechamentos.</div>
    </div>
    <div class="card">
      <div class="ct"><span class="dot dg"></span>Origem das Vendas (${d.totalVendas} vendas)</div>
      <div class="ch"><canvas id="cOV"></canvas></div>
    </div>
  </div>

  <!-- ORIGEM LEADS + MODELOS -->
  <div class="g2">
    <div class="card">
      <div class="ct"><span class="dot dp"></span>Origem Leads Ativos (${d.totalAtivos})</div>
      <div class="ch"><canvas id="cOL"></canvas></div>
    </div>
    <div class="card">
      <div class="ct"><span class="dot" style="background:#14b8a6"></span>Modelos Mais Procurados (pipeline)</div>
      <div style="padding-top:4px">${barrasHtml}</div>
    </div>
  </div>

  <!-- PRÓXIMOS CONTATOS -->
  <div class="card" style="margin-bottom:18px">
    <div class="ct"><span class="dot dy"></span>Próximos Contatos Agendados</div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px">${proxContatosHtml}</div>
  </div>

  <!-- ALERTAS + TABELA -->
  <div class="g3">
    <div class="card">
      <div class="ct"><span class="dot dr"></span>Alertas Críticos</div>
      ${d.naoRespondem > 0 ? `<div class="alerta"><strong>${d.naoRespondem} leads NÃO RESPONDEM</strong><br><span style="font-size:10px">${d.naoRespondemNomes}</span></div>` : ''}
      <div class="alerta"><strong>Funil sem meio</strong><br>Nenhum lead em Negociação ou Proposta.</div>
      <div style="margin-top:14px;font-size:11px;color:#64748b;border-top:1px solid #2d3148;padding-top:10px">
        <div style="color:#94a3b8;font-weight:600;margin-bottom:6px">Perdidos para concorrente:</div>
        ${d.concorrente.map(l => `<div style="margin-bottom:3px">• <span style="color:#f87171">${l.nome||l.id}</span> — ${l.carro||l.obs||'—'}</div>`).join('')}
      </div>
    </div>
    <div class="card">
      <div class="ct"><span class="dot db"></span>Pipeline Ativo — ${d.totalAtivos} Leads</div>
      <div style="overflow:auto;max-height:400px">
        <table>
          <thead><tr><th>ID</th><th>Cliente</th><th>Modelo</th><th>Origem</th><th>Status</th><th>Próx. Contato</th></tr></thead>
          <tbody>${linhasAtivos}</tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- VENDAS POR MÊS + MODELOS VENDIDOS -->
  <div class="g2">
    <div class="card">
      <div class="ct"><span class="dot dg"></span>Vendas por Período</div>
      <div class="ch"><canvas id="cVM"></canvas></div>
    </div>
    <div class="card">
      <div class="ct"><span class="dot do"></span>Modelos Mais Vendidos</div>
      <div class="ch"><canvas id="cMV"></canvas></div>
    </div>
  </div>

</div>
<div style="text-align:center;padding:16px;color:#334155;font-size:11px;border-top:1px solid #1e2030">
  Dashboard CRM • Vendedor Cesar • Atualização automática via Google Sheets + GitHub Pages
</div>

<script>
Chart.defaults.color='#94a3b8';
Chart.defaults.borderColor='#2d3148';
Chart.defaults.font.family='Segoe UI';
const CORES=['#3b82f6','#22c55e','#8b5cf6','#f97316','#ec4899','#14b8a6','#eab308','#64748b'];

function doughnut(id,labels,data){
  new Chart(document.getElementById(id),{
    type:'doughnut',
    data:{labels,datasets:[{data,backgroundColor:CORES,borderWidth:2,borderColor:'#1e2030'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{boxWidth:11,padding:12,font:{size:11}}}}}
  });
}
function barV(id,labels,data){
  new Chart(document.getElementById(id),{
    type:'bar',
    data:{labels,datasets:[{data,backgroundColor:CORES,borderRadius:5,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{stepSize:1}},x:{grid:{display:false}}}}
  });
}
function barH(id,labels,data){
  new Chart(document.getElementById(id),{
    type:'bar',
    data:{labels,datasets:[{data,backgroundColor:CORES,borderRadius:5,borderSkipped:false}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,ticks:{stepSize:1}},y:{grid:{display:false}}}}
  });
}

doughnut('cOV',${origemVendasLabels},${origemVendasData});
doughnut('cOL',${origemAtivosLabels},${origemAtivosData});
barV('cVM',${vendasMesLabels},${vendasMesData});
barH('cMV',${JSON.stringify(topModelosVendidos.map(x=>x[0]))},${JSON.stringify(topModelosVendidos.map(x=>x[1]))});
<\/script>
</body>
</html>`;
}

// Badge de origem
function badgeOrigem(origem) {
  const o = mapearOrigem(origem);
  const map = {
    'META':'meta','Whats Site':'site','Ligação':'lig','Carteira':'cart',
    'Balcão':'balc','Indicação':'ind','OLX':'olx'
  };
  const cls = map[o] || 'o';
  return `<span class="bm bm-${cls}">${o}</span>`;
}

// ── PUBLICAR NO GITHUB ─────────────────────────────────────
function publicarGitHub(htmlContent) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    Logger.log('❌ GITHUB_TOKEN não configurado. Vá em Projeto > Configurações > Propriedades do script.');
    return;
  }

  const url = `https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_FILE}`;
  const opts = (method, payload) => ({
    method, muteHttpExceptions: true,
    headers: { 'Authorization': `token ${token}`, 'Content-Type': 'application/json' },
    payload: payload ? JSON.stringify(payload) : undefined
  });

  // Busca SHA atual do arquivo
  const getResp = UrlFetchApp.fetch(url, opts('GET'));
  const sha = getResp.getResponseCode() === 200
    ? JSON.parse(getResp.getContentText()).sha
    : null;

  const payload = {
    message: `auto: dashboard atualizado em ${new Date().toLocaleString('pt-BR')}`,
    content: Utilities.base64Encode(htmlContent, Utilities.Charset.UTF_8),
  };
  if (sha) payload.sha = sha;

  const putResp = UrlFetchApp.fetch(url, opts('PUT', payload));
  if (putResp.getResponseCode() === 200 || putResp.getResponseCode() === 201) {
    Logger.log('✅ Dashboard publicado com sucesso no GitHub Pages!');
  } else {
    Logger.log('❌ Erro ao publicar: ' + putResp.getContentText());
  }
}
