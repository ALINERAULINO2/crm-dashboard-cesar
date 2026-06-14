// ============================================================
// DASHBOARD CRM — Apps Script
// Gera HTML completo e publica no GitHub Pages
// Como usar:
//   1. Cole este código no Apps Script da planilha
//   2. Em Configurações do script, adicione a propriedade GITHUB_TOKEN
//   3. Execute instalarTrigger() UMA VEZ para ativar automação
// ============================================================

const GITHUB_REPO = 'ALINERAULINO2/crm-dashboard-cesar';
const GITHUB_FILE = 'index.html';
const SHEET_GID   = 848653007;

// ── ENTRADA PRINCIPAL ──────────────────────────────────────
function atualizarDashboard() {
  const leads = lerLeads();
  const dados = processarDados(leads);
  const html  = gerarHTML(dados);
  publicarGitHub(html);
}

// ── TRIGGER: execute uma única vez ────────────────────────
function instalarTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('atualizarDashboard')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit().create();
  ScriptApp.newTrigger('atualizarDashboard')
    .timeBased().everyHours(1).create();
  Logger.log('✅ Triggers instalados! onEdit + a cada 1h');
}

// ── PUBLICAR NO GITHUB ─────────────────────────────────────
function publicarGitHub(htmlContent) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) { Logger.log('❌ GITHUB_TOKEN não configurado'); return; }

  const url     = `https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_FILE}`;
  const headers = { Authorization: 'token ' + token, Accept: 'application/vnd.github.v3+json' };

  // Busca SHA atual
  let sha = '';
  try {
    const r = UrlFetchApp.fetch(url, { headers, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) sha = JSON.parse(r.getContentText()).sha;
  } catch(e) {}

  const payload = JSON.stringify({
    message: 'dashboard: atualização automática ' + fmtDataHora(new Date()),
    content: Utilities.base64Encode(Utilities.newBlob(htmlContent).getBytes()),
    sha: sha || undefined
  });

  const res = UrlFetchApp.fetch(url, {
    method: 'PUT', headers: { ...headers, 'Content-Type': 'application/json' },
    payload, muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  Logger.log(code === 200 || code === 201 ? '✅ GitHub atualizado!' : '❌ Erro ' + code + ': ' + res.getContentText());
}

// ── LEITURA DA PLANILHA ────────────────────────────────────
function lerLeads() {
  const ss   = SpreadsheetApp.getActive();
  const abas = ss.getSheets();
  let aba    = abas[0];
  for (const a of abas) {
    if (a.getSheetId() === SHEET_GID) { aba = a; break; }
  }

  const valores = aba.getDataRange().getValues();
  const leads   = [];
  let cabecalho = false;

  for (const row of valores) {
    const id = String(row[0] || '').trim();
    if (id === 'ID_do_Atendimento' || id.includes('ID_do_Atend')) { cabecalho = true; continue; }
    if (!cabecalho || !id || id === 'Etapa' || id === 'Chave') continue;
    leads.push({
      id:          id,
      origem:      String(row[3] || '').trim(),
      nome:        String(row[4] || '').trim(),
      carro:       String(row[6] || '').trim(),
      etapa:       String(row[8] || '').trim().toUpperCase(),
      obs:         String(row[9] || '').trim(),
      proxContato: row[12] instanceof Date ? row[12] : null,
      dataVenda:   row[15] instanceof Date ? row[15] : null,
    });
  }

  const vistos = new Set();
  return leads.filter(l => { if (vistos.has(l.id)) return false; vistos.add(l.id); return true; });
}

// ── PROCESSAMENTO ──────────────────────────────────────────
function processarDados(leads) {
  const hoje        = new Date();
  const ativos      = leads.filter(l => l.etapa === 'ATIVO');
  const comprou     = leads.filter(l => l.etapa === 'COMPROU AQUI');
  const concorrente = leads.filter(l => l.etapa === 'COMPROU CONCORRENTE');

  const naoRespondem = ativos.filter(l => {
    const o = l.obs.toUpperCase();
    return o.includes('NAO RESPONDE') || o.includes('NÃO RESPONDE');
  });

  const em13Dias = new Date(hoje);
  em13Dias.setDate(em13Dias.getDate() + 13);
  const contatosUrgentes = ativos.filter(l => l.proxContato && l.proxContato <= em13Dias).length;

  // Próximos contatos agrupados por data
  const proxMap = {};
  for (const l of ativos) {
    if (!l.proxContato) continue;
    const key = fmtData(l.proxContato);
    if (!proxMap[key]) proxMap[key] = [];
    proxMap[key].push({ nome: l.nome || l.id, carro: l.carro });
  }
  const proxContatos = Object.keys(proxMap)
    .sort((a, b) => parseData(a) - parseData(b))
    .slice(0, 15)
    .map(k => ({ data: k, leads: proxMap[k] }));

  return {
    dataAtualizacao:  fmtDataHora(hoje),
    totalAtivos:      ativos.length,
    totalVendas:      comprou.length,
    naoRespondem:     naoRespondem.length,
    naoRespondemNomes: naoRespondem.map(l => (l.nome || l.id).split(' ')[0]).join(', '),
    totalConcorrente: concorrente.length,
    contatosUrgentes,
    origemAtivos:    contarPor(ativos,  l => mapOrigem(l.origem)),
    origemVendas:    contarPor(comprou, l => mapOrigem(l.origem)),
    modelosAtivos:   contarPor(ativos,  l => normModelo(l.carro)),
    modelosVendidos: contarPor(comprou, l => normModelo(l.carro)),
    proxContatos,
    ativos: ativos.map(l => ({
      nome: l.nome || l.id, carro: l.carro,
      origem: mapOrigem(l.origem), obs: l.obs,
      proxContato: l.proxContato ? fmtData(l.proxContato) : '',
      urgente: l.proxContato && l.proxContato <= em13Dias,
    })),
    concorrente: concorrente.map(l => ({ nome: l.nome || l.id, carro: l.carro, obs: l.obs })),
  };
}

// ── GERAÇÃO DO HTML ────────────────────────────────────────
function gerarHTML(d) {
  const kpiCards = [
    { icon:'🟢', val: d.totalAtivos,      label:'Leads Ativos',          cls:'blue'   },
    { icon:'🏆', val: d.totalVendas,      label:'Vendas Concluídas',     cls:'green'  },
    { icon:'🔕', val: d.naoRespondem,     label:'Não Respondem',
      sub: d.naoRespondemNomes,           cls: d.naoRespondem > 3 ? 'red' : 'yellow' },
    { icon:'📅', val: d.contatosUrgentes, label:'Contatos Urgentes (13d)',cls:'yellow' },
    { icon:'❌', val: d.totalConcorrente, label:'Comprou Concorrente',   cls:'red'    },
  ].map(k => `
    <div class="kpi-card ${k.cls}">
      <div class="icon">${k.icon}</div>
      <div class="val">${k.val}</div>
      <div class="label">${k.label}</div>
      ${k.sub ? `<div class="sub">${k.sub}</div>` : ''}
    </div>`).join('');

  const tabelaAtivos = d.ativos.map(l => `
    <tr>
      <td>${esc(l.nome)}</td>
      <td>${esc(l.carro)}</td>
      <td>${badgeOrigem(l.origem)}</td>
      <td class="${l.urgente ? 'prox-alert' : ''}">${l.proxContato || '—'}</td>
      <td class="obs">${esc(l.obs)}</td>
    </tr>`).join('');

  const tabelaConcorrente = d.concorrente.map(l => `
    <tr>
      <td>${esc(l.nome)}</td>
      <td>${esc(l.carro)}</td>
      <td class="obs">${esc(l.obs)}</td>
    </tr>`).join('');

  const timeline = d.proxContatos.length ? `
    <section>
      <div class="section-title">Próximos Contatos</div>
      <div class="timeline">
        ${d.proxContatos.map(g => `
        <div class="tl-item">
          <div class="tl-date">${g.data}</div>
          <div class="tl-leads">
            ${g.leads.map(l => `<div class="tl-lead">${esc(l.nome)} <span>${esc(l.carro)}</span></div>`).join('')}
          </div>
        </div>`).join('')}
      </div>
    </section>` : '';

  // Serializa dados para Chart.js
  const json = (obj) => JSON.stringify(obj);

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard CRM — Cesar</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"><\/script>
<style>
:root{--bg:#0f1117;--card:#1a1d27;--card2:#21253a;--accent:#4f8ef7;--green:#22c55e;
  --yellow:#f59e0b;--red:#ef4444;--text:#e2e8f0;--muted:#94a3b8;--border:rgba(255,255,255,.07)}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;background:var(--bg);color:var(--text);min-height:100vh}
header{background:var(--card);border-bottom:1px solid var(--border);padding:18px 28px;
  display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px}
.logo{display:flex;align-items:center;gap:10px}
.logo-icon{width:36px;height:36px;background:var(--accent);border-radius:8px;
  display:flex;align-items:center;justify-content:center;font-size:18px}
.logo-text h1{font-size:17px;font-weight:700}
.logo-text p{font-size:12px;color:var(--muted)}
.header-right{display:flex;align-items:center;gap:14px}
#last-update{font-size:12px;color:var(--muted)}
#btn-atualizar{background:var(--accent);color:#fff;border:none;padding:9px 20px;
  border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;
  display:flex;align-items:center;gap:7px;transition:opacity .2s}
#btn-atualizar:active{opacity:.75}
#btn-atualizar .spinner{display:none;width:14px;height:14px;
  border:2px solid rgba(255,255,255,.3);border-top-color:#fff;
  border-radius:50%;animation:spin .7s linear infinite}
#btn-atualizar.loading .spinner{display:block}
#btn-atualizar.loading .label-text{display:none}
@keyframes spin{to{transform:rotate(360deg)}}
main{padding:24px 28px;max-width:1280px;margin:0 auto}
section{margin-bottom:28px}
.section-title{font-size:13px;font-weight:700;text-transform:uppercase;
  letter-spacing:.08em;color:var(--muted);margin-bottom:14px}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:14px}
.kpi-card{background:var(--card);border:1px solid var(--border);border-radius:12px;
  padding:18px 16px;display:flex;flex-direction:column;gap:6px}
.kpi-card .icon{font-size:20px;margin-bottom:2px}
.kpi-card .val{font-size:32px;font-weight:800;line-height:1}
.kpi-card .label{font-size:12px;color:var(--muted)}
.kpi-card .sub{font-size:11px;color:var(--muted);margin-top:2px}
.kpi-card.blue .val{color:var(--accent)}.kpi-card.green .val{color:var(--green)}
.kpi-card.red .val{color:var(--red)}.kpi-card.yellow .val{color:var(--yellow)}
.charts-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:18px}
.chart-card{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:20px}
.chart-card h3{font-size:13px;font-weight:600;color:var(--muted);margin-bottom:14px}
.chart-wrap{position:relative;height:180px}
.timeline{display:flex;flex-direction:column;gap:10px}
.tl-item{background:var(--card);border:1px solid var(--border);border-radius:10px;
  padding:12px 16px;display:flex;align-items:flex-start;gap:14px}
.tl-date{background:var(--accent);color:#fff;border-radius:6px;padding:4px 10px;
  font-size:12px;font-weight:700;white-space:nowrap;min-width:90px;text-align:center}
.tl-leads{display:flex;flex-wrap:wrap;gap:6px}
.tl-lead{background:var(--card2);border-radius:6px;padding:3px 10px;font-size:12px}
.tl-lead span{color:var(--muted)}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.07em;
  color:var(--muted);padding:8px 10px;border-bottom:1px solid var(--border)}
td{padding:9px 10px;border-bottom:1px solid var(--border);vertical-align:top}
tr:last-child td{border-bottom:none}
tr:hover td{background:rgba(79,142,247,.05)}
.badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600}
.badge-meta{background:#3b3f7a;color:#8b9ff7}.badge-site{background:#1e3a3a;color:#34d399}
.badge-ligacao{background:#3b2e1e;color:#f59e0b}.badge-carteira{background:#2a1e3b;color:#a78bfa}
.badge-balcao{background:#1e2f3b;color:#67e8f9}.badge-indicacao{background:#1f3027;color:#86efac}
.badge-olx{background:#3b2020;color:#fca5a5}.badge-outros{background:#2a2a2a;color:#94a3b8}
.obs{color:var(--muted);font-size:12px;max-width:240px}
.prox-alert{color:var(--yellow);font-weight:600}
@media(max-width:640px){main{padding:16px}header{padding:14px 16px}
  .kpi-grid{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<header>
  <div class="logo">
    <div class="logo-icon">🚗</div>
    <div class="logo-text">
      <h1>Dashboard CRM</h1>
      <p>Cesar — Acompanhamento de leads</p>
    </div>
  </div>
  <div class="header-right">
    <span id="last-update">Atualizado: ${d.dataAtualizacao}</span>
    <button id="btn-atualizar" onclick="recarregar()">
      <div class="spinner"></div>
      <span class="label-text">↻ Atualizar</span>
    </button>
  </div>
</header>
<main>
  <section>
    <div class="section-title">Visão Geral</div>
    <div class="kpi-grid">${kpiCards}</div>
  </section>
  <section>
    <div class="section-title">Distribuição</div>
    <div class="charts-grid">
      <div class="chart-card"><h3>ORIGEM — LEADS ATIVOS</h3><div class="chart-wrap"><canvas id="cOA"></canvas></div></div>
      <div class="chart-card"><h3>MODELOS — LEADS ATIVOS</h3><div class="chart-wrap"><canvas id="cMA"></canvas></div></div>
      <div class="chart-card"><h3>ORIGEM DAS VENDAS</h3><div class="chart-wrap"><canvas id="cOV"></canvas></div></div>
      <div class="chart-card"><h3>MODELOS VENDIDOS</h3><div class="chart-wrap"><canvas id="cMV"></canvas></div></div>
    </div>
  </section>
  ${timeline}
  <section>
    <div class="section-title">Leads Ativos</div>
    <div class="chart-card" style="padding:0;overflow:hidden">
      <table><thead><tr><th>Nome</th><th>Modelo</th><th>Origem</th><th>Próx. Contato</th><th>Obs</th></tr></thead>
      <tbody>${tabelaAtivos}</tbody></table>
    </div>
  </section>
  ${d.concorrente.length ? `
  <section>
    <div class="section-title">Comprou no Concorrente</div>
    <div class="chart-card" style="padding:0;overflow:hidden">
      <table><thead><tr><th>Nome</th><th>Modelo</th><th>Obs</th></tr></thead>
      <tbody>${tabelaConcorrente}</tbody></table>
    </div>
  </section>` : ''}
</main>
<script>
const C=['#4f8ef7','#22c55e','#f59e0b','#ef4444','#a78bfa','#34d399','#fb923c','#67e8f9','#fbbf24','#f472b6'];
function mkChart(id,obj,type){
  const labels=Object.keys(obj),values=Object.values(obj);
  if(!labels.length)return;
  const ctx=document.getElementById(id).getContext('2d');
  if(type==='doughnut'){
    new Chart(ctx,{type:'doughnut',data:{labels,datasets:[{data:values,backgroundColor:C,borderWidth:0}]},
      options:{plugins:{legend:{position:'right',labels:{color:'#94a3b8',font:{size:11},boxWidth:12}}},maintainAspectRatio:false}});
  } else {
    new Chart(ctx,{type:'bar',data:{labels,datasets:[{data:values,backgroundColor:C,borderRadius:4,borderSkipped:false}]},
      options:{indexAxis:'y',plugins:{legend:{display:false}},
        scales:{x:{grid:{color:'rgba(255,255,255,.05)'},ticks:{color:'#94a3b8',font:{size:11}}},
                y:{grid:{display:false},ticks:{color:'#94a3b8',font:{size:11}}}},maintainAspectRatio:false}});
  }
}
mkChart('cOA',${json(d.origemAtivos)},'doughnut');
mkChart('cMA',${json(d.modelosAtivos)},'bar');
mkChart('cOV',${json(d.origemVendas)},'doughnut');
mkChart('cMV',${json(d.modelosVendidos)},'bar');

function recarregar(){
  const btn=document.getElementById('btn-atualizar');
  btn.classList.add('loading'); btn.disabled=true;
  setTimeout(()=>location.reload(),500);
}
<\/script>
</body>
</html>`;
}

// ── HELPERS ────────────────────────────────────────────────
const MES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

function fmtData(d) { return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()}`; }
function fmtDataHora(d) { return `${fmtData(d)} ${pad(d.getHours())}:${pad(d.getMinutes())}`; }
function pad(n) { return String(n).padStart(2,'0'); }
function parseData(s) { const [dd,mm,yyyy]=s.split('/'); return new Date(`${yyyy}-${mm}-${dd}`); }
function esc(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

function contarPor(arr, fn) {
  const r={};
  for(const item of arr){const k=fn(item);if(!k)continue;r[k]=(r[k]||0)+1;}
  return r;
}

function mapOrigem(o) {
  if(!o) return 'Outros';
  const s=o.toLowerCase();
  if(s.includes('meta')||s.includes('instagram')||s.includes('facebook')) return 'META';
  if(s.includes('site'))     return 'Whats Site';
  if(s.includes('liga'))     return 'Ligação';
  if(s.includes('carteira')) return 'Carteira';
  if(s.includes('balc')||s.includes('loja')) return 'Balcão';
  if(s.includes('indica'))   return 'Indicação';
  if(s.includes('olx'))      return 'OLX';
  return 'Outros';
}

function normModelo(m) {
  if(!m) return 'Sem modelo';
  const s=m.toUpperCase();
  if(s.includes('NIVUS'))   return 'Nivus';
  if(s.includes('TERA'))    return 'Tera';
  if(s.includes('T-CROSS')||s.includes('TCROSS')) return 'T-Cross';
  if(s.includes('TAOS'))    return 'Taos';
  if(s.includes('VIRTUS'))  return 'Virtus';
  if(s.includes('POLO'))    return 'Polo';
  if(s.includes('GOLF'))    return 'Golf GTI';
  if(s.includes('TIGUAN'))  return 'Tiguan';
  if(s.includes('SAVEIRO')) return 'Saveiro';
  if(s.includes('AMAROK'))  return 'Amarok';
  if(s.includes('COROLLA')) return 'Corolla';
  if(s.includes('KWID'))    return 'Kwid';
  if(s.includes('HB20')||s.includes('HB-20')) return 'HB-20';
  return m.length>18 ? m.substring(0,18)+'…' : m;
}

function badgeOrigem(o) {
  const map={'META':'badge-meta','Whats Site':'badge-site','Ligação':'badge-ligacao',
    'Carteira':'badge-carteira','Balcão':'badge-balcao','Indicação':'badge-indicacao',
    'OLX':'badge-olx','Outros':'badge-outros'};
  const cls=map[o]||'badge-outros';
  return `<span class="badge ${cls}">${o}</span>`;
}
