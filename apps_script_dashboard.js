// ============================================================
// DASHBOARD CRM — Apps Script
// Funciona como API JSON + trigger automático de atualização
// ============================================================

const SHEET_GID = 848653007;

// ── API: chamada pelo dashboard via fetch ──────────────────
function doGet(e) {
  const dados = serializarDados(processarDados(lerLeads()));
  const output = ContentService.createTextOutput(JSON.stringify(dados));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── TRIGGER: instale chamando esta função uma única vez ────
function instalarTrigger() {
  // Remove triggers antigos
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  // Novo trigger por hora (backup caso doGet falhe)
  ScriptApp.newTrigger('doGet')
    .timeBased().everyHours(1).create();
  Logger.log('✅ Trigger instalado!');
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
  let cabecalho = null;

  for (let i = 0; i < valores.length; i++) {
    const row = valores[i];
    const id  = String(row[0] || '').trim();

    if (id === 'ID_do_Atendimento' || id.includes('ID_do_Atend')) {
      cabecalho = true; continue;
    }
    if (!cabecalho || !id || id === 'Etapa' || id === 'Chave') continue;

    leads.push({
      id:          id,
      dataEntrada: row[1] instanceof Date ? row[1] : null,
      origem:      String(row[3] || '').trim(),
      nome:        String(row[4] || '').trim(),
      carro:       String(row[6] || '').trim(),
      etapa:       String(row[8] || '').trim().toUpperCase(),
      obs:         String(row[9] || '').trim(),
      proxContato: row[12] instanceof Date ? row[12] : null,
      tipoVeiculo: String(row[14] || '').trim(),
      dataVenda:   row[15] instanceof Date ? row[15] : null,
    });
  }

  // Remove duplicatas pelo ID
  const vistos = new Set();
  return leads.filter(l => {
    if (vistos.has(l.id)) return false;
    vistos.add(l.id); return true;
  });
}

// ── PROCESSAMENTO ──────────────────────────────────────────
function processarDados(leads) {
  const hoje = new Date();

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

  // Próximos contatos agrupados por data (ordena pela data)
  const proxMap = {};
  for (const l of ativos) {
    if (!l.proxContato) continue;
    const key = fmtData(l.proxContato);
    if (!proxMap[key]) proxMap[key] = [];
    proxMap[key].push({ id: l.id, nome: l.nome, carro: l.carro });
  }
  const proxContatos = Object.keys(proxMap)
    .sort((a, b) => parseData(a) - parseData(b))
    .map(k => ({ data: k, leads: proxMap[k] }));

  return {
    dataAtualizacao: fmtDataHora(hoje),
    totalAtivos:     ativos.length,
    totalVendas:     comprou.length,
    naoRespondem:    naoRespondem.length,
    naoRespondemNomes: naoRespondem.map(l => (l.nome || l.id).split(' ')[0]).join(', '),
    totalConcorrente: concorrente.length,
    contatosUrgentes,
    origemAtivos:    contarPor(ativos,   l => mapOrigem(l.origem)),
    origemVendas:    contarPor(comprou,  l => mapOrigem(l.origem)),
    modelosAtivos:   contarPor(ativos,   l => normModelo(l.carro)),
    modelosVendidos: contarPor(comprou,  l => normModelo(l.carro)),
    vendasMes:       contarPor(comprou,  l => l.dataVenda ? `${MES[l.dataVenda.getMonth()]}/${l.dataVenda.getFullYear()}` : null),
    ativos: ativos.map(l => ({
      id: l.id, nome: l.nome, carro: l.carro,
      origem: mapOrigem(l.origem), obs: l.obs,
      proxContato: l.proxContato ? fmtData(l.proxContato) : ''
    })),
    concorrente: concorrente.map(l => ({
      id: l.id, nome: l.nome, carro: l.carro, obs: l.obs
    })),
    proxContatos,
  };
}

// ── SERIALIZAÇÃO (remove objetos Date para JSON seguro) ────
function serializarDados(d) {
  return JSON.parse(JSON.stringify(d));
}

// ── HELPERS ────────────────────────────────────────────────
const MES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

function fmtData(d) {
  if (!d) return '';
  return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()}`;
}
function fmtDataHora(d) {
  return `${fmtData(d)} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function pad(n) { return String(n).padStart(2,'0'); }
function parseData(s) {
  const [dd,mm,yyyy] = s.split('/');
  return new Date(`${yyyy}-${mm}-${dd}`);
}

function contarPor(arr, fn) {
  const r = {};
  for (const item of arr) {
    const k = fn(item);
    if (!k) continue;
    r[k] = (r[k] || 0) + 1;
  }
  return r;
}

function mapOrigem(o) {
  if (!o) return 'Outros';
  const s = o.toLowerCase();
  if (s.includes('meta') || s.includes('instagram') || s.includes('facebook')) return 'META';
  if (s.includes('site'))       return 'Whats Site';
  if (s.includes('liga'))       return 'Ligação';
  if (s.includes('carteira'))   return 'Carteira';
  if (s.includes('balc') || s.includes('loja')) return 'Balcão';
  if (s.includes('indica'))     return 'Indicação';
  if (s.includes('olx'))        return 'OLX';
  return 'Outros';
}

function normModelo(m) {
  if (!m) return 'Sem modelo';
  const s = m.toUpperCase();
  if (s.includes('NIVUS'))                       return 'Nivus';
  if (s.includes('TERA'))                        return 'Tera';
  if (s.includes('T-CROSS') || s.includes('TCROSS')) return 'T-Cross';
  if (s.includes('TAOS'))                        return 'Taos';
  if (s.includes('VIRTUS'))                      return 'Virtus';
  if (s.includes('POLO'))                        return 'Polo';
  if (s.includes('GOLF'))                        return 'Golf GTI';
  if (s.includes('TIGUAN'))                      return 'Tiguan';
  if (s.includes('SAVEIRO'))                     return 'Saveiro';
  if (s.includes('AMAROK'))                      return 'Amarok';
  if (s.includes('COROLLA'))                     return 'Corolla';
  if (s.includes('KWID'))                        return 'Kwid';
  if (s.includes('HB20') || s.includes('HB-20')) return 'HB-20';
  return m.length > 18 ? m.substring(0,18)+'…' : m;
}
