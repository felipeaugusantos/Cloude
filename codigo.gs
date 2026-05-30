// ─── Constants ────────────────────────────────────────────────────────────────
const CONFIG = {
  SHEET_GERAL:   "Geral",
  SHEET_KPI:     "KPI_Historico",
  SHEET_SESSOES: "KPI_Sessoes",
  SHEET_LOG:     "SYS_LOG",
  MAX_ROWS:      10000,
  PAGE_SIZE:     100,
  CACHE_SECONDS: 120,
  // ALLOWED_USERS: lista de e-mails separados por vírgula. Se vazio, não restringe acesso.
  // POST_SECRET: defina em Arquivo → Propriedades do projeto (PropertiesService).
};

// Novo schema de 15 colunas (SESSION_ID na posição 0)
const KPI_HEADER = [
  "SESSION_ID","TIMESTAMP","ALT_VERSAO","ABERTA","ANDAMENTO","CORRIGINDO","CORRIGIDO",
  "CONFERIDO","TOTAL","KPI_ABERTA_ANDAMENTO","KPI_CORRIGINDO_CORRIGIDO",
  "KPI_CONFERIDO","SEMAFORO_ABERTA_ANDAMENTO","SEMAFORO_CORRIGINDO_CORRIGIDO",
  "SEMAFORO_CONFERIDO"
];

const KPI_NUMERIC_FIELDS = [
  "ABERTA","ANDAMENTO","CORRIGINDO","CORRIGIDO","CONFERIDO","TOTAL",
  "KPI_ABERTA_ANDAMENTO","KPI_CORRIGINDO_CORRIGIDO","KPI_CONFERIDO"
];

const LOG_HEADER     = ["TIMESTAMP","NIVEL","ORIGEM","FUNCAO","MENSAGEM","DETALHE","USUARIO"];
const SESSOES_HEADER = ["SESSION_ID","TIMESTAMP","TOTAL_REGISTROS","ORIGEM","STATUS","ERRO","HASH_PAYLOAD"];

// Mapeamento explícito de colunas da aba Geral
const GERAL_COLUMNS = {
  data: {
    required: true, fallback: 0,
    names: ["data_liberacao","data liberacao","data liberação","dt_liberacao",
            "dt liberacao","dt liberação","data_origem","data origem",
            "data_lib","data_","dt_","data","date","dt"]
  },
  versao: {
    required: true, fallback: 1,
    names: ["alt_versao","versao_alt","alt versao","versao alt","versao","version","ver_"]
  },
  requisito: {
    required: false, fallback: 2,
    names: ["nr_requisito","num_requisito","nr requisito","num requisito","requisito","requirement","req_"]
  },
  cliente: {
    required: true, fallback: 3,
    names: ["nm_cliente","nome_cliente","nm cliente","nome cliente","cliente","client","customer"]
  },
  status: {
    required: false, fallback: 5,
    names: ["ds_status","ds status","status","situacao","situação","state"]
  },
  caminho: {
    required: false, fallback: 7,
    names: ["caminho_rede","caminho_completo","caminho rede","caminho completo",
            "caminho","path","diretorio","diretório","dir_"]
  },
  revisao: {
    required: false, fallback: 8,
    names: ["nr_revisao","num_revisao","nr revisao","num revisao","revisao","revisão","revision","rev_"]
  }
};

// ─── Spreadsheet accessor ─────────────────────────────────────────────────────
// Works for both container-bound and standalone scripts.
// For standalone scripts, set SPREADSHEET_ID in Project Properties.
function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) throw new Error(
    "Planilha não encontrada. Defina SPREADSHEET_ID nas Propriedades do Script."
  );
  return SpreadsheetApp.openById(id);
}

function getScriptProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || "";
}

function getActiveUserEmail_() {
  try { return (Session.getActiveUser().getEmail() || "").toLowerCase().trim(); }
  catch(e) { return ""; }
}

function getAllowedUsers_() {
  return getScriptProp_("ALLOWED_USERS")
    .split(",")
    .map(email => email.toLowerCase().trim())
    .filter(Boolean);
}

function assertAuthorized_(action) {
  const allowed = getAllowedUsers_();
  if (!allowed.length) return;
  const email = getActiveUserEmail_();
  if (email && allowed.includes(email)) return;
  logEvent_("WARN", "AUTH", action || "assertAuthorized_", "Acesso negado", email || "sem email");
  throw new Error("Usuário não autorizado");
}

function getIgnoredDecisionClients_() {
  try {
    const raw = PropertiesService.getUserProperties().getProperty("IGNORED_DECISION_CLIENTS") || "[]";
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.map(String).filter(Boolean).sort() : [];
  } catch(e) {
    return [];
  }
}

function saveIgnoredDecisionClients(clientsJson) {
  try {
    assertAuthorized_("saveIgnoredDecisionClients");
    const parsed = JSON.parse(clientsJson || "[]");
    if (!Array.isArray(parsed)) throw new Error("Lista inválida");
    const clients = [...new Set(parsed.map(v => String(v || "").trim()).filter(Boolean))].sort();
    PropertiesService.getUserProperties().setProperty("IGNORED_DECISION_CLIENTS", JSON.stringify(clients));
    return { ok: true, ignoredClients: clients };
  } catch(e) {
    logEvent_("ERROR", "SISTEMA", "saveIgnoredDecisionClients", "Erro ao salvar clientes ignorados", e.message);
    return { ok: false, error: e.message };
  }
}

// ─── Numeric validation ───────────────────────────────────────────────────────
function isValidNumberValue(v) {
  if (v === null || v === undefined) return false;
  if (typeof v === "string" && v.trim() === "") return false;
  return Number.isFinite(Number(v));
}

function parseSheetDate_(value, tz) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return {
      dataOrigem: Utilities.formatDate(value, tz, "yyyy-MM-dd"),
      dataExibicao: Utilities.formatDate(value, tz, "dd/MM/yyyy"),
    };
  }
  const raw = String(value || "").trim();
  if (!raw) return { dataOrigem: "", dataExibicao: "" };

  let match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    return {
      dataOrigem: match[1] + "-" + match[2] + "-" + match[3],
      dataExibicao: match[3] + "/" + match[2] + "/" + match[1],
    };
  }

  match = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (match) {
    const d = ("0" + match[1]).slice(-2);
    const m = ("0" + match[2]).slice(-2);
    const y = match[3];
    return { dataOrigem: y + "-" + m + "-" + d, dataExibicao: d + "/" + m + "/" + y };
  }

  return { dataOrigem: "", dataExibicao: "" };
}

// ─── Logging ──────────────────────────────────────────────────────────────────
function logEvent_(nivel, origem, funcao, mensagem, detalhe) {
  try {
    const ss  = getSpreadsheet_();
    const tz  = ss.getSpreadsheetTimeZone();
    const sht = ensureSheet_(ss, CONFIG.SHEET_LOG, LOG_HEADER);
    let usuario = "";
    try { usuario = Session.getActiveUser().getEmail() || ""; } catch(e) {}
    const ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    sht.appendRow([ts, nivel || "INFO", origem || "", funcao || "",
                   mensagem || "", String(detalhe || "").substring(0, 500), usuario]);
  } catch(e) { /* never let log failure break main flow */ }
}

// ─── Sheet helpers ────────────────────────────────────────────────────────────
function ensureSheet_(ss, name, header) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    try { sheet.setFrozenRows(1); } catch(e) {}
  }
  return sheet;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── Entry points ─────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile("Index").evaluate()
    .setTitle("Sistema de Gestão de Cópias v2026")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    if (!e || !e.postData) {
      logEvent_("WARN", "POST", "doPost", "POST recebido sem dados", "");
      return jsonResp({ ok: false, error: "Sem dados" });
    }
    const contentLen = (e.postData.contents || "").length;
    if (contentLen > 524288) {
      logEvent_("WARN", "POST", "doPost", "Payload rejeitado: excede 512 KB", "tamanho=" + contentLen);
      return jsonResp({ ok: false, error: "Payload excede 512 KB" });
    }
    let payload;
    try { payload = JSON.parse(e.postData.contents); }
    catch(parseErr) {
      logEvent_("WARN", "POST", "doPost", "JSON malformado", parseErr.message);
      return jsonResp({ ok: false, error: "Payload inválido" });
    }
    const secret = getScriptProp_("POST_SECRET");
    if (!secret || payload.secret !== secret) {
      logEvent_("WARN", "POST", "doPost", "Secret inválido — acesso negado", "");
      return jsonResp({ ok: false, error: "Não autorizado" });
    }
    if (!Array.isArray(payload.data) || payload.data.length > 500) {
      logEvent_("WARN", "POST", "doPost", "payload.data inválido",
                "length=" + (Array.isArray(payload.data) ? payload.data.length : "not-array"));
      return jsonResp({ ok: false, error: "payload.data deve ser array com até 500 itens" });
    }
    return jsonResp(saveKPIDataManual(JSON.stringify(payload.data), "POST"));
  } catch(err) {
    logEvent_("ERROR", "POST", "doPost", "Erro inesperado", err.message);
    return jsonResp({ ok: false, error: "Erro interno" });
  }
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Column mapping ───────────────────────────────────────────────────────────
function normalizeColName(h) {
  return String(h || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[\s_\-]+/g, " ")
    .trim();
}

function findColByNames_(normHeaders, names, fallback) {
  const normNames = names.map(normalizeColName);
  for (const name of normNames) {
    const idx = normHeaders.findIndex(h => h === name);
    if (idx >= 0) return { idx, usedFallback: false };
  }
  for (const name of normNames) {
    const idx = normHeaders.findIndex(h => h.startsWith(name));
    if (idx >= 0) return { idx, usedFallback: false };
  }
  for (const name of normNames) {
    const idx = normHeaders.findIndex(h => h.includes(name));
    if (idx >= 0) return { idx, usedFallback: false };
  }
  return { idx: fallback, usedFallback: true };
}

// ─── KPI row parser (handles legacy 14-col and new 15-col formats) ────────────
function parseKpiRow_(row, hasSessionId, tz) {
  if (hasSessionId) {
    // col0=SESSION_ID, col1=TIMESTAMP, col2=ALT_VERSAO, col3=ABERTA ...
    const ts = row[1] instanceof Date
      ? Utilities.formatDate(row[1], tz, "yyyy-MM-dd HH:mm:ss")
      : String(row[1] || "").trim();
    return {
      SESSION_ID:                    String(row[0]  || ""),
      TIMESTAMP:                     ts,
      ALT_VERSAO:                    String(row[2]  || ""),
      ABERTA:                        Number(row[3]  || 0),
      ANDAMENTO:                     Number(row[4]  || 0),
      CORRIGINDO:                    Number(row[5]  || 0),
      CORRIGIDO:                     Number(row[6]  || 0),
      CONFERIDO:                     Number(row[7]  || 0),
      TOTAL:                         Number(row[8]  || 0),
      KPI_ABERTA_ANDAMENTO:          Number(row[9]  || 0),
      KPI_CORRIGINDO_CORRIGIDO:      Number(row[10] || 0),
      KPI_CONFERIDO:                 Number(row[11] || 0),
      SEMAFORO_ABERTA_ANDAMENTO:     String(row[12] || ""),
      SEMAFORO_CORRIGINDO_CORRIGIDO: String(row[13] || ""),
      SEMAFORO_CONFERIDO:            String(row[14] || ""),
    };
  }
  // Legacy: col0=TIMESTAMP, col1=ALT_VERSAO ... (14 or fewer cols)
  const numCols = row.length;
  const isNew14 = numCols >= 14;
  const ts = row[0] instanceof Date
    ? Utilities.formatDate(row[0], tz, "yyyy-MM-dd HH:mm:ss")
    : String(row[0] || "").trim();
  return {
    SESSION_ID:                    "",
    TIMESTAMP:                     ts,
    ALT_VERSAO:                    String(row[1]  || ""),
    ABERTA:                        Number(row[2]  || 0),
    ANDAMENTO:                     Number(row[3]  || 0),
    CORRIGINDO:                    Number(row[4]  || 0),
    CORRIGIDO:                     Number(row[5]  || 0),
    CONFERIDO:                     Number(row[6]  || 0),
    TOTAL:                         Number(row[7]  || 0),
    KPI_ABERTA_ANDAMENTO:          Number(row[8]  || 0),
    KPI_CORRIGINDO_CORRIGIDO:      isNew14 ? Number(row[9]  || 0) : null,
    KPI_CONFERIDO:                 isNew14 ? Number(row[10] || 0) : Number(row[9]  || 0),
    SEMAFORO_ABERTA_ANDAMENTO:     isNew14 ? String(row[11] || "") : String(row[10] || ""),
    SEMAFORO_CORRIGINDO_CORRIGIDO: isNew14 ? String(row[12] || "") : null,
    SEMAFORO_CONFERIDO:            isNew14 ? String(row[13] || "") : String(row[11] || ""),
  };
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function buildAllItems(ss) {
  const tz    = ss.getSpreadsheetTimeZone();
  const sheet = ss.getSheetByName(CONFIG.SHEET_GERAL);
  if (!sheet) {
    logEvent_("WARN", "SISTEMA", "buildAllItems", "Aba '" + CONFIG.SHEET_GERAL + "' não encontrada", "");
    return [];
  }
  const last = Math.min(sheet.getLastRow(), CONFIG.MAX_ROWS + 1);
  if (last < 2) return [];
  const numCols    = sheet.getLastColumn();
  const rawHeaders = sheet.getRange(1, 1, 1, numCols).getValues()[0];
  const normHeaders = rawHeaders.map(normalizeColName);

  const colMap = {};
  for (const [field, cfg] of Object.entries(GERAL_COLUMNS)) {
    const result = findColByNames_(normHeaders, cfg.names, cfg.fallback);
    if (result.usedFallback && cfg.required) {
      logEvent_("WARN", "SISTEMA", "buildAllItems",
        "Coluna obrigatória '" + field + "' não encontrada — usando posição " + cfg.fallback,
        "cabeçalhos: " + rawHeaders.slice(0, 10).join(" | "));
    }
    colMap[field] = result.idx;
  }

  const rows = sheet.getRange(2, 1, last - 1, numCols).getValues();
  return rows.map(row => {
    const parsedDate = parseSheetDate_(row[colMap.data], tz);
    return {
      dataOrigem:   parsedDate.dataOrigem,
      dataExibicao: parsedDate.dataExibicao,
      versao:    String(row[colMap.versao]    || ""),
      requisito: String(row[colMap.requisito] || ""),
      cliente:   String(row[colMap.cliente]   || ""),
      status:    String(row[colMap.status]    || ""),
      caminho:   String(row[colMap.caminho]   || ""),
      revisao:   String(row[colMap.revisao]   || ""),
    };
  }).filter(r => r.dataOrigem !== "");
}

function applyFilters(items, filters) {
  return items.filter(r => {
    if (filters.cliente    && r.cliente    !== filters.cliente)                                      return false;
    if (filters.dataInicio && r.dataOrigem <  filters.dataInicio)                                   return false;
    if (filters.dataFim    && r.dataOrigem >  filters.dataFim)                                      return false;
    if (filters.versao     && !r.versao.toLowerCase().includes(filters.versao.toLowerCase()))       return false;
    if (filters.revisao    && !r.revisao.toLowerCase().includes(filters.revisao.toLowerCase()))     return false;
    if (filters.requisito  && !r.requisito.toLowerCase().includes(filters.requisito.toLowerCase())) return false;
    return true;
  });
}

function applySort(items, campo, dir) {
  const allowed = ["dataOrigem","versao","requisito","cliente","status","revisao"];
  if (!allowed.includes(campo)) return items;
  const asc = dir !== "desc";
  return items.slice().sort((a, b) => {
    const av = a[campo] || "", bv = b[campo] || "";
    return asc ? av.localeCompare(bv) : bv.localeCompare(av);
  });
}

// ─── Dashboard helpers ────────────────────────────────────────────────────────
function buildCards_(items, today, tz) {
  const total      = items.length;
  const hoje       = items.filter(r => r.dataOrigem === today).length;
  const semCaminho = items.filter(r => !r.caminho || !r.caminho.trim()).length;

  function mediaUltimosNDias(n) {
    const days = {};
    for (let i = 0; i < n; i++) {
      const d = new Date(); d.setDate(d.getDate() - i);
      days[Utilities.formatDate(d, tz, "yyyy-MM-dd")] = 0;
    }
    items.forEach(r => { if (days[r.dataOrigem] !== undefined) days[r.dataOrigem]++; });
    const vals = Object.values(days);
    return vals.length ? (vals.reduce((a,b)=>a+b,0) / vals.length).toFixed(1) : "0.0";
  }

  return {
    total, hoje, semCaminho,
    media7d:  mediaUltimosNDias(7),
    media30d: mediaUltimosNDias(30),
    clientes: new Set(items.map(r => r.cliente).filter(Boolean)).size,
  };
}

function buildCharts_(items, today, tz) {
  const hojeItems = items.filter(r => r.dataOrigem === today);

  const cliHojeMap = {};
  hojeItems.forEach(r => { cliHojeMap[r.cliente] = (cliHojeMap[r.cliente]||0)+1; });
  const cliHoje = Object.entries(cliHojeMap)
    .sort((a,b)=>b[1]-a[1]).slice(0,10)
    .reduce((o,[k,v])=>{ o.labels.push(k); o.data.push(v); return o; },{labels:[],data:[]});

  const versoesHojeMap = {};
  hojeItems.forEach(r => { versoesHojeMap[r.versao] = (versoesHojeMap[r.versao]||0)+1; });

  const trend30Raw = {};
  for (let i = 29; i >= 0; i--) {
    const d = new Date(); d.setDate(d.getDate() - i);
    trend30Raw[Utilities.formatDate(d, tz, "yyyy-MM-dd")] = 0;
  }
  items.forEach(r => { if (trend30Raw[r.dataOrigem] !== undefined) trend30Raw[r.dataOrigem]++; });

  const weekMap   = { "Dom":0,"Seg":0,"Ter":0,"Qua":0,"Qui":0,"Sex":0,"Sáb":0 };
  const weekNames = ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"];
  items.forEach(r => {
    if (!r.dataOrigem) return;
    const p = r.dataOrigem.split("-");
    weekMap[weekNames[new Date(Number(p[0]), Number(p[1])-1, Number(p[2])).getDay()]]++;
  });

  const cliTotalMap = {};
  items.forEach(r => { cliTotalMap[r.cliente] = (cliTotalMap[r.cliente]||0)+1; });
  const top10 = Object.entries(cliTotalMap)
    .sort((a,b)=>b[1]-a[1]).slice(0,10)
    .reduce((o,[k,v])=>{ o.labels.push(k); o.data.push(v); return o; },{labels:[],data:[]});

  const trend30 = Object.entries(trend30Raw)
    .reduce((o,[k,v])=>{ o.labels.push(k.substring(5)); o.data.push(v); return o; },{labels:[],data:[]});

  return { cliHoje, versoesHojeMap, trend30, weekMap, top10, cliTotalMap, trend30Raw };
}

function buildReports_(items, cliTotalMap, trend30Raw, tz) {
  const total = items.length;

  const mensalMap = {};
  for (let i = 11; i >= 0; i--) {
    const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - i);
    mensalMap[Utilities.formatDate(d, tz, "yyyy-MM")] = 0;
  }
  items.forEach(r => {
    const m = r.dataOrigem.substring(0, 7);
    if (mensalMap[m] !== undefined) mensalMap[m]++;
  });

  const revMap = {};
  items.forEach(r => { if (r.revisao) revMap[r.revisao] = (revMap[r.revisao]||0)+1; });
  const topRevisoes = Object.entries(revMap).sort((a,b)=>b[1]-a[1]).slice(0,10)
    .reduce((o,[k,v])=>{ o.labels.push(k); o.data.push(v); return o; },{labels:[],data:[]});

  const sortedCli = Object.entries(cliTotalMap).sort((a,b)=>b[1]-a[1]).slice(0,15);
  const pareto    = sortedCli.reduce((o,[k,v])=>{ o.labels.push(k); o.data.push(v); return o; },{labels:[],data:[]});
  let acc = 0;
  const totalP = pareto.data.reduce((a,b)=>a+b,0);
  pareto.cumulative = pareto.data.map(v => { acc += v; return totalP ? Math.round(acc*100/totalP) : 0; });

  const cliRevMap = {};
  items.forEach(r => {
    if (!cliRevMap[r.cliente]) cliRevMap[r.cliente] = { total:0, comRev:0 };
    cliRevMap[r.cliente].total++;
    if (r.revisao && r.revisao.trim()) cliRevMap[r.cliente].comRev++;
  });
  const taxaRevisao = Object.entries(cliRevMap)
    .map(([k,v])=>({ cliente:k, taxa: v.total ? Math.round(v.comRev*100/v.total) : 0 }))
    .sort((a,b)=>b.taxa-a.taxa).slice(0,10)
    .reduce((o,e)=>{ o.labels.push(e.cliente); o.data.push(e.taxa); return o; },{labels:[],data:[]});

  const comRevisao        = items.filter(r => r.revisao && r.revisao.trim()).length;
  const taxaRevisaoGlobal = total ? Math.round(comRevisao * 100 / total) : 0;

  const concentracao = total
    ? Math.round((Object.values(cliTotalMap).sort((a,b)=>b-a).slice(0,3).reduce((a,b)=>a+b,0))*100/total)
    : 0;

  const semCaminho     = items.filter(r => !r.caminho || !r.caminho.trim()).length;
  const semCaminhoPerc = total ? Math.round(semCaminho*100/total) : 0;

  const u7 = [], a7 = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(); d.setDate(d.getDate() - i);
    u7.push(trend30Raw[Utilities.formatDate(d, tz, "yyyy-MM-dd")] || 0);
  }
  for (let i = 13; i >= 7; i--) {
    const d = new Date(); d.setDate(d.getDate() - i);
    a7.push(trend30Raw[Utilities.formatDate(d, tz, "yyyy-MM-dd")] || 0);
  }
  const sumU7      = u7.reduce((a,b)=>a+b,0);
  const sumA7      = a7.reduce((a,b)=>a+b,0);
  const tendencia7 = sumA7 > 0 ? Math.round((sumU7-sumA7)*100/sumA7) : (sumU7>0?100:0);

  return {
    mensal: Object.entries(mensalMap).reduce((o,[k,v])=>{ o.labels.push(k); o.data.push(v); return o; },{labels:[],data:[]}),
    topRevisoes, pareto, taxaRevisao, concentracao, taxaRevisaoGlobal,
    tendencia7, ultimos7: sumU7, anteriores7: sumA7,
    semCaminhoCount: semCaminho, semCaminhoPerc,
    mesAtualKey: Utilities.formatDate(new Date(), tz, "yyyy-MM"),
  };
}

// ─── Dashboard ────────────────────────────────────────────────────────────────
function filterDecisionPeriod_(items, period, today, tz) {
  period = period || "30d";
  const refParts = String(today || Utilities.formatDate(new Date(), tz, "yyyy-MM-dd")).split("-");
  const refDate = new Date(Number(refParts[0]), Number(refParts[1]) - 1, Number(refParts[2]));
  const startDate = new Date(refDate);
  if (period === "7d") startDate.setDate(startDate.getDate() - 6);
  else if (period === "15d") startDate.setDate(startDate.getDate() - 14);
  else if (period === "3m") startDate.setMonth(startDate.getMonth() - 3);
  else startDate.setDate(startDate.getDate() - 29);
  const startKey = Utilities.formatDate(startDate, tz, "yyyy-MM-dd");
  return items.filter(r => r.dataOrigem >= startKey && r.dataOrigem <= today);
}

function buildDecision_(items, today, tz, ignoredClients, decisionPeriod) {
  ignoredClients = ignoredClients || [];
  const ignoredSet = new Set(ignoredClients);
  const periodItems = filterDecisionPeriod_(items, decisionPeriod, today, tz);
  const decisionItems = periodItems.filter(r => !ignoredSet.has(r.cliente));
  const totalOriginal = items.length;
  const total = decisionItems.length;
  const hoje = decisionItems.filter(r => r.dataOrigem === today).length;
  const semCaminho = decisionItems.filter(r => !r.caminho || !r.caminho.trim()).length;
  const semCliente = decisionItems.filter(r => !r.cliente || !r.cliente.trim()).length;
  const semVersao = decisionItems.filter(r => !r.versao || !r.versao.trim()).length;
  const semRequisito = decisionItems.filter(r => !r.requisito || !r.requisito.trim()).length;
  const comRevisao = decisionItems.filter(r => r.revisao && r.revisao.trim()).length;
  const taxaRevisao = total ? Math.round(comRevisao * 100 / total) : 0;
  const dataQuality = total
    ? Math.max(0, Math.round(100 - ((semCaminho + semCliente + semVersao + semRequisito) * 100 / (total * 4))))
    : 100;
  const rawCharts = buildCharts_(decisionItems, today, tz);
  const reports = buildReports_(decisionItems, rawCharts.cliTotalMap, rawCharts.trend30Raw, tz);

  const versionMap = {};
  items.forEach(r => {
    const key = r.versao || "Sem versão";
    if (!versionMap[key]) versionMap[key] = { total:0, semCaminho:0, revisao:0 };
    versionMap[key].total++;
    if (!r.caminho || !r.caminho.trim()) versionMap[key].semCaminho++;
    if (r.revisao && r.revisao.trim()) versionMap[key].revisao++;
  });

  const topClientEntry = Object.entries(rawCharts.cliTotalMap || {}).sort((a,b)=>b[1]-a[1])[0] || ["-", 0];
  const topVersionEntry = Object.entries(versionMap).sort((a,b)=>b[1].total-a[1].total)[0] || ["-", { total:0, semCaminho:0, revisao:0 }];

  const topClients = Object.entries(rawCharts.cliTotalMap || {})
    .sort((a,b)=>b[1]-a[1])
    .slice(0, 8)
    .map(([cliente, count]) => {
      const cliItems = items.filter(r => r.cliente === cliente);
      const cliSemCaminho = cliItems.filter(r => !r.caminho || !r.caminho.trim()).length;
      const cliComRev = cliItems.filter(r => r.revisao && r.revisao.trim()).length;
      const pct = total ? Math.round(count * 100 / total) : 0;
      let status = "Normal";
      if (pct >= 35 || cliSemCaminho >= 10) status = "Crítico";
      else if (pct >= 20 || cliSemCaminho >= 3 || (count && cliComRev * 100 / count >= 50)) status = "Atenção";
      return { cliente, total: count, pct, semCaminho: cliSemCaminho, taxaRevisao: count ? Math.round(cliComRev * 100 / count) : 0, status };
    });

  const trendValues = rawCharts.trend30Raw ? Object.values(rawCharts.trend30Raw) : [];
  const avg30 = trendValues.length ? trendValues.reduce((a,b)=>a+b,0) / trendValues.length : 0;
  const spikePct = avg30 ? Math.round((hoje - avg30) * 100 / avg30) : 0;
  const riskScore =
    (reports.semCaminhoPerc >= 10 ? 35 : reports.semCaminhoPerc >= 5 ? 20 : 0) +
    (reports.concentracao >= 70 ? 30 : reports.concentracao >= 50 ? 15 : 0) +
    (reports.tendencia7 >= 30 ? 20 : reports.tendencia7 >= 15 ? 10 : 0) +
    (dataQuality < 85 ? 20 : dataQuality < 95 ? 10 : 0);
  const riskLevel = riskScore >= 60 ? "Crítico" : riskScore >= 30 ? "Atenção" : "Normal";

  const recommendations = [];
  if (semCaminho > 0) recommendations.push({ level: semCaminho >= 10 ? "Crítico" : "Atenção", title: "Corrigir caminhos ausentes", detail: semCaminho + " registro(s) sem caminho podem bloquear conferência ou comunicação." });
  if (topClientEntry[1] && total && Math.round(topClientEntry[1] * 100 / total) >= 30) recommendations.push({ level: "Atenção", title: "Monitorar concentração por cliente", detail: topClientEntry[0] + " concentra " + Math.round(topClientEntry[1] * 100 / total) + "% do volume filtrado." });
  if (reports.tendencia7 >= 20) recommendations.push({ level: "Atenção", title: "Preparar capacidade operacional", detail: "Os últimos 7 dias estão " + reports.tendencia7 + "% acima dos 7 dias anteriores." });
  if (spikePct >= 50 && hoje >= 5) recommendations.push({ level: "Crítico", title: "Investigar pico diário", detail: "Hoje está " + spikePct + "% acima da média diária dos últimos 30 dias." });
  if (taxaRevisao >= 40) recommendations.push({ level: "Atenção", title: "Revisar causas de retrabalho", detail: "A taxa de revisão está em " + taxaRevisao + "% no período filtrado." });
  if (!recommendations.length) recommendations.push({ level: "Normal", title: "Operação estável", detail: "Nenhum desvio relevante foi identificado para o período filtrado." });

  const anomalies = [];
  if (spikePct >= 50 && hoje >= 5) anomalies.push("Volume de hoje acima do padrão recente: +" + spikePct + "%.");
  if (reports.semCaminhoPerc >= 10) anomalies.push("Percentual de registros sem caminho acima de 10%.");
  if (reports.concentracao >= 70) anomalies.push("Top 3 clientes concentram " + reports.concentracao + "% do volume.");
  if (dataQuality < 90) anomalies.push("Qualidade da base abaixo de 90%.");

  return {
    riskLevel, riskScore,
    topClient: { name: topClientEntry[0], total: topClientEntry[1], pct: total ? Math.round(topClientEntry[1] * 100 / total) : 0 },
    topVersion: { name: topVersionEntry[0], total: topVersionEntry[1].total, semCaminho: topVersionEntry[1].semCaminho },
    dataQuality, tendencia7: reports.tendencia7,
    semCaminho, semCaminhoPerc: reports.semCaminhoPerc, taxaRevisao,
    recommendations: recommendations.slice(0, 5),
    anomalies, topClients,
    health: { semCaminho, semCliente, semVersao, semRequisito },
    ignoredClients,
    ignoredCount: periodItems.length - total,
    decisionPeriod: decisionPeriod || "30d",
  };
}

function getDashboardData(filtersJson) {
  try {
    assertAuthorized_("getDashboardData");
    const filters = filtersJson ? JSON.parse(filtersJson) : {};
    const cache = CacheService.getScriptCache();
    const ignoredDecisionClients = getIgnoredDecisionClients_();
    const cacheKey = "dashboard:" + Utilities.base64EncodeWebSafe((filtersJson || "{}") + "|" + ignoredDecisionClients.join("|")).substring(0, 180);
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    const ss      = getSpreadsheet_();
    const tz      = ss.getSpreadsheetTimeZone();
    const today   = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

    const allItems  = buildAllItems(ss);
    const items     = applyFilters(allItems, filters);
    const cards     = buildCards_(items, today, tz);
    const rawCharts = buildCharts_(items, today, tz);
    const charts    = {
      cliHoje:        rawCharts.cliHoje,
      versoesHojeMap: rawCharts.versoesHojeMap,
      trend30:        rawCharts.trend30,
      weekMap:        rawCharts.weekMap,
      top10:          rawCharts.top10,
    };
    const reports    = buildReports_(items, rawCharts.cliTotalMap, rawCharts.trend30Raw, tz);
    const decision   = buildDecision_(items, today, tz, ignoredDecisionClients, filters.decisionPeriod || "30d");
    const clientList = [...new Set(allItems.map(r => r.cliente).filter(Boolean))].sort();
    const sortCampo  = filters.sortCampo || "dataOrigem";
    const sortDir    = filters.sortDir   || "desc";
    const sorted     = applySort(items, sortCampo, sortDir);
    const pageRows   = sorted.slice(0, CONFIG.PAGE_SIZE);
    const totalPages = Math.ceil(sorted.length / CONFIG.PAGE_SIZE);

    const response = {
      ok: true, cards, charts, reports, decision,
      table: { rows: pageRows, totalPages, totalRows: sorted.length, page: 1 },
      clientList,
    };
    cache.put(cacheKey, JSON.stringify(response), CONFIG.CACHE_SECONDS);
    return response;
  } catch(e) {
    logEvent_("ERROR", "SISTEMA", "getDashboardData", "Erro no dashboard", e.message);
    return { ok: false, error: e.message };
  }
}

// ─── Paginated table ──────────────────────────────────────────────────────────
function getTableData(page, filtersJson) {
  try {
    assertAuthorized_("getTableData");
    const filters = filtersJson ? JSON.parse(filtersJson) : {};
    const ss = getSpreadsheet_();
    let items = applyFilters(buildAllItems(ss), filters);
    const sortCampo = filters.sortCampo || "dataOrigem";
    const sortDir   = filters.sortDir   || "desc";
    items = applySort(items, sortCampo, sortDir);
    const p          = Math.max(1, Number(page) || 1);
    const start      = (p - 1) * CONFIG.PAGE_SIZE;
    const pageRows   = items.slice(start, start + CONFIG.PAGE_SIZE);
    const totalPages = Math.ceil(items.length / CONFIG.PAGE_SIZE);
    return { ok: true, rows: pageRows, page: p, totalPages, totalRows: items.length };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ─── CSV Export ───────────────────────────────────────────────────────────────
function exportCSVData(filtersJson) {
  try {
    assertAuthorized_("exportCSVData");
    const filters = filtersJson ? JSON.parse(filtersJson) : {};
    const ss = getSpreadsheet_();
    let items = applyFilters(buildAllItems(ss), filters);
    const sortCampo = filters.sortCampo || "dataOrigem";
    const sortDir   = filters.sortDir   || "desc";
    items = applySort(items, sortCampo, sortDir);
    const header = ["Data","Versão","Requisito","Cliente","Status","Caminho","Revisão"];
    const rows   = items.map(r =>
      [r.dataExibicao,r.versao,r.requisito,r.cliente,r.status,r.caminho,r.revisao]
        .map(v => '"' + String(v||"").replace(/"/g,'""') + '"').join(";"));
    return { ok: true, csv: [header.join(";"), ...rows].join("\r\n") };
  } catch(e) {
    logEvent_("ERROR", "SISTEMA", "exportCSVData", "Erro ao exportar CSV", e.message);
    return { ok: false, error: e.message };
  }
}

// ─── KPI Save ─────────────────────────────────────────────────────────────────
function hashPayload_(jsonString) {
  try {
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, jsonString || "");
    return digest.map(b => ("0" + (b & 0xFF).toString(16)).slice(-2)).join("");
  } catch(e) {
    return "";
  }
}

function legacySessionIdFromTs_(ts) {
  const clean = String(ts || "").replace(/[^0-9]/g, "");
  return "LEGADO_" + (clean || Utilities.getUuid());
}

function isKPIHistoricoNewFormat_(sheet) {
  if (!sheet || sheet.getLastRow() === 0) return false;
  const firstHeader = String(sheet.getRange(1, 1).getValue() || "").toUpperCase().trim();
  return firstHeader === "SESSION_ID";
}

function ensureKPIHistoricoNewFormat_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_KPI);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_KPI);
    sheet.getRange(1, 1, 1, KPI_HEADER.length).setValues([KPI_HEADER]);
    try { sheet.setFrozenRows(1); } catch(e) {}
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, KPI_HEADER.length).setValues([KPI_HEADER]);
    try { sheet.setFrozenRows(1); } catch(e) {}
    return sheet;
  }

  if (isKPIHistoricoNewFormat_(sheet)) return sheet;

  const tz = ss.getSpreadsheetTimeZone();
  const lastRow = sheet.getLastRow();
  sheet.insertColumnBefore(1);
  sheet.getRange(1, 1, 1, KPI_HEADER.length).setValues([KPI_HEADER]);
  try { sheet.setFrozenRows(1); } catch(e) {}

  if (lastRow >= 2) {
    const tsValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const sessionValues = tsValues.map(row => {
      const raw = row[0];
      const ts = raw instanceof Date
        ? Utilities.formatDate(raw, tz, "yyyy-MM-dd HH:mm:ss")
        : String(raw || "").trim();
      return [legacySessionIdFromTs_(ts)];
    });
    sheet.getRange(2, 1, sessionValues.length, 1).setValues(sessionValues);
  }

  logEvent_("INFO", "SISTEMA", "ensureKPIHistoricoNewFormat_",
    "KPI_Historico migrado para schema com SESSION_ID", "linhas=" + Math.max(0, lastRow - 1));
  return sheet;
}

function saveKPIDataManual(jsonString, origem) {
  const lock = LockService.getScriptLock();
  let locked = false;

  // Declare all vars before try so catch block can reference them for audit logging.
  // Initialization happens inside try to keep lock always released via finally.
  let ss, tz, ts, sessionId;
  const origemStr   = origem || "MANUAL";
  let   hashPayload = "";

  try {
    if (origemStr !== "POST") assertAuthorized_("saveKPIDataManual");
    lock.waitLock(30000);
    locked = true;

    ss        = getSpreadsheet_();
    tz        = ss.getSpreadsheetTimeZone();
    ts        = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    sessionId = Utilities.getUuid();
    const data = JSON.parse(jsonString);
    if (!Array.isArray(data) || data.length === 0) throw new Error("JSON inválido ou vazio");

    hashPayload = hashPayload_(jsonString).substring(0, 16);
    const kpiSheet = ensureKPIHistoricoNewFormat_(ss);

    const requiredFields = KPI_HEADER.slice(2); // exclude SESSION_ID and TIMESTAMP
    const rows = data.map((item, idx) => {
      const label = item.ALT_VERSAO || "?";

      const missing = requiredFields.filter(f => item[f] === undefined || item[f] === null);
      if (missing.length)
        throw new Error("Item " + (idx+1) + " (" + label + "): campos ausentes: " + missing.join(", "));

      const invalidNum = KPI_NUMERIC_FIELDS.filter(f => !isValidNumberValue(item[f]));
      if (invalidNum.length)
        throw new Error("Item " + (idx+1) + " (" + label + "): valor inválido em: " +
          invalidNum.map(f => f + '="' + item[f] + '"').join(", "));

      const baseRow = [
        String(item.ALT_VERSAO   || ""),
        Number(item.ABERTA),
        Number(item.ANDAMENTO),
        Number(item.CORRIGINDO),
        Number(item.CORRIGIDO),
        Number(item.CONFERIDO),
        Number(item.TOTAL),
        Number(item.KPI_ABERTA_ANDAMENTO),
        Number(item.KPI_CORRIGINDO_CORRIGIDO),
        Number(item.KPI_CONFERIDO),
        String(item.SEMAFORO_ABERTA_ANDAMENTO      || ""),
        String(item.SEMAFORO_CORRIGINDO_CORRIGIDO  || ""),
        String(item.SEMAFORO_CONFERIDO             || ""),
      ];
      return [sessionId, ts, ...baseRow];
    });

    const lastRow  = Math.max(kpiSheet.getLastRow(), 1);
    kpiSheet.getRange(lastRow + 1, 1, rows.length, KPI_HEADER.length).setValues(rows);

    // Audit record
    const sessoesSheet = ensureSheet_(ss, CONFIG.SHEET_SESSOES, SESSOES_HEADER);
    sessoesSheet.appendRow([sessionId, ts, rows.length, origemStr, "SUCESSO", "", hashPayload]);

    logEvent_("INFO", origemStr, "saveKPIDataManual",
      "KPI salvo: " + rows.length + " registros", "session_id=" + sessionId);

    return { ok: true, saved: rows.length, timestamp: ts, sessionId };

  } catch(e) {
    try {
      const sessoesSheet = ensureSheet_(ss, CONFIG.SHEET_SESSOES, SESSOES_HEADER);
      sessoesSheet.appendRow([sessionId, ts, 0, origemStr, "ERRO", e.message, hashPayload]);
    } catch(e2) {}
    logEvent_("ERROR", origemStr, "saveKPIDataManual", "Erro ao salvar KPI", e.message);
    return { ok: false, error: e.message };
  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch(e) {}
    }
  }
}

// ─── KPI History ──────────────────────────────────────────────────────────────
function getKPIHistorySessions() {
  try {
    assertAuthorized_("getKPIHistorySessions");
    const ss      = getSpreadsheet_();
    const tz      = ss.getSpreadsheetTimeZone();
    const sessions = [];
    const knownTs  = new Set();

    // Primary source: KPI_Sessoes (new audit table)
    const sessoesSheet = ss.getSheetByName(CONFIG.SHEET_SESSOES);
    if (sessoesSheet && sessoesSheet.getLastRow() >= 2) {
      const numCols = Math.min(sessoesSheet.getLastColumn(), SESSOES_HEADER.length);
      sessoesSheet.getRange(2, 1, sessoesSheet.getLastRow() - 1, numCols).getValues()
        .forEach(row => {
          const sid = String(row[0] || "").trim();
          const ts  = row[1] instanceof Date
            ? Utilities.formatDate(row[1], tz, "yyyy-MM-dd HH:mm:ss")
            : String(row[1] || "").trim();
          if (!ts) return;
          knownTs.add(ts);
          sessions.push({
            id:     sid || ts,
            ts,
            count:  Number(row[2] || 0),
            origem: String(row[3] || "MANUAL"),
            status: String(row[4] || "SUCESSO"),
          });
        });
    }

    // Fallback: legacy timestamps from KPI_Historico not already in KPI_Sessoes
    const kpiSheet = ss.getSheetByName(CONFIG.SHEET_KPI);
    if (kpiSheet && kpiSheet.getLastRow() >= 2) {
      const numCols   = Math.min(kpiSheet.getLastColumn(), 2);
      const headerRow = kpiSheet.getRange(1, 1, 1, numCols).getValues()[0];
      const hasSessionId = String(headerRow[0] || "").toUpperCase().trim() === "SESSION_ID";
      const tsColIdx     = hasSessionId ? 1 : 0;
      const tsData       = kpiSheet.getRange(2, tsColIdx + 1, kpiSheet.getLastRow() - 1, 1).getValues();

      const legacyCount = {};
      tsData.forEach(row => {
        const raw = row[0];
        const ts  = raw instanceof Date
          ? Utilities.formatDate(raw, tz, "yyyy-MM-dd HH:mm:ss")
          : String(raw || "").trim();
        if (ts && !knownTs.has(ts)) legacyCount[ts] = (legacyCount[ts] || 0) + 1;
      });
      Object.entries(legacyCount).forEach(([ts, count]) => {
        sessions.push({ id: ts, ts, count, origem: "LEGADO", status: "SUCESSO" });
      });
    }

    sessions.sort((a, b) => b.ts.localeCompare(a.ts));
    return { ok: true, sessions };
  } catch(e) {
    logEvent_("ERROR", "SISTEMA", "getKPIHistorySessions", "Erro ao carregar sessões", e.message);
    return { ok: false, error: e.message };
  }
}

function getKPIDataBySession(identifier) {
  try {
    assertAuthorized_("getKPIDataBySession");
    const ss    = getSpreadsheet_();
    const sheet = ss.getSheetByName(CONFIG.SHEET_KPI);
    if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };
    const tz = ss.getSpreadsheetTimeZone();

    const numCols      = sheet.getLastColumn();
    const headerRow    = sheet.getRange(1, 1, 1, Math.min(numCols, 2)).getValues()[0];
    const hasSessionId = String(headerRow[0] || "").toUpperCase().trim() === "SESSION_ID";

    const lookupKeys = {};
    const addLookupKey = key => {
      key = String(key || "").trim();
      if (key) lookupKeys[key] = true;
    };

    addLookupKey(identifier);

    const sessoesSheet = ss.getSheetByName(CONFIG.SHEET_SESSOES);
    if (sessoesSheet && sessoesSheet.getLastRow() >= 2) {
      const sessoesRows = sessoesSheet.getRange(2, 1, sessoesSheet.getLastRow() - 1, 2).getValues();
      const found = sessoesRows.find(row => String(row[0] || "").trim() === String(identifier || "").trim());
      if (found) {
        const sessionTs = found[1] instanceof Date
          ? Utilities.formatDate(found[1], tz, "yyyy-MM-dd HH:mm:ss")
          : String(found[1] || "").trim();
        addLookupKey(sessionTs);
        addLookupKey(legacySessionIdFromTs_(sessionTs));
      }
    }

    const raw  = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
    const rows = raw.filter(row => {
      if (hasSessionId) {
        const sid = String(row[0] || "").trim();
        const ts  = row[1] instanceof Date
          ? Utilities.formatDate(row[1], tz, "yyyy-MM-dd HH:mm:ss")
          : String(row[1] || "").trim();
        return !!(lookupKeys[sid] || lookupKeys[ts] || lookupKeys[legacySessionIdFromTs_(ts)]);
      }
      const ts = row[0] instanceof Date
        ? Utilities.formatDate(row[0], tz, "yyyy-MM-dd HH:mm:ss")
        : String(row[0] || "").trim();
      return !!(lookupKeys[ts] || lookupKeys[legacySessionIdFromTs_(ts)]);
    }).map(row => parseKpiRow_(row, hasSessionId, tz));

    return { ok: true, data: rows };
  } catch(e) {
    logEvent_("ERROR", "SISTEMA", "getKPIDataBySession", "Erro ao carregar dados de sessão", e.message);
    return { ok: false, error: e.message };
  }
}

function getLastKPI() {
  try {
    assertAuthorized_("getLastKPI");
    const sessRes = getKPIHistorySessions();
    if (!sessRes.ok) return sessRes;
    const successSessions = (sessRes.sessions || [])
      .filter(s => String(s.status || "SUCESSO").toUpperCase() === "SUCESSO");
    if (!successSessions.length) {
      return { ok: true, data: [], session: null, sessionId: null, sessions: sessRes.sessions || [] };
    }
    const lastSess = successSessions[0];
    const dataRes  = getKPIDataBySession(lastSess.id);
    if (!dataRes.ok) return dataRes;
    return {
      ok: true,
      data:      dataRes.data || [],
      session:   lastSess.ts,
      sessionId: lastSess.id,
      sessions:  sessRes.sessions,
    };
  } catch(e) {
    logEvent_("ERROR", "SISTEMA", "getLastKPI", "Erro ao carregar último KPI", e.message);
    return { ok: false, error: e.message };
  }
}

// getKPITrendData — lê a planilha uma única vez para evitar leituras repetidas
function getKPITrendData() {
  try {
    assertAuthorized_("getKPITrendData");
    const sessRes = getKPIHistorySessions();
    if (!sessRes.ok || !sessRes.sessions.length) return { ok: true, trend: [] };

    const ss    = getSpreadsheet_();
    const sheet = ss.getSheetByName(CONFIG.SHEET_KPI);
    if (!sheet || sheet.getLastRow() < 2) return { ok: true, trend: [] };
    const tz = ss.getSpreadsheetTimeZone();

    const numCols      = sheet.getLastColumn();
    const headerRow    = sheet.getRange(1, 1, 1, Math.min(numCols, 2)).getValues()[0];
    const hasSessionId = String(headerRow[0] || "").toUpperCase().trim() === "SESSION_ID";

    // Single read of entire KPI sheet
    const allRaw = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();

    // Build lookup map: session key → parsed rows
    const byKey = {};
    allRaw.forEach(row => {
      const parsed = parseKpiRow_(row, hasSessionId, tz);
      const keys = [];
      if (hasSessionId) {
        keys.push(String(row[0] || "").trim());
        keys.push(parsed.TIMESTAMP);
        keys.push(legacySessionIdFromTs_(parsed.TIMESTAMP));
      } else {
        const ts = row[0] instanceof Date
          ? Utilities.formatDate(row[0], tz, "yyyy-MM-dd HH:mm:ss")
          : String(row[0] || "").trim();
        keys.push(ts);
        keys.push(legacySessionIdFromTs_(ts));
      }
      keys.forEach(key => {
        if (!key) return;
        if (!byKey[key]) byKey[key] = [];
        byKey[key].push(parsed);
      });
    });

    const last10 = sessRes.sessions
      .filter(s => String(s.status || "SUCESSO").toUpperCase() === "SUCESSO")
      .slice(0, 10)
      .reverse();
    const trend  = [];

    for (const sess of last10) {
      const sessRows = byKey[sess.id] || byKey[sess.ts] || [];
      const totalRow = sessRows.find(r => r.ALT_VERSAO === "TOTAL GERAL");
      if (!totalRow) continue;
      trend.push({
        ts:            sess.ts,
        conferidoPerc: totalRow.KPI_CONFERIDO          || 0,
        abeAndPerc:    totalRow.KPI_ABERTA_ANDAMENTO   || 0,
        corrCorrPerc:  totalRow.KPI_CORRIGINDO_CORRIGIDO || 0,
      });
    }

    return { ok: true, trend };
  } catch(e) {
    logEvent_("ERROR", "SISTEMA", "getKPITrendData", "Erro ao calcular tendência", e.message);
    return { ok: false, error: e.message };
  }
}
