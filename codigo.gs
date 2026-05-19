const CONFIG = {
  SHEET_GERAL: "Geral",
  SHEET_KPI:   "KPI_Historico",
  MAX_ROWS:    10000,
  PAGE_SIZE:   100,
  // POST_SECRET é lido via Arquivo → Propriedades do projeto (PropertiesService).
  // Defina a propriedade de script "POST_SECRET" com o mesmo valor usado em kpi_sync.py.
};

// 14-column KPI schema (ordem e nomes são contrato público com o frontend)
const KPI_HEADER = [
  "TIMESTAMP","ALT_VERSAO","ABERTA","ANDAMENTO","CORRIGINDO","CORRIGIDO",
  "CONFERIDO","TOTAL","KPI_ABERTA_ANDAMENTO","KPI_CORRIGINDO_CORRIGIDO",
  "KPI_CONFERIDO","SEMAFORO_ABERTA_ANDAMENTO","SEMAFORO_CORRIGINDO_CORRIGIDO",
  "SEMAFORO_CONFERIDO"
];

// Campos que devem ser números; zero é válido, "abc" não é
const KPI_NUMERIC_FIELDS = [
  "ABERTA","ANDAMENTO","CORRIGINDO","CORRIGIDO","CONFERIDO","TOTAL",
  "KPI_ABERTA_ANDAMENTO","KPI_CORRIGINDO_CORRIGIDO","KPI_CONFERIDO"
];

// ─── Entry points ────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Sistema de Gestão de Cópias v2026")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    if (!e || !e.postData) return jsonResp({ ok: false, error: "Sem dados" });
    if ((e.postData.contents || "").length > 524288)          // 512 KB
      return jsonResp({ ok: false, error: "Payload excede 512 KB" });
    const payload = JSON.parse(e.postData.contents);
    const secret  = PropertiesService.getScriptProperties().getProperty("POST_SECRET") || "";
    if (!secret || payload.secret !== secret)
      return jsonResp({ ok: false, error: "Não autorizado" });
    if (!Array.isArray(payload.data) || payload.data.length > 500)
      return jsonResp({ ok: false, error: "payload.data deve ser array com até 500 itens" });
    return jsonResp(saveKPIDataManual(JSON.stringify(payload.data)));
  } catch(err) {
    return jsonResp({ ok: false, error: err.message });
  }
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Helpers ─────────────────────────────────────────────────────────────────
function getSheetData(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const last = Math.min(sheet.getLastRow(), CONFIG.MAX_ROWS + 1);
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
}

// Normaliza header para comparação: minúsculas + sem acentos
function normalizeHeader(h) {
  return String(h || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").trim();
}

function buildAllItems(ss) {
  const tz    = ss.getSpreadsheetTimeZone();
  const sheet = ss.getSheetByName(CONFIG.SHEET_GERAL);
  if (!sheet) return [];
  const last = Math.min(sheet.getLastRow(), CONFIG.MAX_ROWS + 1);
  if (last < 2) return [];
  const numCols = sheet.getLastColumn();

  const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0].map(normalizeHeader);

  function findCol(keywords, fallback) {
    const normKeys = keywords.map(normalizeHeader);
    // Prioridade: match exato → startsWith → includes
    for (const strategy of [
      (h, k) => h === k,
      (h, k) => h.startsWith(k),
      (h, k) => h.includes(k),
    ]) {
      const idx = headers.findIndex(h => normKeys.some(k => strategy(h, k)));
      if (idx >= 0) return idx;
    }
    return fallback;
  }

  const c = {
    data:      findCol(["data_liberacao", "data_lib", "data_", "dt_", "data", "date", "dt"],        0),
    versao:    findCol(["alt_versao", "versao_alt", "versao", "version", "ver_"],                    1),
    requisito: findCol(["nr_requisito", "num_requisito", "requisito", "requirement", "req_"],        2),
    cliente:   findCol(["nm_cliente", "nome_cliente", "cliente", "client", "customer"],              3),
    status:    findCol(["ds_status", "status", "situacao", "state"],                                 5),
    caminho:   findCol(["caminho_rede", "caminho_completo", "caminho", "path", "diretorio", "dir_"], 7),
    revisao:   findCol(["nr_revisao", "num_revisao", "revisao", "revision", "rev_"],                 8),
  };

  const rows = sheet.getRange(2, 1, last - 1, numCols).getValues();
  return rows.map(row => ({
    dataOrigem:   row[c.data] instanceof Date ? Utilities.formatDate(row[c.data], tz, "yyyy-MM-dd") : "",
    dataExibicao: row[c.data] instanceof Date ? Utilities.formatDate(row[c.data], tz, "dd/MM/yyyy") : "",
    versao:    String(row[c.versao]    || ""),
    requisito: String(row[c.requisito] || ""),
    cliente:   String(row[c.cliente]   || ""),
    status:    String(row[c.status]    || ""),
    caminho:   String(row[c.caminho]   || ""),
    revisao:   String(row[c.revisao]   || ""),
  })).filter(r => r.dataOrigem !== "");
}

function applyFilters(items, filters) {
  return items.filter(r => {
    if (filters.cliente    && r.cliente    !== filters.cliente)       return false;
    if (filters.dataInicio && r.dataOrigem <  filters.dataInicio)     return false;
    if (filters.dataFim    && r.dataOrigem >  filters.dataFim)        return false;
    if (filters.revisao    && !r.revisao.toLowerCase().includes(filters.revisao.toLowerCase())) return false;
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

// ─── Dashboard helpers ───────────────────────────────────────────────────────
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
    total,
    hoje,
    semCaminho,
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

  // cliTotalMap e trend30Raw são passados para buildReports_ via getDashboardData
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

  // Pareto (top 15 clientes, com % acumulado)
  const sortedCli = Object.entries(cliTotalMap).sort((a,b)=>b[1]-a[1]).slice(0,15);
  const pareto    = sortedCli.reduce((o,[k,v])=>{ o.labels.push(k); o.data.push(v); return o; },{labels:[],data:[]});
  let acc = 0;
  const totalP = pareto.data.reduce((a,b)=>a+b,0);
  pareto.cumulative = pareto.data.map(v => { acc += v; return totalP ? Math.round(acc*100/totalP) : 0; });

  // Taxa de revisão por cliente (% de itens com revisão preenchida)
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

  // Tendência 7d vs 7d anteriores
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
    topRevisoes,
    pareto,
    taxaRevisao,
    concentracao,
    taxaRevisaoGlobal,
    tendencia7,
    ultimos7:        sumU7,
    anteriores7:     sumA7,
    semCaminhoCount: semCaminho,
    semCaminhoPerc,
    // Chave do mês atual calculada no backend (fuso horário da planilha)
    mesAtualKey: Utilities.formatDate(new Date(), tz, "yyyy-MM"),
  };
}

// ─── Dashboard (dados agregados — nenhum dado bruto é transferido) ────────────
function getDashboardData(filtersJson) {
  try {
    const filters = filtersJson ? JSON.parse(filtersJson) : {};
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
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
    const clientList = [...new Set(allItems.map(r => r.cliente).filter(Boolean))].sort();

    const sortCampo  = filters.sortCampo || "dataOrigem";
    const sortDir    = filters.sortDir   || "desc";
    const sorted     = applySort(items, sortCampo, sortDir);
    const pageRows   = sorted.slice(0, CONFIG.PAGE_SIZE);
    const totalPages = Math.ceil(sorted.length / CONFIG.PAGE_SIZE);

    return {
      ok: true,
      cards,
      charts,
      reports,
      table: { rows: pageRows, totalPages, totalRows: sorted.length, page: 1 },
      clientList,
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ─── Paginated table ─────────────────────────────────────────────────────────
function getTableData(page, filtersJson) {
  try {
    const filters = filtersJson ? JSON.parse(filtersJson) : {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
    const filters = filtersJson ? JSON.parse(filtersJson) : {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let items = applyFilters(buildAllItems(ss), filters);
    const sortCampo = filters.sortCampo || "dataOrigem";
    const sortDir   = filters.sortDir   || "desc";
    items = applySort(items, sortCampo, sortDir);
    const header = ["Data","Versão","Requisito","Cliente","Status","Caminho","Revisão"];
    const rows   = items.map(r =>
      [r.dataExibicao,r.versao,r.requisito,r.cliente,r.status,r.caminho,r.revisao]
        .map(v => `"${String(v||"").replace(/"/g,'""')}"`).join(";"));
    return { ok: true, csv: [header.join(";"), ...rows].join("\r\n") };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ─── KPI Functions ────────────────────────────────────────────────────────────
function saveKPIDataManual(jsonString) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = JSON.parse(jsonString);
    if (!Array.isArray(data) || data.length === 0) throw new Error("JSON inválido ou vazio");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_KPI);
    if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_KPI);
    // Garante cabeçalho quando a aba existe mas está vazia (criação nova ou limpeza manual)
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 14).setValues([KPI_HEADER]);
    }

    const tz = ss.getSpreadsheetTimeZone();
    const ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    const requiredFields = KPI_HEADER.slice(1); // todos exceto TIMESTAMP
    const rows = data.map((item, idx) => {
      const label = item.ALT_VERSAO || "?";

      const missing = requiredFields.filter(f => item[f] === undefined || item[f] === null);
      if (missing.length)
        throw new Error(`Item ${idx+1} (${label}): campos ausentes: ${missing.join(", ")}`);

      // Rejeita valores não numéricos; zero é válido
      const invalidNum = KPI_NUMERIC_FIELDS.filter(f => isNaN(Number(item[f])));
      if (invalidNum.length)
        throw new Error(`Item ${idx+1} (${label}): valor não numérico em: ${invalidNum.join(", ")}`);

      return [
        ts,
        String(item.ALT_VERSAO || ""),
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
    });

    const lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange(lastRow + 1, 1, rows.length, 14).setValues(rows);
    return { ok: true, saved: rows.length, timestamp: ts };
  } catch(e) {
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getKPIHistorySessions() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_KPI);
    if (!sheet || sheet.getLastRow() < 2) return { ok: true, sessions: [] };
    const tz = ss.getSpreadsheetTimeZone();

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    const countMap = {};
    data.forEach(row => {
      const raw = row[0];
      const ts  = raw instanceof Date
        ? Utilities.formatDate(raw, tz, "yyyy-MM-dd HH:mm:ss")
        : String(raw || "").trim();
      if (ts) countMap[ts] = (countMap[ts] || 0) + 1;
    });

    const sessions = Object.entries(countMap)
      .sort((a, b) => b[0].localeCompare(a[0]))
      .map(([ts, count]) => ({ ts, count }));

    return { ok: true, sessions };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getKPIDataBySession(timestampStr) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_KPI);
    if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };
    const tz = ss.getSpreadsheetTimeZone();

    const numCols = sheet.getLastColumn();
    const raw     = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
    const isNew   = numCols >= 14;

    const rows = raw
      .filter(row => {
        const v  = row[0];
        const ts = v instanceof Date
          ? Utilities.formatDate(v, tz, "yyyy-MM-dd HH:mm:ss")
          : String(v || "").trim();
        return ts === timestampStr;
      })
      .map(row => ({
        ALT_VERSAO:                    String(row[1]  || ""),
        ABERTA:                        Number(row[2]  || 0),
        ANDAMENTO:                     Number(row[3]  || 0),
        CORRIGINDO:                    Number(row[4]  || 0),
        CORRIGIDO:                     Number(row[5]  || 0),
        CONFERIDO:                     Number(row[6]  || 0),
        TOTAL:                         Number(row[7]  || 0),
        KPI_ABERTA_ANDAMENTO:          Number(row[8]  || 0),
        KPI_CORRIGINDO_CORRIGIDO:      isNew ? Number(row[9]  || 0) : null,
        KPI_CONFERIDO:                 isNew ? Number(row[10] || 0) : Number(row[9]  || 0),
        SEMAFORO_ABERTA_ANDAMENTO:     isNew ? String(row[11] || "") : String(row[10] || ""),
        SEMAFORO_CORRIGINDO_CORRIGIDO: isNew ? String(row[12] || "") : null,
        SEMAFORO_CONFERIDO:            isNew ? String(row[13] || "") : String(row[11] || ""),
      }));

    return { ok: true, data: rows };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getLastKPI() {
  try {
    const sessRes = getKPIHistorySessions();
    if (!sessRes.ok) return sessRes;
    if (!sessRes.sessions.length) return { ok: true, data: [], session: null, sessions: [] };
    const lastTs  = sessRes.sessions[0].ts;
    const dataRes = getKPIDataBySession(lastTs);
    if (!dataRes.ok) return dataRes;
    return { ok: true, data: dataRes.data || [], session: lastTs, sessions: sessRes.sessions };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getKPITrendData() {
  try {
    const sessRes = getKPIHistorySessions();
    if (!sessRes.ok || !sessRes.sessions.length) return { ok: true, trend: [] };

    const last10 = sessRes.sessions.slice(0, 10).reverse();
    const trend  = [];

    for (const { ts } of last10) {
      const dataRes = getKPIDataBySession(ts);
      if (!dataRes.ok) continue;
      const totalRow = dataRes.data.find(r => r.ALT_VERSAO === "TOTAL GERAL");
      if (!totalRow) continue;
      trend.push({
        ts,
        conferidoPerc: totalRow.KPI_CONFERIDO || 0,
        abeAndPerc:    totalRow.KPI_ABERTA_ANDAMENTO || 0,
        corrCorrPerc:  totalRow.KPI_CORRIGINDO_CORRIGIDO || 0,
      });
    }

    return { ok: true, trend };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}
