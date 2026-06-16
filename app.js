const APP_CONFIG = {
  authRequired: false,
  tenantId: "",
  clientId: "",
  redirectUri: window.location.href.split("#")[0],
};

const MONTHS = [
  ["janeiro", "jan"],
  ["fevereiro", "fev"],
  ["marco", "mar"],
  ["abril", "abr"],
  ["maio", "mai"],
  ["junho", "jun"],
  ["julho", "jul"],
  ["agosto", "ago"],
  ["setembro", "set"],
  ["outubro", "out"],
  ["novembro", "nov"],
  ["dezembro", "dez"],
];

const state = {
  msalClient: null,
  account: null,
  result: null,
};

const uploadForm = document.querySelector("#uploadForm");
const fileInput = document.querySelector("#fileInput");
const processButton = document.querySelector("#processButton");
const statusText = document.querySelector("#statusText");
const summaryPanel = document.querySelector("#summaryPanel");
const downloadPanel = document.querySelector("#downloadPanel");
const downloadButton = document.querySelector("#downloadButton");
const tabs = document.querySelector("#tabs");
const tablePanel = document.querySelector("#tablePanel");
const loginButton = document.querySelector("#loginButton");
const logoutButton = document.querySelector("#logoutButton");
const authStatus = document.querySelector("#authStatus");

window.addEventListener("DOMContentLoaded", init);

function init() {
  uploadForm.addEventListener("submit", processSelectedFile);
  downloadButton.addEventListener("click", downloadProcessedWorkbook);

  if (APP_CONFIG.authRequired) {
    loginButton.hidden = false;
    processButton.disabled = true;
    authStatus.textContent = "Entre com Microsoft para processar.";
    loginButton.addEventListener("click", login);
    logoutButton.addEventListener("click", logout);
  }
}

async function login() {
  if (!window.msal) {
    setStatus("Biblioteca Microsoft ainda nao carregou.", true);
    return;
  }
  state.msalClient = new msal.PublicClientApplication({
    auth: {
      clientId: APP_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${APP_CONFIG.tenantId}`,
      redirectUri: APP_CONFIG.redirectUri,
    },
  });
  const response = await state.msalClient.loginPopup({ scopes: ["User.Read"] });
  state.account = response.account;
  authStatus.textContent = response.account?.username || "Usuario autenticado";
  loginButton.hidden = true;
  logoutButton.hidden = false;
  processButton.disabled = false;
}

async function logout() {
  if (state.msalClient && state.account) {
    await state.msalClient.logoutPopup({ account: state.account });
  }
  state.account = null;
  loginButton.hidden = false;
  logoutButton.hidden = true;
  processButton.disabled = true;
  authStatus.textContent = "Entre com Microsoft para processar.";
}

async function processSelectedFile(event) {
  event.preventDefault();
  if (APP_CONFIG.authRequired && !state.account) {
    setStatus("Autenticacao obrigatoria.", true);
    return;
  }
  if (!window.XLSX) {
    setStatus("Biblioteca de Excel ainda nao carregou.", true);
    return;
  }
  const file = fileInput.files[0];
  if (!file) return;

  try {
    processButton.disabled = true;
    clearResults();
    setStatus("Lendo arquivo no navegador...");
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true, dense: false });
    setStatus("Processando Base, C, D e F...");
    state.result = runPipeline(workbook);
    renderResult(state.result);
    setStatus("Processamento concluido. Nenhum dado foi enviado para servidor.");
  } catch (error) {
    console.error(error);
    setStatus(error.message || "Falha ao processar arquivo.", true);
  } finally {
    processButton.disabled = APP_CONFIG.authRequired && !state.account;
  }
}

function runPipeline(workbook) {
  const alerts = [];
  const extracted = extractWorkbook(workbook, alerts);
  const normalized = normalizeData(extracted, alerts);
  const transformed = transformData(normalized, alerts);
  addValidationAlerts(normalized, transformed, alerts);
  const summary = buildSummary(normalized, transformed, alerts);
  const outputs = {
    Resumo: objectToRows(summary),
    Alertas: alerts.map((alert) => alertRecord(alert)),
    "Calendario atividades": transformed.calendarioAtividades,
    "Meses Trabalhados": transformed.mesesTrabalhados,
    "Meses Recebidos ($)": transformed.mesesRecebidos,
    "Comparacao 1": transformed.comparacao1,
    "Comparacao 2": transformed.comparacao2,
    "F - Equipe tratada": normalized.baseF,
  };
  return { summary, outputs, alerts };
}

function extractWorkbook(workbook, alerts) {
  const baseSheet = findSheet(workbook, ["BASE", "Base"]);
  const cSheet = findSheet(workbook, ["C - DADOS ATIVIDADES", "C - Atividades"]);
  const dSheet = findSheet(workbook, ["D - DADOS EQUIPE", "D - Dados Equipe"]);
  const fSheet = findSheet(workbook, ["F - Desp. EQUIPE", "F - Equipe ($)"]);

  if (!baseSheet) alerts.push(alert("Erro", "Layout", "Extracao", "Aba Base nao encontrada.", { campo: "Base" }));
  if (!cSheet) alerts.push(alert("Erro", "Layout", "Extracao", "Aba C nao encontrada.", { campo: "C" }));
  if (!dSheet) alerts.push(alert("Erro", "Layout", "Extracao", "Aba D nao encontrada.", { campo: "D" }));
  if (!fSheet) alerts.push(alert("Erro", "Layout", "Extracao", "Aba F nao encontrada.", { campo: "F" }));

  return {
    baseRaw: baseSheet ? extractBase(sheetRows(baseSheet)) : [],
    cRaw: cSheet ? extractC(sheetRows(cSheet)) : [],
    dRaw: dSheet ? extractD(sheetRows(dSheet)) : [],
    fRaw: fSheet ? extractF(sheetRows(fSheet)) : [],
  };
}

function extractBase(rows) {
  const records = [];
  rows.forEach((row, index) => {
    const label = row[0];
    const key = normalizeText(label);
    if (key.startsWith("base.") || key.includes("data de inicio projeto")) {
      records.push({ Campo: label, Valor: firstNonBlank(row.slice(1, 4)), Linha: index + 1 });
    }
  });
  return records;
}

function extractC(rows) {
  const header = findHeaderRow(rows, ["C.1", "ATIVIDADE"], 15);
  if (header < 0) return [];
  const records = [];
  for (let r = header + 1; r < rows.length; r += 1) {
    const row = rows[r];
    const activity = row[0];
    if (!looksLikeActivity(activity)) continue;
    records.push({
      "C.1 - Atividade": activity,
      "C.2 - Detalhamento": row[1],
      "C.3 - Ano": row[2],
      "C.4 - Mes": row[3],
      "C.5 - Duracao": row[4],
      "Data de inicio": row[5],
      "Data de fim": row[6],
      "Mes de inicio": row[7],
      "Mes de fim": row[8],
      Linha: r + 1,
    });
  }
  return records;
}

function extractD(rows) {
  const header = findHeaderRow(rows, ["D.1", "NOME"], 12);
  if (header < 0) return [];
  const records = [];
  for (let r = header + 1; r < rows.length; r += 1) {
    const row = rows[r];
    if (isBlank(row[0])) continue;
    const record = {
      Participante: clean(row[0]),
      CPF: row[1],
      Funcao: row[2],
      Formacao: row[3],
      "Periodo informado": toInt(row[4]),
      Linha: r + 1,
    };
    for (let i = 1; i <= 30; i += 1) record[`ATIV. ${i}`] = row[4 + i];
    records.push(record);
  }
  return records;
}

function extractF(rows) {
  const header = findHeaderRow(rows, ["F.1", "UNIDADE"], 12);
  if (header < 0) return [];
  const records = [];
  for (let r = header + 1; r < rows.length; r += 1) {
    const row = rows[r];
    if (row.slice(1, 11).every(isBlank)) continue;
    records.push({
      Rubrica: row[0],
      Unidade: row[1],
      "Tipo remuneracao": row[2],
      Participante: row[3],
      Referencia: row[4],
      Justificativa: row[5],
      "Inicio periodo": row[6],
      "Termino periodo": row[7],
      Horas: row[8],
      Remuneracao: row[9],
      Encargos: row[10],
      Linha: r + 1,
    });
  }
  return records;
}

function normalizeData(extracted, alerts) {
  const { projectStart, projectEnd } = projectDates(extracted, alerts);
  const calendario = buildCalendar(projectStart, projectEnd);
  const baseC = normalizeC(extracted.cRaw, projectStart, alerts);
  const baseD = normalizeD(extracted.dRaw, alerts);
  const baseF = normalizeF(extracted.fRaw, projectStart, alerts);
  return { projectStart, projectEnd, calendario, baseC, baseD, baseF };
}

function projectDates(extracted, alerts) {
  let projectStart = null;
  let projectEnd = null;
  extracted.baseRaw.forEach((row) => {
    const key = normalizeText(row.Campo);
    if (key.includes("data de contrat") || key.includes("data de inicio")) projectStart = toDate(row.Valor);
    if (key.includes("data de conclus") || key.includes("data de encerramento")) projectEnd = toDate(row.Valor);
  });

  if (!projectStart) {
    const dates = [
      ...extracted.cRaw.map((row) => toDate(row["Data de inicio"]) || dateFromYearMonth(row["C.3 - Ano"], row["C.4 - Mes"])),
      ...extracted.fRaw.map((row) => toDate(row["Inicio periodo"])),
    ].filter(Boolean);
    if (dates.length) {
      projectStart = minDate(dates);
      alerts.push(alert("Atencao", "Datas", "Base", "Data de inicio ausente; usando menor data encontrada em C/F.", { campo: "Base" }));
    }
  }

  if (!projectEnd) {
    const dates = [];
    extracted.cRaw.forEach((row) => {
      const start = toDate(row["Data de inicio"]) || dateFromYearMonth(row["C.3 - Ano"], row["C.4 - Mes"]);
      const duration = toInt(row["C.5 - Duracao"]);
      const explicitEnd = toDate(row["Data de fim"]);
      if (explicitEnd) dates.push(explicitEnd);
      else if (start && duration !== null) dates.push(addMonths(start, duration));
    });
    extracted.fRaw.forEach((row) => {
      const end = toDate(row["Termino periodo"]);
      if (end) dates.push(end);
    });
    if (dates.length) {
      projectEnd = maxDate(dates);
      alerts.push(alert("Atencao", "Datas", "Base", "Data de conclusao ausente; usando maior data encontrada em C/F.", { campo: "Base" }));
    }
  }

  if (!projectStart || !projectEnd) {
    alerts.push(alert("Erro", "Datas", "Base", "Nao foi possivel determinar inicio/conclusao do projeto.", { campo: "Base" }));
  }
  return { projectStart, projectEnd };
}

function buildCalendar(start, end) {
  if (!start || !end || end < start) return [];
  const lastSeq = monthIndex(start, end);
  const rows = [];
  for (let seq = 0; seq <= lastSeq; seq += 1) {
    const current = addMonths(monthStart(start), seq);
    const mesAno = monthYearLabel(current);
    rows.push({
      "Mes Sequencial": seq,
      "Mes Ano": mesAno,
      Rotulo_Mes: `${String(seq).padStart(2, "0")} - ${mesAno}`,
      "Data inicio do mes": current,
    });
  }
  return rows;
}

function normalizeC(rows, projectStart, alerts) {
  if (!rows.length) alerts.push(alert("Erro", "Layout", "C", "Nenhuma atividade C encontrada.", { campo: "C" }));
  return rows
    .filter((row) => !isBlank(row["C.1 - Atividade"]))
    .map((row) => {
      const start = toDate(row["Data de inicio"]) || dateFromYearMonth(row["C.3 - Ano"], row["C.4 - Mes"]);
      const duration = toInt(row["C.5 - Duracao"]);
      const end = toDate(row["Data de fim"]) || (start && duration !== null ? addMonths(start, duration) : null);
      if (!start || !end) {
        alerts.push(alert("Erro", "Datas", "C", "Atividade com data inicial/final invalida.", { atividade: clean(row["C.1 - Atividade"]), campo: "Data" }));
      }
      return {
        Atividade: clean(row["C.1 - Atividade"]),
        Detalhamento: row["C.2 - Detalhamento"],
        Ano: toInt(row["C.3 - Ano"]),
        Mes: row["C.4 - Mes"],
        Duracao: duration,
        "Data de inicio": start,
        "Data de fim": end,
        "Mes de inicio": projectStart && start ? monthIndex(projectStart, start) : toInt(row["Mes de inicio"]),
        "Mes de fim": projectStart && end ? monthIndex(projectStart, end) : toInt(row["Mes de fim"]),
        Atividade_Ordenada: activitySortName(row["C.1 - Atividade"]),
      };
    });
}

function normalizeD(rows, alerts) {
  if (!rows.length) alerts.push(alert("Erro", "Layout", "D", "Nenhum participante D encontrado.", { campo: "D" }));
  return rows.map((row) => {
    if (isBlank(row.CPF)) {
      alerts.push(alert("Erro", "Dados obrigatorios", "D", "Participante sem CPF informado.", { participante: row.Participante, campo: "CPF" }));
    }
    return row;
  });
}

function normalizeF(rows, projectStart, alerts) {
  if (!rows.length) alerts.push(alert("Erro", "Layout", "F", "Nenhuma linha F valida encontrada.", { campo: "F" }));
  return rows
    .filter((row) => !isBlank(row.Participante))
    .map((row) => {
      const start = toDate(row["Inicio periodo"]);
      const end = toDate(row["Termino periodo"]);
      const hours = toNumber(row.Horas);
      const pay = toNumber(row.Remuneracao);
      const charges = toNumber(row.Encargos);
      const startSeq = projectStart && start ? monthIndex(projectStart, start) : null;
      const endSeq = projectStart && end ? monthIndex(projectStart, end) : null;
      const months = inclusiveMonths(startSeq, endSeq).length || null;
      if (!start || !end || end < start) {
        alerts.push(alert("Erro", "Datas", "F", "Linha remunerada com datas invalidas.", { participante: clean(row.Participante), campo: "F.6/F.7" }));
      }
      return {
        Rubrica: row.Rubrica,
        Unidade: row.Unidade,
        "Tipo remuneracao": row["Tipo remuneracao"],
        Participante: clean(row.Participante),
        Referencia: row.Referencia,
        Justificativa: row.Justificativa,
        "Inicio periodo": start,
        "Termino periodo": end,
        Horas: hours,
        Remuneracao: pay,
        Encargos: charges,
        "TOTAL (REMUN. + ENCARGOS)": (pay || 0) + (charges || 0),
        "Meses recebidos": months,
        HH: safeDiv(pay, hours),
        "Salario mensal": safeDiv(pay, months),
        "Horas trabalhadas/mes": safeDiv(hours, months),
        "% de encargos": safeDiv(charges, pay),
        "Mes de inicio": startSeq,
        "Mes de fim": endSeq,
      };
    });
}

function transformData(data, alerts) {
  const intC = [];
  const calendarBySeq = new Map(data.calendario.map((row) => [row["Mes Sequencial"], row]));
  data.baseC.forEach((activity) => {
    inclusiveMonths(activity["Mes de inicio"], activity["Mes de fim"]).forEach((seq) => {
      const cal = calendarBySeq.get(seq);
      if (!cal) return;
      intC.push({ Atividade: activity.Atividade, Atividade_Ordenada: activity.Atividade_Ordenada, ...cal });
    });
  });

  const intDPart = [];
  data.baseD.forEach((person) => {
    for (let i = 1; i <= 30; i += 1) {
      if (!isBlank(person[`ATIV. ${i}`])) {
        intDPart.push({ ...person, Atividade: `ATIV. ${i}` });
      }
    }
  });

  const activityMonths = groupBy(intC, "Atividade");
  const intD = [];
  intDPart.forEach((row) => {
    const months = activityMonths.get(row.Atividade) || [];
    if (!months.length) {
      alerts.push(alert("Atencao", "Atividade", "D x C", "Atividade preenchida na D nao encontrada na C.", { participante: row.Participante, atividade: row.Atividade }));
    }
    months.forEach((month) => intD.push({ ...row, Atividade_Ordenada: month.Atividade_Ordenada, ...month }));
  });

  const intF = [];
  data.baseF.forEach((row) => {
    inclusiveMonths(row["Mes de inicio"], row["Mes de fim"]).forEach((seq) => {
      const cal = calendarBySeq.get(seq);
      if (!cal) return;
      intF.push({ Participante: row.Participante, ...cal, Marcador: 1 });
    });
  });

  return {
    intC,
    intD,
    intF,
    calendarioAtividades: outputCalendarioAtividades(intC, data.calendario),
    mesesTrabalhados: outputMesesTrabalhados(intD, data.calendario),
    mesesRecebidos: outputMesesRecebidos(intF, data.calendario),
    comparacao1: outputComparacao1(data.baseD, intD, intF),
    comparacao2: outputComparacao2(intF, intD, data.calendario),
  };
}

function outputCalendarioAtividades(intC, calendar) {
  const months = orderedMonths(calendar);
  const byActivity = new Map();
  intC.forEach((row) => {
    const key = row.Atividade;
    if (!byActivity.has(key)) byActivity.set(key, { Atividades: row.Atividade, _sort: row.Atividade_Ordenada });
    byActivity.get(key)[row.Rotulo_Mes] = 1;
  });
  return Array.from(byActivity.values())
    .sort((a, b) => String(a._sort).localeCompare(String(b._sort)))
    .map((row) => completeMonthRow(removePrivate(row), months));
}

function outputMesesTrabalhados(intD, calendar) {
  const months = orderedMonths(calendar);
  const byParticipant = new Map();
  intD.forEach((row) => {
    if (!byParticipant.has(row.Participante)) byParticipant.set(row.Participante, { Participante: row.Participante });
    const current = byParticipant.get(row.Participante)[row.Rotulo_Mes] || [];
    current.push(row.Atividade);
    byParticipant.get(row.Participante)[row.Rotulo_Mes] = current;
  });
  return Array.from(byParticipant.values()).map((row) => {
    const output = { Participante: row.Participante };
    months.forEach((month) => {
      output[month] = Array.isArray(row[month]) ? combineActivities(row[month]) : "";
    });
    return output;
  });
}

function outputMesesRecebidos(intF, calendar) {
  const months = orderedMonths(calendar);
  const byParticipant = new Map();
  intF.forEach((row) => {
    if (!byParticipant.has(row.Participante)) byParticipant.set(row.Participante, { Participante: row.Participante });
    byParticipant.get(row.Participante)[row.Rotulo_Mes] = 1;
  });
  return Array.from(byParticipant.values()).map((row) => completeMonthRow(row, months));
}

function outputComparacao1(baseD, intD, intF) {
  const calculated = countDistinctMonths(intD);
  const received = countDistinctMonths(intF);
  return baseD.map((row) => {
    const participant = row.Participante;
    const informed = row["Periodo informado"];
    const calc = calculated.get(participant) || 0;
    const rec = received.get(participant) || 0;
    return {
      Participante: participant,
      "Meses de trabalho (calculado)": calc,
      "Meses de trabalho (informado)": informed,
      "Calculado >= Informado?": informed !== null && calc < informed ? "Problema" : "Ok",
      "Meses Recebidos ($)": rec,
      "Recebido <= Informado?": informed !== null && rec > informed ? "Problema" : "Ok",
    };
  });
}

function outputComparacao2(intF, intD, calendar) {
  const months = orderedMonths(calendar);
  const planned = new Set(intD.map((row) => `${row.Participante}||${row["Mes Sequencial"]}`));
  const byParticipant = new Map();
  uniqueBy(intF, (row) => `${row.Participante}||${row["Mes Sequencial"]}`).forEach((row) => {
    if (!byParticipant.has(row.Participante)) byParticipant.set(row.Participante, { Participante: row.Participante });
    const key = `${row.Participante}||${row["Mes Sequencial"]}`;
    byParticipant.get(row.Participante)[row.Rotulo_Mes] = planned.has(key) ? "Regular" : "Irregular";
  });
  return Array.from(byParticipant.values()).map((row) => completeMonthRow(row, months));
}

function addValidationAlerts(data, transformed, alerts) {
  const planned = new Set(data.baseD.map((row) => row.Participante));
  const paid = new Set(data.baseF.map((row) => row.Participante));
  paid.forEach((participant) => {
    if (!planned.has(participant)) alerts.push(alert("Erro", "Participante", "F x D", "Participante remunerado na F nao encontrado na D.", { participante: participant }));
  });
  planned.forEach((participant) => {
    if (!paid.has(participant)) alerts.push(alert("Atencao", "Participante", "D x F", "Participante planejado na D nao encontrado na F.", { participante: participant }));
  });
  transformed.comparacao2.forEach((row) => {
    Object.entries(row).forEach(([column, value]) => {
      if (column !== "Participante" && value === "Irregular") {
        alerts.push(alert("Erro", "F x D", "Comparacao 2", "Participante recebeu no mes, mas nao possui atividade planejada na D.", { participante: row.Participante, mes: column }));
      }
    });
  });
}

function buildSummary(data, transformed, alerts) {
  return {
    Atividades: data.baseC.length,
    "Participantes planejados": new Set(data.baseD.map((row) => row.Participante)).size,
    "Participantes remunerados": new Set(data.baseF.map((row) => row.Participante)).size,
    "Registros atividade x mes": transformed.intC.length,
    "Registros trabalho planejado x mes": transformed.intD.length,
    "Registros remuneracao x mes": transformed.intF.length,
    Alertas: alerts.length,
    Erros: alerts.filter((item) => item.severidade === "Erro").length,
    Atencoes: alerts.filter((item) => item.severidade === "Atencao").length,
  };
}

function renderResult(result) {
  summaryPanel.hidden = false;
  summaryPanel.innerHTML = Object.entries(result.summary)
    .map(([label, value]) => `<article class="summary-card"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></article>`)
    .join("");

  downloadPanel.hidden = false;
  tabs.hidden = false;
  tabs.innerHTML = Object.keys(result.outputs)
    .map((name, index) => `<button class="tab-button${index === 0 ? " active" : ""}" type="button" data-tab="${escapeHtml(name)}">${escapeHtml(name)}</button>`)
    .join("");
  tabs.querySelectorAll("button").forEach((button) => button.addEventListener("click", () => selectTab(button.dataset.tab)));
  selectTab(Object.keys(result.outputs)[0]);
}

function selectTab(name) {
  tabs.querySelectorAll("button").forEach((button) => button.classList.toggle("active", button.dataset.tab === name));
  tablePanel.hidden = false;
  tablePanel.innerHTML = renderTable(state.result.outputs[name] || []);
}

function renderTable(rows) {
  if (!rows.length) return "<p>Nenhum registro.</p>";
  const columns = Object.keys(rows[0]);
  const body = rows
    .slice(0, 250)
    .map((row) => `<tr class="${escapeHtml(row.Severidade || "")}">${columns.map((column) => `<td>${formatDisplay(row[column])}</td>`).join("")}</tr>`)
    .join("");
  return `<table><thead><tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr></thead><tbody>${body}</tbody></table>`;
}

function downloadProcessedWorkbook() {
  if (!state.result) return;
  const workbook = XLSX.utils.book_new();
  Object.entries(state.result.outputs).forEach(([name, rows]) => {
    const worksheet = XLSX.utils.json_to_sheet(rows.map(serializeRow));
    XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName(name));
  });
  XLSX.writeFile(workbook, `controle_hh_processado_${timestamp()}.xlsx`);
}

function findSheet(workbook, aliases) {
  const byName = new Map(workbook.SheetNames.map((name) => [normalizeText(name), workbook.Sheets[name]]));
  for (const alias of aliases) {
    const sheet = byName.get(normalizeText(alias));
    if (sheet) return sheet;
  }
  return null;
}

function sheetRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
}

function findHeaderRow(rows, tokens, maxRows) {
  const wanted = tokens.map(normalizeText);
  for (let i = 0; i < Math.min(rows.length, maxRows); i += 1) {
    const text = normalizeText(rows[i].join(" "));
    if (wanted.every((token) => text.includes(token))) return i;
  }
  return -1;
}

function alert(severidade, categoria, origem, mensagem, extra = {}) {
  return { severidade, categoria, origem, mensagem, ...extra };
}

function alertRecord(item) {
  return {
    Severidade: item.severidade,
    Categoria: item.categoria,
    Origem: item.origem,
    Participante: item.participante || "",
    Atividade: item.atividade || "",
    Mes: item.mes || "",
    Campo: item.campo || "",
    Mensagem: item.mensagem,
  };
}

function objectToRows(object) {
  return Object.entries(object).map(([Indicador, Valor]) => ({ Indicador, Valor }));
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function clean(value) {
  return String(value ?? "").trim();
}

function firstNonBlank(values) {
  return values.find((value) => !isBlank(value)) ?? "";
}

function toNumber(value) {
  if (isBlank(value)) return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  let text = String(value).replace("R$", "").replace("%", "").trim();
  if (text.includes(",") && text.includes(".")) text = text.replace(/\./g, "").replace(",", ".");
  else text = text.replace(",", ".");
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}

function toInt(value) {
  const number = toNumber(value);
  return number === null ? null : Math.trunc(number);
}

function toDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return dateOnly(value);
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  if (isBlank(value)) return null;
  const text = String(value).trim();
  const br = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (br) return new Date(Number(br[3]), Number(br[2]) - 1, Number(br[1]));
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  const fallback = new Date(text);
  return Number.isNaN(fallback.getTime()) ? null : dateOnly(fallback);
}

function dateOnly(value) {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate());
}

function dateFromYearMonth(yearValue, monthValue) {
  const year = toInt(yearValue);
  const month = monthNumber(monthValue);
  if (!year || !month) return null;
  return new Date(year, month - 1, 1);
}

function monthNumber(value) {
  const number = toInt(value);
  if (number && number >= 1 && number <= 12) return number;
  const key = normalizeText(value);
  const index = MONTHS.findIndex(([full, short]) => normalizeText(full) === key || normalizeText(short) === key);
  return index >= 0 ? index + 1 : null;
}

function monthStart(value) {
  return new Date(value.getFullYear(), value.getMonth(), 1);
}

function addMonths(value, count) {
  return new Date(value.getFullYear(), value.getMonth() + count, 1);
}

function monthIndex(projectStart, value) {
  const start = monthStart(projectStart);
  const current = monthStart(value);
  return (current.getFullYear() - start.getFullYear()) * 12 + current.getMonth() - start.getMonth();
}

function monthYearLabel(value) {
  return `${MONTHS[value.getMonth()][1]}/${String(value.getFullYear()).slice(-2)}`;
}

function inclusiveMonths(start, end) {
  if (start === null || start === undefined || end === null || end === undefined || end < start) return [];
  return Array.from({ length: end - start + 1 }, (_, index) => start + index);
}

function activitySortName(value) {
  const match = clean(value).match(/(\d+)/);
  return match ? `ATIV. ${String(Number(match[1])).padStart(2, "0")}` : clean(value);
}

function activityShort(value) {
  const match = clean(value).match(/(\d+)/);
  return match ? String(Number(match[1])) : clean(value);
}

function combineActivities(values) {
  const numbers = values.map(activityShort).map(Number).filter(Number.isFinite);
  return [...new Set(numbers)].sort((a, b) => a - b).join(", ");
}

function safeDiv(left, right) {
  return left === null || right === null || right === 0 ? null : left / right;
}

function orderedMonths(calendar) {
  return [...calendar].sort((a, b) => a["Mes Sequencial"] - b["Mes Sequencial"]).map((row) => row.Rotulo_Mes);
}

function completeMonthRow(row, months) {
  const output = { ...row };
  months.forEach((month) => {
    if (!(month in output)) output[month] = "";
  });
  return output;
}

function countDistinctMonths(rows) {
  const map = new Map();
  rows.forEach((row) => {
    if (!map.has(row.Participante)) map.set(row.Participante, new Set());
    map.get(row.Participante).add(row["Mes Sequencial"]);
  });
  return new Map(Array.from(map.entries()).map(([key, value]) => [key, value.size]));
}

function groupBy(rows, key) {
  const map = new Map();
  rows.forEach((row) => {
    const value = row[key];
    if (!map.has(value)) map.set(value, []);
    map.get(value).push(row);
  });
  return map;
}

function uniqueBy(rows, keyFn) {
  const seen = new Set();
  return rows.filter((row) => {
    const key = keyFn(row);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function removePrivate(row) {
  return Object.fromEntries(Object.entries(row).filter(([key]) => !key.startsWith("_")));
}

function minDate(values) {
  return values.reduce((min, value) => (value < min ? value : min), values[0]);
}

function maxDate(values) {
  return values.reduce((max, value) => (value > max ? value : max), values[0]);
}

function looksLikeActivity(value) {
  const text = normalizeText(value);
  return text.startsWith("ativ.") || text.startsWith("ativ ");
}

function serializeRow(row) {
  return Object.fromEntries(Object.entries(row).map(([key, value]) => [key, value instanceof Date ? value : value ?? ""]));
}

function safeSheetName(name) {
  return name.replace(/[\[\]:*?/\\]/g, "_").slice(0, 31);
}

function timestamp() {
  const now = new Date();
  return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
}

function formatDisplay(value) {
  if (value instanceof Date) return escapeHtml(value.toLocaleDateString("pt-BR"));
  return escapeHtml(value ?? "");
}

function clearResults() {
  summaryPanel.hidden = true;
  downloadPanel.hidden = true;
  tabs.hidden = true;
  tablePanel.hidden = true;
}

function setStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.style.color = isError ? "var(--red)" : "var(--muted)";
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

