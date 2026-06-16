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

const VIEW_TABS = [
  ["overview", "Visão geral", "resumo"],
  ["alerts", "Alertas", "resumo"],
  ["calendar", "Calendário", "cronogramas"],
  ["worked", "Trabalho", "cronogramas"],
  ["paid", "Recebimentos", "cronogramas"],
  ["comparisons", "Comparações", "dados"],
  ["f-treated", "F tratada", "dados"],
];

const ALERT_EXPLAINS = {
  "F x D": "Recebe sem trabalhar: participante remunerado sem atividades previstas.",
  "D x F": "Planejado sem receber: participante presente apenas na aba D, sem remuneração na F.",
  "Comparacao 2": "Recebeu em meses sem atividade planejada na D.",
  "Horas/mes": "Carga mensal acima de 176h.",
  "HH": "Valor-hora (HH) acima de R$ 250.",
  "Encargos 80": "Encargos acima de 80% da remuneração.",
  "Encargos 100": "Encargos acima de 100% da remuneração.",
};

const state = {
  msalClient: null,
  account: null,
  result: null,
  currentView: "overview",
};

const uploadForm = document.querySelector("#uploadForm");
const fileInput = document.querySelector("#fileInput");
const processButton = document.querySelector("#processButton");
const statusText = document.querySelector("#statusText");
const resultHeader = document.querySelector("#resultHeader");
const resultTitle = document.querySelector("#resultTitle");
const resultMeta = document.querySelector("#resultMeta");
const downloadButton = document.querySelector("#downloadButton");
const tabs = document.querySelector("#tabs");
const viewPanel = document.querySelector("#viewPanel");
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
    processButton.textContent = "Processando…";
    clearResults();
    setStatus("Lendo arquivo no navegador…", "busy");
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true, dense: false });
    setStatus("Processando Base, C, D e F…", "busy");
    state.result = runPipeline(workbook, file.name);
    renderResult(state.result);
    setStatus("Concluído. Nenhum dado foi enviado para servidor.");
  } catch (error) {
    console.error(error);
    setStatus(error.message || "Falha ao processar arquivo.", "error");
  } finally {
    processButton.disabled = APP_CONFIG.authRequired && !state.account;
    processButton.textContent = "Processar";
  }
}

function runPipeline(workbook, fileName) {
  const alerts = [];
  const extracted = extractWorkbook(workbook, alerts);
  const normalized = normalizeData(extracted, alerts);
  const transformed = transformData(normalized, alerts);
  addValidationAlerts(normalized, transformed, alerts);

  const groupedAlerts = groupAlerts(alerts);
  const summary = buildSummary(normalized, transformed, groupedAlerts);
  const exportOutputs = {
    Resumo: objectToRows(summary),
    Alertas: groupedAlerts.map(alertGroupRecord),
    "Calendario atividades": transformed.calendarioAtividades,
    "Meses Trabalhados": transformed.mesesTrabalhados,
    "Meses Recebidos ($)": transformed.mesesRecebidos,
    "Comparacao 1": transformed.comparacao1,
    "Comparacao 2": transformed.comparacao2,
    "F - Equipe tratada": normalized.baseF,
  };

  return {
    fileName,
    generatedAt: new Date(),
    normalized,
    transformed,
    alerts,
    groupedAlerts,
    summary,
    exportOutputs,
  };
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
    if (!planned.has(participant)) alerts.push(alert("Erro", "Participante", "F x D", "Participante remunerado na F nao encontrado na D.", { participante: participant, explainKey: "F x D" }));
  });
  planned.forEach((participant) => {
    if (!paid.has(participant)) alerts.push(alert("Atencao", "Participante", "D x F", "Participante planejado na D nao encontrado na F.", { participante: participant, explainKey: "D x F" }));
  });

  transformed.comparacao2.forEach((row) => {
    const months = Object.entries(row)
      .filter(([column, value]) => column !== "Participante" && value === "Irregular")
      .map(([column]) => column);
    if (months.length) {
      alerts.push(
        alert("Erro", "F x D", "Comparacao 2", "Participante recebeu em meses sem atividade planejada na D.", {
          participante: row.Participante,
          meses: months,
          explainKey: "Comparacao 2",
        }),
      );
    }
  });

  addPayrollAlerts(data.baseF, alerts);
}

function addPayrollAlerts(baseF, alerts) {
  baseF.forEach((row) => {
    const who = clean(row.Participante);
    if (!who) return;
    const hoursPerMonth = toNumber(row["Horas trabalhadas/mes"]);
    const hh = toNumber(row.HH);
    const charges = toNumber(row["% de encargos"]);
    if (hoursPerMonth > 176) {
      alerts.push(alert("Atencao", "Carga horária", "Horas/mes", "Carga mensal acima de 176h.", { participante: who, explainKey: "Horas/mes" }));
    }
    if (hh > 250) {
      alerts.push(alert("Atencao", "Valor-hora", "HH", "Valor-hora (HH) acima de R$ 250.", { participante: who, explainKey: "HH" }));
    }
    if (charges > 1) {
      alerts.push(alert("Erro", "Encargos", "Encargos 100", "Encargos acima de 100% da remuneração.", { participante: who, explainKey: "Encargos 100" }));
    } else if (charges > 0.8) {
      alerts.push(alert("Atencao", "Encargos", "Encargos 80", "Encargos acima de 80% da remuneração.", { participante: who, explainKey: "Encargos 80" }));
    }
  });
}

function buildSummary(data, transformed, groupedAlerts) {
  const errorCount = groupedAlerts.filter((item) => item.severidade === "Erro").length;
  const warningCount = groupedAlerts.filter((item) => item.severidade === "Atencao").length;
  return {
    Atividades: data.baseC.length,
    "Participantes planejados": new Set(data.baseD.map((row) => row.Participante)).size,
    "Participantes remunerados": new Set(data.baseF.map((row) => row.Participante)).size,
    "Registros atividade x mes": transformed.intC.length,
    "Registros trabalho planejado x mes": transformed.intD.length,
    "Registros remuneracao x mes": transformed.intF.length,
    Alertas: groupedAlerts.length,
    Erros: errorCount,
    Atencoes: warningCount,
  };
}

function groupAlerts(alerts) {
  const map = new Map();
  alerts.forEach((item) => {
    const months = item.meses || (item.mes ? [item.mes] : []);
    const key = [item.severidade, item.categoria, item.origem, item.participante || "", item.atividade || "", item.campo || "", item.mensagem].join("||");
    if (!map.has(key)) {
      map.set(key, {
        ...item,
        meses: [],
        count: 0,
      });
    }
    const group = map.get(key);
    group.count += 1;
    months.forEach((month) => {
      if (!group.meses.includes(month)) group.meses.push(month);
    });
  });
  return Array.from(map.values()).sort((a, b) => severityRank(a.severidade) - severityRank(b.severidade) || String(a.categoria).localeCompare(String(b.categoria)));
}

function renderResult(result) {
  resultHeader.hidden = false;
  resultTitle.textContent = result.summary.Erros ? "Processado com pontos críticos" : "Processamento concluído";
  resultMeta.textContent = `${result.fileName} · ${formatDateTime(result.generatedAt)} · ${result.summary.Alertas} alertas agrupados`;
  tabs.hidden = false;
  let lastGroup = null;
  tabs.innerHTML = VIEW_TABS.map(([id, label, group], index) => {
    const sep = group !== lastGroup && index > 0 ? '<span class="tab-sep" aria-hidden="true"></span>' : "";
    lastGroup = group;
    return `${sep}<button class="tab-button${index === 0 ? " active" : ""}" type="button" data-view="${id}">${escapeHtml(label)}</button>`;
  }).join("");
  tabs.querySelectorAll("button").forEach((button) => button.addEventListener("click", () => selectView(button.dataset.view)));
  state.currentView = "overview";
  selectView("overview");
}

function selectView(viewId) {
  state.currentView = viewId;
  tabs.querySelectorAll("button").forEach((button) => button.classList.toggle("active", button.dataset.view === viewId));
  viewPanel.hidden = false;
  viewPanel.innerHTML = renderView(viewId);
  wireViewInteractions(viewId);
}

function renderView(viewId) {
  if (!state.result) return "";
  if (viewId === "overview") return renderOverview();
  if (viewId === "alerts") return renderAlertsView();
  if (viewId === "calendar") return renderGanttView("Calendário de atividades", state.result.transformed.calendarioAtividades, "Atividades", "activity");
  if (viewId === "worked") return renderGanttView("Meses trabalhados", state.result.transformed.mesesTrabalhados, "Participante", "worked");
  if (viewId === "paid") return renderGanttView("Meses recebidos", state.result.transformed.mesesRecebidos, "Participante", "paid");
  if (viewId === "comparisons") return renderComparisons();
  if (viewId === "f-treated") return `<div class="view-content">${renderDataTable(state.result.normalized.baseF)}</div>`;
  return "";
}

function renderOverview() {
  const { summary, groupedAlerts } = state.result;
  const hasErrors = summary.Erros > 0;
  const topGroups = summarizeAlertGroups(groupedAlerts).slice(0, 5);
  return `
    <div class="view-content overview-layout">
      <section class="status-block">
        <p class="eyebrow">Status</p>
        <h2>${hasErrors ? "Revisão necessária" : "Pronto para revisão"}</h2>
        <div class="status-line">
          <span class="status-dot${hasErrors ? " has-errors" : ""}"></span>
          <span>${summary.Erros} críticos · ${summary.Atencoes} atenções · ${summary.Alertas} grupos</span>
        </div>
        <div class="compact-metrics">
          ${metricRow("Atividades", summary.Atividades)}
          ${metricRow("Participantes planejados", summary["Participantes planejados"])}
          ${metricRow("Participantes remunerados", summary["Participantes remunerados"])}
          ${metricRow("Trabalho x mês", summary["Registros trabalho planejado x mes"])}
          ${metricRow("Remuneração x mês", summary["Registros remuneracao x mes"])}
        </div>
      </section>
      <section class="section-block">
        <p class="eyebrow">O que revisar primeiro</p>
        <div class="alert-summary-list">
          ${
            topGroups.length
              ? topGroups.map(renderAlertSummaryGroup).join("")
              : '<div class="empty-state">Nenhum alerta agrupado encontrado.</div>'
          }
        </div>
      </section>
    </div>
  `;
}

function renderAlertsView() {
  const { summary } = state.result;
  const hasErrors = summary.Erros > 0;
  return `
    <div class="view-content">
      <div class="alert-verdict ${hasErrors ? "is-error" : summary.Atencoes ? "is-warning" : "is-ok"}">
        <span class="status-dot${hasErrors ? " has-errors" : ""}"></span>
        <div>
          <div class="alert-verdict-title">${hasErrors ? "Há pontos críticos" : summary.Atencoes ? "Pontos de atenção" : "Tudo regular"}</div>
          <div class="alert-verdict-meta">${summary.Erros} críticos · ${summary.Atencoes} atenções · ${summary.Alertas} grupos</div>
        </div>
      </div>
      <div class="filters">
        <input id="alertSearch" class="filter-input" type="search" placeholder="Buscar participante, atividade ou mensagem" />
        <select id="severityFilter" class="filter-select">
          <option value="">Todas severidades</option>
          <option value="Erro">Críticos</option>
          <option value="Atencao">Atenções</option>
        </select>
        <select id="originFilter" class="filter-select">
          <option value="">Todas origens</option>
          ${uniqueValues(state.result.groupedAlerts.map((item) => item.origem)).map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("")}
        </select>
      </div>
      <div id="alertList" class="alert-list">${renderAlertList(state.result.groupedAlerts)}</div>
    </div>
  `;
}

function renderAlertList(alerts) {
  if (!alerts.length) return '<div class="empty-state">Nenhum alerta para os filtros atuais.</div>';
  const groups = new Map();
  alerts.forEach((item) => {
    const key = `${item.categoria}||${item.origem}`;
    if (!groups.has(key)) groups.set(key, { categoria: item.categoria, origem: item.origem, explainKey: item.explainKey, items: [], errors: 0, warnings: 0 });
    const g = groups.get(key);
    g.items.push(item);
    if (item.severidade === "Erro") g.errors += 1;
    else g.warnings += 1;
  });
  return Array.from(groups.values())
    .sort((a, b) => b.errors - a.errors || b.items.length - a.items.length)
    .map(renderAlertGroup)
    .join("");
}

function renderAlertGroup(group) {
  const total = group.items.length;
  const badgeClass = group.errors ? "error" : "warning";
  const badgeText = group.errors
    ? `${group.errors} crítico(s)`
    : `${group.warnings} atenção(ões)`;
  return `
    <section class="alert-group-block">
      <button type="button" class="alert-group-toggle" aria-expanded="false">
        <span class="chev" aria-hidden="true">›</span>
        <span class="dot ${group.errors ? "error" : "warning"}"></span>
        <span class="alert-title">${escapeHtml(group.categoria)}</span>
        <span class="alert-group-spacer"></span>
        <span class="alert-meta">${escapeHtml(group.origem)}</span>
        <span class="badge ${badgeClass}">× ${total}</span>
      </button>
      <div class="alert-detail-list" hidden>
        ${group.explainKey && ALERT_EXPLAINS[group.explainKey] ? `<div class="alert-explain"><span class="alert-explain-icon" aria-hidden="true">i</span><span>${escapeHtml(ALERT_EXPLAINS[group.explainKey])}</span></div>` : ""}
        ${group.items.map(renderAlertItem).join("")}
      </div>
    </section>
  `;
}

function renderAlertSummaryGroup(group) {
  return `
    <article class="alert-group">
      <div class="alert-group-header">
        <div>
          <div class="alert-title">${escapeHtml(group.categoria)}</div>
          <div class="alert-meta">${escapeHtml(group.origem)} · ${group.total} ocorrência(s)</div>
        </div>
        <span class="badge ${group.errors ? "error" : "warning"}">${group.errors} críticos · ${group.warnings} atenções</span>
      </div>
    </article>
  `;
}

function renderAlertItem(item) {
  const who = [item.participante, item.atividade, item.campo].filter(Boolean).join(" · ") || item.mensagem;
  const months = item.meses?.length
    ? `<span class="alert-detail-months">${item.meses.map((m) => `<span class="badge">${escapeHtml(m)}</span>`).join("")}</span>`
    : "";
  const occ = item.count > 1 && !item.meses?.length ? `<span class="badge">× ${item.count}</span>` : "";
  return `
    <div class="alert-detail-row" data-search="${escapeHtml(normalizeText([item.severidade, item.categoria, item.origem, item.participante, item.atividade, item.campo, item.mensagem, (item.meses || []).join(" ")].join(" ")))}" data-severity="${escapeHtml(item.severidade)}" data-origin="${escapeHtml(item.origem)}">
      <span class="who">${escapeHtml(who)}</span>
      <span class="where">${months || occ}</span>
    </div>
  `;
}

function renderGanttView(title, rows, labelKey, type) {
  const months = monthColumns(rows);
  if (!rows.length || !months.length) {
    return `<div class="view-content"><h2>${escapeHtml(title)}</h2><div class="empty-state">Sem dados para exibir.</div></div>`;
  }
  const legend = {
    activity: ["activity", "Atividade ativa"],
    worked: ["worked", "Mês trabalhado (passe o mouse para ver atividades)"],
    paid: ["paid", "Mês recebido"],
  }[type] || ["activity", "Ativo"];
  return `
    <div class="view-content">
      <p class="eyebrow">Visualização mensal</p>
      <h2>${escapeHtml(title)}</h2>
      <div class="gantt-legend">
        <span class="key"><span class="swatch ${legend[0]}"></span>${escapeHtml(legend[1])}</span>
      </div>
      <div class="gantt-shell">
        <table class="gantt">
          <thead>
            <tr>
              <th class="label-col">${escapeHtml(labelKey)}</th>
              ${months.map((month) => `<th>${escapeHtml(shortMonthHeader(month))}<span class="month-index">${escapeHtml(monthIndexLabel(month))}</span></th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${rows.map((row) => renderGanttRow(row, months, labelKey, type)).join("")}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function monthIndexLabel(label) {
  const match = /^(\d{2}) - /.exec(label);
  return match ? String(Number(match[1])) : "";
}

function renderGanttRow(row, months, labelKey, type) {
  const active = months.map((month) => !isBlank(row[month]));
  return `
    <tr>
      <td class="label-col">${escapeHtml(row[labelKey] || "")}</td>
      ${months.map((month, i) => renderGanttCell(row, month, type, runPosition(active, i))).join("")}
    </tr>
  `;
}

function runPosition(active, i) {
  if (!active[i]) return "";
  const left = i > 0 && active[i - 1];
  const right = i < active.length - 1 && active[i + 1];
  if (left && right) return "mid";
  if (right) return "start";
  if (left) return "end";
  return "solo";
}

function renderGanttCell(row, month, type, pos) {
  const value = row[month];
  if (isBlank(value)) return '<td class="month-cell"></td>';
  const label = row.Atividades || row.Participante || "";
  const cls = `mark ${type} run-${pos}`;
  if (type === "worked") {
    const tip = `${shortMonthHeader(month)} · atividade(s) ${value}`;
    return `<td class="month-cell"><span class="${cls}" data-tip="${escapeHtml(tip)}" tabindex="0"></span></td>`;
  }
  const tip = `${label} · ${shortMonthHeader(month)}`;
  return `<td class="month-cell"><span class="${cls}" data-tip="${escapeHtml(tip)}" tabindex="0"></span></td>`;
}

function renderComparisons() {
  return `
    <div class="view-content">
      <section class="section-block">
        <p class="eyebrow">Resumo por participante</p>
        ${renderDataTable(state.result.transformed.comparacao1)}
      </section>
      <section class="section-block" style="margin-top: 16px;">
        <p class="eyebrow">Regularidade mês a mês</p>
        ${renderComparison2Gantt(state.result.transformed.comparacao2)}
      </section>
    </div>
  `;
}

function renderComparison2Gantt(rows) {
  const months = monthColumns(rows);
  if (!rows.length || !months.length) return '<div class="empty-state">Sem dados para exibir.</div>';
  return `
    <div class="gantt-toolbar">
      <div class="gantt-legend">
        <span class="key"><span class="swatch regular"></span>Regular</span>
        <span class="key"><span class="swatch irregular"></span>Irregular</span>
        <span class="info-icon" tabindex="0" data-tip="Irregular: o membro recebe e não trabalha naquele mês." aria-label="O que significa irregular">i</span>
      </div>
      <label class="irregular-toggle"><input type="checkbox" id="onlyIrregular" /> Mostrar apenas irregulares</label>
    </div>
    <div class="gantt-shell">
      <table class="gantt">
        <thead>
          <tr><th class="label-col">Participante</th>${months.map((month) => `<th>${escapeHtml(shortMonthHeader(month))}<span class="month-index">${escapeHtml(monthIndexLabel(month))}</span></th>`).join("")}</tr>
        </thead>
        <tbody>
          ${rows.map((row) => renderComparison2Row(row, months)).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderComparison2Row(row, months) {
  const hasIrregular = months.some((month) => row[month] === "Irregular");
  const irregularFlags = months.map((month) => row[month] === "Irregular");
  const regularFlags = months.map((month) => row[month] === "Regular");
  const cells = months
    .map((month, i) => {
      const value = row[month];
      if (isBlank(value)) return '<td class="month-cell"></td>';
      const flags = value === "Irregular" ? irregularFlags : regularFlags;
      const pos = runPosition(flags, i);
      const tip = `${escapeHtml(row.Participante || "")} · ${shortMonthHeader(month)} · ${value}`;
      return `<td class="month-cell"><span class="status-cell ${escapeHtml(value)} run-${pos}" data-tip="${tip}" tabindex="0"></span></td>`;
    })
    .join("");
  return `<tr data-irregular="${hasIrregular ? "1" : "0"}"><td class="label-col">${escapeHtml(row.Participante || "")}</td>${cells}</tr>`;
}

function renderDataTable(rows) {
  if (!rows.length) return '<div class="empty-state">Sem dados para exibir.</div>';
  const columns = Object.keys(rows[0]);
  const numericCols = columns.filter((c) => columnFormat(c) !== "text");
  const body = rows
    .slice(0, 300)
    .map((row) => `<tr class="${escapeHtml(row.Severidade || "")}">${columns.map((column) => `<td class="${numericCols.includes(column) ? "num" : ""}">${formatDisplay(row[column], column)}</td>`).join("")}</tr>`)
    .join("");
  return `<div class="table-shell"><table class="data-table"><thead><tr>${columns.map((column) => `<th class="${numericCols.includes(column) ? "num" : ""}">${escapeHtml(column)}</th>`).join("")}</tr></thead><tbody>${body}</tbody></table></div>`;
}

function wireViewInteractions(viewId) {
  if (viewId === "comparisons") {
    const toggle = document.querySelector("#onlyIrregular");
    if (toggle) {
      toggle.addEventListener("change", () => {
        document.querySelectorAll("tr[data-irregular]").forEach((tr) => {
          tr.hidden = toggle.checked && tr.dataset.irregular === "0";
        });
      });
    }
    return;
  }
  if (viewId !== "alerts") return;
  const search = document.querySelector("#alertSearch");
  const severity = document.querySelector("#severityFilter");
  const origin = document.querySelector("#originFilter");
  [search, severity, origin].forEach((control) => control.addEventListener("input", applyAlertFilters));
  wireAlertGroupToggles();
}

function wireAlertGroupToggles() {
  document.querySelectorAll(".alert-group-toggle").forEach((toggle) => {
    toggle.addEventListener("click", () => {
      const detail = toggle.nextElementSibling;
      const open = toggle.getAttribute("aria-expanded") === "true";
      toggle.setAttribute("aria-expanded", String(!open));
      detail.hidden = open;
    });
  });
}

function applyAlertFilters() {
  const term = normalizeText(document.querySelector("#alertSearch").value);
  const severity = document.querySelector("#severityFilter").value;
  const origin = document.querySelector("#originFilter").value;
  const filtered = state.result.groupedAlerts.filter((item) => {
    const searchable = normalizeText([item.severidade, item.categoria, item.origem, item.participante, item.atividade, item.campo, item.mensagem, (item.meses || []).join(" ")].join(" "));
    return (!term || searchable.includes(term)) && (!severity || item.severidade === severity) && (!origin || item.origem === origin);
  });
  document.querySelector("#alertList").innerHTML = renderAlertList(filtered);
  wireAlertGroupToggles();
}

function downloadProcessedWorkbook() {
  if (!state.result) return;
  const workbook = XLSX.utils.book_new();
  Object.entries(state.result.exportOutputs).forEach(([name, rows]) => {
    const worksheet = styledSheet(rows.map(serializeRow), name);
    XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName(name));
  });
  XLSX.writeFile(workbook, `controle_hh_processado_${timestamp()}.xlsx`);
}

function styledSheet(rows, sheetName) {
  const worksheet = XLSX.utils.json_to_sheet(rows.length ? rows : [{}], { cellDates: true });
  const range = worksheet["!ref"] ? XLSX.utils.decode_range(worksheet["!ref"]) : null;
  if (!range) return worksheet;
  worksheet["!autofilter"] = { ref: worksheet["!ref"] };
  worksheet["!freeze"] = { xSplit: 0, ySplit: 1 };
  worksheet["!cols"] = columnWidths(rows, range);

  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      if (!worksheet[address]) continue;
      const header = row === 0;
      worksheet[address].s = cellStyle(worksheet[address].v, header, sheetName, col);
    }
  }
  return worksheet;
}

function cellStyle(value, header, sheetName, col) {
  const border = {
    top: { style: "thin", color: { rgb: "E3E9F0" } },
    bottom: { style: "thin", color: { rgb: "E3E9F0" } },
    left: { style: "thin", color: { rgb: "E3E9F0" } },
    right: { style: "thin", color: { rgb: "E3E9F0" } },
  };
  if (header) {
    return {
      font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
      fill: { fgColor: { rgb: "2B6CB0" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border,
    };
  }
  const monthCol = isMonthSheet(sheetName) && col > 0;
  const style = {
    alignment: { vertical: "top", wrapText: true, horizontal: monthCol ? "center" : "left" },
    border,
  };
  if (value === "Erro" || value === "Problema" || value === "Irregular") style.fill = { fgColor: { rgb: "FDECEA" } };
  if (value === "Atencao") style.fill = { fgColor: { rgb: "FDF2D8" } };
  if (value === "Ok" || value === "Regular") style.fill = { fgColor: { rgb: "E1F3EA" } };
  if (monthCol && !isBlank(value)) style.fill = { fgColor: { rgb: sheetName.includes("Recebidos") ? "E1F3EA" : "E7F0FB" } };
  return style;
}

function columnWidths(rows, range) {
  const widths = [];
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    let max = 10;
    rows.slice(0, 300).forEach((row) => {
      const value = Object.values(row)[col];
      max = Math.max(max, Math.min(36, String(value ?? "").length + 2));
    });
    widths.push({ wch: max });
  }
  return widths;
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

function alertGroupRecord(item) {
  return {
    Severidade: item.severidade,
    Categoria: item.categoria,
    Origem: item.origem,
    Participante: item.participante || "",
    Atividade: item.atividade || "",
    Meses: item.meses?.join(", ") || "",
    Campo: item.campo || "",
    Ocorrencias: item.count || 1,
    Mensagem: item.mensagem,
  };
}

function objectToRows(object) {
  return Object.entries(object).map(([Indicador, Valor]) => ({ Indicador, Valor }));
}

function summarizeAlertGroups(alerts) {
  const map = new Map();
  alerts.forEach((item) => {
    const key = `${item.categoria}||${item.origem}`;
    if (!map.has(key)) map.set(key, { categoria: item.categoria, origem: item.origem, errors: 0, warnings: 0, total: 0 });
    const group = map.get(key);
    group.total += 1;
    if (item.severidade === "Erro") group.errors += 1;
    else group.warnings += 1;
  });
  return Array.from(map.values()).sort((a, b) => b.errors - a.errors || b.total - a.total);
}

function metricRow(label, value) {
  return `<div class="metric-row"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`;
}

function monthColumns(rows) {
  if (state.result?.normalized?.calendario?.length) {
    return orderedMonths(state.result.normalized.calendario);
  }
  if (!rows.length) return [];
  return Object.keys(rows[0])
    .filter(isMonthColumn)
    .sort((a, b) => monthColumnIndex(a) - monthColumnIndex(b));
}

function shortMonthHeader(label) {
  return label.replace(/^\d{2} - /, "");
}

function uniqueValues(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b)));
}

function severityRank(value) {
  return value === "Erro" ? 0 : value === "Atencao" ? 1 : 2;
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
  const output = {};
  Object.entries(row).forEach(([key, value]) => {
    if (!isMonthColumn(key)) output[key] = value;
  });
  months.forEach((month) => {
    output[month] = row[month] ?? "";
  });
  return output;
}

function isMonthColumn(key) {
  return /^\d{2} - /.test(key);
}

function monthColumnIndex(label) {
  const match = /^(\d{2}) - /.exec(label);
  return match ? Number(match[1]) : Number.MAX_SAFE_INTEGER;
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

function isMonthSheet(name) {
  return ["Calendario atividades", "Meses Trabalhados", "Meses Recebidos ($)"].includes(name);
}

function timestamp() {
  const now = new Date();
  return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
}

function formatDateTime(value) {
  return value.toLocaleString("pt-BR", { dateStyle: "short", timeStyle: "short" });
}

function columnFormat(column) {
  const c = normalizeText(column);
  if (/(remunerac|salario|total|valor|encargos\b|hh)/.test(c) && !c.includes("% de")) return "money";
  if (c.includes("% de") || c.includes("percent")) return "percent";
  if (c.includes("hora")) return "hours";
  if (/(meses|mes de|linha)/.test(c)) return "int";
  return "text";
}

function formatDisplay(value, column) {
  if (value instanceof Date) return escapeHtml(value.toLocaleDateString("pt-BR"));
  if (column && typeof value === "number" && Number.isFinite(value)) {
    const kind = columnFormat(column);
    if (kind === "money") return escapeHtml(value.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
    if (kind === "percent") return escapeHtml(`${(value * 100).toLocaleString("pt-BR", { maximumFractionDigits: 1 })}%`);
    if (kind === "hours") return escapeHtml(value.toLocaleString("pt-BR", { maximumFractionDigits: 1 }));
    if (kind === "int") return escapeHtml(String(Math.round(value)));
  }
  return escapeHtml(value ?? "");
}

function clearResults() {
  resultHeader.hidden = true;
  tabs.hidden = true;
  viewPanel.hidden = true;
}

function setStatus(message, kind = "idle") {
  statusText.textContent = message;
  statusText.classList.remove("is-error", "is-busy");
  if (kind === true || kind === "error") statusText.classList.add("is-error");
  else if (kind === "busy") statusText.classList.add("is-busy");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
