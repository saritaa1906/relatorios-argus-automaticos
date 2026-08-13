import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import sharp from "sharp";

const U = {
  nao: "N\u00c3O", ligacao: "LIGA\u00c7\u00c3O", conversao: "CONVERS\u00c3O",
  credito: "CR\u00c9DITO", e: "\u00c9", tabulacao: "TABULA\u00c7\u00c3O",
};

export const RULES = new Map([
  ["CLIENTE AUSENTE", `${U.nao} CONVERSOU`],
  ["CLIENTE DESLIGOU - RETORNAR", `CONVERSOU MAS SEM ${U.conversao}`],
  [`${U.ligacao} ROBOTIZADA`, `${U.nao} CONVERSOU`],
  [`EM NEGOCIA\u00c7\u00c3O`, U.conversao],
  ["CAIXA POSTAL / MENSAGEM OPERADORA", `${U.nao} CONVERSOU`],
  [`${U.ligacao} MUDA`, `${U.nao} CONVERSOU`],
  ["CONTRATO FECHADO", U.conversao],
  [`CLIENTE ${U.nao} TEM INTERESSE`, `CONVERSOU MAS SEM ${U.conversao}`],
  [`DESLIGOU - ${U.nao} RETORNAR`, `CONVERSOU MAS SEM ${U.conversao}`],
  ["CLIENTE NEGATIVO", `${U.nao} CONVERSOU`],
  ["FALECIDO", `${U.nao} CONVERSOU`],
  [`SEM LIMITE DE ${U.credito}`, `${U.nao} CONVERSOU`],
  ["CONTRATO REFINANCIADO", `CONVERSOU MAS SEM ${U.conversao}`],
  ["IDADE SUPERIOR AO PERMITIDO", `CONVERSOU MAS SEM ${U.conversao}`],
  [`TEL. ${U.nao} ${U.e} DO CLIENTE - FINALIZAR LEAD`, `${U.nao} CONVERSOU`],
]);

const C = { navy: "0B2239", blue: "1677A5", teal: "008E8E", orange: "F59E0B", green: "18864B", red: "CF4545", pale: "F4F7FA" };
const normalize = (value) => String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\u00a0/g, " ").trim().toUpperCase();
const normalizedRules = new Map([...RULES].map(([key, value]) => [normalize(key), value]));
const fmtDate = (iso) => iso.split("-").reverse().join("/");
const pt = (value) => Number(value).toLocaleString("pt-BR");
const resultOf = (row) => row.resultadoLigacao ?? row.resultado ?? "SEM RESULTADO";
const tabOf = (row) => row.tabulacao ?? row.tabulacaoLigacao ?? row.categoriaTabulacao ?? "";
const leadOf = (row) => row.nrLead ?? row.idLead ?? row.lead ?? "";
const dateOf = (row) => row.dataHoraLigacao ?? row.envioLigacao ?? row.dataLigacao ?? "";
const groupOf = (row) => normalizedRules.get(normalize(tabOf(row))) || (tabOf(row) ? "RESULTADO FORA DAS 15 REGRAS" : "SEM EXPLICA\u00c7\u00c3O");

function header(cell, color = C.navy) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${color}` } };
  cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
}

function border(cell) {
  cell.border = { top: { style: "thin", color: { argb: "FFD9E2E8" } }, left: { style: "thin", color: { argb: "FFD9E2E8" } }, bottom: { style: "thin", color: { argb: "FFD9E2E8" } }, right: { style: "thin", color: { argb: "FFD9E2E8" } } };
}

async function chart(labels, values, title, width = 760, height = 300) {
  const max = Math.max(...values, 1);
  const gap = (width - 60) / Math.max(values.length, 1);
  const bars = values.map((value, index) => {
    const h = Math.max(2, (value / max) * (height - 100));
    const x = 45 + index * gap + gap * 0.18;
    const y = height - 48 - h;
    return `<rect x="${x}" y="${y}" width="${gap * 0.64}" height="${h}" rx="4" fill="#1677A5"/><text x="${x + gap * 0.32}" y="${y - 7}" text-anchor="middle" font-family="Arial" font-size="12" fill="#233746">${value}</text><text x="${x + gap * 0.32}" y="${height - 27}" text-anchor="middle" font-family="Arial" font-size="11" fill="#52606D">${labels[index]}</text>`;
  }).join("");
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}"><rect width="100%" height="100%" fill="white"/><text x="${width / 2}" y="27" text-anchor="middle" font-family="Arial" font-size="18" font-weight="bold" fill="#172B4D">${title}</text><line x1="35" y1="${height - 48}" x2="${width - 15}" y2="${height - 48}" stroke="#D8DEE4"/>${bars}</svg>`;
  return sharp(Buffer.from(svg)).png().toBuffer();
}

export function shortName(name) {
  return String(name || "CARTEIRA").replace(/[^\p{L}\p{N} ._-]/gu, " ").replace(/\s+/g, " ").trim().split(" ").slice(0, 7).join(" ");
}

export function reportName(mailing, start, end) {
  const s = start.split("-"); const e = end.split("-");
  return `Relatorio Argus - ${shortName(mailing.loteDesc ?? mailing.nomeLote ?? mailing.lote)} - ${s[2]}-${s[1]} a ${e[2]}-${e[1]}-${e[0]}.xlsx`;
}

export async function buildReport({ mailing, rows, start, end, outputPath }) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Automacao Argus"; wb.created = new Date(); wb.calcProperties.fullCalcOnLoad = true;
  const dash = wb.addWorksheet("01 Resumo", { views: [{ showGridLines: false }] });
  const tabs = wb.addWorksheet("02 Tabulacoes", { views: [{ showGridLines: false, state: "frozen", ySplit: 1 }] });
  const results = wb.addWorksheet("03 Resultados", { views: [{ showGridLines: false, state: "frozen", ySplit: 1 }] });
  const base = wb.addWorksheet("04 Base da API", { views: [{ showGridLines: false, state: "frozen", xSplit: 3, ySplit: 1 }] });
  const rules = wb.addWorksheet("05 Regras", { views: [{ showGridLines: false }] });
  const support = wb.addWorksheet("06 Apoio", { views: [{ showGridLines: false }] });

  const people = new Set(rows.map(leadOf).filter(Boolean)).size;
  const resultsNorm = rows.map((row) => normalize(resultOf(row)));
  const attended = resultsNorm.filter((value) => value.includes("ATENDIMENTO") || value.includes("ATENDIDA")).length;
  const contracts = rows.filter((row) => normalize(tabOf(row)) === "CONTRATO FECHADO").length;
  const groups = rows.map(groupOf);
  const conversations = groups.filter((value) => value === U.conversao || value === `CONVERSOU MAS SEM ${U.conversao}`).length;
  const negotiations = rows.filter((row) => normalize(tabOf(row)) === "EM NEGOCIACAO").length;

  dash.columns = Array.from({ length: 14 }, () => ({ width: 12 }));
  dash.mergeCells("A1:N3"); dash.getCell("A1").value = "RELATORIO DA CARTEIRA - RESUMO FACIL"; header(dash.getCell("A1"), C.navy); dash.getCell("A1").font = { bold: true, color: { argb: "FFFFFFFF" }, size: 22 };
  dash.mergeCells("A4:N4"); dash.getCell("A4").value = `${shortName(mailing.loteDesc ?? mailing.nomeLote)}  |  ${fmtDate(start)} a ${fmtDate(end)}  |  PERIODO COMPLETO: 14 DIAS`; header(dash.getCell("A4"), C.blue);
  const cards = [["A6:C10", "PESSOAS NA LISTA", people, C.blue], ["D6:F10", "VEZES QUE LIGAMOS", rows.length, C.teal], ["G6:I10", "LIGACOES ATENDIDAS", attended, C.orange], ["J6:L10", "CONTRATOS FECHADOS", contracts, C.green], ["M6:N10", "DIAS", 14, C.red]];
  for (const [range, title, value, color] of cards) {
    const [a, b] = range.split(":"); const startRow = Number(a.match(/\d+/)[0]); const endRow = Number(b.match(/\d+/)[0]); const startCol = a.match(/[A-Z]+/)[0]; const endCol = b.match(/[A-Z]+/)[0];
    dash.mergeCells(`${startCol}${startRow}:${endCol}${startRow + 1}`); dash.getCell(`${startCol}${startRow}`).value = title; header(dash.getCell(`${startCol}${startRow}`), color);
    dash.mergeCells(`${startCol}${startRow + 2}:${endCol}${endRow}`); const cell = dash.getCell(`${startCol}${startRow + 2}`); cell.value = value; cell.font = { bold: true, size: 22, color: { argb: "FF233746" } }; cell.alignment = { horizontal: "center", vertical: "middle" }; border(cell);
  }
  dash.mergeCells("A12:N14"); dash.getCell("A12").value = `EM UMA FRASE: Ligamos ${pt(rows.length)} vezes para ${pt(people)} pessoas. ${pt(attended)} ligacoes foram atendidas e ${pt(contracts)} contratos foram fechados.`; dash.getCell("A12").font = { bold: true, size: 14, color: { argb: "FF145C3E" } }; dash.getCell("A12").alignment = { horizontal: "center", vertical: "middle", wrapText: true }; dash.getCell("A12").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF7EF" } };

  const dailyLabels = []; const dailyValues = [];
  for (let cursor = new Date(`${start}T12:00:00Z`); cursor <= new Date(`${end}T12:00:00Z`); cursor.setUTCDate(cursor.getUTCDate() + 1)) {
    const iso = cursor.toISOString().slice(0, 10); dailyLabels.push(`${iso.slice(8, 10)}/${iso.slice(5, 7)}`); dailyValues.push(rows.filter((row) => String(dateOf(row)).slice(0, 10) === iso).length);
  }
  const outcomeMap = new Map(); for (const row of rows) { const key = String(resultOf(row)); outcomeMap.set(key, (outcomeMap.get(key) || 0) + 1); }
  const top = [...outcomeMap].sort((a, b) => b[1] - a[1]).slice(0, 7);
  const dailyImage = wb.addImage({ buffer: await chart(dailyLabels, dailyValues, "LIGACOES POR DIA"), extension: "png" }); dash.addImage(dailyImage, { tl: { col: 0, row: 15 }, ext: { width: 720, height: 290 } });
  const topImage = wb.addImage({ buffer: await chart(top.map(([name]) => normalize(name).slice(0, 12)), top.map(([, value]) => value), "PRINCIPAIS RESULTADOS"), extension: "png" }); dash.addImage(topImage, { tl: { col: 7, row: 15 }, ext: { width: 720, height: 290 } });
  const pathImage = wb.addImage({ buffer: await chart(["LIGAMOS", "ATENDERAM", "CONVERSARAM", "NEGOCIARAM", "FECHARAM"], [rows.length, attended, conversations, negotiations, contracts], "CAMINHO: LIGAMOS > ATENDERAM > FECHARAM", 980, 310), extension: "png" }); dash.addImage(pathImage, { tl: { col: 2, row: 31 }, ext: { width: 980, height: 310 } });
  dash.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1 };

  tabs.addRow([U.tabulacao, "SIGNIFICADO SIMPLES", "QUANTIDADE", "% DAS LIGACOES"]); tabs.getRow(1).eachCell((cell) => header(cell, C.blue));
  for (const [name, meaning] of RULES) tabs.addRow([name, meaning, rows.filter((row) => normalize(tabOf(row)) === normalize(name)).length, rows.length ? rows.filter((row) => normalize(tabOf(row)) === normalize(name)).length / rows.length : 0]);
  tabs.getColumn(4).numFmt = "0.0%"; tabs.columns = [{ width: 48 }, { width: 34 }, { width: 16 }, { width: 18 }];

  results.addRow(["RESULTADO DA LIGACAO", "QUANTIDADE", "% DO TOTAL"]); results.getRow(1).eachCell((cell) => header(cell, C.blue));
  for (const [name, value] of [...outcomeMap].sort((a, b) => b[1] - a[1])) results.addRow([name, value, rows.length ? value / rows.length : 0]);
  results.getColumn(3).numFmt = "0.0%"; results.columns = [{ width: 48 }, { width: 16 }, { width: 18 }];

  const apiKeys = [...new Set(rows.flatMap((row) => Object.keys(row)))];
  base.addRow([...apiKeys, "CLASSIFICACAO DO RELATORIO"]); base.getRow(1).eachCell((cell) => header(cell));
  for (let index = 0; index < rows.length; index += 1000) base.addRows(rows.slice(index, index + 1000).map((row) => [...apiKeys.map((key) => row[key] == null ? "" : typeof row[key] === "object" ? JSON.stringify(row[key]) : row[key]), groupOf(row)]));
  base.columns.forEach((column) => { column.width = 20; });

  rules.addRow([U.tabulacao, "CLASSIFICACAO"]); rules.getRow(1).eachCell((cell) => header(cell)); for (const item of RULES) rules.addRow(item); rules.columns = [{ width: 50 }, { width: 38 }];
  support.addRows([["CAMPO", "VALOR"], ["ID DO LOTE", mailing.idLote], ["CARTEIRA", mailing.loteDesc ?? mailing.nomeLote], ["INICIO", fmtDate(start)], ["FIM", fmtDate(end)], ["DIAS ANALISADOS", 14], ["REGRA DO PERIODO", "A data inicial conta como dia 1; a data final e o inicio + 13 dias."], ["JANELAS DA API", "2 periodos consecutivos de 7 dias, com toda a paginacao."]]); support.getRow(1).eachCell((cell) => header(cell)); support.columns = [{ width: 28 }, { width: 90 }];

  for (const sheet of wb.worksheets) { sheet.autoFilter = sheet.name === "04 Base da API" && apiKeys.length ? { from: { row: 1, column: 1 }, to: { row: 1, column: apiKeys.length + 1 } } : undefined; sheet.eachRow((row) => row.eachCell((cell) => { cell.font = { ...cell.font, name: "Arial" }; border(cell); })); }
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await wb.xlsx.writeFile(outputPath);
}
