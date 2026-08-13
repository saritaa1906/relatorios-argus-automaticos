import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { buildReport, reportName } from "../src/report.mjs";

test("gera Excel valido, com 14 dias e sem xl/tables", async () => {
  const temp = await fs.mkdtemp(path.join(os.tmpdir(), "argus-report-"));
  const output = path.join(temp, "relatorio.xlsx");
  const rows = [
    { idLigacao: 1, dataHoraLigacao: "2026-08-01T10:00:00-03:00", nrLead: 101, resultadoLigacao: "ATENDIMENTO", tabulacao: "CONTRATO FECHADO", telefone: "31999999999" },
    { idLigacao: 2, dataHoraLigacao: "2026-08-14T11:00:00-03:00", nrLead: 102, resultadoLigacao: "NAO ATENDE", tabulacao: "CLIENTE AUSENTE", telefone: "31988888888" },
  ];
  await buildReport({ mailing: { idLote: 4210, loteDesc: "PE 4 REFIN INSS C6 4210" }, rows, start: "2026-08-01", end: "2026-08-14", outputPath: output });
  assert.ok((await fs.stat(output)).size > 10_000);
  const zip = await JSZip.loadAsync(await fs.readFile(output));
  const listing = Object.keys(zip.files).join("\n");
  assert.doesNotMatch(listing, /xl\/tables\//);
  assert.match(listing, /xl\/worksheets\/sheet1.xml/);
  assert.equal(reportName({ loteDesc: "PE 4 REFIN INSS C6 4210" }, "2026-08-01", "2026-08-14"), "Relatorio Argus - PE 4 REFIN INSS C6 4210 - 01-08 a 14-08-2026.xlsx");
});
