const BASE = "https://argus.app.br/apiargus";

async function request(endpoint, token, body) {
  const response = await fetch(`${BASE}${endpoint}`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "Token-Signature": token },
    body: JSON.stringify(body),
    signal: AbortSignal.timeout(120_000),
  });
  if (!response.ok) throw new Error(`Argus HTTP ${response.status} em ${endpoint}`);
  const data = await response.json();
  if (Number(data.codStatus) !== 1) throw new Error(`Argus: ${data.descStatus || "falha desconhecida"}`);
  return data;
}

export async function listMailings(token) {
  const data = await request("/report/mailingsdiscador", token, { idCampanha: 1 });
  return data.mailings || [];
}

async function fetchWindow(token, idLote, start, end) {
  const rows = [];
  let ultimoId = 0;
  do {
    const body = { idCampanha: 1, idLote, periodoInicial: start, periodoFinal: end };
    if (ultimoId) body.ultimoId = ultimoId;
    const data = await request("/report/ligacoesdetalhadas", token, body);
    rows.push(...(data.ligacoesDetalhadas || []));
    ultimoId = Number(data.idProxPagina || 0);
    if (data.endOfTable || ultimoId === 0) break;
  } while (true);
  return rows;
}

function isoDate(date) {
  return date.toISOString().slice(0, 10);
}

export async function fetchFourteenDays(token, mailing) {
  const start = new Date(`${mailing.loteInicio.slice(0, 10)}T12:00:00Z`);
  const d7 = new Date(start); d7.setUTCDate(d7.getUTCDate() + 6);
  const d8 = new Date(start); d8.setUTCDate(d8.getUTCDate() + 7);
  const end = new Date(start); end.setUTCDate(end.getUTCDate() + 13);
  const first = await fetchWindow(token, mailing.idLote, `${isoDate(start)}T00:00:00-03:00`, `${isoDate(d7)}T23:59:59-03:00`);
  const second = await fetchWindow(token, mailing.idLote, `${isoDate(d8)}T00:00:00-03:00`, `${isoDate(end)}T23:59:59-03:00`);
  return { rows: [...first, ...second], start: isoDate(start), end: isoDate(end) };
}

