// Sincroniza el calendario de la FBPA con la lista Calendario del portal y genera los .ics públicos.
// Permiso Sites.Selected (write) concedido sobre PortalCBP el 2026-11-08.
import { writeFileSync, mkdirSync } from 'node:fs';
import { leerPartidos } from './fbpa.mjs';

const CFG = {
  fbpa: 'https://www.fbpa.es/resultados-club-7954/cb-pumarin',
  tenant: process.env.TENANT_ID || '92d41dc0-249a-4073-bd9e-d966005c3a94',
  clientId: process.env.CLIENT_ID || '17a79756-e595-41f7-a55e-2d616c4a9d6c',
  site: 'gijoncbp.sharepoint.com,fbb8e2b4-653c-490e-8463-9062314c2408,429a1159-19a6-4d88-bfb6-d7f6ed32f48c',
  listaCalendario: '86cc3ca3-467f-43ca-8b33-eb994839eb5f',
  listaAvisos: 'e451e830-802b-4d55-a207-f38b7d0d9e29',
  seco: process.argv.includes('--seco')
};

const COMP = (c) => c.replace(/ - ÚNICO$| - UNICO$/i, '').replace(/LIGA REGULAR/i, 'liga regular').replace(/1ª FASE/i, '1ª fase');

/* ---------- Hora de Madrid -> UTC ---------- */
function madridUtc({ y, m, d, h, mi }) {
  const guess = Date.UTC(y, m - 1, d, h, mi);
  const parts = new Intl.DateTimeFormat('en-GB', { timeZone: 'Europe/Madrid', hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })
    .formatToParts(new Date(guess)).reduce((a, p) => ((a[p.type] = p.value), a), {});
  const asMadrid = Date.UTC(+parts.year, +parts.month - 1, +parts.day, +parts.hour, +parts.minute);
  return new Date(guess - (asMadrid - guess));
}

/* ---------- Microsoft Graph con identidad federada de GitHub ---------- */
async function token() {
  const u = process.env.ACTIONS_ID_TOKEN_REQUEST_URL; const t = process.env.ACTIONS_ID_TOKEN_REQUEST_TOKEN;
  if (!u || !t || !CFG.clientId) { throw new Error('Faltan credenciales federadas (ejecutar desde GitHub Actions con CLIENT_ID)'); }
  const oidc = await fetch(u + '&audience=api://AzureADTokenExchange', { headers: { Authorization: 'bearer ' + t } }).then((r) => r.json());
  const body = new URLSearchParams({ client_id: CFG.clientId, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials',
    client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer', client_assertion: oidc.value });
  const r = await fetch(`https://login.microsoftonline.com/${CFG.tenant}/oauth2/v2.0/token`, { method: 'POST', body });
  const j = await r.json();
  if (!j.access_token) { throw new Error('No se obtuvo token de Microsoft: ' + JSON.stringify(j).slice(0, 300)); }
  return j.access_token;
}
let TK = '';
async function graph(path, opt = {}) {
  const url = path.startsWith('http') ? path : 'https://graph.microsoft.com/v1.0' + path;
  for (let i = 0; i < 4; i++) {
    const r = await fetch(url, { ...opt, headers: { Authorization: 'Bearer ' + TK, 'Content-Type': 'application/json', ...(opt.headers || {}) } });
    if (r.status === 429 || r.status >= 500) { await new Promise((s) => setTimeout(s, (+(r.headers.get('retry-after')) || 3) * 1000)); continue; }
    const txt = await r.text();
    if (!r.ok) { throw new Error(`Graph ${r.status} ${path}: ${txt.slice(0, 300)}`); }
    return txt ? JSON.parse(txt) : null;
  }
  throw new Error('Graph sin respuesta: ' + path);
}
async function todosLosElementos(lista, campos) {
  let url = `/sites/${CFG.site}/lists/${lista}/items?$expand=fields($select=${campos})&$top=500`;
  const out = [];
  while (url) { const j = await graph(url); out.push(...j.value); url = j['@odata.nextLink'] || null; }
  return out;
}

/* ---------- iCalendar ---------- */
const icsF = (d) => d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
const icsT = (s) => String(s || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\n/g, '\\n');
function ics(nombre, evs) {
  const L = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Gijon CBP//Sync//ES', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH', 'X-WR-CALNAME:' + icsT(nombre), 'X-WR-TIMEZONE:Europe/Madrid', 'REFRESH-INTERVAL;VALUE=DURATION:PT6H', 'X-PUBLISHED-TTL:PT6H'];
  const st = icsF(new Date());
  for (const e of evs) {
    const ini = new Date(e.Fecha); const fin = new Date(ini.getTime() + (e.TipoEvento === 'Partido' ? 105 : 90) * 60000);
    const local = /\(local\)|· Local/.test((e.Title || '') + ' ' + (e.Notas || ''));
    const res = (e.Notas || '').match(/Resultado: (\d+)-(\d+)/);
    const sum = e.TipoEvento === 'Partido' ? `${e.Equipo} ${local ? 'vs' : 'en'} ${e.Rival}${res ? ` (${res[1]}-${res[2]})` : ''}` : e.Title;
    // Las notas internas no se publican: solo competición y jornada
    const desc = e.TipoEvento === 'Partido' ? (e.Notas || '').split(' · Resultado')[0].replace(/ · Fuente: FBPA/, '') : '';
    L.push('BEGIN:VEVENT', `UID:cbp-${e.id}@gijoncbp`, 'DTSTAMP:' + st, 'DTSTART:' + icsF(ini), 'DTEND:' + icsF(fin), 'SUMMARY:' + icsT(sum), 'LOCATION:' + icsT(e.Lugar), 'DESCRIPTION:' + icsT(desc), 'END:VEVENT');
  }
  L.push('END:VCALENDAR');
  return L.join('\r\n');
}
const slug = (s) => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');

/* ---------- Principal ---------- */
async function main() {
  const html = await fetch(CFG.fbpa, { headers: { 'User-Agent': 'GijonCBP-Portal-Sync/1.0 (+https://cbpumarin.com)' } }).then((r) => { if (!r.ok) { throw new Error('FBPA ' + r.status); } return r.text(); });
  const partidos = leerPartidos(html);
  console.log(`FBPA: ${partidos.length} partidos del club`);
  if (partidos.length < 50) { throw new Error('Muy pocos partidos leídos: la página de la FBPA puede haber cambiado. No se toca nada.'); }
  if (CFG.seco) { console.log(JSON.stringify(partidos.slice(0, 3), null, 2)); return; }

  TK = await token();
  const campos = 'Title,Fecha,Equipo,TipoEvento,Rival,Lugar,Notas,FbpaKey';
  const items = await todosLosElementos(CFG.listaCalendario, campos);
  console.log(`Portal: ${items.length} eventos en Calendario`);
  const porClave = new Map(); const sinClave = [];
  for (const it of items) { if (it.fields.FbpaKey) { porClave.set(it.fields.FbpaKey, it); } else if (it.fields.TipoEvento === 'Partido') { sinClave.push(it); } }

  let nuevos = 0, cambios = 0; const avisos = [];
  const ahora = Date.now();
  for (const p of partidos) {
    const clave = `${p.equipo}|${p.competicion}|J${p.jornada}`;
    const fecha = p.fecha ? madridUtc(p.fecha).toISOString() : null;
    let notas = `${COMP(p.competicion)} · Jornada ${p.jornada} · ${p.local ? 'Local' : 'Visitante'}`;
    if (p.pf !== null && p.pa !== null) { notas += ` · Resultado: ${p.pf}-${p.pa} (${p.pf > p.pa ? 'victoria' : 'derrota'})`; }
    notas += ' · Fuente: FBPA';
    const lugar = p.campo || (p.local ? 'Por confirmar (local)' : 'Por confirmar (pista del rival)');
    const campos = { Title: `${p.equipo} · ${p.rival} (${p.local ? 'local' : 'visitante'})`.slice(0, 255), Equipo: p.equipo, TipoEvento: 'Partido', Rival: p.rival.slice(0, 255), Lugar: lugar.slice(0, 255), Notas: notas, FbpaKey: clave };
    if (fecha) { campos.Fecha = fecha; }

    let it = porClave.get(clave);
    if (!it) {
      const i = sinClave.findIndex((x) => x.fields.Equipo === p.equipo && (x.fields.Notas || '').includes(`Jornada ${p.jornada} `) && (x.fields.Rival || '') === p.rival);
      if (i > -1) { it = sinClave.splice(i, 1)[0]; }
    }
    if (!it) {
      await graph(`/sites/${CFG.site}/lists/${CFG.listaCalendario}/items`, { method: 'POST', body: JSON.stringify({ fields: campos }) });
      nuevos++; continue;
    }
    const f = it.fields; const dif = {};
    for (const k of Object.keys(campos)) {
      if (k === 'Fecha') { if (!f.Fecha || new Date(f.Fecha).getTime() !== new Date(campos.Fecha).getTime()) { dif.Fecha = campos.Fecha; } }
      else if ((f[k] || '') !== campos[k]) { dif[k] = campos[k]; }
    }
    if (!Object.keys(dif).length) { continue; }
    await graph(`/sites/${CFG.site}/lists/${CFG.listaCalendario}/items/${it.id}/fields`, { method: 'PATCH', body: JSON.stringify(dif) });
    cambios++;
    const cambioReal = (dif.Fecha && f.Fecha) || (dif.Lugar && f.Lugar && !/Por confirmar/.test(dif.Lugar));
    const t = new Date(campos.Fecha || f.Fecha).getTime();
    if (cambioReal && t > ahora && t - ahora < 10 * 86400000) {
      const d = new Date(t).toLocaleString('es-ES', { timeZone: 'Europe/Madrid', weekday: 'long', day: 'numeric', month: 'long', hour: '2-digit', minute: '2-digit' });
      avisos.push({ Title: `Cambio: ${p.equipo} vs ${p.rival}`.slice(0, 255), Detalle: `La FBPA ha cambiado el partido. Nuevo horario: ${d} en ${lugar}.`, Prioridad: 'Importante', FechaLimite: new Date(t).toISOString() });
    }
  }
  for (const a of avisos) { await graph(`/sites/${CFG.site}/lists/${CFG.listaAvisos}/items`, { method: 'POST', body: JSON.stringify({ fields: a }) }); }
  console.log(`Cambios: ${nuevos} nuevos, ${cambios} actualizados, ${avisos.length} avisos de cambio`);

  // Calendarios .ics públicos (solo partidos, torneos e hitos; nunca reuniones ni notas internas)
  const finales = (await todosLosElementos(CFG.listaCalendario, campos)).map((x) => ({ id: x.id, ...x.fields }))
    .filter((e) => e.Fecha && ['Partido', 'Torneo', 'Hito'].includes(e.TipoEvento)).sort((a, b) => new Date(a.Fecha) - new Date(b.Fecha));
  mkdirSync('docs/ics', { recursive: true });
  const equipos = [...new Set(finales.map((e) => e.Equipo).filter((e) => e && e !== 'Todo el club'))].sort();
  writeFileSync('docs/ics/todo-el-club.ics', ics('Gijón CBP – Todo el club', finales));
  for (const eq of equipos) { writeFileSync(`docs/ics/${slug(eq)}.ics`, ics('Gijón CBP – ' + eq, finales.filter((e) => e.Equipo === eq || e.Equipo === 'Todo el club'))); }
  writeFileSync('docs/equipos.json', JSON.stringify(equipos.map((e) => ({ equipo: e, archivo: `ics/${slug(e)}.ics` })), null, 2));
  writeFileSync('docs/ultima-actualizacion.txt', new Date().toISOString() + `\n${partidos.length} partidos FBPA · ${nuevos} nuevos · ${cambios} actualizados\n`);
  console.log(`ICS: ${equipos.length + 1} calendarios publicados`);
}

main().catch((e) => { console.error('ERROR:', e.message); process.exit(1); });
