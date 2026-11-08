// Directorio de clubes de la FBPA: nombre, dirección y qué equipos (nombres FBPA) pertenecen a cada club.
// La FBPA no publica correos ni teléfonos: esos campos se rellenan a mano en el portal y aquí nunca se tocan.
import { texto, norm } from './fbpa.mjs';

const BASE = 'https://www.fbpa.es';
const UA = { 'User-Agent': 'GijonCBP-Portal-Sync/1.0 (+https://cbpumarin.com)' };

function filas(html) {
  const i = html.indexOf('id="nav-calendario"'); if (i < 0) { return []; }
  const sec = html.slice(i);
  return (sec.match(/<tr[\s\S]*?<\/tr>/g) || []).map((f) => (f.match(/<td[\s\S]*?<\/td>/g) || []).map(texto)).filter((t) => t.length >= 6).map((t) => [t[0], t[3]]);
}
function direccion(html, nombre) {
  const cab = html.slice(0, Math.max(0, html.indexOf('id="nav-calendario"')) || 200000);
  const m = texto(cab).match(new RegExp(nombre.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s+([^|]{5,160}?, Asturias)'));
  return m ? m[1].trim() : '';
}

const VACIAS = new Set(['C', 'B', 'D', 'CB', 'CD', 'CDB', 'CLUB', 'DE', 'DEL', 'LA', 'EL', 'LOS', 'LAS', 'Y', 'BALONCESTO', 'BASKET', 'BASKETBALL', 'AD', 'ACD', 'ASOCIACION', 'DEPORTIVO', 'DEPORTIVA',
  'AGRUPACION', 'COL', 'COLEGIO', 'IES', 'SAD', 'N', 'NO', 'A', 'U18', 'U22', 'GB', 'ANTIGUOS', 'ALUMNOS', 'SOCIEDAD', 'CULTURAL']);
const palabras = (s) => norm(s).replace(/[^A-Z0-9Ñ ]+/g, ' ').split(' ').filter((w) => w && !VACIAS.has(w));
function siglas(club) {
  const ws = norm(club).replace(/[^A-Z0-9Ñ ]+/g, ' ').split(' ').filter((w) => w && !['DE', 'DEL', 'LA', 'EL', 'Y'].includes(w));
  const letras = ws.filter((w) => !/^\d+$/.test(w)).map((w) => w[0]).join(''); const num = ws.filter((w) => /^\d+$/.test(w)).join('');
  return [letras + num, letras].filter((x) => x.length >= 3);
}
/** 0..1: cuánto del nombre del equipo coincide con el del club */
export function parecido(equipo, club) {
  const e = palabras(equipo); if (!e.length) { return 0; }
  const c = new Set(palabras(club));
  let p = e.filter((w) => w.length >= 3 && c.has(w)).length / e.length;
  const sg = siglas(club); const en = norm(equipo).replace(/[^A-Z0-9Ñ ]+/g, ' ').split(' ');
  if (sg.some((s) => en.includes(s))) { p = Math.max(p, 0.9); }
  return p;
}

export async function leerClubes() {
  const idx = await fetch(BASE + '/resultados-clubes', { headers: UA }).then((r) => r.text());
  const enlaces = [];
  for (const m of idx.matchAll(/href="(\/resultados-club-\d+\/[^"]+)"[^>]*>([\s\S]*?)<\/a>/g)) { if (!enlaces.some((e) => e.url === m[1])) { enlaces.push({ url: m[1], nombre: texto(m[2]) }); } }
  const paginas = [];
  for (let i = 0; i < enlaces.length; i += 6) {
    const lote = enlaces.slice(i, i + 6);
    const res = await Promise.all(lote.map(async (e) => {
      try { const h = await fetch(BASE + e.url, { headers: UA }).then((r) => r.text()); return { ...e, dir: direccion(h, e.nombre), filas: filas(h) }; }
      catch (x) { return { ...e, dir: '', filas: [] }; }
    }));
    paginas.push(...res);
  }
  // Cada equipo se asigna al club con el que más se parece su nombre (palabras y siglas),
  // entre los clubes en cuya página aparece; si no hay parecido, al de mayor peso en su página.
  const cands = new Map();
  for (const p of paginas) {
    if (!p.filas.length) { continue; }
    const c = new Map();
    for (const [a, b] of p.filas) { for (const n of new Set([a, b])) { if (!/DESCANSA/i.test(n)) { c.set(n, (c.get(n) || 0) + 1); } } }
    for (const [n, k] of c) { const key = norm(n); if (!cands.has(key)) { cands.set(key, { nombre: n, opciones: [] }); } cands.get(key).opciones.push({ url: p.url, club: p.nombre, k, peso: k / p.filas.length }); }
  }
  const asign = new Map();
  for (const [, v] of cands) {
    let mejor = null;
    for (const o of v.opciones) {
      const s = parecido(v.nombre, o.club) * 10 + o.peso;
      if (!mejor || s > mejor.s) { mejor = { s, url: o.url }; }
    }
    if (mejor) { asign.set(v.nombre, mejor.url); }
  }
  return paginas.map((p) => ({ url: p.url, nombre: p.nombre, dir: p.dir,
    equipos: [...asign.entries()].filter(([, u]) => u === p.url).map(([n]) => n).sort() }));
}

/** Crea o actualiza la lista «Partidos · Clubes» (sin tocar los datos de contacto) */
export async function actualizarClubes(graph, site, lista) {
  const clubes = await leerClubes();
  const items = [];
  let url = `/sites/${site}/lists/${lista}/items?$expand=fields($select=Title,UrlFBPA,EquiposFBPA,Direccion)&$top=500`;
  while (url) { const j = await graph(url); items.push(...j.value); url = j['@odata.nextLink'] || null; }
  let n = 0, a = 0;
  for (const c of clubes) {
    if (/cb-pumarin/.test(c.url)) { continue; }
    const campos = { Title: c.nombre, UrlFBPA: c.url, Direccion: c.dir, EquiposFBPA: c.equipos.join('\n') };
    const it = items.find((x) => x.fields.UrlFBPA === c.url);
    if (!it) { await graph(`/sites/${site}/lists/${lista}/items`, { method: 'POST', body: JSON.stringify({ fields: campos }) }); n++; continue; }
    const dif = {}; for (const k of Object.keys(campos)) { if ((it.fields[k] || '') !== campos[k] && (k !== 'EquiposFBPA' || campos[k])) { dif[k] = campos[k]; } }
    if (Object.keys(dif).length) { await graph(`/sites/${site}/lists/${lista}/items/${it.id}/fields`, { method: 'PATCH', body: JSON.stringify(dif) }); a++; }
  }
  return { total: clubes.length, nuevos: n, actualizados: a };
}
