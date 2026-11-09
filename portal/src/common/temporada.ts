import { SPHttpClient } from '@microsoft/sp-http';
import { esc } from './cbp';
import { spGet, spPost, listUrl } from './ui';
import { IJug } from './jugadores';

/* Temporadas: empiezan el 1 de agosto. «2026-27» = agosto 2026 a julio 2027 */
export function inicioTemporada(d?: Date): number { const h = d || new Date(); return h.getMonth() >= 7 ? h.getFullYear() : h.getFullYear() - 1; }
export function nombreTemporada(inicio: number): string { return inicio + '-' + String(inicio + 1).slice(-2); }
export function temporadaActual(d?: Date): string { return nombreTemporada(inicioTemporada(d)); }
export function inicioDe(t: string): number { return +t.slice(0, 4); }

/* Categoría FBPA por año de nacimiento para la temporada que empieza en «inicio» */
export const CATEGORIAS = ['Benjamín', 'Alevín', 'Infantil', 'Cadete', 'Junior', 'Sénior'];
export function categoriaPorEdad(anioNac: number, inicio: number): string {
  const d = inicio - anioNac; // 2026-27: 2017-18 → 8-9 benjamín, 2015-16 alevín, 2013-14 infantil, 2011-12 cadete, 2009-10 junior
  if (d <= 9) { return 'Benjamín'; } if (d <= 11) { return 'Alevín'; } if (d <= 13) { return 'Infantil'; }
  if (d <= 15) { return 'Cadete'; } if (d <= 17) { return 'Junior'; } return 'Sénior';
}
export function categoriaDe(equipo: string): string { return CATEGORIAS.filter((c: string) => equipo.indexOf(c) === 0)[0] || ''; }

/** Equipo propuesto para la temporada siguiente: misma rama y letra en la categoría que le toca por edad */
export function sugerirEquipo(j: IJug, inicioNueva: number, lista: string[]): string {
  const cat = categoriaDe(j.equipo);
  if (!cat || !j.nac) { return j.equipo; } // madres, padres, tecnificación o sin fecha: se queda
  const nueva = categoriaPorEdad(j.nac.getFullYear(), inicioNueva);
  if (CATEGORIAS.indexOf(nueva) <= CATEGORIAS.indexOf(cat)) { return j.equipo; } // juega por encima de su edad o le toca seguir
  const resto = j.equipo.slice(cat.length).trim(); // «Masculino A», «Femenino»
  const rama = resto.split(' ')[0]; const letra = resto.split(' ')[1] || 'A';
  const candidatos = lista.filter((e: string) => e.indexOf(nueva + ' ' + rama) === 0);
  if (!candidatos.length) { return j.equipo; }
  // misma letra; si no existe, la letra más cercana por debajo (C → B), y si no, el último equipo de la rama
  const conLetra = candidatos.filter((e: string) => (e.split(' ')[2] || 'A') <= letra).sort();
  return conLetra.length ? conLetra[conLetra.length - 1] : candidatos.sort()[candidatos.length - 1];
}

/* ---------- Trayectoria ---------- */
export interface ITramo {
  id: number; jugador: number; temporada: string; equipo: string; habitual: string; ocasional: string; dorsal?: number;
  estado: string; sesiones: number; presencias: number; convocatorias: number; notas: string;
}
const SEL_T = 'Id,JugadorT,Temporada,EquipoT,HabitualT,OcasionalT,DorsalT,EstadoT,Sesiones,Presencias,Convocatorias,NotasT';
function mapT(i: any): ITramo { // eslint-disable-line @typescript-eslint/no-explicit-any
  return { id: i.Id, jugador: i.JugadorT, temporada: i.Temporada || '', equipo: i.EquipoT || '', habitual: i.HabitualT || '', ocasional: i.OcasionalT || '',
    dorsal: i.DorsalT === null || i.DorsalT === undefined ? undefined : i.DorsalT, estado: i.EstadoT || '', sesiones: i.Sesiones || 0, presencias: i.Presencias || 0,
    convocatorias: i.Convocatorias || 0, notas: i.NotasT || '' };
}
export function cargarTrayectoria(c: SPHttpClient, w: string, jugador: number): Promise<ITramo[]> {
  return spGet(c, listUrl(w, 'Lists/Trayectoria') + '/items?$select=' + SEL_T + '&$filter=JugadorT eq ' + jugador + '&$top=100')
    .then((j: any) => (j.value || []).map(mapT).sort((a: ITramo, b: ITramo) => b.temporada.localeCompare(a.temporada))); // eslint-disable-line @typescript-eslint/no-explicit-any
}
export function cargarTemporada(c: SPHttpClient, w: string, temporada: string): Promise<ITramo[]> {
  return spGet(c, listUrl(w, 'Lists/Trayectoria') + "/items?$select=" + SEL_T + "&$filter=Temporada eq '" + temporada + "'&$top=5000").then((j: any) => (j.value || []).map(mapT)); // eslint-disable-line @typescript-eslint/no-explicit-any
}

export function pct(t: { sesiones: number; presencias: number }): string { return t.sesiones ? Math.round(100 * t.presencias / t.sesiones) + '%' : '–'; }

/** Bloque «Trayectoria en el club» para la ficha del jugador: temporada en curso + temporadas cerradas */
export function htmlTrayectoria(j: IJug, tramos: ITramo[]): string {
  const actual = temporadaActual();
  interface IFila { temporada: string; equipo: string; habitual: string; dorsal?: number; estado: string; sesiones: number; presencias: number; convocatorias: number; enCurso: boolean; }
  const filas: IFila[] = [{ temporada: actual, equipo: j.equipo, habitual: j.habitual.join(', '), dorsal: j.dorsal, estado: 'En curso', sesiones: 0, presencias: 0, convocatorias: 0, enCurso: true }];
  tramos.filter((t: ITramo) => t.temporada !== actual).forEach((t: ITramo) => filas.push({ ...t, enCurso: false }));
  return '<div class="cbp-tray">' + filas.map((t: IFila) =>
    '<div class="cbp-tray-f' + (t.enCurso ? ' on' : '') + '"><span class="tem">' + esc(t.temporada) + '</span><span class="eq"><b>' + esc(t.equipo) + '</b>' +
    (t.dorsal !== undefined && t.dorsal !== null ? ' · dorsal ' + t.dorsal : '') + (t.habitual ? '<br><small>También con ' + esc(t.habitual) + '</small>' : '') + '</span>' +
    '<span class="dat">' + (t.enCurso ? 'En curso' : (t.sesiones ? pct(t) + ' asistencia · ' : '') + t.convocatorias + ' convocatoria' + (t.convocatorias === 1 ? '' : 's') + (t.estado && t.estado !== 'Completada' ? ' · ' + esc(t.estado) : '')) + '</span></div>').join('') +
    (tramos.length ? '' : '<p class="cbp-nota" style="margin:6px 0 0">Las temporadas anteriores aparecerán aquí cuando coordinación cierre cada temporada.</p>') + '</div>';
}
export const CSS_TRAY = `
.cbp-tray{display:flex;flex-direction:column;gap:6px}
.cbp-tray-f{display:grid;grid-template-columns:64px minmax(0,1fr) auto;gap:10px;align-items:center;border:1px solid #E6E3DC;border-radius:10px;padding:8px 12px;font-size:13px}
.cbp-tray-f.on{border-color:#FF9F00;background:#FFF8EA}
.cbp-tray-f .tem{font-family:Anton,Impact,sans-serif;font-size:15px;letter-spacing:.02em}
.cbp-tray-f small{color:#5c5c5c}
.cbp-tray-f .dat{color:#5c5c5c;font-size:12px;text-align:right}
.cbp-h3{font-family:Anton,Impact,sans-serif;font-weight:400;font-size:20px;letter-spacing:.01em;text-transform:uppercase}
.cbp-det{border:1px solid #E6E3DC;border-radius:12px;padding:8px 12px;margin:0 0 8px;background:#fff}
.cbp-det summary{cursor:pointer;font-size:14px;padding:4px 0}
.cbp-sel-sm{width:100%;min-height:34px;border:1px solid #D9D5CC;border-radius:8px;padding:4px 8px;font:inherit;font-size:13px;background:#fff}
@media (max-width:640px){.cbp-tray-f{grid-template-columns:56px minmax(0,1fr)}.cbp-tray-f .dat{grid-column:2;text-align:left}}
`;

/* ---------- Cierre de temporada: foto de cada jugador activo ---------- */
export interface ICierre { jugador: IJug; sesiones: number; presencias: number; convocatorias: number; }
function parseIds(s: string): number[] { try { const a = JSON.parse(s || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; } }

export function calcularCierre(c: SPHttpClient, w: string, jugs: IJug[], temporada: string): Promise<ICierre[]> {
  const ini = inicioDe(temporada); const desde = ini + '-08-01T00:00:00Z'; const hasta = (ini + 1) + '-08-01T00:00:00Z';
  const asis = spGet(c, listUrl(w, 'Lists/Asistencia') + "/items?$select=EquipoA,Presentes,Ausentes,Justificadas&$filter=FechaA ge datetime'" + desde + "' and FechaA lt datetime'" + hasta + "'&$top=5000").then((j: any) => j.value || []).catch(() => []); // eslint-disable-line @typescript-eslint/no-explicit-any
  const conv = spGet(c, listUrl(w, 'Lists/Convocatorias') + "/items?$select=JugadoresC&$filter=FechaPartido ge datetime'" + desde + "' and FechaPartido lt datetime'" + hasta + "'&$top=5000").then((j: any) => j.value || []).catch(() => []); // eslint-disable-line @typescript-eslint/no-explicit-any
  return Promise.all([asis, conv]).then((r: any[][]) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    const ses: { [id: number]: number } = {}; const pre: { [id: number]: number } = {}; const cv: { [id: number]: number } = {};
    r[0].forEach((s: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
      const p = parseIds(s.Presentes); const todos = p.concat(parseIds(s.Ausentes), parseIds(s.Justificadas));
      todos.forEach((id: number) => { ses[id] = (ses[id] || 0) + 1; }); p.forEach((id: number) => { pre[id] = (pre[id] || 0) + 1; });
    });
    r[1].forEach((x: any) => parseIds(x.JugadoresC).forEach((id: number) => { cv[id] = (cv[id] || 0) + 1; })); // eslint-disable-line @typescript-eslint/no-explicit-any
    return jugs.map((j: IJug): ICierre => ({ jugador: j, sesiones: ses[j.id] || 0, presencias: pre[j.id] || 0, convocatorias: cv[j.id] || 0 }));
  });
}

/** Escribe la foto de la temporada. Si un jugador ya tiene fila en esa temporada, la actualiza (se puede repetir sin duplicar). */
export function guardarCierre(c: SPHttpClient, w: string, temporada: string, datos: ICierre[], existentes: ITramo[], progreso?: (n: number) => void): Promise<number> {
  const url = listUrl(w, 'Lists/Trayectoria') + '/items';
  let hechos = 0; let i = 0;
  const uno = (d: ICierre): Promise<void> => {
    const ya = existentes.filter((t: ITramo) => t.jugador === d.jugador.id)[0];
    const body = { Title: d.jugador.nombre, JugadorT: d.jugador.id, Temporada: temporada, EquipoT: d.jugador.equipo, HabitualT: d.jugador.habitual.join(', '),
      OcasionalT: d.jugador.ocasional.join(', '), DorsalT: d.jugador.dorsal === undefined ? null : d.jugador.dorsal, EstadoT: d.jugador.activo ? 'Completada' : 'Baja durante la temporada',
      Sesiones: d.sesiones, Presencias: d.presencias, Convocatorias: d.convocatorias };
    return spPost(c, url + (ya ? '(' + ya.id + ')' : ''), body, !!ya).then(() => { hechos++; if (progreso) { progreso(hechos); } });
  };
  const lote = (): Promise<void> => { const parte = datos.slice(i, i + 8); i += 8; return parte.length ? Promise.all(parte.map(uno)).then(lote) : Promise.resolve(); };
  return lote().then(() => hechos);
}

export function esCoordinacion(c: SPHttpClient, w: string): Promise<boolean> {
  return spGet(c, w + '/_api/web/currentuser?$select=IsSiteAdmin&$expand=Groups').then((u: any) => // eslint-disable-line @typescript-eslint/no-explicit-any
    !!u.IsSiteAdmin || (u.Groups || []).some((g: any) => g.Title === 'Coordinación deportiva' || g.Title === 'Junta Directiva')).catch(() => false); // eslint-disable-line @typescript-eslint/no-explicit-any
}
