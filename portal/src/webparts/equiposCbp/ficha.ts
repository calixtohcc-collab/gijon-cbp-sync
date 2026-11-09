import { SPHttpClient } from '@microsoft/sp-http';
import { IEquipo, IEvento, esc, icon } from '../../common/cbp';
import { filaEvento } from '../../common/render';
import { spGet, spPost, listUrl, abrirModal, cerrarModal, aviso } from '../../common/ui';
import { enlazarEventos } from '../../common/evento';
import { slugEquipo, WEB_CALENDARIOS } from '../../common/estilos3';
import { IJug, cargarJugs, fichaJugador, faltan, telLlamar, telLink, cargarEstilosJug } from '../../common/jugadores';
interface IAsis { presentes: number[]; ausentes: number[]; justificadas: number[]; }

const DIA_ABR: { [k: string]: string } = { lunes: 'LUN', martes: 'MAR', 'miércoles': 'MIÉ', miercoles: 'MIÉ', jueves: 'JUE', viernes: 'VIE', 'sábado': 'SÁB', sabado: 'SÁB', domingo: 'DOM' };

function arr(s: string): number[] { try { const v = JSON.parse(s || '[]'); return Array.isArray(v) ? v : []; } catch (e) { return []; } }

export interface IFichaOpts {
  client: SPHttpClient;
  w: string;
  equipo: IEquipo;
  eventos: IEvento[];
  volver: () => void;
  alCambiarEquipo: (e: IEquipo) => void;
}

/** Ficha completa de un equipo, dentro del propio portal */
export function renderFicha(el: HTMLElement, o: IFichaOpts): void {
  const e = o.equipo;
  el.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:20px;gap:18px"><div class="cbp-cargando">Cargando ficha de ' + esc(e.nombre) + '…</div></div></div>';
  const qEq = encodeURIComponent(e.nombre.replace(/'/g, "''"));
  Promise.all([
    cargarJugs(o.client, o.w).catch(() => []),
    spGet(o.client, listUrl(o.w, 'Lists/Asistencia') + "/items?$select=Presentes,Ausentes,Justificadas&$filter=EquipoA eq '" + qEq + "'&$top=500").catch(() => ({ value: [] }))
  ]).then((r: any[]) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    cargarEstilosJug();
    const todos: IJug[] = (r[0] as IJug[]).filter((j: IJug) => j.activo);
    const vinc = todos.filter((j: IJug) => j.equipo !== e.nombre && (j.habitual.indexOf(e.nombre) >= 0 || j.ocasional.indexOf(e.nombre) >= 0));
    const jugs: IJug[] = todos.filter((j: IJug) => j.equipo === e.nombre)
      .sort((a: IJug, b: IJug) => (a.dorsal === undefined ? 999 : a.dorsal) - (b.dorsal === undefined ? 999 : b.dorsal) || a.nombre.localeCompare(b.nombre));
    const asis: IAsis[] = (r[1].value || []).map((i: any): IAsis => ({ presentes: arr(i.Presentes), ausentes: arr(i.Ausentes), justificadas: arr(i.Justificadas) })); // eslint-disable-line @typescript-eslint/no-explicit-any
    pintar(el, o, jugs, asis, vinc);
  });
}

function pintar(el: HTMLElement, o: IFichaOpts, jugs: IJug[], asis: IAsis[], vinc: IJug[]): void {
  const e = o.equipo; const ahora = new Date();
  const evs = o.eventos.filter((v: IEvento) => v.equipo === e.nombre);
  const prox = evs.filter((v: IEvento) => v.fecha >= ahora).slice(0, 6);
  const jugados = evs.filter((v: IEvento) => v.tipo === 'Partido' && v.pf !== undefined && v.pa !== undefined && v.fecha < ahora);
  const recientes = jugados.slice().reverse();
  const vict = jugados.filter((v: IEvento) => (v.pf as number) > (v.pa as number)).length;
  const der = jugados.length - vict;
  const dif = jugados.reduce((s: number, v: IEvento) => s + (v.pf as number) - (v.pa as number), 0);
  const siguiente = evs.filter((v: IEvento) => v.tipo === 'Partido' && v.fecha >= ahora)[0];
  const tieneFbpa = evs.some((v: IEvento) => v.tipo === 'Partido');
  const icsUrl = WEB_CALENDARIOS + 'ics/' + slugEquipo(e.nombre) + '.ics';

  // Estadísticas
  const stats = jugados.length
    ? '<div class="cbp-fstat"><b>' + vict + '–' + der + '</b><span>Balance</span></div>' +
      '<div class="cbp-fstat"><b>' + (dif > 0 ? '+' : '') + dif + '</b><span>Diferencia</span></div>' +
      '<div class="cbp-fstat"><b>' + Math.round((vict / jugados.length) * 100) + '%</b><span>Victorias</span></div>'
    : '<div class="cbp-fstat"><b>' + jugs.length + '</b><span>Jugadores</span></div>' +
      '<div class="cbp-fstat"><b>' + prox.length + '</b><span>Próximos</span></div>' +
      '<div class="cbp-fstat"><b>' + asis.length + '</b><span>Sesiones</span></div>';

  const dato = (ico: string, etiqueta: string, valor: string, campo: string): string => valor
    ? '<span class="cbp-fdato">' + icon(ico, 14) + '<b>' + etiqueta + '</b> ' + esc(valor) + '</span>'
    : '<button type="button" class="cbp-fdato vacio" data-editar="' + campo + '">+ Añadir ' + etiqueta.toLowerCase() + '</button>';

  const hero = '<section class="cbp-fhero"><div style="position:relative;z-index:1">' +
    '<span class="cbp-over">' + esc(e.categoria) + ' · ' + esc(e.rama) + '</span><h2>' + esc(e.nombre) + '</h2>' +
    (e.competicion ? '<p>' + esc(e.competicion) + '</p>' : '') +
    '<div class="cbp-fdatos">' + dato('junta', 'Cuerpo técnico', e.entrenadores, 'Entrenadores') + dato('pin', 'Pista', e.pista, 'Pista') + '</div></div>' +
    '<div class="cbp-fstats">' + stats + '</div></section>';

  const herr = (tab: string, txt: string, pri?: boolean, extra?: string): string =>
    '<a class="cbp-btn' + (pri ? ' pri' : '') + '" data-herr="' + tab + '" href="' + esc(o.w + '/SitePages/Herramientas.aspx#' + tab + (extra || '')) + '"' + (pri ? ' style="color:#FF9F00"' : '') + '>' + txt + '</a>';
  const acciones = '<div class="cbp-facciones">' +
    (siguiente ? herr('convocatorias', 'Convocar próximo partido', true, '=' + siguiente.id) : '') +
    herr('asistencia', 'Pasar lista') + herr('sesiones', 'Preparar sesión') +
    '<a class="cbp-btn" href="' + esc(tieneFbpa ? icsUrl.replace(/^https:/, 'webcal:') : WEB_CALENDARIOS) + '" target="_blank" rel="noopener">Calendario en el móvil</a>' +
    '<button type="button" class="cbp-btn" data-editar="todo">Editar ficha</button></div>';

  // Columna izquierda: partidos
  const proxHtml = prox.length ? prox.map((v: IEvento) => filaEvento(v, o.w)).join('') : '<p class="cbp-nota">No hay partidos ni eventos programados.</p>';
  const racha = recientes.slice(0, 8).map((v: IEvento) => { const g = (v.pf as number) > (v.pa as number); return '<span class="' + (g ? 'v' : 'd') + '" title="' + esc(v.rival + ' ' + v.pf + '–' + v.pa) + '">' + (g ? 'V' : 'D') + '</span>'; }).join('');
  const resHtml = recientes.length
    ? (racha ? '<div class="cbp-racha" aria-label="Últimos resultados">' + racha + '</div>' : '') + recientes.slice(0, 5).map((v: IEvento) => filaEvento(v, o.w, true)).join('')
    : '<p class="cbp-nota">Todavía no hay resultados esta temporada.</p>';

  // Columna derecha: horario y plantilla
  const lineas = e.horario.split(/\n/).filter((l: string) => l.trim());
  const horario = lineas.length ? '<div class="cbp-horario">' + lineas.map((l: string) => {
    const m = l.match(/^(\S+)\s+(\S+)\s*[·\-–]\s*(.*)$/) || l.match(/^(\S+)\s+(.*)$/);
    const d = m ? DIA_ABR[m[1].toLowerCase()] : '';
    if (!m || !d) { return '<div class="cbp-hdia"><b>·</b><span>' + esc(l) + '</span></div>'; }
    return '<div class="cbp-hdia"><b>' + d + '</b><span>' + esc(m[2]) + (m[3] ? '<i>' + esc(m[3]) + '</i>' : '') + '</span></div>';
  }).join('') + '</div>' : '<p class="cbp-nota">Sin horario de entrenamiento. <button type="button" class="cbp-btn sm" data-editar="Horario">Añadir</button></p>';

  const pct = (id: number): string => {
    let p = 0; let t = 0;
    asis.forEach((a: IAsis) => { const enP = a.presentes.indexOf(id) >= 0; const enA = a.ausentes.indexOf(id) >= 0 || a.justificadas.indexOf(id) >= 0; if (enP || enA) { t++; } if (enP) { p++; } });
    if (!t) { return ''; }
    const v = Math.round((p / t) * 100);
    return '<span class="cbp-pct" title="Asistencia a entrenamientos: ' + p + ' de ' + t + '"><i><u style="width:' + v + '%"></u></i>' + v + '%</span>';
  };
  const tel = (j: IJug): string => { const t = telLlamar(j); return t ? '<a class="cbp-tel" href="' + esc(telLink(t.tel)) + '" title="Llamar a ' + esc(t.quien + ' · ' + t.tel) + '" aria-label="Llamar a ' + esc(t.quien) + '">☎</a>' : '<span></span>'; };
  const fila = (j: IJug, sube?: string): string => { const f = faltan(j);
    return '<div class="cbp-jug" data-jug="' + j.id + '" style="cursor:pointer;grid-template-columns:40px minmax(0,1fr) auto auto"><span class="cbp-dorsal">' + (j.dorsal === undefined ? '–' : j.dorsal) + '</span>' +
      '<span class="nom">' + esc(j.nombre) + (sube ? ' <span class="cbp-sube' + (sube === 'oca' ? ' oca' : '') + '">' + (sube === 'oca' ? 'Ocasional' : 'Sube') + ' · ' + esc(j.equipo) + '</span>' : '') +
      (f.length ? '<br><span class="cbp-falta">Falta: ' + esc(f.join(', ')) + '</span>' : '') + '</span>' + (sube ? '<span></span>' : (pct(j.id) || '<span></span>')) + tel(j) + '</div>'; };
  const subenHab = vinc.filter((j: IJug) => j.habitual.indexOf(e.nombre) >= 0); const subenOca = vinc.filter((j: IJug) => j.habitual.indexOf(e.nombre) < 0);
  const roster = (jugs.length ? '<div class="cbp-roster">' + jugs.map((j: IJug) => fila(j)).join('') + '</div>'
    : '<p class="cbp-nota">Aún no hay jugadores. Añádelos aquí o desde Herramientas → Plantillas; aparecen en ambos sitios.</p>') +
    (vinc.length ? '<h4 style="margin:10px 0 0;font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#5c5c5c">Suben de otros equipos</h4><div class="cbp-roster">' +
      subenHab.map((j: IJug) => fila(j, 'hab')).join('') + subenOca.map((j: IJug) => fila(j, 'oca')).join('') + '</div>' : '') +
    '<button type="button" class="cbp-btn sm pri" data-nuevo-jug="1" style="align-self:flex-start;color:#FF9F00">+ Añadir jugador</button>';

  const panel = (t: string, sub: string, cuerpo: string): string => '<section class="cbp-panel"><div class="cbp-panel-h"><h3>' + t + '</h3>' + (sub ? '<small>' + sub + '</small>' : '') + '</div><div class="cbp-panel-b">' + cuerpo + '</div></section>';

  el.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:16px;gap:18px">' +
    '<div><button type="button" class="cbp-volver" data-volver="1">← Todos los equipos</button></div>' + hero + acciones +
    '<div class="cbp-fgrid"><div class="cbp-fcol">' +
    panel('Próximos', prox.length ? 'siguientes ' + prox.length : '', proxHtml) +
    panel('Resultados', jugados.length ? jugados.length + ' jugados' : '', resHtml) +
    '</div><div class="cbp-fcol">' +
    panel('Plantilla', jugs.length + ' jugador' + (jugs.length === 1 ? '' : 'es'), roster) +
    panel('Entrenamientos', '', horario) +
    '</div></div></div></div>';

  enlazarEventos(el, o.eventos, o.w);
  const q = (s: string): NodeListOf<HTMLElement> => el.querySelectorAll(s) as NodeListOf<HTMLElement>;
  (el.querySelector('[data-volver]') as HTMLElement).addEventListener('click', o.volver);
  q('[data-herr]').forEach((a: HTMLElement) => a.addEventListener('click', () => { try { window.localStorage.setItem('cbp-equipo', e.nombre); } catch (x) { /* sin almacenamiento */ } }));
  q('[data-editar]').forEach((b: HTMLElement) => b.addEventListener('click', () => editar(o, b.getAttribute('data-editar') || 'todo')));
  const todos = jugs.concat(vinc);
  q('[data-jug]').forEach((b: HTMLElement) => b.addEventListener('click', (ev: Event) => {
    if ((ev.target as HTMLElement).closest('a')) { return; }
    const j = todos.filter((x: IJug) => x.id === +(b.getAttribute('data-jug') || 0))[0];
    if (j) { fichaJugador(o.client, o.w, j, e.nombre, () => renderFicha(el, o)); }
  }));
  q('[data-nuevo-jug]').forEach((b: HTMLElement) => b.addEventListener('click', () => fichaJugador(o.client, o.w, undefined, e.nombre, () => renderFicha(el, o))));
}


function editar(o: IFichaOpts, foco: string): void {
  const e = o.equipo;
  const c = abrirModal('Editar ' + e.nombre,
    '<form class="cbp-form" data-f="1">' +
    '<label class="cbp-field full">Cuerpo técnico<input name="Entrenadores" value="' + esc(e.entrenadores) + '" placeholder="Ej.: Ana López (principal), Luis Pérez (ayudante)"></label>' +
    '<label class="cbp-field full">Pista habitual<input name="Pista" value="' + esc(e.pista) + '"></label>' +
    '<label class="cbp-field full">Competición<input name="Competicion" value="' + esc(e.competicion) + '"></label>' +
    '<label class="cbp-field full">Horario de entrenamiento <span style="font-weight:400;color:#5c5c5c">(una línea por día: «Lunes 19:00–20:30 · Pabellón»)</span><textarea name="Horario" rows="4">' + esc(e.horario) + '</textarea></label>' +
    '<div class="cbp-btns full"><button type="button" class="cbp-btn" data-c="1">Cancelar</button><button type="submit" class="cbp-btn pri">Guardar</button></div></form>', 560);
  const f = c.querySelector('[data-f]') as HTMLFormElement;
  const campo = f.querySelector('[name="' + foco + '"]') as HTMLElement; if (campo) { setTimeout(() => campo.focus(), 50); }
  (c.querySelector('[data-c]') as HTMLElement).addEventListener('click', cerrarModal);
  f.addEventListener('submit', (ev: Event) => {
    ev.preventDefault();
    const v = (n: string): string => (f.querySelector('[name="' + n + '"]') as HTMLInputElement).value.trim();
    const body = { Entrenadores: v('Entrenadores'), Pista: v('Pista'), Competicion: v('Competicion'), Horario: v('Horario') };
    spPost(o.client, listUrl(o.w, 'Lists/Equipos') + '/items(' + e.id + ')', body, true).then(() => {
      const nuevo: IEquipo = { ...e, entrenadores: body.Entrenadores, pista: body.Pista, competicion: body.Competicion, horario: body.Horario };
      cerrarModal(); aviso('Ficha guardada'); o.alCambiarEquipo(nuevo);
    }).catch((x: Error) => aviso('Error: ' + x.message, true));
  });
}
