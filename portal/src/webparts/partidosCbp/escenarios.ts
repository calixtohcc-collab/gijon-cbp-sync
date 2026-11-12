import { esc } from '../../common/cbp';
import { aviso } from '../../common/ui';
import { IPartido, IAlerta, IPista, ICand, IEscenario, IResolverOpts, IPeticion, IHueco, resolver, sugerir, enPlazo, norm, minDia, hm, DIA_TXT } from './motor';
import { IGCtx, IPartidoG, esFijo, guardarGest, justificacion, fTxt } from './gestion';

/* Motor de escenarios en la pestaña Semana (solo coordinación):
   calcula varias soluciones para toda la semana, las enseña sobre el mismo gantt
   y deja fijar, descartar o cambiar un cambio concreto y recalcular el resto. */

export interface IEscEstado { res: IEscenario[] | null; sel: string; forzados: { [id: number]: ICand }; calculando: boolean; semanaKey: string; pendiente: boolean; }
export function nuevoEstado(): IEscEstado { return { res: null, sel: '', forzados: {}, calculando: false, semanaKey: '', pendiente: false }; }

function opciones(ctx: IGCtx, forzados: { [id: number]: ICand }): IResolverOpts {
  const fijos: { [id: number]: boolean } = {}; const peticiones: { [id: number]: IPeticion } = {};
  ctx.semana.forEach((p: IPartidoG) => {
    const g = ctx.gest[p.id];
    if (esFijo(g) || (g && g.estado === 'Propuesto')) { fijos[p.id] = true; } // trabaja sobre lo que coordinación ya ha propuesto o bloqueado en la tabla
    if (g && g.petHora && (g.estado === 'Oficial' || g.estado === 'Rechazado' || g.estado === 'Propuesto')) { peticiones[p.id] = { fecha: g.petHora, nota: g.petNota, de: g.petDe }; }
  });
  return { reglas: ctx.reglas, pistas: ctx.pistas, margen: ctx.margen, jugs: ctx.jugs, prefsDe: ctx.eqInfo, fijos, forzados, peticiones, plazo: ctx.plazo };
}

export function calcular(ctx: IGCtx, est: IEscEstado, listo: () => void): void {
  est.calculando = true; est.pendiente = false; listo();
  window.setTimeout(() => {
    try {
      const t0 = Date.now();
      est.res = resolver(ctx.semana, opciones(ctx, est.forzados));
      if (!est.sel || !est.res.some((e: IEscenario) => e.clave === est.sel)) { est.sel = est.res.slice().sort((a: IEscenario, b: IEscenario) => a.criticas - b.criticas || a.cambios.length - b.cambios.length)[0].clave; }
      if (Date.now() - t0 > 8000) { aviso('El cálculo ha tardado ' + Math.round((Date.now() - t0) / 1000) + ' s'); }
    } catch (x) { aviso('No se pudo calcular: ' + (x as Error).message, true); est.res = null; }
    est.calculando = false; listo();
  }, 30);
}

/** Barra de escenarios + tabla de cambios del escenario elegido */
export function vEscenarios(ctx: IGCtx, est: IEscEstado, alActual: IAlerta[]): string {
  const crit = alActual.filter((a: IAlerta) => a.grav === 'critica').length;
  const peticiones = ctx.semana.filter((p: IPartidoG) => { const g = ctx.gest[p.id]; return g && g.petHora && (g.estado === 'Oficial' || g.estado === 'Rechazado'); }).length;
  const fijos = ctx.semana.filter((p: IPartidoG) => esFijo(ctx.gest[p.id])).length;
  const props = ctx.semana.filter((p: IPartidoG) => ctx.gest[p.id] && ctx.gest[p.id].estado === 'Propuesto').length;
  const intro = '<p class="cbp-nota" style="margin:0">El motor parte de la tabla de arriba: respeta las propuestas que ya has guardado (' + props + '), los partidos inamovibles y los cambios en trámite (' + fijos + ')' +
    (peticiones ? ', y tiene en cuenta ' + peticiones + ' petici' + (peticiones === 1 ? 'ón' : 'ones') + ' de hora de entrenadores' : '') + '. Busca soluciones para toda la semana a la vez. Los partidos sin franjas cargadas (p. ej. competición nacional) solo se mueven a mano. Puedes aceptar o rechazar cambios concretos y pulsar «Recalcular» cuando quieras.</p>';
  if (est.calculando) { return '<section class="cbp-box cbp-esc"><div class="cbp-box-h"><h3 class="cbp-h2" style="font-size:24px;margin:0">Resolver la semana</h3></div><div class="cbp-cargando">Calculando escenarios…</div></section>'; }
  if (!est.res) {
    return '<section class="cbp-box cbp-esc"><div class="cbp-box-h" style="flex-wrap:wrap;gap:8px"><h3 class="cbp-h2" style="font-size:24px;margin:0">Resolver la semana</h3>' +
      '<button type="button" class="cbp-btn pri" data-esc-calc="1" style="color:#FF9F00">' + (crit ? 'Calcular escenarios sin conflictos' : 'Calcular escenarios') + '</button></div>' + intro + '</section>';
  }
  const chip = (clave: string, nombre: string, sub: string, mal: boolean): string => '<button type="button" class="cbp-esc-chip' + (est.sel === clave ? ' on' : '') + (mal ? ' mal' : '') + '" data-esc="' + clave + '"><b>' + esc(nombre) + '</b><small>' + sub + '</small></button>';
  const chips = chip('', 'Situación actual', crit + ' conflicto' + (crit === 1 ? '' : 's'), crit > 0) + est.res.map((e: IEscenario) =>
    chip(e.clave, e.nombre, e.cambios.length + ' cambio' + (e.cambios.length === 1 ? '' : 's') + ' · ' + (e.criticas ? e.criticas + ' conflicto' + (e.criticas === 1 ? '' : 's') : 'sin conflictos') + (e.holguraMin >= 0 ? ' · holgura mín. ' + e.holguraMin + ' min' : ''), e.criticas > 0)).join('');
  const forz = Object.keys(est.forzados).length;
  let cuerpo = '';
  const e = est.res.filter((x: IEscenario) => x.clave === est.sel)[0];
  if (e) {
    const pid = (id: number): IPartidoG => ctx.semana.filter((p: IPartidoG) => p.id === id)[0];
    const filas = e.cambios.map((c) => {
      const p = pid(c.id); const f = !!est.forzados[c.id];
      const pista = c.pista ? (ctx.pistas.filter((k: IPista) => k.clave === c.pista)[0] || { inst: c.lugar, modulo: '' } as IPista) : null;
      const etiquetas = [p.local ? (c.enPlazo ? '<span class="cbp-plazo ok">En plazo: sin conformidad del rival</span>' : '<span class="cbp-plazo">Fuera de plazo: conformidad del rival</span>') : '<span class="cbp-plazo">Fuera: lo decide el club local</span>'];
      if (c.cambiaDia) { etiquetas.push('<span class="cbp-plazo">Cambia de día</span>'); }
      if (c.tipoFranja === 'conformidad' || c.tipoFranja === 'condicion') { etiquetas.push('<span class="cbp-plazo">Franja con conformidad</span>'); }
      if (c.peticion) { etiquetas.push('<span class="cbp-plazo ok">Atiende la petición del entrenador</span>'); }
      return '<tr' + (f ? ' class="fij"' : '') + '><td><b>' + esc(p.equipo) + '</b><br><small style="color:#5c5c5c">' + (p.local ? 'vs ' : 'en ') + esc(p.rival) + '</small></td>' +
        '<td><s style="color:#5c5c5c">' + esc(fTxt(c.desde)) + '</s></td><td><b>' + esc(fTxt(c.hasta)) + '</b>' + (pista ? '<br><small>' + esc(pista.inst + (pista.modulo ? ' ' + pista.modulo : '')) + '</small>' : '') + '</td>' +
        '<td><div style="display:flex;gap:4px;flex-wrap:wrap">' + etiquetas.join('') + '</div></td>' +
        '<td><div style="display:flex;gap:6px;flex-wrap:wrap">' + (f ? '<span class="cbp-plazo ok">Aceptado</span><button type="button" class="cbp-btn sm" data-esc-sol="' + c.id + '">Quitar</button>'
          : '<button type="button" class="cbp-btn sm" data-esc-fij="' + c.id + '" title="Me vale: se queda así en los próximos cálculos">Aceptar este cambio</button><button type="button" class="cbp-btn sm" data-esc-orig="' + c.id + '" title="Este partido se queda a su hora; el motor buscará otra solución para el resto">No mover este partido</button><button type="button" class="cbp-btn sm" data-esc-otra="' + c.id + '">Otra hora…</button>') + '</div></td></tr>';
    }).join('');
    const quedan = e.alertas.filter((a: IAlerta) => a.grav === 'critica');
    const ajustes = Object.keys(est.forzados).map((k: string) => { const p = pid(+k); const f = est.forzados[+k]; if (!p) { return ''; } const igual = f.fecha.getTime() === p.fecha.getTime();
      return '<span class="cbp-ajuste">' + esc(p.equipo) + ': ' + (igual ? 'no se mueve' : esc(fTxt(f.fecha))) + '<button type="button" data-esc-sol="' + p.id + '" aria-label="Quitar ajuste">✕</button></span>'; }).join('');
    cuerpo = '<p class="cbp-nota" style="margin:0"><b>' + esc(e.nombre) + ':</b> ' + esc(e.desc) + '</p>' +
      (ajustes ? '<div class="cbp-ajustes"><b>Tus ajustes' + (est.pendiente ? ' (sin aplicar)' : '') + ':</b> ' + ajustes + '</div>' : '') +
      (quedan.length ? '<div class="cbp-odd" style="margin:0;background:#FFF6F6"><b>Quedan ' + quedan.length + ' conflicto' + (quedan.length === 1 ? '' : 's') + ' sin resolver:</b>\n' + quedan.map((a: IAlerta) => '• ' + esc(a.titulo) + ' — ' + esc(a.accion)).join('\n') + '</div>' : '') +
      (e.cambios.length ? '<div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Partido</th><th>FBPA</th><th>Propuesta</th><th>Qué implica</th><th></th></tr></thead><tbody>' + filas + '</tbody></table></div>'
        : '<div class="cbp-vacio">Este escenario no mueve ningún partido.</div>') +
      '<div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">' +
      (e.cambios.length ? '<button type="button" class="cbp-btn pri" data-esc-guardar="1" style="color:#FF9F00">Guardar como propuesta (' + e.cambios.length + ')</button>' : '') +
      '<button type="button" class="cbp-btn' + (est.pendiente ? ' pri" style="color:#FF9F00' : '') + '" data-esc-calc="1">' + (est.pendiente ? 'Recalcular con mis ajustes' : 'Recalcular') + '</button>' +
      (forz ? '<button type="button" class="cbp-btn" data-esc-limpiar="1">Quitar mis ajustes (' + forz + ')</button>' : '') +
      '<span class="cbp-nota">Al guardar, cada cambio queda «Propuesto»: el entrenador da el OK, coordinación valida y se carga en la FBPA o se pide al rival.</span></div>';
  } else {
    cuerpo = '<p class="cbp-nota" style="margin:0">Estás viendo los horarios actuales. Elige un escenario para verlo sobre el gantt.</p>';
  }
  return '<section class="cbp-box cbp-esc"><div class="cbp-box-h" style="flex-wrap:wrap;gap:8px"><h3 class="cbp-h2" style="font-size:24px;margin:0">Resolver la semana</h3></div>' + intro +
    '<div class="cbp-esc-chips" role="tablist">' + chips + '</div>' + cuerpo + '</section>';
}

/** Partidos y alertas que hay que pintar en el gantt según el escenario elegido */
export function vista(est: IEscEstado, ps: IPartido[], al: IAlerta[]): { ps: IPartido[]; al: IAlerta[]; movidos: { [id: number]: Date } } {
  const e = est.res && est.sel ? est.res.filter((x: IEscenario) => x.clave === est.sel)[0] : undefined;
  if (!e) { return { ps, al, movidos: {} }; }
  const movidos: { [id: number]: Date } = {}; e.cambios.forEach((c) => { movidos[c.id] = c.desde; });
  return { ps: e.partidos, al: e.alertas, movidos };
}

export function enlazarEscenarios(root: HTMLElement, ctx: IGCtx, est: IEscEstado, al: IAlerta[], pintar: () => void): void {
  const q = (s: string): NodeListOf<HTMLElement> => root.querySelectorAll(s) as NodeListOf<HTMLElement>;
  const recalc = (): void => calcular(ctx, est, pintar);
  const escSel = (): IEscenario | undefined => est.res ? est.res.filter((x: IEscenario) => x.clave === est.sel)[0] : undefined;
  const cambio = (id: number) => { const e = escSel(); return e ? e.cambios.filter((c) => c.id === id)[0] : undefined; }; // eslint-disable-line @typescript-eslint/explicit-function-return-type
  const pid = (id: number): IPartidoG => ctx.semana.filter((p: IPartidoG) => p.id === id)[0];
  q('[data-esc-calc]').forEach((b: HTMLElement) => b.addEventListener('click', recalc));
  q('[data-esc]').forEach((b: HTMLElement) => b.addEventListener('click', () => { est.sel = b.getAttribute('data-esc') || ''; pintar(); }));
  q('[data-esc-limpiar]').forEach((b: HTMLElement) => b.addEventListener('click', () => { est.forzados = {}; est.pendiente = true; pintar(); }));
  q('[data-esc-fij]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
    const id = +(b.getAttribute('data-esc-fij') || 0); const c = cambio(id); if (!c) { return; }
    est.forzados[id] = { fecha: c.hasta, lugar: c.lugar, pista: c.pista, tipoFranja: c.tipoFranja }; est.pendiente = true; pintar();
  }));
  q('[data-esc-sol]').forEach((b: HTMLElement) => b.addEventListener('click', () => { delete est.forzados[+(b.getAttribute('data-esc-sol') || 0)]; est.pendiente = true; pintar(); }));
  q('[data-esc-orig]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
    const p = pid(+(b.getAttribute('data-esc-orig') || 0)); if (!p) { return; }
    est.forzados[p.id] = { fecha: p.fecha, lugar: p.lugar, pista: '', tipoFranja: '' }; est.pendiente = true; pintar();
  }));
  q('[data-esc-otra]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
    const e = escSel(); const p = pid(+(b.getAttribute('data-esc-otra') || 0)); if (!e || !p) { return; }
    const actual = e.partidos.filter((x: IPartido) => x.id === p.id)[0] || p;
    const asign: { [id: number]: string } = {}; e.cambios.forEach((c) => { if (c.pista) { asign[c.id] = c.pista; } });
    const info = ctx.eqInfo(p);
    const hs: IHueco[] = sugerir(actual, e.partidos, ctx.reglas, ctx.pistas, ctx.margen, ctx.jugs, info.pref, info.alt, asign);
    const sel = document.createElement('select'); sel.className = 'cbp-select'; sel.style.cssText = 'min-height:34px;font-size:13px;max-width:260px';
    sel.innerHTML = '<option value="">Elige otra hora…</option>' + hs.map((h: IHueco, i: number) => '<option value="' + i + '">' + esc(DIA_TXT[h.fecha.getDay()].slice(0, 3) + ' ' + hm(minDia(h.fecha)) + (h.pista ? ' · ' + h.pista : '') + ' · ' + h.nota) + '</option>').join('') + (hs.length ? '' : '<option disabled>Sin huecos libres con el resto del escenario</option>');
    b.replaceWith(sel); sel.focus();
    sel.addEventListener('change', () => {
      const h = hs[+sel.value]; if (!h) { return; }
      const k = h.pista ? ctx.pistas.filter((x: IPista) => norm(x.clave) === norm(h.pista))[0] : undefined;
      est.forzados[p.id] = { fecha: h.fecha, lugar: k ? k.alias : p.lugar, pista: h.pista, tipoFranja: h.tipoFranja }; est.pendiente = true; pintar();
    });
  }));
  q('[data-esc-guardar]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
    const e = escSel(); if (!e) { return; }
    if (!window.confirm('Se guardarán ' + e.cambios.length + ' cambios como propuesta. ¿Seguir?')) { return; }
    (b as HTMLButtonElement).disabled = true;
    e.cambios.reduce((pr: Promise<void>, c) => pr.then(() => {
      const p = pid(c.id);
      const body: { [k: string]: any } = { FechaPropuesta: c.hasta.toISOString(), EstadoG: 'Propuesto', Motivo: justificacion(p, al) || 'Ajuste de horarios de la semana para evitar coincidencias.', OKEntrenador: '' }; // eslint-disable-line @typescript-eslint/no-explicit-any
      if (c.pista) { body.Pista = c.pista; }
      if (c.peticion) { body.PeticionHora = null; }
      return guardarGest(ctx, p, body, 'Escenario «' + e.nombre + '»: ' + fTxt(c.hasta) + (p.local && enPlazo({ ...p, fecha: p.oficial }, ctx.plazo) ? ' (en plazo)' : ''));
    }), Promise.resolve()).then(() => { est.res = null; est.forzados = {}; est.sel = ''; aviso('Escenario guardado: ' + e.cambios.length + ' propuestas pendientes del OK'); ctx.recargar(); })
      .catch((x: Error) => { (b as HTMLButtonElement).disabled = false; aviso('Error: ' + x.message, true); });
  }));
}

export const CSS_ESC = `
.cbp-esc{display:flex;flex-direction:column;gap:12px;padding:16px 18px}
.cbp-esc-chips{display:flex;gap:8px;flex-wrap:wrap}
.cbp-esc-chip{display:flex;flex-direction:column;align-items:flex-start;gap:2px;border:1.5px solid #E6E3DC;background:#fff;border-radius:12px;padding:8px 14px;cursor:pointer;font:inherit;text-align:left;min-width:160px}
.cbp-esc-chip b{font-size:14px}.cbp-esc-chip small{font-size:12px;color:#5c5c5c}
.cbp-esc-chip.mal small{color:#B00020}
.cbp-esc-chip.on{border-color:#090909;background:#090909;color:#fff}.cbp-esc-chip.on small{color:#FF9F00}
.cbp-plazo{display:inline-block;font-size:11px;font-weight:700;border-radius:999px;padding:2px 8px;background:#F1EEE8;color:#5c5c5c;white-space:nowrap}
.cbp-plazo.ok{background:#E8F5E9;color:#1B5E20}
.cbp-pet{color:#8A5300;font-weight:600}
.cbp-lock{border:1px solid #E6E3DC;background:#fff;border-radius:8px;width:28px;height:28px;cursor:pointer;font-size:13px;flex:none}
.cbp-lock.on{background:#090909;border-color:#090909}
.cbp-tabla tr.fij td{background:#FFF8EA}
.cbp-ajustes{display:flex;gap:6px;flex-wrap:wrap;align-items:center;font-size:13px}
.cbp-ajuste{display:inline-flex;align-items:center;gap:4px;background:#FFF1D6;color:#5a3800;border-radius:999px;padding:3px 4px 3px 10px;font-weight:600}
.cbp-ajuste button{border:0;background:transparent;cursor:pointer;font-size:12px;padding:0 6px}
.cbp-al-ico{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;font-weight:800;font-size:13px;color:#fff;background:#B00020}
.cbp-al-ico.aviso{background:#FF9F00;color:#090909}.cbp-al-ico.ok{background:#E8F5E9;color:#1B5E20}
.cbp-tabla.pt tr.con td{background:#FFF6F6}
.cbp-tabla.pt tr[data-fijo] td{background:#F6F5F1}
.cbp-cp{display:inline-flex;align-items:center;gap:6px}.cbp-cp code{font-family:inherit;font-weight:700;font-size:14px;background:#F6F5F1;border-radius:6px;padding:2px 6px}
.cbp-cp-b{border:1px solid #D9D5CC;background:#fff;border-radius:6px;width:28px;height:26px;cursor:pointer;font-size:13px}
.cbp-cp-b.ok{background:#E8F5E9;border-color:#1B5E20;color:#1B5E20}
.cbp-tl-p.mov{outline:2px dashed #FF9F00;outline-offset:1px}
.cbp-tl-p.ghost{opacity:.28;pointer-events:none;border-style:dashed}
`;
