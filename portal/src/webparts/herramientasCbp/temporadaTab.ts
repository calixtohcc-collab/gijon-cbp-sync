import { esc } from '../../common/cbp';
import { spPost, listUrl, aviso } from '../../common/ui';
import { EQUIPOS_LISTA, IJug } from '../../common/jugadores';
import { temporadaActual, nombreTemporada, inicioDe, sugerirEquipo, cargarTemporada, calcularCierre, guardarCierre, esCoordinacion, pct, ICierre, ITramo } from '../../common/temporada';
import { ICtx } from './datos';

/* Asistente de cambio de temporada (solo coordinación y junta):
   1) Cerrar: guarda en «Trayectoria» la foto de cada jugador (equipo, vínculos, dorsal, asistencia, convocatorias).
   2) Preparar: propone el equipo de la temporada siguiente por año de nacimiento y lo aplica a la plantilla. */

let temporadaSel = '';
let permiso: boolean | undefined;

export function renderTemporada(el: HTMLElement, ctx: ICtx): void {
  if (permiso === undefined) {
    el.innerHTML = '<div class="cbp-cargando">Comprobando permisos…</div>';
    esCoordinacion(ctx.client, ctx.w).then((p: boolean) => { permiso = p; renderTemporada(el, ctx); });
    return;
  }
  if (!permiso) { el.innerHTML = '<div class="cbp-vacio">El cambio de temporada lo hacen coordinación deportiva y la junta.</div>'; return; }
  const actual = temporadaActual();
  const anterior = nombreTemporada(inicioDe(actual) - 1);
  if (!temporadaSel) { temporadaSel = actual; }
  const siguiente = nombreTemporada(inicioDe(temporadaSel) + 1);
  el.innerHTML = '<div class="cbp-sec" style="gap:18px">' +
    '<p style="margin:0;color:#5c5c5c">Al acabar cada temporada, guarda la trayectoria de cada jugador y prepara la plantilla de la siguiente. Hazlo en verano, cuando ya no queden partidos: el paso 2 cambia los equipos de la plantilla actual.</p>' +
    '<label class="cbp-field" style="max-width:260px">Temporada que se cierra<select data-tem="1">' + [actual, anterior].map((t: string) => '<option' + (t === temporadaSel ? ' selected' : '') + '>' + t + '</option>').join('') + '</select></label>' +
    '<div class="cbp-card" style="padding:16px 18px;box-shadow:none"><h3 class="cbp-h3" style="margin:0 0 6px">1 · Cerrar la temporada ' + esc(temporadaSel) + '</h3>' +
    '<p class="cbp-nota" style="margin:0 0 10px">Guarda para cada jugador su equipo, vínculos, dorsal, entrenamientos, presencias y convocatorias de la temporada. Se puede repetir: actualiza sin duplicar.</p>' +
    '<div data-cierre="1"><button type="button" class="cbp-btn pri" data-calc="1">Calcular resumen</button></div></div>' +
    '<div class="cbp-card" style="padding:16px 18px;box-shadow:none"><h3 class="cbp-h3" style="margin:0 0 6px">2 · Preparar la temporada ' + esc(siguiente) + '</h3>' +
    '<p class="cbp-nota" style="margin:0 0 10px">Propone a cada jugador el equipo que le toca por año de nacimiento (misma rama y letra). Revisa, cambia lo que haga falta y aplica. Los vínculos con otros equipos se vacían: se marcan de nuevo cada temporada.</p>' +
    '<div data-prep="1"><button type="button" class="cbp-btn" data-prop="1">Ver propuesta</button></div></div>' +
    '</div>';
  (el.querySelector('[data-tem]') as HTMLSelectElement).addEventListener('change', (e: Event) => { temporadaSel = (e.target as HTMLSelectElement).value; renderTemporada(el, ctx); });
  (el.querySelector('[data-calc]') as HTMLElement).addEventListener('click', () => cierre(el.querySelector('[data-cierre]') as HTMLElement, ctx));
  (el.querySelector('[data-prop]') as HTMLElement).addEventListener('click', () => propuesta(el.querySelector('[data-prep]') as HTMLElement, ctx, siguiente, () => renderTemporada(el, ctx)));
}

function cierre(box: HTMLElement, ctx: ICtx): void {
  box.innerHTML = '<div class="cbp-cargando">Calculando asistencia y convocatorias…</div>';
  const t = temporadaSel;
  Promise.all([calcularCierre(ctx.client, ctx.w, ctx.jugadores, t), cargarTemporada(ctx.client, ctx.w, t)]).then((r: [ICierre[], ITramo[]]) => {
    // Activos, y quienes se dieron de baja pero tuvieron actividad esa temporada
    const datos = r[0].filter((d: ICierre) => d.jugador.activo || d.sesiones || d.convocatorias);
    const yaHay = r[1].length;
    const porEq: { [e: string]: ICierre[] } = {};
    datos.forEach((d: ICierre) => { (porEq[d.jugador.equipo] = porEq[d.jugador.equipo] || []).push(d); });
    const eqs = EQUIPOS_LISTA.filter((e: string) => porEq[e]).concat(Object.keys(porEq).filter((e: string) => EQUIPOS_LISTA.indexOf(e) < 0));
    const tot = (ds: ICierre[]): { sesiones: number; presencias: number } => ds.reduce((a, d) => ({ sesiones: a.sesiones + d.sesiones, presencias: a.presencias + d.presencias }), { sesiones: 0, presencias: 0 });
    box.innerHTML = '<div class="cbp-scroll"><table class="cbp-tabla"><thead><tr><th>Equipo</th><th>Jugadores</th><th>Asistencia media</th><th>Convocatorias</th></tr></thead><tbody>' +
      eqs.map((e: string) => { const ds = porEq[e]; return '<tr><td><b>' + esc(e) + '</b></td><td>' + ds.length + '</td><td>' + pct(tot(ds)) + '</td><td>' + ds.reduce((a: number, d: ICierre) => a + d.convocatorias, 0) + '</td></tr>'; }).join('') +
      '</tbody></table></div>' +
      '<div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:12px"><button type="button" class="cbp-btn pri" data-guardar="1">Guardar trayectoria de ' + datos.length + ' jugadores</button>' +
      (yaHay ? '<span class="cbp-nota">Ya hay ' + yaHay + ' guardados de ' + esc(t) + ': se actualizarán.</span>' : '') + '</div>';
    (box.querySelector('[data-guardar]') as HTMLButtonElement).addEventListener('click', (ev: Event) => {
      const b = ev.currentTarget as HTMLButtonElement; b.disabled = true;
      guardarCierre(ctx.client, ctx.w, t, datos, r[1], (n: number) => { b.textContent = 'Guardando… ' + n + '/' + datos.length; })
        .then((n: number) => { aviso('Trayectoria ' + t + ' guardada: ' + n + ' jugadores'); b.textContent = 'Guardado ✓'; })
        .catch((x: Error) => { b.disabled = false; b.textContent = 'Reintentar'; aviso('Error: ' + x.message, true); });
    });
  }).catch((x: Error) => { box.innerHTML = '<div class="cbp-vacio">No se pudo calcular: ' + esc(x.message) + '</div>'; });
}

function propuesta(box: HTMLElement, ctx: ICtx, siguiente: string, rer: () => void): void {
  box.innerHTML = '<div class="cbp-cargando">Comprobando que la temporada está cerrada…</div>';
  cargarTemporada(ctx.client, ctx.w, temporadaSel).then((cerrados: ITramo[]) => {
    const activos = ctx.jugadores.filter((j: IJug) => j.activo);
    if (!cerrados.length) { box.innerHTML = '<div class="cbp-vacio">Primero cierra la temporada ' + esc(temporadaSel) + ' (paso 1) para no perder su trayectoria.</div>'; return; }
    const ini = inicioDe(siguiente);
    const prop: { [id: number]: string } = {};
    activos.forEach((j: IJug) => { prop[j.id] = sugerirEquipo(j, ini, EQUIPOS_LISTA); });
    const cambian = activos.filter((j: IJug) => prop[j.id] !== j.equipo).length;
    const porEq: { [e: string]: IJug[] } = {};
    activos.forEach((j: IJug) => { (porEq[j.equipo] = porEq[j.equipo] || []).push(j); });
    const eqs = EQUIPOS_LISTA.filter((e: string) => porEq[e]);
    const opciones = (sel: string): string => '<option value="__baja">Deja el club</option>' + EQUIPOS_LISTA.map((e: string) => '<option' + (e === sel ? ' selected' : '') + '>' + esc(e) + '</option>').join('');
    box.innerHTML = '<p class="cbp-nota" style="margin:0 0 10px"><b>' + cambian + '</b> de ' + activos.length + ' jugadores cambian de equipo según su edad. Marca «Deja el club» para quien no renueva.</p>' +
      eqs.map((e: string) => '<details class="cbp-det"' + (porEq[e].some((j: IJug) => prop[j.id] !== j.equipo) ? ' open' : '') + '><summary><b>' + esc(e) + '</b> · ' + porEq[e].length + ' jugadores</summary>' +
        '<div class="cbp-scroll"><table class="cbp-tabla"><tbody>' + porEq[e].sort((a: IJug, b: IJug) => a.nombre.localeCompare(b.nombre)).map((j: IJug) =>
          '<tr><td>' + esc(j.nombre) + (j.nac ? ' <small style="color:#5c5c5c">' + j.nac.getFullYear() + '</small>' : '') + '</td><td style="width:260px"><select class="cbp-sel-sm" data-j="' + j.id + '">' + opciones(prop[j.id]) + '</select></td></tr>').join('') +
        '</tbody></table></div></details>').join('') +
      '<div style="display:flex;gap:10px;align-items:center;margin-top:12px"><button type="button" class="cbp-btn pri" data-aplicar="1">Aplicar a la plantilla ' + esc(siguiente) + '</button></div>';
    (box.querySelector('[data-aplicar]') as HTMLButtonElement).addEventListener('click', (ev: Event) => {
      const cambios: { j: IJug; a: string }[] = [];
      box.querySelectorAll('select[data-j]').forEach((s: Element) => {
        const j = activos.filter((x: IJug) => x.id === +((s as HTMLElement).getAttribute('data-j') || 0))[0]; const a = (s as HTMLSelectElement).value;
        if (j && (a !== j.equipo || j.habitual.length || j.ocasional.length)) { cambios.push({ j, a }); }
      });
      const bajas = cambios.filter((c) => c.a === '__baja').length;
      if (!window.confirm('Se actualizarán ' + cambios.length + ' fichas (' + bajas + ' bajas) para la temporada ' + siguiente + '. ¿Seguir?')) { return; }
      const b = ev.currentTarget as HTMLButtonElement; b.disabled = true;
      const url = listUrl(ctx.w, 'Lists/Jugadores') + '/items';
      let n = 0; let i = 0;
      const uno = (c: { j: IJug; a: string }): Promise<void> => spPost(ctx.client, url + '(' + c.j.id + ')', c.a === '__baja' ? { Activo: false } :
        { EquipoJ: c.a, EquiposHabitual: [], EquiposOcasional: [], Dorsal: c.a === c.j.equipo ? (c.j.dorsal === undefined ? null : c.j.dorsal) : null }, true).then(() => { n++; b.textContent = 'Aplicando… ' + n + '/' + cambios.length; });
      const lote = (): Promise<void> => { const p = cambios.slice(i, i + 8); i += 8; return p.length ? Promise.all(p.map(uno)).then(lote) : Promise.resolve(); };
      lote().then(() => ctx.recargarJugadores()).then(() => { aviso('Plantilla ' + siguiente + ' preparada: ' + n + ' fichas actualizadas'); rer(); })
        .catch((x: Error) => { b.disabled = false; b.textContent = 'Reintentar'; aviso('Error: ' + x.message, true); });
    });
  }).catch((x: Error) => { box.innerHTML = '<div class="cbp-vacio">Error: ' + esc(x.message) + '</div>'; });
}
