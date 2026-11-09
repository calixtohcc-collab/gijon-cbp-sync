import { esc } from '../../common/cbp';
import { spPost, spGet, listUrl, aviso, fechaInput } from '../../common/ui';
import { ICtx, IJugador, plantilla, selectorEquipo, enlazarComunes, sinPlantilla, parse } from './datos';

interface ISesionA { id: number; fecha: string; p: number[]; a: number[]; j: number[]; }
let fechaSel = fechaInput(new Date());

export function renderAsistencia(el: HTMLElement, ctx: ICtx): void {
  const rer = (): void => renderAsistencia(el, ctx);
  const eq = ctx.equipo;
  el.innerHTML = '<div class="cbp-sec" style="gap:16px"><div style="display:flex;gap:14px;flex-wrap:wrap;align-items:flex-end">' + selectorEquipo(eq) +
    (eq ? '<label class="cbp-field" style="max-width:200px">Fecha<input type="date" data-fecha="1" value="' + fechaSel + '"></label>' : '') + '</div><div data-cuerpo="1"></div></div>';
  enlazarComunes(el, ctx, rer);
  const fi = el.querySelector('[data-fecha]') as HTMLInputElement;
  if (fi) { fi.addEventListener('change', () => { fechaSel = fi.value || fechaInput(new Date()); rer(); }); }
  const cuerpo = el.querySelector('[data-cuerpo]') as HTMLElement;
  if (!eq) { cuerpo.innerHTML = '<div class="cbp-vacio">Elige un equipo para pasar lista.</div>'; return; }
  const jug = plantilla(ctx, eq);
  if (!jug.length) { cuerpo.innerHTML = sinPlantilla(ctx); enlazarComunes(cuerpo, ctx, rer); return; }
  cuerpo.innerHTML = '<div class="cbp-cargando">Cargando…</div>';
  spGet(ctx.client, listUrl(ctx.w, 'Lists/Asistencia') + "/items?$select=Id,FechaA,Presentes,Ausentes,Justificadas&$filter=EquipoA eq '" + encodeURIComponent(eq.replace(/'/g, "''")) + "'&$orderby=FechaA desc&$top=500")
    .then((j: any) => (j.value || []).map((i: any): ISesionA => ({ id: i.Id, fecha: (i.FechaA || '').slice(0, 10), // eslint-disable-line @typescript-eslint/no-explicit-any
      p: parse<number[]>(i.Presentes, []), a: parse<number[]>(i.Ausentes, []), j: parse<number[]>(i.Justificadas, []) })))
    .then((ses: ISesionA[]) => pintar(cuerpo, ctx, jug, ses, rer))
    .catch(() => pintar(cuerpo, ctx, jug, [], rer));
}

function claveLocal(iso: string): string {
  // FechaA viene en UTC; la fecha local puede ser el día anterior a medianoche UTC
  if (!iso) { return ''; }
  const d = new Date(iso.length === 10 ? iso + 'T12:00:00Z' : iso);
  return fechaInput(d);
}

function pintar(el: HTMLElement, ctx: ICtx, jug: IJugador[], sesiones: ISesionA[], rer: () => void): void {
  const hoy = sesiones.filter((s: ISesionA) => claveLocal(s.fecha) === fechaSel)[0];
  const estado: { [id: number]: string } = {};
  jug.forEach((j: IJugador) => { estado[j.id] = 'p'; });
  if (hoy) { hoy.a.forEach((id: number) => { estado[id] = 'a'; }); hoy.j.forEach((id: number) => { estado[id] = 'j'; }); hoy.p.forEach((id: number) => { estado[id] = 'p'; }); }
  const fila = (j: IJugador): string => '<div class="cbp-player"><span class="dor">' + (j.dorsal === undefined ? '–' : j.dorsal) + '</span><span class="nom">' + esc(j.nombre) + '</span>' +
    '<span class="cbp-tri" role="group" aria-label="Asistencia de ' + esc(j.nombre) + '">' +
    ['p', 'a', 'j'].map((k: string) => '<button type="button" class="' + k + '" data-j="' + j.id + '" data-k="' + k + '" aria-pressed="' + (estado[j.id] === k ? 'true' : 'false') + '" title="' + ({ p: 'Presente', a: 'Ausente', j: 'Justificada' } as { [k: string]: string })[k] + '">' + k.toUpperCase() + '</button>').join('') + '</span></div>';

  // Estadísticas
  const tot = sesiones.length;
  const stats = jug.map((j: IJugador) => {
    const p = sesiones.filter((s: ISesionA) => s.p.indexOf(j.id) > -1).length;
    const jj = sesiones.filter((s: ISesionA) => s.j.indexOf(j.id) > -1).length;
    const a = sesiones.filter((s: ISesionA) => s.a.indexOf(j.id) > -1).length;
    const base = p + a + jj; const pct = base ? Math.round(p * 100 / base) : 0;
    return { j, p, a, jj, pct, base };
  }).sort((x: { pct: number }, y: { pct: number }) => x.pct - y.pct);

  el.innerHTML = '<p style="margin:0;font-size:13px;color:#5c5c5c">' + (hoy ? 'Ya hay una lista guardada para este día: puedes corregirla.' : 'Por defecto todos aparecen como presentes. Marca A (ausente) o J (justificada) donde corresponda.') + '</p>' +
    '<div class="cbp-players">' + jug.map(fila).join('') + '</div>' +
    '<div class="cbp-btns"><button type="button" class="cbp-btn pri" data-guardar="1">Guardar asistencia</button></div>' +
    '<h3 class="cbp-h2" style="font-size:24px;margin-top:10px">Resumen de la temporada</h3>' +
    (tot ? '<p style="margin:0;color:#5c5c5c">' + tot + ' sesiones registradas</p><div class="cbp-scroll"><table class="cbp-tabla"><thead><tr><th>Jugador/a</th><th>Asistencia</th><th>P</th><th>A</th><th>J</th></tr></thead><tbody>' +
      stats.map((s: { j: IJugador; p: number; a: number; jj: number; pct: number; base: number }) => '<tr><td>' + esc(s.j.nombre) + '</td><td style="min-width:150px"><div style="display:flex;gap:8px;align-items:center"><div class="cbp-barra" style="flex:1"><i style="width:' + s.pct + '%"></i></div><b>' + (s.base ? s.pct + '%' : '–') + '</b></div></td><td>' + s.p + '</td><td>' + s.a + '</td><td>' + s.jj + '</td></tr>').join('') +
      '</tbody></table></div>' : '<div class="cbp-vacio">Aún no hay sesiones registradas para este equipo.</div>');

  const bts = el.querySelectorAll('[data-k]');
  for (let i = 0; i < bts.length; i++) {
    bts[i].addEventListener('click', (e: Event) => {
      const b = e.currentTarget as HTMLElement; const id = +(b.getAttribute('data-j') || 0); estado[id] = b.getAttribute('data-k') || 'p';
      const grupo = b.parentNode as HTMLElement; const hs = grupo.querySelectorAll('button');
      for (let k = 0; k < hs.length; k++) { hs[k].setAttribute('aria-pressed', hs[k] === b ? 'true' : 'false'); }
    });
  }
  (el.querySelector('[data-guardar]') as HTMLElement).addEventListener('click', () => {
    const ids = (k: string): number[] => jug.filter((j: IJugador) => estado[j.id] === k).map((j: IJugador) => j.id);
    const body = { Title: ctx.equipo + ' · ' + fechaSel, EquipoA: ctx.equipo, FechaA: fechaSel + 'T12:00:00Z', Presentes: JSON.stringify(ids('p')), Ausentes: JSON.stringify(ids('a')), Justificadas: JSON.stringify(ids('j')) };
    const url = listUrl(ctx.w, 'Lists/Asistencia') + '/items' + (hoy ? '(' + hoy.id + ')' : '');
    spPost(ctx.client, url, body, !!hoy).then(() => { aviso('Asistencia guardada: ' + ids('p').length + ' presentes'); rer(); }).catch((e: Error) => aviso('Error: ' + e.message, true));
  });
}
