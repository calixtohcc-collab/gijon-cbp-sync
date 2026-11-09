import { esc, IEvento, DIAS_LARGO, MESES_LARGO, hora } from '../../common/cbp';
import { spPost, spGet, listUrl, aviso } from '../../common/ui';
import { ICtx, IJugador, plantilla, selectorEquipo, enlazarComunes, sinPlantilla, parse } from './datos';

interface IConv { id: number; jugadores: number[]; citacion: string; notas: string; }

let partidoSel = 0;
export function preseleccionarPartido(id: number): void { partidoSel = id; }

export function renderConvocatorias(el: HTMLElement, ctx: ICtx): void {
  const rer = (): void => renderConvocatorias(el, ctx);
  if (partidoSel) {
    const p = ctx.eventos.filter((e: IEvento) => e.id === partidoSel)[0];
    if (p && p.equipo !== ctx.equipo) { ctx.setEquipo(p.equipo); }
  }
  const eq = ctx.equipo;
  const ahora = new Date(Date.now() - 6 * 3600000);
  const partidos = ctx.eventos.filter((e: IEvento) => e.equipo === eq && e.tipo === 'Partido' && e.fecha >= ahora).slice(0, 12);
  if (eq && partidos.length && !partidos.some((p: IEvento) => p.id === partidoSel)) { partidoSel = partidos[0].id; }
  const partido = partidos.filter((p: IEvento) => p.id === partidoSel)[0];
  const jug = eq ? plantilla(ctx, eq, true) : [];

  el.innerHTML = '<div class="cbp-sec" style="gap:16px">' +
    '<div style="display:flex;gap:14px;flex-wrap:wrap;align-items:flex-end">' + selectorEquipo(eq) +
    (eq ? '<label class="cbp-field" style="min-width:280px;flex:1">Partido<select data-partido="1">' + (partidos.length ? partidos.map((p: IEvento) =>
      '<option value="' + p.id + '"' + (p.id === partidoSel ? ' selected' : '') + '>' + p.fecha.getDate() + '/' + (p.fecha.getMonth() + 1) + ' ' + hora(p.fecha) + ' · ' + (p.local ? 'vs ' : 'en ') + esc(p.rival) + '</option>').join('') : '<option value="">Sin partidos próximos</option>') + '</select></label>' : '') +
    '</div><div data-cuerpo="1"></div></div>';
  enlazarComunes(el, ctx, () => { partidoSel = 0; rer(); });
  const sp = el.querySelector('[data-partido]') as HTMLSelectElement;
  if (sp) { sp.addEventListener('change', () => { partidoSel = +sp.value; rer(); }); }
  const cuerpo = el.querySelector('[data-cuerpo]') as HTMLElement;
  if (!eq) { cuerpo.innerHTML = '<div class="cbp-vacio">Elige un equipo para preparar la convocatoria.</div>'; return; }
  if (!jug.length) { cuerpo.innerHTML = sinPlantilla(ctx); enlazarComunes(cuerpo, ctx, rer); return; }
  if (!partido) { cuerpo.innerHTML = '<div class="cbp-vacio">Este equipo no tiene partidos próximos en el calendario.</div>'; return; }
  cuerpo.innerHTML = '<div class="cbp-cargando">Cargando…</div>';

  spGet(ctx.client, listUrl(ctx.w, 'Lists/Convocatorias') + '/items?$select=Id,JugadoresC,Citacion,NotasC&$filter=PartidoId eq ' + partido.id + '&$top=1')
    .then((j: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
      const it = (j.value || [])[0];
      const conv: IConv = it ? { id: it.Id, jugadores: parse<number[]>(it.JugadoresC, []), citacion: it.Citacion || '', notas: it.NotasC || '' } : { id: 0, jugadores: [], citacion: '', notas: '' };
      pintar(cuerpo, ctx, partido, jug, conv, rer);
    }).catch(() => pintar(cuerpo, ctx, partido, jug, { id: 0, jugadores: [], citacion: '', notas: '' }, rer));
}

function pintar(el: HTMLElement, ctx: ICtx, p: IEvento, jug: IJugador[], conv: IConv, rer: () => void): void {
  const sel: { [id: number]: boolean } = {};
  conv.jugadores.forEach((id: number) => { sel[id] = true; });
  const cuenta = (): number => jug.filter((j: IJugador) => sel[j.id]).length;
  const defCit = conv.citacion || (hora(new Date(p.fecha.getTime() - 45 * 60000)) + ' en ' + (p.local ? (p.lugar || 'el pabellón') : 'Pabellón La Perchera (salida)'));
  el.innerHTML = '<div class="cbp-card" style="padding:16px 18px;display:flex;gap:14px;align-items:center;flex-wrap:wrap">' +
    '<div class="cbp-fecha" aria-hidden="true"><small>' + ['DOM', 'LUN', 'MAR', 'MIÉ', 'JUE', 'VIE', 'SÁB'][p.fecha.getDay()] + '</small><b>' + p.fecha.getDate() + '</b><span>' + ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'][p.fecha.getMonth()] + '</span></div>' +
    '<div style="flex:1;min-width:200px"><div style="font-weight:700;font-size:17px">' + (p.local ? 'vs ' : 'en casa de ') + esc(p.rival) + '</div>' +
    '<div style="color:#5c5c5c;font-size:13px">' + hora(p.fecha) + ' · ' + esc(p.lugar || 'Lugar por confirmar') + '</div></div>' +
    '<div style="font-weight:700"><span data-cuenta="1">' + cuenta() + '</span> convocados</div></div>' +
    '<div style="display:flex;gap:8px;flex-wrap:wrap"><button type="button" class="cbp-btn sm" data-todos="1">Marcar todos</button><button type="button" class="cbp-btn sm" data-ninguno="1">Quitar todos</button></div>' +
    '<div class="cbp-players">' + jug.map((j: IJugador) => '<label class="cbp-player' + (sel[j.id] ? ' sel' : '') + '"><input type="checkbox" data-j="' + j.id + '"' + (sel[j.id] ? ' checked' : '') + '>' +
      '<span class="dor">' + (j.dorsal === undefined ? '–' : j.dorsal) + '</span><span class="nom">' + esc(j.nombre) + (j.sube ? ' <span class="cbp-sube' + (j.sube === 'oca' ? ' oca' : '') + '">' + esc(j.equipo) + '</span>' : '') + '</span></label>').join('') + '</div>' +
    '<div class="cbp-form"><label class="cbp-field">Citación<input data-cit="1" value="' + esc(defCit) + '"></label>' +
    '<label class="cbp-field">Notas para las familias<input data-notas="1" value="' + esc(conv.notas) + '" placeholder="Equipación negra, traer agua…"></label></div>' +
    '<div class="cbp-btns"><button type="button" class="cbp-btn" data-copiar="1">Copiar para WhatsApp</button><button type="button" class="cbp-btn pri" data-guardar="1">Guardar convocatoria</button></div>';

  const checks = el.querySelectorAll('[data-j]');
  const sync = (): void => { (el.querySelector('[data-cuenta]') as HTMLElement).textContent = String(cuenta()); };
  for (let i = 0; i < checks.length; i++) {
    checks[i].addEventListener('change', (e: Event) => {
      const c = e.currentTarget as HTMLInputElement; sel[+(c.getAttribute('data-j') || 0)] = c.checked;
      (c.parentNode as HTMLElement).classList.toggle('sel', c.checked); sync();
    });
  }
  const todos = (v: boolean): void => { for (let i = 0; i < checks.length; i++) { const c = checks[i] as HTMLInputElement; c.checked = v; sel[+(c.getAttribute('data-j') || 0)] = v; (c.parentNode as HTMLElement).classList.toggle('sel', v); } sync(); };
  (el.querySelector('[data-todos]') as HTMLElement).addEventListener('click', () => todos(true));
  (el.querySelector('[data-ninguno]') as HTMLElement).addEventListener('click', () => todos(false));
  const texto = (): string => {
    const lista = jug.filter((j: IJugador) => sel[j.id]).map((j: IJugador) => '• ' + (j.dorsal !== undefined ? j.dorsal + ' ' : '') + j.nombre).join('\n');
    const cit = (el.querySelector('[data-cit]') as HTMLInputElement).value; const notas = (el.querySelector('[data-notas]') as HTMLInputElement).value;
    return '🏀 CONVOCATORIA ' + p.equipo.toUpperCase() + '\n' + DIAS_LARGO[p.fecha.getDay()] + ' ' + p.fecha.getDate() + ' de ' + MESES_LARGO[p.fecha.getMonth()] + ' · ' + hora(p.fecha) + '\n' +
      (p.local ? 'vs ' : 'en ') + p.rival + '\n📍 ' + (p.lugar || 'por confirmar') + '\n⏰ Citación: ' + cit + '\n\n' + lista + (notas ? '\n\n' + notas : '') + '\n\nGijón CBP · Pumatas Basketball';
  };
  (el.querySelector('[data-copiar]') as HTMLElement).addEventListener('click', () => {
    const t = texto();
    if (navigator.clipboard) { navigator.clipboard.writeText(t).then(() => aviso('Convocatoria copiada. Pégala en el grupo.'), () => aviso('No se pudo copiar', true)); }
  });
  (el.querySelector('[data-guardar]') as HTMLElement).addEventListener('click', () => {
    const ids = jug.filter((j: IJugador) => sel[j.id]).map((j: IJugador) => j.id);
    const body = { Title: p.equipo + ' · ' + p.rival + ' · ' + p.fecha.getDate() + '/' + (p.fecha.getMonth() + 1), EquipoC: p.equipo, PartidoId: p.id, FechaPartido: p.fecha.toISOString(), RivalC: p.rival,
      Citacion: (el.querySelector('[data-cit]') as HTMLInputElement).value, NotasC: (el.querySelector('[data-notas]') as HTMLInputElement).value, JugadoresC: JSON.stringify(ids) };
    const url = listUrl(ctx.w, 'Lists/Convocatorias') + '/items' + (conv.id ? '(' + conv.id + ')' : '');
    spPost(ctx.client, url, body, !!conv.id).then(() => { aviso('Convocatoria guardada (' + ids.length + ' jugadores)'); rer(); }).catch((e: Error) => aviso('Error: ' + e.message, true));
  });
}
