import { esc, MESES_LARGO, DIAS_LARGO } from '../../common/cbp';
import { spPost, spGet, listUrl, aviso, fechaInput } from '../../common/ui';
import { ICtx, selectorEquipo, enlazarComunes, parse } from './datos';

interface IBloqueS { fase: string; ejercicio: string; matId?: number; min: number; notas: string; }
interface ISesion { id: number; titulo: string; fecha: string; dur: number; objetivo: string; bloques: IBloqueS[]; }
interface IMatLite { id: number; titulo: string; ruta: string; tipo: string; cat: string; }

const FASES = ['Calentamiento', 'Técnica individual', 'Táctica', 'Juego / competición', 'Vuelta a la calma'];
let editando: ISesion | undefined;
let materiales: IMatLite[] | undefined;

function nueva(): ISesion {
  return { id: 0, titulo: '', fecha: fechaInput(new Date()), dur: 90, objetivo: '', bloques: [
    { fase: 'Calentamiento', ejercicio: '', min: 15, notas: '' }, { fase: 'Técnica individual', ejercicio: '', min: 20, notas: '' },
    { fase: 'Táctica', ejercicio: '', min: 25, notas: '' }, { fase: 'Juego / competición', ejercicio: '', min: 20, notas: '' }, { fase: 'Vuelta a la calma', ejercicio: '', min: 10, notas: '' }] };
}

export function renderSesiones(el: HTMLElement, ctx: ICtx): void {
  const rer = (): void => renderSesiones(el, ctx);
  const eq = ctx.equipo;
  el.innerHTML = '<div class="cbp-sec" style="gap:16px">' + selectorEquipo(eq) + '<div data-cuerpo="1"></div></div>';
  enlazarComunes(el, ctx, () => { editando = undefined; rer(); });
  const cuerpo = el.querySelector('[data-cuerpo]') as HTMLElement;
  if (!eq) { cuerpo.innerHTML = '<div class="cbp-vacio">Elige un equipo para planificar sus sesiones.</div>'; return; }
  cuerpo.innerHTML = '<div class="cbp-cargando">Cargando…</div>';
  const pMat = materiales ? Promise.resolve(materiales) : spGet(ctx.client, listUrl(ctx.w, 'Materiales') + '/items?$select=Id,Title,FileLeafRef,FileRef,TipoMaterial,Categoria,FSObjType&$top=1000')
    .then((j: any) => (materiales = (j.value || []).filter((i: any) => +i.FSObjType === 0).map((i: any): IMatLite => ({ id: i.Id, titulo: i.Title || i.FileLeafRef, ruta: i.FileRef, tipo: i.TipoMaterial || '', cat: i.Categoria || '' })))) // eslint-disable-line @typescript-eslint/no-explicit-any
    .catch(() => (materiales = []));
  const pSes = spGet(ctx.client, listUrl(ctx.w, 'Lists/Sesiones') + "/items?$select=Id,Title,FechaS,Duracion,Objetivo,Bloques&$filter=EquipoS eq '" + encodeURIComponent(eq.replace(/'/g, "''")) + "'&$orderby=FechaS desc&$top=200")
    .then((j: any) => (j.value || []).map((i: any): ISesion => ({ id: i.Id, titulo: i.Title || '', fecha: (i.FechaS || '').slice(0, 10), dur: i.Duracion || 0, objetivo: i.Objetivo || '', bloques: parse<IBloqueS[]>(i.Bloques, []) }))) // eslint-disable-line @typescript-eslint/no-explicit-any
    .catch(() => []);
  Promise.all([pMat, pSes]).then((r: [IMatLite[], ISesion[]]) => pintar(cuerpo, ctx, r[0], r[1], rer));
}

function pintar(el: HTMLElement, ctx: ICtx, mats: IMatLite[], sesiones: ISesion[], rer: () => void): void {
  const s = editando || nueva();
  const ejercicios = mats.filter((m: IMatLite) => m.tipo === 'Ejercicio' || m.tipo === 'Sistema' || m.tipo === 'Planificación');
  const total = (): number => s.bloques.reduce((a: number, b: IBloqueS) => a + (+b.min || 0), 0);
  const bloqueHtml = (b: IBloqueS, i: number): string => '<div class="cbp-bloque" data-b="' + i + '">' +
    '<label class="cbp-field">Fase<select data-f="fase">' + FASES.map((f: string) => '<option' + (f === b.fase ? ' selected' : '') + '>' + f + '</option>').join('') + '</select></label>' +
    '<label class="cbp-field">Ejercicio<input data-f="ejercicio" list="cbp-ejs" value="' + esc(b.ejercicio) + '" placeholder="Elige de la biblioteca o escribe"></label>' +
    '<label class="cbp-field">Minutos<input data-f="min" type="number" min="0" max="180" inputmode="numeric" value="' + b.min + '"></label>' +
    '<button type="button" class="cbp-btn sm" data-quitar="' + i + '" aria-label="Quitar bloque" style="height:44px">✕</button>' +
    '<label class="cbp-field" style="grid-column:1/-1">Consignas<input data-f="notas" value="' + esc(b.notas) + '"></label></div>';
  el.innerHTML = '<div class="cbp-card cbp-panel" style="display:flex;flex-direction:column;gap:14px">' +
    '<div class="cbp-sec-head"><h3 class="cbp-h2" style="font-size:24px">' + (s.id ? 'Editar sesión' : 'Nueva sesión') + '</h3><span>Total: <b class="cbp-total" data-total="1">' + total() + '</b> min</span></div>' +
    '<div class="cbp-form"><label class="cbp-field">Fecha<input type="date" data-s="fecha" value="' + esc(s.fecha) + '"></label>' +
    '<label class="cbp-field">Objetivo<input data-s="objetivo" value="' + esc(s.objetivo) + '" placeholder="Ej.: Defensa del bloqueo directo"></label></div>' +
    '<datalist id="cbp-ejs">' + ejercicios.map((m: IMatLite) => '<option value="' + esc(m.titulo) + '">' + esc(m.tipo + ' · ' + m.cat) + '</option>').join('') + '</datalist>' +
    '<div style="display:flex;flex-direction:column;gap:10px" data-bloques="1">' + s.bloques.map(bloqueHtml).join('') + '</div>' +
    '<div class="cbp-btns" style="justify-content:space-between"><button type="button" class="cbp-btn" data-add="1">+ Añadir bloque</button>' +
    '<span style="display:flex;gap:10px;flex-wrap:wrap">' + (s.id ? '<button type="button" class="cbp-btn" data-nueva="1">Nueva sesión</button>' : '') +
    '<button type="button" class="cbp-btn" data-print="1">Imprimir</button><button type="button" class="cbp-btn pri" data-guardar="1">Guardar sesión</button></span></div></div>' +
    '<h3 class="cbp-h2" style="font-size:24px;margin-top:8px">Sesiones guardadas</h3>' +
    (sesiones.length ? '<div class="cbp-card cbp-lista">' + sesiones.map((x: ISesion) => {
      const d = new Date(x.fecha + 'T12:00:00');
      return '<button type="button" class="cbp-partido" data-abrir="' + x.id + '" style="border:0;background:none;width:100%;text-align:left;font:inherit;cursor:pointer;border-top:1px solid #EFECE6">' +
        '<div class="cbp-fecha" aria-hidden="true"><small>' + ['DOM', 'LUN', 'MAR', 'MIÉ', 'JUE', 'VIE', 'SÁB'][d.getDay()] + '</small><b>' + d.getDate() + '</b><span>' + MESES_LARGO[d.getMonth()].slice(0, 3).toUpperCase() + '</span></div>' +
        '<div class="cbp-pinfo"><span class="cbp-riv">' + esc(x.objetivo || x.titulo) + '</span><span class="cbp-meta">' + x.dur + ' min · ' + x.bloques.length + ' bloques</span></div></button>';
    }).join('') + '</div>' : '<div class="cbp-vacio">Todavía no hay sesiones guardadas para este equipo.</div>');

  const leer = (): void => {
    s.fecha = (el.querySelector('[data-s=fecha]') as HTMLInputElement).value; s.objetivo = (el.querySelector('[data-s=objetivo]') as HTMLInputElement).value;
    const bs = el.querySelectorAll('[data-b]');
    s.bloques = [];
    for (let i = 0; i < bs.length; i++) {
      const g = (f: string): string => (bs[i].querySelector('[data-f=' + f + ']') as HTMLInputElement).value;
      const ej = g('ejercicio'); const m = ejercicios.filter((x: IMatLite) => x.titulo === ej)[0];
      s.bloques.push({ fase: g('fase'), ejercicio: ej, matId: m ? m.id : undefined, min: +g('min') || 0, notas: g('notas') });
    }
    s.dur = total();
  };
  const actualizarTotal = (): void => { leer(); (el.querySelector('[data-total]') as HTMLElement).textContent = String(total()); };
  const mins = el.querySelectorAll('[data-f=min]');
  for (let i = 0; i < mins.length; i++) { mins[i].addEventListener('input', actualizarTotal); }
  const on = (sel: string, fn: (b: HTMLElement) => void): void => { const e = el.querySelectorAll(sel); for (let i = 0; i < e.length; i++) { e[i].addEventListener('click', (ev: Event) => fn(ev.currentTarget as HTMLElement)); } };
  on('[data-add]', () => { leer(); s.bloques.push({ fase: 'Táctica', ejercicio: '', min: 10, notas: '' }); editando = s; pintar(el, ctx, mats, sesiones, rer); });
  on('[data-quitar]', (b: HTMLElement) => { leer(); s.bloques.splice(+(b.getAttribute('data-quitar') || 0), 1); editando = s; pintar(el, ctx, mats, sesiones, rer); });
  on('[data-nueva]', () => { editando = undefined; pintar(el, ctx, mats, sesiones, rer); });
  on('[data-abrir]', (b: HTMLElement) => { editando = sesiones.filter((x: ISesion) => x.id === +(b.getAttribute('data-abrir') || 0))[0]; pintar(el, ctx, mats, sesiones, rer); window.scrollTo({ top: el.getBoundingClientRect().top + window.pageYOffset - 120, behavior: 'smooth' }); });
  on('[data-print]', () => { leer(); imprimir(s, ctx.equipo, mats, ctx.w); });
  on('[data-guardar]', () => {
    leer();
    const body = { Title: ctx.equipo + ' · ' + s.fecha, EquipoS: ctx.equipo, FechaS: s.fecha + 'T12:00:00Z', Duracion: s.dur, Objetivo: s.objetivo, Bloques: JSON.stringify(s.bloques) };
    spPost(ctx.client, listUrl(ctx.w, 'Lists/Sesiones') + '/items' + (s.id ? '(' + s.id + ')' : ''), body, !!s.id)
      .then(() => { aviso('Sesión guardada'); editando = undefined; rer(); }).catch((e: Error) => aviso('Error: ' + e.message, true));
  });
}

function imprimir(s: ISesion, equipo: string, mats: IMatLite[], w: string): void {
  const d = new Date(s.fecha + 'T12:00:00'); const origin = window.location.protocol + '//' + window.location.host;
  const filas = s.bloques.map((b: IBloqueS) => {
    const m = mats.filter((x: IMatLite) => x.id === b.matId)[0];
    return '<tr><td><b>' + esc(b.fase) + '</b></td><td>' + (m ? '<a href="' + esc(origin + m.ruta) + '">' + esc(b.ejercicio) + '</a>' : esc(b.ejercicio)) + (b.notas ? '<div class="n">' + esc(b.notas) + '</div>' : '') + '</td><td class="m">' + b.min + '\'</td></tr>';
  }).join('');
  const html = '<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Sesión ' + esc(equipo) + '</title><link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Anton&family=Inter:wght@400;600;700&display=swap"><style>' +
    'body{font-family:Inter,Arial,sans-serif;margin:32px;color:#090909}h1{font-family:Anton,Impact,sans-serif;font-weight:400;text-transform:uppercase;font-size:34px;margin:0}' +
    '.top{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:4px solid #FF9F00;padding-bottom:10px;margin-bottom:16px}.o{font-size:13px;color:#555}' +
    'table{width:100%;border-collapse:collapse;font-size:14px}td{padding:10px 8px;border-bottom:1px solid #ddd;vertical-align:top}td.m{text-align:right;font-weight:700;width:60px}.n{font-size:12px;color:#555;margin-top:3px}' +
    'tfoot td{font-weight:800;border-top:2px solid #090909;border-bottom:0}a{color:#090909}</style></head><body>' +
    '<div class="top"><div><div class="o">GIJÓN CBP · SESIÓN DE ENTRENAMIENTO</div><h1>' + esc(equipo) + '</h1><div>' + DIAS_LARGO[d.getDay()] + ' ' + d.getDate() + ' de ' + MESES_LARGO[d.getMonth()] + ' de ' + d.getFullYear() + '</div></div>' +
    '<div style="text-align:right"><div class="o">Objetivo</div><b>' + esc(s.objetivo || '—') + '</b></div></div>' +
    '<table><tbody>' + filas + '</tbody><tfoot><tr><td></td><td>Total</td><td class="m">' + s.dur + '\'</td></tr></tfoot></table>' +
    '<script>window.onload=function(){setTimeout(function(){window.print()},400)}</script></body></html>';
  const win = window.open('', '_blank');
  if (win) { win.document.write(html); win.document.close(); } else { aviso('Permite las ventanas emergentes para imprimir', true); }
  void w;
}
