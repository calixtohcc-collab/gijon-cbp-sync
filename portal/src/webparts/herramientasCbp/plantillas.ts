import { esc } from '../../common/cbp';
import { spPost, listUrl, aviso } from '../../common/ui';
import { fichaJugador, faltan, telLlamar, telLink, edad, cargarEstilosJug } from '../../common/jugadores';
import { ICtx, IJugador, plantilla, selectorEquipo, enlazarComunes } from './datos';

export function renderPlantillas(el: HTMLElement, ctx: ICtx): void {
  cargarEstilosJug();
  const rer = (): void => renderPlantillas(el, ctx);
  const recargar = (): void => { ctx.recargarJugadores().then(rer); };
  const eq = ctx.equipo;
  const jug = eq ? plantilla(ctx, eq, true) : [];
  const propios = jug.filter((j: IJugador) => !j.sube);
  const incompletos = propios.filter((j: IJugador) => faltan(j).length).length;
  const fila = (j: IJugador): string => {
    const f = faltan(j); const t = telLlamar(j); const e = edad(j.nac);
    return '<div class="cbp-player" data-jug="' + j.id + '" style="cursor:pointer;grid-template-columns:auto minmax(0,1fr) auto auto"><span class="dor">' + (j.dorsal === undefined ? '–' : j.dorsal) + '</span>' +
      '<span class="nom">' + esc(j.nombre) + (j.sube ? ' <span class="cbp-sube' + (j.sube === 'oca' ? ' oca' : '') + '">' + (j.sube === 'oca' ? 'Ocasional' : 'Sube') + ' · ' + esc(j.equipo) + '</span>' : '') +
      '<br><small style="color:#5c5c5c;font-weight:500">' + (e !== undefined ? e + ' años' : '') + (j.habitual.length ? (e !== undefined ? ' · ' : '') + 'también ' + esc(j.habitual.join(', ')) : '') + '</small>' +
      (f.length && !j.sube ? '<span class="cbp-falta" style="display:block">Falta: ' + esc(f.join(', ')) + '</span>' : '') + '</span>' +
      (t ? '<a class="cbp-tel" href="' + esc(telLink(t.tel)) + '" title="Llamar a ' + esc(t.quien + ' · ' + t.tel) + '">☎</a>' : '<span></span>') + '<span></span></div>';
  };
  el.innerHTML = '<div class="cbp-sec" style="gap:16px">' +
    '<p style="margin:0;color:#5c5c5c">Plantilla de cada equipo: nombre y primer apellido, dorsal, fecha de nacimiento y contacto (un adulto en menores; teléfono propio desde los 16). Marca también con qué otros equipos juega cada persona: el módulo de partidos usa esos vínculos para detectar coincidencias. Solo lo ven el cuerpo técnico y la junta.</p>' +
    selectorEquipo(eq) +
    (eq ? '<div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center"><button type="button" class="cbp-btn pri" data-nuevo="1">+ Añadir jugador</button>' +
      '<label class="cbp-btn" style="cursor:pointer" title="Pega una lista con un nombre por línea («7 Laura García»)">Pegar lista<textarea data-pegar="1" style="position:absolute;opacity:0;width:1px;height:1px"></textarea></label>' +
      (incompletos ? '<span class="cbp-falta">' + incompletos + ' ficha' + (incompletos === 1 ? '' : 's') + ' incompleta' + (incompletos === 1 ? '' : 's') + '</span>' : '') + '</div>' +
      (jug.length ? '<div class="cbp-players">' + jug.map(fila).join('') + '</div>' : '<div class="cbp-vacio">Sin jugadores todavía.</div>') : '') +
    '</div>';
  enlazarComunes(el, ctx, rer);
  const nuevo = el.querySelector('[data-nuevo]') as HTMLElement;
  if (nuevo) { nuevo.addEventListener('click', () => fichaJugador(ctx.client, ctx.w, undefined, eq, recargar)); }
  el.querySelectorAll('[data-jug]').forEach((b: Element) => b.addEventListener('click', (ev: Event) => {
    if ((ev.target as HTMLElement).closest('a')) { return; }
    const j = jug.filter((x: IJugador) => x.id === +((b as HTMLElement).getAttribute('data-jug') || 0))[0];
    if (j) { fichaJugador(ctx.client, ctx.w, j, eq, recargar); }
  }));
  const pegar = el.querySelector('[data-pegar]') as HTMLTextAreaElement;
  if (pegar) {
    pegar.addEventListener('paste', (e: ClipboardEvent) => {
      const t = e.clipboardData ? e.clipboardData.getData('text') : '';
      e.preventDefault();
      const items = t.split(/\r?\n/).map((l: string) => l.trim()).filter((l: string) => l).map((l: string) => { const m = l.match(/^(\d{1,2})[\s.\-–]+(.+)$/); return m ? { Title: m[2], Dorsal: +m[1] } : { Title: l }; });
      if (!items.length) { return; }
      Promise.all(items.map((it: { Title: string; Dorsal?: number }) => spPost(ctx.client, listUrl(ctx.w, 'Lists/Jugadores') + '/items', { ...it, EquipoJ: eq, Activo: true })))
        .then(() => { aviso(items.length + ' jugadores añadidos. Completa su fecha de nacimiento y contacto.'); recargar(); })
        .catch((er: Error) => aviso('Error: ' + er.message, true));
    });
    pegar.parentElement!.addEventListener('click', () => { pegar.focus(); aviso('Ahora pega la lista (Ctrl+V)'); });
  }
}
