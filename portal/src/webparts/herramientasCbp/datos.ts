import { SPHttpClient } from '@microsoft/sp-http';
import { IEvento } from '../../common/cbp';
import { opcionesEquipo } from '../../common/ui';
import { IJug, cargarJugs } from '../../common/jugadores';

export interface IJugador extends IJug { sube?: 'hab' | 'oca'; }

export interface ICtx {
  client: SPHttpClient;
  w: string;
  eventos: IEvento[];
  jugadores: IJugador[];
  equipo: string;
  setEquipo: (e: string) => void;
  recargarJugadores: () => Promise<void>;
  irA: (tab: string) => void;
}

export function cargarJugadores(client: SPHttpClient, w: string): Promise<IJugador[]> {
  return cargarJugs(client, w).then((js: IJug[]) => js.sort((a: IJug, b: IJug) => (a.dorsal === undefined ? 999 : a.dorsal) - (b.dorsal === undefined ? 999 : b.dorsal) || a.nombre.localeCompare(b.nombre)));
}

/** Plantilla de un equipo. Con «conVinculados» añade a quienes suben de otros equipos (habitual primero). */
export function plantilla(ctx: ICtx, equipo: string, conVinculados?: boolean): IJugador[] {
  const base: IJugador[] = ctx.jugadores.filter((j: IJugador) => j.equipo === equipo && j.activo);
  if (!conVinculados) { return base; }
  const hab = ctx.jugadores.filter((j: IJugador) => j.activo && j.equipo !== equipo && j.habitual.indexOf(equipo) >= 0).map((j: IJugador): IJugador => ({ ...j, sube: 'hab' }));
  const oca = ctx.jugadores.filter((j: IJugador) => j.activo && j.equipo !== equipo && j.habitual.indexOf(equipo) < 0 && j.ocasional.indexOf(equipo) >= 0).map((j: IJugador): IJugador => ({ ...j, sube: 'oca' }));
  return base.concat(hab, oca);
}

export function selectorEquipo(sel: string): string {
  return '<label class="cbp-field" style="max-width:320px">Equipo<select data-eq="1"><option value="">Elige un equipo…</option>' + opcionesEquipo(sel) + '</select></label>';
}

export function sinPlantilla(ctx: ICtx): string {
  return '<div class="cbp-vacio">Este equipo todavía no tiene jugadores dados de alta. <button type="button" class="cbp-btn sm" data-ira="plantillas" style="margin-left:8px">Añadir plantilla</button></div>';
}

export function enlazarComunes(el: HTMLElement, ctx: ICtx, rerender: () => void): void {
  const s = el.querySelector('[data-eq]') as HTMLSelectElement;
  if (s) { s.addEventListener('change', () => { ctx.setEquipo(s.value); rerender(); }); }
  const ir = el.querySelectorAll('[data-ira]');
  for (let i = 0; i < ir.length; i++) { ir[i].addEventListener('click', (e: Event) => ctx.irA((e.currentTarget as HTMLElement).getAttribute('data-ira') || '')); }
}

export function parse<T>(s: string, def: T): T { try { return s ? JSON.parse(s) as T : def; } catch (e) { return def; } }
