import { SPHttpClient } from '@microsoft/sp-http';
import { esc } from './cbp';
import { spPost, spGet, listUrl, abrirModal, cerrarModal, aviso, fechaInput } from './ui';

/* Jugadores: datos mínimos, contacto para menores y vínculos con otros equipos */
export const EQUIPOS_LISTA: string[] = ['Sénior Femenino A', 'Sénior Femenino B', 'Sénior Masculino A', 'Junior Femenino A', 'Junior Masculino A', 'Junior Masculino B',
  'Cadete Femenino A', 'Cadete Femenino B', 'Cadete Masculino A', 'Cadete Masculino B', 'Cadete Masculino C', 'Infantil Femenino A', 'Infantil Masculino A', 'Infantil Masculino B', 'Infantil Masculino C',
  'Alevín Femenino A', 'Alevín Femenino B', 'Alevín Masculino A', 'Alevín Masculino B', 'Alevín Masculino C', 'Benjamín Femenino', 'Benjamín Masculino A', 'Benjamín Masculino B', 'Benjamín Masculino C',
  'Tecnificación Minibasket', 'Tecnificación Canasta Grande', 'Equipo de Madres', 'Equipo de Padres'];

export interface IJug {
  id: number; nombre: string; equipo: string; dorsal?: number; activo: boolean; nac?: Date;
  contacto: string; contactoTel: string; tel: string; habitual: string[]; ocasional: string[];
}

export const SELECT_JUG = 'Id,Title,EquipoJ,Dorsal,Activo,FechaNac,ContactoNombre,ContactoTel,TelefonoPropio,EquiposHabitual,EquiposOcasional';
export function mapJug(i: any): IJug { // eslint-disable-line @typescript-eslint/no-explicit-any
  return { id: i.Id, nombre: i.Title || '', equipo: i.EquipoJ || '', dorsal: i.Dorsal === null || i.Dorsal === undefined ? undefined : i.Dorsal, activo: i.Activo !== false,
    nac: i.FechaNac ? new Date(i.FechaNac) : undefined, contacto: i.ContactoNombre || '', contactoTel: i.ContactoTel || '', tel: i.TelefonoPropio || '',
    habitual: i.EquiposHabitual || [], ocasional: i.EquiposOcasional || [] };
}
export function cargarJugs(client: SPHttpClient, w: string): Promise<IJug[]> {
  return spGet(client, listUrl(w, 'Lists/Jugadores') + '/items?$select=' + SELECT_JUG + '&$top=3000').then((j: any) => (j.value || []).map(mapJug)); // eslint-disable-line @typescript-eslint/no-explicit-any
}

export function edad(nac?: Date, ref?: Date): number | undefined {
  if (!nac) { return undefined; }
  const h = ref || new Date(); let e = h.getFullYear() - nac.getFullYear();
  if (h.getMonth() < nac.getMonth() || (h.getMonth() === nac.getMonth() && h.getDate() < nac.getDate())) { e--; }
  return e;
}
/** Qué falta según la edad: menores → contacto adulto; datos básicos siempre */
export function faltan(j: IJug): string[] {
  const f: string[] = []; const e = edad(j.nac);
  if (!j.nac) { f.push('fecha de nacimiento'); }
  if (e === undefined || e < 18) { if (!j.contacto || !j.contactoTel) { f.push('contacto adulto'); } }
  else if (!j.tel) { f.push('teléfono'); }
  return f;
}
export function telLink(t: string): string { return 'tel:' + t.replace(/[^\d+]/g, ''); }
/** Teléfono al que llamar: el propio si es mayor de 16 y lo tiene; si no, el del contacto */
export function telLlamar(j: IJug): { tel: string; quien: string } | undefined {
  const e = edad(j.nac);
  if (j.tel && e !== undefined && e >= 16) { return { tel: j.tel, quien: j.nombre }; }
  if (j.contactoTel) { return { tel: j.contactoTel, quien: j.contacto || 'Contacto' }; }
  if (j.tel) { return { tel: j.tel, quien: j.nombre }; }
  return undefined;
}

/** Ficha de alta o edición de un jugador */
export function fichaJugador(client: SPHttpClient, w: string, j: IJug | undefined, equipoDefecto: string, alGuardar: () => void): void {
  const marcas = (n: string, sel: string[], principal: string): string => '<div class="cbp-vinc">' + EQUIPOS_LISTA.filter((e: string) => e !== principal).map((e: string) =>
    '<label><input type="checkbox" name="' + n + '" value="' + esc(e) + '"' + (sel.indexOf(e) >= 0 ? ' checked' : '') + '>' + esc(e) + '</label>').join('') + '</div>';
  const eq = j ? j.equipo : equipoDefecto;
  const html = '<form class="cbp-form" data-fj="1" style="grid-template-columns:repeat(2,minmax(0,1fr))">' +
    '<label class="cbp-field full">Nombre y primer apellido<input name="nombre" required value="' + esc(j ? j.nombre : '') + '" autocomplete="off"></label>' +
    '<label class="cbp-field">Dorsal<input name="dorsal" type="number" min="0" max="99" inputmode="numeric" value="' + (j && j.dorsal !== undefined ? j.dorsal : '') + '"></label>' +
    '<label class="cbp-field">Fecha de nacimiento<input name="nac" type="date" required value="' + (j && j.nac ? fechaInput(j.nac) : '') + '"></label>' +
    '<label class="cbp-field full">Equipo principal<select name="equipo">' + EQUIPOS_LISTA.map((e: string) => '<option' + (e === eq ? ' selected' : '') + '>' + esc(e) + '</option>').join('') + '</select></label>' +
    '<p class="cbp-nota full" data-edad="1" style="margin:0"></p>' +
    '<fieldset class="full cbp-fs" data-menor="1"><legend>Contacto adulto (obligatorio en menores)</legend><div class="cbp-form" style="grid-template-columns:repeat(2,minmax(0,1fr))">' +
    '<label class="cbp-field">Nombre<input name="contacto" value="' + esc(j ? j.contacto : '') + '" placeholder="Madre, padre o tutor"></label>' +
    '<label class="cbp-field">Teléfono<input name="contactoTel" type="tel" value="' + esc(j ? j.contactoTel : '') + '" inputmode="tel"></label></div></fieldset>' +
    '<label class="cbp-field full" data-propio="1">Teléfono propio <span style="font-weight:400;color:#5c5c5c">(a partir de 16 años)</span><input name="tel" type="tel" value="' + esc(j ? j.tel : '') + '" inputmode="tel"></label>' +
    '<fieldset class="full cbp-fs"><legend>Juega también con · habitual</legend><p class="cbp-nota" style="margin:0 0 6px">Equipos con los que juega casi siempre. Cuentan como conflicto si los partidos se pisan.</p>' + marcas('hab', j ? j.habitual : [], eq) + '</fieldset>' +
    '<fieldset class="full cbp-fs"><legend>Convocatoria ocasional</legend><p class="cbp-nota" style="margin:0 0 6px">Equipos que le llaman a veces. Salen como aviso.</p>' + marcas('oca', j ? j.ocasional : [], eq) + '</fieldset>' +
    '<div class="cbp-btns full">' + (j ? '<button type="button" class="cbp-btn" data-baja="1" style="margin-right:auto;color:#B00020;border-color:#B00020">Dar de baja</button>' : '') +
    '<button type="button" class="cbp-btn" data-c="1">Cancelar</button><button type="submit" class="cbp-btn pri">' + (j ? 'Guardar' : 'Dar de alta') + '</button></div></form>';
  const b = abrirModal(j ? j.nombre : 'Nuevo jugador', html, 660);
  const f = b.querySelector('[data-fj]') as HTMLFormElement;
  const val = (n: string): string => (f.querySelector('[name="' + n + '"]') as HTMLInputElement).value.trim();
  const ajustar = (): void => {
    const s = val('nac'); const e = s ? edad(new Date(s + 'T12:00:00')) : undefined;
    const menor = e === undefined || e < 18; const propio = e !== undefined && e >= 16;
    (f.querySelector('[data-menor]') as HTMLElement).style.display = menor ? '' : 'none';
    (f.querySelector('[data-propio]') as HTMLElement).style.display = propio ? '' : 'none';
    (f.querySelector('[name="contacto"]') as HTMLInputElement).required = menor && e !== undefined;
    (f.querySelector('[name="contactoTel"]') as HTMLInputElement).required = menor && e !== undefined;
    (f.querySelector('[data-edad]') as HTMLElement).textContent = e === undefined ? 'Indica la fecha de nacimiento para saber qué datos de contacto hacen falta.' : e + ' años · ' + (e < 16 ? 'menor: solo contacto adulto.' : e < 18 ? 'menor: contacto adulto obligatorio y puede tener teléfono propio.' : 'mayor de edad: teléfono propio.');
  };
  (f.querySelector('[name="nac"]') as HTMLInputElement).addEventListener('change', ajustar); ajustar();
  (f.querySelector('[data-c]') as HTMLElement).addEventListener('click', cerrarModal);
  const url = listUrl(w, 'Lists/Jugadores') + '/items';
  const baja = f.querySelector('[data-baja]') as HTMLElement;
  if (baja && j) {
    baja.addEventListener('click', () => {
      if (!window.confirm('¿Dar de baja a ' + j.nombre + '? Se conserva su historial de asistencia.')) { return; }
      spPost(client, url + '(' + j.id + ')', { Activo: false }, true).then(() => { cerrarModal(); aviso(j.nombre + ' dado de baja'); alGuardar(); }).catch((x: Error) => aviso('Error: ' + x.message, true));
    });
  }
  f.addEventListener('submit', (ev: Event) => {
    ev.preventDefault();
    const s = val('nac'); const e = s ? edad(new Date(s + 'T12:00:00')) : undefined;
    const chk = (n: string): string[] => { const o: string[] = []; f.querySelectorAll('[name="' + n + '"]:checked').forEach((x: Element) => o.push((x as HTMLInputElement).value)); return o; };
    const principal = val('equipo');
    const hab = chk('hab').filter((x: string) => x !== principal); const oca = chk('oca').filter((x: string) => x !== principal && hab.indexOf(x) < 0);
    const body: { [k: string]: any } = { // eslint-disable-line @typescript-eslint/no-explicit-any
      Title: val('nombre'), EquipoJ: principal, FechaNac: s ? s + 'T12:00:00Z' : null, Activo: true,
      ContactoNombre: e !== undefined && e >= 18 ? '' : val('contacto'), ContactoTel: e !== undefined && e >= 18 ? '' : val('contactoTel'),
      TelefonoPropio: e !== undefined && e >= 16 ? val('tel') : '', EquiposHabitual: hab, EquiposOcasional: oca,
      Dorsal: val('dorsal') === '' ? null : +val('dorsal')
    };
    spPost(client, url + (j ? '(' + j.id + ')' : ''), body, !!j).then(() => { cerrarModal(); aviso(j ? 'Ficha guardada' : body.Title + ' dado de alta'); alGuardar(); })
      .catch((x: Error) => aviso('Error: ' + x.message, true));
  });
}

export const CSS_JUG = `
.cbp-fs{border:1px solid #E6E3DC;border-radius:12px;padding:10px 14px 12px;margin:0;min-width:0}
.cbp-fs legend{font-size:13px;font-weight:700;padding:0 6px}
.cbp-vinc{display:flex;flex-wrap:wrap;gap:6px}
.cbp-vinc label{display:inline-flex;align-items:center;gap:6px;font-size:12.5px;font-weight:600;border:1px solid #E6E3DC;border-radius:999px;padding:5px 10px;cursor:pointer;background:#fff}
.cbp-vinc label:has(input:checked){background:#090909;color:#FF9F00;border-color:#090909}
.cbp-vinc input{accent-color:#FF9F00;margin:0}
.cbp-sube{font-size:11px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;border-radius:999px;padding:2px 8px;background:#FFF1D6;color:#8A5300;white-space:nowrap}
.cbp-sube.oca{background:#F1EEE8;color:#5c5c5c}
.cbp-falta{font-size:11px;font-weight:700;color:#B00020;white-space:nowrap}
.cbp-tel{display:inline-flex;align-items:center;justify-content:center;width:32px;height:32px;border-radius:8px;border:1px solid #E6E3DC;color:#090909;text-decoration:none;font-size:15px}
.cbp-tel:hover{border-color:#090909}
`;
export function cargarEstilosJug(): void {
  if (document.getElementById('cbp-cssj')) { return; }
  const s = document.createElement('style'); s.id = 'cbp-cssj'; s.textContent = CSS_JUG; document.head.appendChild(s);
}
