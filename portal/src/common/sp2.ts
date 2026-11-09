import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { esc, MESES_LARGO, DIAS_LARGO } from './cbp';
import { spPost, abrirModal, cerrarModal, aviso, fechaInput } from './ui';

/* ---------- Acceso a listas por título ---------- */
export function porTitulo(w: string, titulo: string): string {
  return w + "/_api/web/lists/getbytitle('" + encodeURIComponent(titulo.replace(/'/g, "''")) + "')";
}

export function getJ(client: SPHttpClient, url: string): Promise<any> { // eslint-disable-line @typescript-eslint/no-explicit-any
  return client.get(url, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' } })
    .then((r: SPHttpClientResponse) => { if (!r.ok) { const e = new Error('HTTP ' + r.status) as Error & { status?: number }; e.status = r.status; throw e; } return r.json(); });
}

/** Sube un archivo a una carpeta y rellena sus columnas */
export function subirArchivo(client: SPHttpClient, w: string, carpeta: string, f: File, campos?: { [k: string]: any }): Promise<void> { // eslint-disable-line @typescript-eslint/no-explicit-any
  const nombre = f.name.replace(/[#%*:<>?/\\|"]/g, '-');
  return client.post(w + "/_api/web/GetFolderByServerRelativeUrl('" + encodeURIComponent(carpeta.replace(/'/g, "''")) + "')/Files/add(url='" + encodeURIComponent(nombre.replace(/'/g, "''")) + "',overwrite=true)?$expand=ListItemAllFields",
    SPHttpClient.configurations.v1, { body: f, headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' } })
    .then((r: SPHttpClientResponse) => { if (!r.ok) { return r.text().then((t: string) => { throw new Error('HTTP ' + r.status + ' ' + t.slice(0, 160)); }); } return r.json(); })
    .then((j: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
      if (!campos || !Object.keys(campos).length) { return undefined; }
      const it = j.ListItemAllFields;
      return spPost(client, w + "/_api/web/GetFolderByServerRelativeUrl('" + encodeURIComponent(carpeta.replace(/'/g, "''")) + "')/Files('" + encodeURIComponent(nombre.replace(/'/g, "''")) + "')/ListItemAllFields", campos, true)
        .then(() => it);
    }).then(() => undefined);
}

/* ---------- Fechas y formato ---------- */
const DIA = 86400000;
export function inicioDia(d: Date): number { return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime(); }
export function diasHasta(d: Date): number { return Math.round((inicioDia(d) - inicioDia(new Date())) / DIA); }
export function fCorta(d: Date): string { return d.getDate() + ' ' + MESES_LARGO[d.getMonth()].slice(0, 3); }
export function fLarga(d: Date): string { return DIAS_LARGO[d.getDay()] + ' ' + d.getDate() + ' de ' + MESES_LARGO[d.getMonth()]; }
export function hhmm(d: Date): string { return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2); }
export function euros(n: number): string { return (Math.round(n)).toLocaleString('es-ES') + ' €'; }
export function relativo(d: Date): string {
  const n = -diasHasta(d);
  if (n <= 0) { return 'hoy'; } if (n === 1) { return 'ayer'; } if (n < 30) { return 'hace ' + n + ' días'; }
  return 'el ' + fCorta(d);
}
export function cuenta(d: Date): string {
  const n = diasHasta(d);
  if (n < 0) { return 'hace ' + (-n) + ' día' + (n === -1 ? '' : 's'); }
  if (n === 0) { return 'hoy'; } if (n === 1) { return 'mañana'; }
  return 'en ' + n + ' días';
}
export function iniciales(n: string): string {
  return n.replace(/\s*-\s*Gij[oó]n CBP.*$/i, '').split(/\s+/).filter((x: string) => x).slice(0, 2).map((x: string) => x.charAt(0).toUpperCase()).join('');
}
export function localDT(d: Date): string { return fechaInput(d) + 'T' + hhmm(d); }

/* ---------- Formularios en ventana ---------- */
export interface ICampo {
  n: string; label: string; tipo?: 'text' | 'area' | 'date' | 'datetime' | 'select' | 'number' | 'check' | 'email';
  valor?: string | number | boolean; opciones?: string[]; full?: boolean; req?: boolean; ayuda?: string; ph?: string;
}
export function formModal(titulo: string, campos: ICampo[], boton: string, guardar: (v: { [k: string]: any }) => Promise<any>, ancho?: number, extra?: string): HTMLElement { // eslint-disable-line @typescript-eslint/no-explicit-any
  const html = '<form class="cbp-form" data-fm="1" style="grid-template-columns:repeat(2,minmax(0,1fr))">' + campos.map((c: ICampo) => {
    const v = c.valor === undefined || c.valor === null ? '' : String(c.valor);
    const lab = esc(c.label) + (c.ayuda ? ' <span style="font-weight:400;color:#5c5c5c">' + esc(c.ayuda) + '</span>' : '');
    const cls = 'cbp-field' + (c.full || c.tipo === 'area' ? ' full' : '');
    const ph = c.ph ? ' placeholder="' + esc(c.ph) + '"' : '';
    const rq = c.req ? ' required' : '';
    if (c.tipo === 'area') { return '<label class="' + cls + '">' + lab + '<textarea name="' + c.n + '" rows="4"' + ph + rq + '>' + esc(v) + '</textarea></label>'; }
    if (c.tipo === 'select') { return '<label class="' + cls + '">' + lab + '<select name="' + c.n + '">' + (c.opciones || []).map((o: string) => '<option' + (o === v ? ' selected' : '') + '>' + esc(o) + '</option>').join('') + '</select></label>'; }
    if (c.tipo === 'check') { return '<label class="cbp-check' + (c.full ? ' full' : '') + '" style="align-self:end;min-height:44px"><input type="checkbox" name="' + c.n + '"' + (c.valor ? ' checked' : '') + '>' + lab + '</label>'; }
    const t = c.tipo === 'datetime' ? 'datetime-local' : (c.tipo || 'text');
    return '<label class="' + cls + '">' + lab + '<input type="' + t + '" name="' + c.n + '" value="' + esc(v) + '"' + ph + rq + (c.tipo === 'number' ? ' step="any"' : '') + '></label>';
  }).join('') + (extra || '') +
    '<div class="cbp-btns full"><button type="button" class="cbp-btn" data-c="1">Cancelar</button><button type="submit" class="cbp-btn pri">' + esc(boton) + '</button></div></form>';
  const body = abrirModal(titulo, html, ancho || 620);
  (body.querySelector('[data-c]') as HTMLElement).addEventListener('click', cerrarModal);
  const f = body.querySelector('[data-fm]') as HTMLFormElement;
  f.addEventListener('submit', (ev: Event) => {
    ev.preventDefault();
    const v: { [k: string]: any } = {}; // eslint-disable-line @typescript-eslint/no-explicit-any
    campos.forEach((c: ICampo) => {
      const el = f.querySelector('[name="' + c.n + '"]') as HTMLInputElement;
      if (c.tipo === 'check') { v[c.n] = el.checked; } else if (c.tipo === 'number') { v[c.n] = el.value === '' ? null : +el.value; } else { v[c.n] = el.value.trim(); }
    });
    const b = f.querySelector('button[type="submit"]') as HTMLButtonElement; b.disabled = true; b.textContent = 'Guardando…';
    guardar(v).then(() => cerrarModal()).catch((x: Error) => { b.disabled = false; b.textContent = boton; aviso('Error: ' + x.message, true); });
  });
  return body;
}
/** Fecha de un input date → valor ISO para SharePoint (o null) */
export function isoDia(s: string): string | null { return s ? s + 'T12:00:00Z' : null; }
export function isoDT(s: string): string | null { return s ? new Date(s).toISOString() : null; }
