import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IEvento, esc } from './cbp';

/* ---------- Escritura en SharePoint ---------- */
const JH = { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'odata-version': '' };

export function spPost(client: SPHttpClient, url: string, body?: any, merge?: boolean): Promise<any> { // eslint-disable-line @typescript-eslint/no-explicit-any
  const headers: { [k: string]: string } = { ...JH };
  if (merge) { headers['X-HTTP-Method'] = 'MERGE'; headers['IF-MATCH'] = '*'; }
  const opts: ISPHttpClientOptions = { headers, body: body !== undefined ? JSON.stringify(body) : undefined };
  return client.post(url, SPHttpClient.configurations.v1, opts).then((r: SPHttpClientResponse) => {
    if (!r.ok) { return r.text().then((t: string) => { throw new Error('HTTP ' + r.status + ' ' + t.slice(0, 200)); }); }
    return r.status === 204 ? null : r.text().then((t: string) => (t ? JSON.parse(t) : null));
  });
}
export function spGet(client: SPHttpClient, url: string): Promise<any> { // eslint-disable-line @typescript-eslint/no-explicit-any
  return client.get(url, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' } })
    .then((r: SPHttpClientResponse) => { if (!r.ok) { throw new Error('HTTP ' + r.status); } return r.json(); });
}
export function listUrl(webUrl: string, rel: string): string {
  const a = document.createElement('a'); a.href = webUrl;
  let p = a.pathname; if (p.charAt(0) !== '/') { p = '/' + p; }
  return webUrl + "/_api/web/GetList('" + encodeURIComponent(p.replace(/\/$/, '') + '/' + rel) + "')";
}
export function webPath(webUrl: string): string {
  const a = document.createElement('a'); a.href = webUrl;
  let p = a.pathname; if (p.charAt(0) !== '/') { p = '/' + p; }
  return p.replace(/\/$/, '');
}

/* ---------- Ventana modal ---------- */
export function abrirModal(titulo: string, html: string, ancho?: number): HTMLElement {
  cerrarModal();
  const ov = document.createElement('div');
  ov.className = 'cbp cbp-ov'; ov.id = 'cbp-modal';
  ov.innerHTML = '<div class="cbp-modal" role="dialog" aria-modal="true" aria-label="' + esc(titulo) + '" style="max-width:' + (ancho || 640) + 'px">' +
    '<div class="cbp-modal-head"><h2>' + esc(titulo) + '</h2><button type="button" class="cbp-x" aria-label="Cerrar">×</button></div>' +
    '<div class="cbp-modal-body">' + html + '</div></div>';
  document.body.appendChild(ov);
  ov.addEventListener('click', (e: MouseEvent) => { if (e.target === ov) { cerrarModal(); } });
  (ov.querySelector('.cbp-x') as HTMLElement).addEventListener('click', cerrarModal);
  document.addEventListener('keydown', escCerrar);
  const f = ov.querySelector('input,select,textarea,button:not(.cbp-x)') as HTMLElement;
  if (f) { setTimeout(() => f.focus(), 30); }
  return ov.querySelector('.cbp-modal-body') as HTMLElement;
}
function escCerrar(e: KeyboardEvent): void { if (e.key === 'Escape') { cerrarModal(); } }
export function cerrarModal(): void {
  const o = document.getElementById('cbp-modal'); if (o) { o.parentNode!.removeChild(o); }
  document.removeEventListener('keydown', escCerrar);
}

export function aviso(msg: string, error?: boolean): void {
  const t = document.createElement('div');
  t.className = 'cbp cbp-toast' + (error ? ' err' : ''); t.setAttribute('role', 'status'); t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => { if (t.parentNode) { t.parentNode.removeChild(t); } }, error ? 9000 : 3500);
}

export function descargar(nombre: string, contenido: string | Blob, mime: string): void {
  const blob = typeof contenido === 'string' ? new Blob([contenido], { type: mime }) : contenido;
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nombre;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.parentNode!.removeChild(a); }, 500);
}

/* ---------- iCalendar ---------- */
function icsFecha(d: Date): string {
  const p = (n: number): string => (n < 10 ? '0' : '') + n;
  return d.getUTCFullYear() + p(d.getUTCMonth() + 1) + p(d.getUTCDate()) + 'T' + p(d.getUTCHours()) + p(d.getUTCMinutes()) + '00Z';
}
function icsTxt(s: string): string { return (s || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\n/g, '\\n'); }
export function generarIcs(evs: IEvento[], nombre: string): string {
  const ahora = icsFecha(new Date());
  const lineas = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Gijon CBP//Portal//ES', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH', 'X-WR-CALNAME:' + icsTxt(nombre), 'X-WR-TIMEZONE:Europe/Madrid'];
  evs.forEach((e: IEvento) => {
    const fin = new Date(e.fecha.getTime() + (e.tipo === 'Partido' ? 105 : 90) * 60000);
    const resumen = e.tipo === 'Partido' ? e.equipo + (e.local ? ' vs ' : ' en ') + e.rival : e.titulo;
    lineas.push('BEGIN:VEVENT', 'UID:cbp-' + e.id + '@gijoncbp', 'DTSTAMP:' + ahora, 'DTSTART:' + icsFecha(e.fecha), 'DTEND:' + icsFecha(fin),
      'SUMMARY:' + icsTxt(resumen), 'LOCATION:' + icsTxt(e.lugar), 'DESCRIPTION:' + icsTxt(e.notas), 'END:VEVENT');
  });
  lineas.push('END:VCALENDAR');
  return lineas.join('\r\n');
}

export function mapa(lugar: string): string {
  return 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(lugar + ', Asturias');
}

export function fechaInput(d: Date): string {
  const p = (n: number): string => (n < 10 ? '0' : '') + n;
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}

export const EQUIPOS: string[] = ['Sénior Femenino A', 'Sénior Femenino B', 'Sénior Masculino A', 'Junior Femenino A', 'Junior Masculino A', 'Junior Masculino B', 'Cadete Femenino A', 'Cadete Femenino B', 'Cadete Masculino A', 'Cadete Masculino B', 'Cadete Masculino C', 'Infantil Femenino A', 'Infantil Masculino A', 'Infantil Masculino B', 'Infantil Masculino C', 'Alevín Femenino A', 'Alevín Femenino B', 'Alevín Masculino A', 'Alevín Masculino B', 'Alevín Masculino C', 'Benjamín Femenino', 'Benjamín Masculino A', 'Benjamín Masculino B', 'Benjamín Masculino C', 'Tecnificación Minibasket', 'Tecnificación Canasta Grande', 'Equipo de Madres', 'Equipo de Padres'];
export function opcionesEquipo(sel: string, conTodo?: boolean): string {
  return (conTodo ? '<option value="Todo el club"' + (sel === 'Todo el club' ? ' selected' : '') + '>Todo el club</option>' : '') +
    EQUIPOS.map((e: string) => '<option' + (e === sel ? ' selected' : '') + '>' + esc(e) + '</option>').join('');
}

/* ---------- Estilos adicionales ---------- */
let ok2 = false;
export function cargarEstilosUi(): void {
  if (ok2 || document.getElementById('cbp-estilos-ui')) { ok2 = true; return; }
  ok2 = true;
  const st = document.createElement('style'); st.id = 'cbp-estilos-ui'; st.textContent = CSS2; document.head.appendChild(st);
}
const CSS2 = `
.cbp-ov{position:fixed;inset:0;background:rgba(9,9,9,.55);z-index:100000;display:flex;align-items:flex-start;justify-content:center;padding:6vh 16px;overflow:auto}
.cbp-modal{background:#fff;border-radius:18px;width:100%;box-shadow:0 20px 60px rgba(0,0,0,.3);overflow:hidden}
.cbp-modal-head{background:#090909;color:#fff;padding:16px 20px;display:flex;align-items:center;justify-content:space-between;gap:12px}
.cbp-modal-head h2{margin:0;font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:24px;text-transform:uppercase;letter-spacing:.01em;color:#fff}
.cbp-x{width:40px;height:40px;border-radius:10px;border:0;background:#222;color:#FF9F00;font-size:24px;line-height:1;cursor:pointer}
.cbp-modal-body{padding:20px;display:flex;flex-direction:column;gap:14px}
.cbp-form{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px 14px}
.cbp-form .full{grid-column:1/-1}
.cbp-field{display:flex;flex-direction:column;gap:5px;font-size:13px;font-weight:600;min-width:0}
.cbp-field input,.cbp-field select,.cbp-field textarea{font:inherit;font-size:15px;font-weight:400;padding:10px 12px;border:1px solid #D6D2C8;border-radius:10px;background:#fff;color:#090909;min-height:44px;width:100%}
.cbp-field textarea{min-height:90px;resize:vertical}
.cbp-field input:focus,.cbp-field select:focus,.cbp-field textarea:focus{outline:2px solid #FF9F00;outline-offset:1px;border-color:#090909}
.cbp-btns{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end}
.cbp-btn{font:inherit;font-size:14px;font-weight:700;min-height:44px;padding:0 18px;border-radius:12px;border:1.5px solid #090909;background:#fff;color:#090909;cursor:pointer;display:inline-flex;align-items:center;gap:8px;text-decoration:none}
.cbp-btn:hover{background:#F3F1EC}
.cbp-btn.pri{background:#090909;color:#FF9F00}
.cbp-btn.pri:hover{background:#222}
.cbp-btn.amb{background:#FF9F00;border-color:#FF9F00;color:#090909}
.cbp-btn[disabled]{opacity:.5;cursor:default}
.cbp-btn.sm{min-height:36px;padding:0 12px;font-size:13px;border-radius:10px}
.cbp-toast{position:fixed;left:50%;bottom:28px;transform:translateX(-50%);background:#090909;color:#fff;padding:12px 18px;border-radius:12px;z-index:100001;font-size:14px;box-shadow:0 8px 24px rgba(0,0,0,.25);border-left:4px solid #FF9F00}
.cbp-toast.err{border-left-color:#d13438}
.cbp-dl{display:grid;grid-template-columns:120px 1fr;gap:8px 14px;margin:0;font-size:14px}
.cbp-dl dt{font-size:11px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#5c5c5c;padding-top:2px}
.cbp-dl dd{margin:0}
.cbp-bigres{display:flex;align-items:center;gap:14px;background:#090909;color:#fff;border-radius:14px;padding:14px 18px}
.cbp-bigres b{font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:40px;line-height:1;color:#FF9F00}
/* Semana */
.cbp-sem{display:grid;grid-template-columns:52px repeat(7,minmax(0,1fr));border:1px solid #E6E3DC;border-radius:14px;overflow:hidden;background:#fff}
.cbp-sem-h{background:#090909;color:#fff;font-size:11px;font-weight:700;letter-spacing:.06em;text-align:center;padding:8px 2px;text-transform:uppercase}
.cbp-sem-h.hoy{color:#FF9F00}
.cbp-sem-horas{position:relative;border-right:1px solid #E6E3DC}
.cbp-sem-horas span{position:absolute;right:6px;font-size:11px;color:#8a8a8a;transform:translateY(-50%)}
.cbp-sem-col{position:relative;border-right:1px solid #EFECE6;background-image:linear-gradient(#EFECE6 1px,transparent 1px);background-size:100% 44px}
.cbp-sem-col:last-child{border-right:0}
.cbp-blk{position:absolute;border-radius:7px;padding:3px 5px;font-size:11px;line-height:1.2;overflow:hidden;cursor:pointer;border:0;text-align:left;font-family:inherit;color:#090909}
.cbp-blk b{display:block;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cbp-blk span{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;opacity:.85}
.cbp-blk.ent{background:#ECEAE4;border-left:3px solid #9b9b9b}
.cbp-blk.casa{background:#FF9F00}
.cbp-blk.fuera{background:#fff;border:1.5px solid #FF9F00}
.cbp-blk.tor{background:#090909;color:#FF9F00}
.cbp-blk.otro{background:#DDE6F5;border-left:3px solid #4a6fa5}
.cbp-dia{cursor:pointer}
.cbp-dia:hover{background:#FFFBF2}
.cbp-scroll{overflow-x:auto}
/* Biblioteca */
.cbp-mats{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:16px}
.cbp-mat{display:flex;flex-direction:column;overflow:hidden;text-decoration:none;transition:transform .15s,box-shadow .15s}
.cbp-mat:hover{transform:translateY(-2px);box-shadow:0 10px 24px rgba(0,0,0,.08)}
.cbp-mat-prev{height:140px;background:#F3F1EC;display:flex;align-items:center;justify-content:center;overflow:hidden;position:relative}
.cbp-mat-prev img{width:100%;height:100%;object-fit:cover}
.cbp-mat-ext{font-family:'Anton','Impact',sans-serif;font-size:30px;color:#090909;letter-spacing:.04em}
.cbp-mat-tipo{position:absolute;left:10px;top:10px;background:#090909;color:#FF9F00;font-size:10px;font-weight:800;letter-spacing:.08em;padding:4px 8px;border-radius:6px;text-transform:uppercase}
.cbp-mat-body{padding:12px 14px 14px;display:flex;flex-direction:column;gap:4px;flex:1}
.cbp-mat-body b{font-size:14px;line-height:1.3}
.cbp-mat-body p{margin:0;font-size:12px;color:#5c5c5c;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.cbp-mat-meta{margin-top:auto;padding-top:6px;font-size:11px;color:#777;display:flex;justify-content:space-between;gap:8px}
.cbp-search{font:inherit;font-size:15px;min-height:44px;padding:0 14px;border-radius:12px;border:1px solid #D6D2C8;width:100%;max-width:360px;background:#fff}
.cbp-drop{border:2px dashed #CFCABF;border-radius:14px;padding:22px;text-align:center;color:#5c5c5c;background:#FAF9F6;cursor:pointer}
.cbp-drop.on{border-color:#FF9F00;background:#FFF6E6}
/* Herramientas */
.cbp-tooltabs{display:flex;gap:8px;flex-wrap:wrap}
.cbp-tooltab{font:inherit;display:inline-flex;align-items:center;gap:8px;min-height:46px;padding:0 16px;border-radius:12px;border:1px solid #E6E3DC;background:#fff;font-weight:700;font-size:14px;cursor:pointer;color:#090909}
.cbp-tooltab[aria-selected="true"]{background:#090909;color:#FF9F00;border-color:#090909}
.cbp-panel{padding:20px}
.cbp-players{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:8px}
.cbp-player{display:flex;align-items:center;gap:10px;padding:8px 10px;border:1px solid #E6E3DC;border-radius:12px;background:#fff;min-height:52px}
.cbp-player .dor{width:34px;height:34px;flex:none;border-radius:9px;background:#090909;color:#FF9F00;font-family:'Anton','Impact',sans-serif;font-size:18px;display:flex;align-items:center;justify-content:center}
.cbp-player .nom{flex:1;min-width:0;font-weight:600;font-size:14px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cbp-player input[type=checkbox]{width:22px;height:22px;accent-color:#090909}
.cbp-player.sel{border-color:#090909;background:#FFF6E6}
.cbp-tri{display:inline-flex;border:1px solid #D6D2C8;border-radius:10px;overflow:hidden;flex:none}
.cbp-tri button{font:inherit;border:0;background:#fff;min-width:38px;min-height:36px;font-weight:800;font-size:12px;cursor:pointer;color:#555}
.cbp-tri button[aria-pressed="true"].p{background:#1f7a4d;color:#fff}
.cbp-tri button[aria-pressed="true"].a{background:#c0392b;color:#fff}
.cbp-tri button[aria-pressed="true"].j{background:#FF9F00;color:#090909}
.cbp-tabla{width:100%;border-collapse:collapse;font-size:14px}
.cbp-tabla th{text-align:left;font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:#5c5c5c;padding:8px;border-bottom:2px solid #090909}
.cbp-tabla td{padding:8px;border-bottom:1px solid #EFECE6}
.cbp-barra{height:8px;border-radius:4px;background:#EFECE6;overflow:hidden;min-width:80px}
.cbp-barra i{display:block;height:100%;background:#FF9F00}
.cbp-pz-wrap{display:flex;flex-direction:column;gap:10px}
.cbp-pz-tools{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.cbp-pz-tool{font:inherit;min-height:42px;min-width:42px;padding:0 10px;border-radius:10px;border:1px solid #D6D2C8;background:#fff;cursor:pointer;font-size:13px;font-weight:700;display:inline-flex;align-items:center;gap:6px;color:#090909}
.cbp-pz-tool[aria-pressed="true"]{background:#090909;color:#FF9F00;border-color:#090909}
.cbp-pz-sep{width:1px;height:28px;background:#D6D2C8;margin:0 4px}
.cbp-pz-canvas{width:100%;height:auto;border-radius:14px;touch-action:none;display:block;background:#E9CFA0;box-shadow:inset 0 0 0 1px rgba(0,0,0,.08)}
.cbp-bloque{display:grid;grid-template-columns:150px minmax(0,1fr) 90px 44px;gap:8px;align-items:end;padding:12px;border:1px solid #E6E3DC;border-radius:12px;background:#fff}
.cbp-bloque .cbp-field{font-size:12px}
.cbp-total{font-family:'Anton','Impact',sans-serif;font-size:26px}
@media (max-width:720px){
 .cbp-form{grid-template-columns:1fr}
 .cbp-dl{grid-template-columns:1fr}
 .cbp-bloque{grid-template-columns:1fr 1fr}
 .cbp-sem{min-width:720px}
}
`;
