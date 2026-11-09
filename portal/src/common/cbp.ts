import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/* ---------- Tipos ---------- */
export interface IEvento {
  id: number;
  titulo: string;
  fecha: Date;
  equipo: string;
  categoria: string;
  tipo: string;
  rival: string;
  lugar: string;
  notas: string;
  local: boolean;
  pf?: number;
  pa?: number;
}
export interface IAviso { id: number; titulo: string; detalle: string; limite?: Date; importante: boolean; }
export interface IEquipo { id: number; nombre: string; rama: string; categoria: string; horario: string; pista: string; competicion: string; entrenadores: string; }

/* ---------- Utilidades ---------- */
export function esc(s: string): string {
  return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
export const MESES: string[] = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'];
export const MESES_LARGO: string[] = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
export const DIAS: string[] = ['DOM', 'LUN', 'MAR', 'MIÉ', 'JUE', 'VIE', 'SÁB'];
export const DIAS_LARGO: string[] = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
export function hora(d: Date): string {
  const h = d.getHours(); const m = d.getMinutes();
  return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
}
export function mismoDia(a: Date, b: Date): boolean {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}
export function claveDia(d: Date): string { return d.getFullYear() + '-' + d.getMonth() + '-' + d.getDate(); }
export function categoriaDe(equipo: string): string {
  const e = equipo || '';
  if (e.indexOf('Tecnificación') === 0) { return 'Escuela'; }
  if (e.indexOf('Equipo de') === 0) { return 'Madres y padres'; }
  if (e === 'Todo el club') { return 'Club'; }
  return e.split(' ')[0];
}
export const CATEGORIAS: string[] = ['Sénior', 'Junior', 'Cadete', 'Infantil', 'Alevín', 'Benjamín', 'Escuela', 'Madres y padres'];

/* ---------- Datos ---------- */
function getJson(client: SPHttpClient, url: string): Promise<any> { // eslint-disable-line @typescript-eslint/no-explicit-any
  return client.get(url, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } })
    .then((r: SPHttpClientResponse) => { if (!r.ok) { throw new Error('HTTP ' + r.status); } return r.json(); });
}

export function cargarEventos(client: SPHttpClient, webUrl: string): Promise<IEvento[]> {
  const url = webUrl + "/_api/web/GetList('" + encodeURIComponent(listPath(webUrl, 'Lists/Calendario')) + "')/items?$select=Id,Title,Fecha,Equipo,TipoEvento,Rival,Lugar,Notas&$orderby=Fecha asc&$top=2000";
  return getJson(client, url).then((j: any) => (j.value || []).map((i: any): IEvento => { // eslint-disable-line @typescript-eslint/no-explicit-any
    const notas: string = i.Notas || '';
    const m = notas.match(/Resultado: (\d+)-(\d+)/);
    const titulo: string = i.Title || '';
    return {
      id: i.Id, titulo, fecha: new Date(i.Fecha), equipo: i.Equipo || '', categoria: categoriaDe(i.Equipo || ''),
      tipo: i.TipoEvento || '', rival: i.Rival || '', lugar: i.Lugar || '', notas,
      local: titulo.indexOf('(local)') > -1 || notas.indexOf('· Local') > -1,
      pf: m ? parseInt(m[1], 10) : undefined, pa: m ? parseInt(m[2], 10) : undefined
    };
  }));
}

export function cargarAvisos(client: SPHttpClient, webUrl: string): Promise<IAviso[]> {
  const url = webUrl + "/_api/web/GetList('" + encodeURIComponent(listPath(webUrl, 'Lists/Avisos')) + "')/items?$select=Id,Title,Detalle,FechaLimite,Prioridad&$top=50";
  return getJson(client, url).then((j: any) => (j.value || []).map((i: any): IAviso => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
    id: i.Id, titulo: i.Title || '', detalle: i.Detalle || '', limite: i.FechaLimite ? new Date(i.FechaLimite) : undefined,
    importante: i.Prioridad === 'Importante'
  })).sort((a: IAviso, b: IAviso) => {
    if (a.importante !== b.importante) { return a.importante ? -1 : 1; }
    const ta = a.limite ? a.limite.getTime() : 9e15; const tb = b.limite ? b.limite.getTime() : 9e15;
    return ta - tb;
  }));
}

export function cargarEquipos(client: SPHttpClient, webUrl: string): Promise<IEquipo[]> {
  const url = webUrl + "/_api/web/GetList('" + encodeURIComponent(listPath(webUrl, 'Lists/Equipos')) + "')/items?$select=Id,Title,Rama,CategoriaEq,Horario,Pista,Competicion,Entrenadores&$orderby=Id asc&$top=200";
  return getJson(client, url).then((j: any) => (j.value || []).map((i: any): IEquipo => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
    id: i.Id, nombre: i.Title || '', rama: i.Rama || '', categoria: i.CategoriaEq || '', horario: i.Horario || '',
    pista: i.Pista || '', competicion: i.Competicion || '', entrenadores: i.Entrenadores || ''
  })));
}

function listPath(webUrl: string, rel: string): string {
  const a = document.createElement('a'); a.href = webUrl;
  let p = a.pathname; if (p.charAt(0) !== '/') { p = '/' + p; }
  return p.replace(/\/$/, '') + '/' + rel;
}

/* ---------- Iconos (trazo) ---------- */
export const ICON: { [k: string]: string } = {
  calendario: '<path d="M3 5h18v16H3z"/><path d="M3 10h18M8 3v4M16 3v4"/>',
  equipos: '<circle cx="8" cy="9" r="3"/><circle cx="16" cy="9" r="3"/><path d="M2 20c1-3.5 3.5-5 6-5s5 1.5 6 5M12 20c1-3.5 3.5-5 6-5s3.5.8 4.5 2.5"/>',
  temporada: '<path d="M4 4h16v4a8 8 0 0 1-16 0z"/><path d="M12 16v4M8 20h8"/>',
  materiales: '<path d="M4 4h6v16H4zM10 4h4v16h-4z"/><path d="M15 5l4-1 3 15-4 1z"/>',
  avisos: '<path d="M3 11v2l13 5V6z"/><path d="M16 9a3 3 0 0 1 0 6M6 13l1 6h3l-1-5"/>',
  herramientas: '<path d="M14 6l4-4 4 4-4 4zM3 21l9-9M12 12l3 3"/><path d="M3 3l5 5"/>',
  junta: '<rect x="4" y="10" width="16" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/>',
  pin: '<path d="M12 21s-7-6.5-7-12a7 7 0 0 1 14 0c0 5.5-7 12-7 12z"/><circle cx="12" cy="9" r="2.5"/>',
  reloj: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  flechaIzq: '<path d="M15 5l-7 7 7 7"/>',
  flechaDer: '<path d="M9 5l7 7-7 7"/>'
};
export function icon(name: string, size?: number): string {
  const s = size || 22;
  return '<svg class="cbp-ico" width="' + s + '" height="' + s + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + (ICON[name] || '') + '</svg>';
}

/* ---------- Estilos de marca ---------- */
let estilosCargados = false;
export function cargarEstilos(): void {
  if (estilosCargados || document.getElementById('cbp-estilos')) { estilosCargados = true; return; }
  estilosCargados = true;
  const l1 = document.createElement('link'); l1.rel = 'preconnect'; l1.href = 'https://fonts.gstatic.com'; l1.crossOrigin = 'anonymous';
  const l2 = document.createElement('link'); l2.rel = 'stylesheet';
  l2.href = 'https://fonts.googleapis.com/css2?family=Anton&family=Inter:wght@400;500;600;700&display=swap';
  document.head.appendChild(l1); document.head.appendChild(l2);
  const st = document.createElement('style'); st.id = 'cbp-estilos'; st.textContent = CSS;
  document.head.appendChild(st);
}

const CSS = `
.cbp{--negro:#090909;--ambar:#FF9F00;--ambar-osc:#8A5600;--gris:#5c5c5c;--linea:#E6E3DC;--fondo:#F6F5F1;--blanco:#fff;
 font-family:'Inter',system-ui,-apple-system,'Segoe UI',sans-serif;color:var(--negro);font-size:15px;line-height:1.45;-webkit-font-smoothing:antialiased}
.cbp *{box-sizing:border-box}
.cbp a{color:inherit}
.cbp .cbp-ico{flex:none;display:block}
.cbp-titular{font-family:'Anton','Impact','Arial Narrow',sans-serif;font-weight:400;text-transform:uppercase;letter-spacing:.01em;line-height:.98;margin:0}
.cbp-over{font-size:12px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--ambar-osc)}
.cbp-dark .cbp-over{color:var(--ambar)}
.cbp-h2{font-family:'Anton','Impact',sans-serif;font-weight:400;text-transform:uppercase;font-size:30px;line-height:1;margin:0;letter-spacing:.01em}
.cbp-sec{display:flex;flex-direction:column;gap:16px}
.cbp-sec-head{display:flex;align-items:flex-end;justify-content:space-between;gap:12px;flex-wrap:wrap}
.cbp-link{font-size:14px;font-weight:600;text-decoration:none;border-bottom:2px solid var(--ambar);padding-bottom:1px}
.cbp-link:hover{color:var(--ambar-osc)}
.cbp-card{background:var(--blanco);border:1px solid var(--linea);border-radius:16px}
.cbp-chips{display:flex;flex-wrap:wrap;gap:6px}
.cbp-chip{font:inherit;font-size:13px;font-weight:600;min-height:38px;padding:0 14px;border-radius:999px;border:1px solid #D6D2C8;background:var(--blanco);color:var(--negro);cursor:pointer}
.cbp-chip:hover{border-color:var(--negro)}
.cbp-chip[aria-pressed="true"]{background:var(--negro);border-color:var(--negro);color:var(--ambar)}
.cbp-select{font:inherit;font-size:14px;min-height:40px;padding:0 12px;border-radius:10px;border:1px solid #D6D2C8;background:var(--blanco);color:var(--negro)}
.cbp-vacio{padding:28px;text-align:center;color:var(--gris);border:1px dashed #CFCABF;border-radius:14px;background:var(--blanco)}
.cbp-cargando{padding:40px;text-align:center;color:var(--gris)}
/* Hero */
.cbp-hero{background:var(--negro);color:#fff;position:relative;overflow:hidden}
.cbp-hero-in{max-width:1240px;margin:0 auto;padding:48px 32px 40px;display:flex;gap:40px;align-items:center;flex-wrap:wrap}
.cbp-hero-txt{flex:999 1 520px;min-width:0;display:flex;flex-direction:column;gap:18px}
.cbp-hero h1{font-size:clamp(44px,6.2vw,84px);color:#fff}
.cbp-hero h1 em{font-style:normal;color:var(--ambar)}
.cbp-hero p{margin:0;color:#CFCFCF;font-size:17px;max-width:560px}
.cbp-hero-logo{flex:1 1 220px;display:flex;justify-content:center}
.cbp-hero-logo img{width:100%;max-width:250px;height:auto}
.cbp-stats{display:flex;gap:12px;flex-wrap:wrap;margin-top:6px}
.cbp-stat{border:1px solid #2b2b2b;border-radius:12px;padding:10px 16px;min-width:120px}
.cbp-stat b{display:block;font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:30px;line-height:1;color:var(--ambar)}
.cbp-stat span{font-size:12px;color:#BDBDBD;letter-spacing:.04em}
.cbp-fases{display:flex;gap:6px;margin-top:8px}
.cbp-fase{display:flex;flex-direction:column;gap:6px;min-width:0}
.cbp-fase i{display:block;height:8px;border-radius:4px;background:#333}
.cbp-fase.on i{background:var(--ambar)}
.cbp-fase span{font-size:12px;color:#A9A9A9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cbp-fase.on span{color:#fff;font-weight:700}
/* Wrap */
.cbp-wrap{max-width:1240px;margin:0 auto;padding:36px 32px 48px;display:flex;flex-direction:column;gap:40px}
/* Accesos */
.cbp-accesos{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:12px}
.cbp-acceso{display:flex;flex-direction:column;gap:14px;justify-content:space-between;min-height:112px;padding:18px;border-radius:16px;background:var(--blanco);border:1px solid var(--linea);text-decoration:none;transition:transform .15s,box-shadow .15s,border-color .15s}
.cbp-acceso:hover{transform:translateY(-2px);border-color:var(--negro);box-shadow:0 8px 20px rgba(0,0,0,.08)}
.cbp-acceso .cbp-acceso-ico{width:42px;height:42px;border-radius:12px;background:var(--negro);color:var(--ambar);display:flex;align-items:center;justify-content:center}
.cbp-acceso b{font-size:15px;font-weight:700}
.cbp-acceso small{display:block;font-size:12px;color:var(--gris);font-weight:500;margin-top:2px}
/* Grid 2 col */
.cbp-dos{display:flex;flex-wrap:wrap;gap:28px;align-items:flex-start}
.cbp-dos > .cbp-main{flex:999 1 560px;min-width:0}
.cbp-dos > .cbp-side{flex:1 1 320px;min-width:0}
/* Partido */
.cbp-lista{display:flex;flex-direction:column}
.cbp-partido{display:flex;align-items:center;gap:14px;padding:14px 18px;border-top:1px solid #EFECE6;text-decoration:none}
.cbp-partido:first-child{border-top:0}
.cbp-partido:hover{background:#FBFAF7}
.cbp-fecha{flex:none;width:60px;height:64px;border-radius:12px;background:var(--negro);color:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px}
.cbp-fecha small{font-size:10px;font-weight:700;letter-spacing:.1em;color:#BDBDBD}
.cbp-fecha b{font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:26px;line-height:1;color:var(--ambar)}
.cbp-fecha span{font-size:10px;font-weight:700;letter-spacing:.1em}
.cbp-pinfo{flex:1;min-width:0;display:flex;flex-direction:column;gap:2px}
.cbp-pinfo .cbp-eq{font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--ambar-osc)}
.cbp-pinfo .cbp-riv{font-size:16px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cbp-pinfo .cbp-meta{display:flex;gap:12px;flex-wrap:wrap;font-size:13px;color:var(--gris)}
.cbp-pinfo .cbp-meta span{display:inline-flex;align-items:center;gap:4px;min-width:0}
.cbp-tag{flex:none;font-size:11px;font-weight:800;letter-spacing:.06em;padding:5px 10px;border-radius:999px;border:1.5px solid var(--ambar);color:var(--negro);background:#fff;white-space:nowrap}
.cbp-tag.casa{background:var(--ambar)}
.cbp-tag.torneo{background:var(--negro);border-color:var(--negro);color:var(--ambar)}
.cbp-tag.otro{border-color:#D6D2C8;background:#F3F1EC}
.cbp-res{flex:none;display:flex;align-items:center;gap:8px}
.cbp-marcador{font-family:'Anton','Impact',sans-serif;font-size:24px;line-height:1;letter-spacing:.02em}
.cbp-wl{width:28px;height:28px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:13px}
.cbp-wl.v{background:var(--ambar);color:var(--negro)}
.cbp-wl.d{background:#EDEBE6;color:#555}
/* Avisos */
.cbp-aviso{padding:16px 18px;border-top:1px solid #EFECE6;display:flex;flex-direction:column;gap:6px}
.cbp-aviso:first-child{border-top:0}
.cbp-aviso h3{margin:0;font-size:15px;font-weight:700;line-height:1.3}
.cbp-aviso p{margin:0;font-size:13px;color:var(--gris);display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.cbp-aviso-top{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.cbp-badge{font-size:10px;font-weight:800;letter-spacing:.08em;padding:3px 8px;border-radius:6px;background:var(--ambar);color:var(--negro)}
.cbp-limite{font-size:12px;font-weight:600;color:var(--negro)}
/* Equipos */
.cbp-equipos{display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:16px}
.cbp-equipo{display:flex;flex-direction:column;overflow:hidden}
.cbp-equipo-head{background:var(--negro);color:#fff;padding:18px 20px 16px;display:flex;flex-direction:column;gap:6px}
.cbp-equipo-head .cbp-over{color:var(--ambar)}
.cbp-equipo-head h3{font-family:'Anton','Impact',sans-serif;font-weight:400;text-transform:uppercase;font-size:26px;line-height:1;margin:0;color:#fff}
.cbp-equipo-body{padding:16px 20px 18px;display:flex;flex-direction:column;gap:12px;flex:1}
.cbp-sesiones{display:flex;flex-direction:column;gap:6px;margin:0;padding:0;list-style:none}
.cbp-sesiones li{display:flex;gap:10px;font-size:13px}
.cbp-sesiones li b{flex:none;width:34px;font-size:11px;letter-spacing:.06em;font-weight:800;color:var(--ambar-osc);padding-top:1px}
.cbp-prox{margin-top:auto;padding-top:12px;border-top:1px solid #EFECE6;font-size:13px;display:flex;flex-direction:column;gap:2px}
.cbp-prox small{font-size:11px;font-weight:700;letter-spacing:.08em;color:var(--gris);text-transform:uppercase}
.cbp-comp{font-size:12px;color:var(--gris)}
/* Calendario mes */
.cbp-mes-head{display:flex;align-items:center;gap:10px}
.cbp-mes-head h3{margin:0;min-width:200px;text-align:center;font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:24px;text-transform:uppercase}
.cbp-nav{width:40px;height:40px;border-radius:10px;border:1px solid #D6D2C8;background:#fff;display:flex;align-items:center;justify-content:center;cursor:pointer;color:var(--negro)}
.cbp-nav:hover{border-color:var(--negro)}
.cbp-mes{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));border:1px solid var(--linea);border-radius:14px;overflow:hidden;background:var(--linea);gap:1px}
.cbp-mes .cbp-dsem{background:var(--negro);color:#fff;font-size:11px;font-weight:700;letter-spacing:.1em;text-align:center;padding:8px 0}
.cbp-dia{background:#fff;min-height:108px;padding:6px;display:flex;flex-direction:column;gap:3px;min-width:0}
.cbp-dia.fuera{background:#FAF9F6}
.cbp-dia.fuera .cbp-num{color:#B5B0A6}
.cbp-dia.hoy .cbp-num{background:var(--ambar);color:var(--negro);border-radius:6px}
.cbp-num{align-self:flex-start;font-size:12px;font-weight:700;padding:1px 6px}
.cbp-ev{font-size:11px;line-height:1.25;padding:3px 6px;border-radius:6px;background:#F3F1EC;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;border-left:3px solid #CFCABF}
.cbp-ev.casa{border-left-color:var(--ambar);background:#FFF4E0}
.cbp-ev.torneo{background:var(--negro);color:var(--ambar);border-left-color:var(--ambar)}
.cbp-mas{font-size:11px;font-weight:700;color:var(--gris);padding:0 6px}
.cbp-diahead{font-size:13px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;padding:16px 18px 6px;color:var(--negro);background:var(--fondo);border-top:1px solid #EFECE6}
.cbp-toolbar{display:flex;flex-wrap:wrap;gap:12px;align-items:center;justify-content:space-between}
.cbp-tabs{display:inline-flex;background:#ECEAE4;border-radius:12px;padding:4px;gap:4px}
.cbp-tab{font:inherit;font-size:14px;font-weight:600;border:0;background:transparent;padding:0 16px;min-height:38px;border-radius:9px;cursor:pointer;color:#333}
.cbp-tab[aria-selected="true"]{background:var(--negro);color:var(--ambar)}
.cbp-leyenda{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--gris)}
.cbp-leyenda span{display:inline-flex;gap:6px;align-items:center}
.cbp-leyenda i{width:12px;height:12px;border-radius:3px;display:inline-block}
@media (max-width:720px){
 .cbp-hero-in{padding:32px 18px 28px;gap:24px}
 .cbp-hero-logo{order:-1;justify-content:flex-start}.cbp-hero-logo img{max-width:110px}
 .cbp-wrap{padding:24px 16px 36px;gap:30px}
 .cbp-partido{padding:12px 12px;gap:10px}.cbp-fecha{width:52px;height:58px}
 .cbp-dia{min-height:64px;padding:3px}.cbp-ev{font-size:0;height:6px;padding:0;border-left-width:0;border-radius:3px}
 .cbp-ev.casa{background:var(--ambar)}.cbp-ev{background:#bbb}
 .cbp-mas{font-size:10px;padding:0 2px}
 .cbp-mes-head h3{min-width:0;font-size:20px}
}
`;
