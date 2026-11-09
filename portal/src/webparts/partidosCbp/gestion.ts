import { SPHttpClient } from '@microsoft/sp-http';
import { esc, MESES } from '../../common/cbp';
import { spPost, abrirModal, cerrarModal, aviso, descargar, fechaInput } from '../../common/ui';
import { porTitulo, getJ, formModal } from '../../common/sp2';
import { libroXlsx } from '../../common/xlsx';
import { IPartido, IRegla, IPista, IAlerta, IJugV, IHueco, sugerir, asignarPistas, analizar, norm, minDia, hm, DIA_TXT } from './motor';

export const ESTADOS = ['Oficial', 'Propuesto', 'Validado', 'Pedido al rival', 'Aceptado', 'Contrapropuesta', 'Rechazado', 'Cargado en FBPA'];
const L_GES = 'Partidos · Gestión'; const L_CLU = 'Partidos · Clubes';

export interface IPartidoG extends IPartido { oficial: Date; comp: string; jornada: string; catPDM: string; }
export interface IGest { id: number; pid: number; fechaProp?: Date; pista: string; estado: string; motivo: string; respuesta: string; ok: string; historial: string; }
export interface IClub { id: number; nombre: string; corto: string; email: string; tel: string; contacto: string; url: string; direccion: string; equipos: string[]; manual: string[]; manualTxt: string; }
export interface IEqInfo { pref: string[]; alt: string[]; }

export interface IGCtx {
  c: SPHttpClient; w: string; usuario: string; esCoord: boolean; margen: number;
  semana: IPartidoG[]; gest: { [pid: number]: IGest }; clubes: IClub[]; pistas: IPista[]; reglas: IRegla[]; jugs: IJugV[];
  eqInfo: (p: IPartido) => IEqInfo; recargar: () => void;
}

/* ---------- Carga ---------- */
export function cargarGestion(c: SPHttpClient, w: string): Promise<{ [pid: number]: IGest }> {
  return getJ(c, porTitulo(w, L_GES) + '/items?$select=Id,PartidoId,FechaPropuesta,Pista,EstadoG,Motivo,Respuesta,OKEntrenador,Historial&$top=3000').then((j: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    const m: { [pid: number]: IGest } = {};
    (j.value || []).forEach((i: any) => { m[i.PartidoId] = { id: i.Id, pid: i.PartidoId, fechaProp: i.FechaPropuesta ? new Date(i.FechaPropuesta) : undefined, pista: i.Pista || '', estado: i.EstadoG || 'Oficial', motivo: i.Motivo || '', respuesta: i.Respuesta || '', ok: i.OKEntrenador || '', historial: i.Historial || '' }; }); // eslint-disable-line @typescript-eslint/no-explicit-any
    return m;
  }).catch(() => ({}));
}
export function cargarClubes(c: SPHttpClient, w: string): Promise<IClub[]> {
  return getJ(c, porTitulo(w, L_CLU) + '/items?$select=Id,Title,NombreCorto,Email,Telefono,Contacto,UrlFBPA,Direccion,EquiposFBPA,EquiposManual&$top=500').then((j: any) => (j.value || []).map((i: any): IClub => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
    id: i.Id, nombre: i.Title || '', corto: i.NombreCorto || '', email: i.Email || '', tel: i.Telefono || '', contacto: i.Contacto || '', url: i.UrlFBPA || '', direccion: i.Direccion || '',
    equipos: (i.EquiposFBPA || '').split('\n').map((x: string) => norm(x)).filter((x: string) => x),
    manual: (i.EquiposManual || '').split('\n').map((x: string) => norm(x)).filter((x: string) => x), manualTxt: i.EquiposManual || '' }))).catch(() => []);
}
export function esCoordinador(c: SPHttpClient, w: string): Promise<boolean> {
  return getJ(c, w + '/_api/web/currentuser?$select=IsSiteAdmin&$expand=Groups').then((u: any) => // eslint-disable-line @typescript-eslint/no-explicit-any
    !!u.IsSiteAdmin || (u.Groups || []).some((g: any) => g.Title === 'Coordinación deportiva')).catch(() => false); // eslint-disable-line @typescript-eslint/no-explicit-any
}
export function clubDe(clubes: IClub[], rival: string): IClub | undefined { const r = norm(rival); return clubes.filter((k: IClub) => k.manual.indexOf(r) >= 0)[0] || clubes.filter((k: IClub) => k.equipos.indexOf(r) >= 0 && !clubes.some((o: IClub) => o.manual.indexOf(r) >= 0))[0]; }

/* ---------- Formato ---------- */
const DIAS_C = ['Dom.', 'Lun.', 'Mar.', 'Mié.', 'Jue.', 'Vie.', 'Sáb.'];
export function fTxt(d: Date): string { return DIA_TXT[d.getDay()] + ' ' + d.getDate() + ' ' + MESES[d.getMonth()].toLowerCase() + ' · ' + hm(minDia(d)); }
function ahoraTxt(): string { const d = new Date(); return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + ' ' + hm(minDia(d)); }
function titulo(s: string): string {
  const min = ['de', 'del', 'la', 'las', 'el', 'los', 'y'];
  return s.toLowerCase().replace(/(^|[\s\-./(])([a-zà-ÿñ])/g, (m: string, a: string, b: string) => a + b.toUpperCase())
    .split(' ').map((w: string, i: number) => (i > 0 && min.indexOf(w.toLowerCase()) >= 0 ? w.toLowerCase() : /^(Cb|Cd|Ad|Acd|Cdb|Ies|Bvm\d*|Rgc|Adba|Ii|Iii)$/i.test(w.replace(/[.]/g, '')) ? w.toUpperCase() : w)).join(' ');
}
function nombreCorto(n: string): string { return n.replace(/\s*-\s*Gij[oó]n CBP.*$/i, ''); }
export function pillEstado(e: string): string {
  const cls = e === 'Oficial' ? '' : e === 'Rechazado' ? ' mal' : e === 'Aceptado' || e === 'Cargado en FBPA' ? ' ok' : ' amb';
  return '<span class="cbp-fpill' + cls + '">' + esc(e) + '</span>';
}

/* ---------- Guardar ---------- */
export function guardarGest(ctx: IGCtx, p: IPartidoG, cambios: { [k: string]: any }, accion: string): Promise<void> { // eslint-disable-line @typescript-eslint/no-explicit-any
  const g = ctx.gest[p.id];
  const hist = (g ? g.historial + '\n' : '') + ahoraTxt() + ' · ' + nombreCorto(ctx.usuario) + ' · ' + accion;
  const body: { [k: string]: any } = { ...cambios, Historial: hist }; // eslint-disable-line @typescript-eslint/no-explicit-any
  if (g) { return spPost(ctx.c, porTitulo(ctx.w, L_GES) + '/items(' + g.id + ')', body, true).then(() => undefined); }
  return spPost(ctx.c, porTitulo(ctx.w, L_GES) + '/items', { Title: p.equipo + ' vs ' + p.rival, PartidoId: p.id, Equipo: p.equipo, Rival: p.rival, Local: p.local, FechaOficial: p.oficial.toISOString(), EstadoG: 'Oficial', ...body }).then(() => undefined);
}

/* ---------- Tabla de partidos de la semana ---------- */
export function vPartidos(ctx: IGCtx, al: IAlerta[]): string {
  const ps = ctx.semana.slice().sort((a: IPartidoG, b: IPartidoG) => a.fecha.getTime() - b.fecha.getTime());
  const grav: { [id: number]: string } = {}; al.forEach((a: IAlerta) => a.partidos.forEach((id: number) => { if (!grav[id] || a.grav === 'critica') { grav[id] = a.grav; } }));
  const fila = (p: IPartidoG): string => {
    const g = ctx.gest[p.id]; const est = g ? g.estado : 'Oficial';
    const cambiado = p.fecha.getTime() !== p.oficial.getTime();
    const enFbpa = g && g.fechaProp && Math.abs(g.fechaProp.getTime() - p.oficial.getTime()) < 60000;
    const mods = ctx.pistas.filter((k: IPista) => norm(k.inst) === norm(p.lugar) || norm(k.alias) === norm(p.lugar));
    const pista = p.local ? (ctx.esCoord && mods.length > 1 ? '<select class="cbp-select" data-pista="' + p.id + '" style="min-height:34px;font-size:13px"><option value="">Sin asignar</option>' + mods.map((k: IPista) => '<option' + (g && g.pista === k.clave ? ' selected' : '') + ' value="' + esc(k.clave) + '">' + esc(k.modulo || k.inst) + '</option>').join('') + '</select>'
      : esc(g && g.pista ? (ctx.pistas.filter((k: IPista) => k.clave === g.pista)[0] || { modulo: g.pista }).modulo : (mods[0] ? mods[0].modulo : ''))) : '';
    const acc: string[] = [];
    const b = (a: string, t: string, pri?: boolean): string => '<button type="button" class="cbp-btn sm' + (pri ? ' pri' : '') + '" data-ac="' + a + '|' + p.id + '"' + (pri ? ' style="color:#FF9F00"' : '') + '>' + t + '</button>';
    if (est === 'Oficial' || est === 'Rechazado') { acc.push(b('proponer', 'Proponer cambio', grav[p.id] === 'critica')); }
    if (est === 'Propuesto') { if (ctx.esCoord) { acc.push(b('validar', 'Validar', true)); } if (!g || !g.ok) { acc.push(b('ok', 'OK entrenador')); } acc.push(b('proponer', 'Editar')); if (ctx.esCoord) { acc.push(b('descartar', 'Descartar')); } }
    if (est === 'Validado' && ctx.esCoord) { acc.push(b('mensaje', p.local ? 'Comunicar al rival' : 'Pedir al rival', true)); acc.push(b('cargado', 'Cargado en FBPA')); }
    if (est === 'Pedido al rival' && ctx.esCoord) { acc.push(b('respuesta', 'Respuesta del rival', true)); acc.push(b('mensaje', 'Reenviar')); }
    if (est === 'Contrapropuesta' && ctx.esCoord) { acc.push(b('aceptarcp', 'Aceptar su hora', true)); acc.push(b('proponer', 'Proponer otra')); }
    if (est === 'Aceptado' && ctx.esCoord) { acc.push(b('cargado', 'Cargado en FBPA', true)); }
    if (est !== 'Oficial') { acc.push(b('historial', 'Historial')); }
    return '<tr' + (grav[p.id] === 'critica' ? ' style="background:#FFF6F6"' : '') + '><td><b>' + esc(DIA_TXT[p.fecha.getDay()].slice(0, 3)) + ' ' + hm(minDia(p.fecha)) + '</b>' + (cambiado ? '<br><small style="text-decoration:line-through;color:#5c5c5c">' + hm(minDia(p.oficial)) + (p.oficial.toDateString() !== p.fecha.toDateString() ? ' ' + DIA_TXT[p.oficial.getDay()].slice(0, 3) : '') + ' FBPA</small>' : '') + '</td>' +
      '<td><b>' + esc(p.equipo) + '</b><br><small style="color:#5c5c5c">' + (p.local ? 'vs ' : 'en ') + esc(p.rival) + '</small></td>' +
      '<td>' + (p.local ? esc(p.lugar) : '<small>Fuera · ' + esc(p.lugar) + '</small>') + '</td><td>' + pista + '</td>' +
      '<td>' + pillEstado(est) + (g && g.ok ? '<br><small style="color:#1B5E20">✓ ' + esc(g.ok) + '</small>' : '') + (est === 'Contrapropuesta' && g ? '<br><small>' + esc(g.respuesta) + '</small>' : '') +
      (est === 'Cargado en FBPA' && !enFbpa ? '<br><small style="color:#B36F00">La FBPA aún no lo refleja</small>' : '') + '</td>' +
      '<td><div style="display:flex;gap:6px;flex-wrap:wrap">' + acc.join('') + '</div></td></tr>';
  };
  const herramientas = '<div class="cbp-box-h" style="flex-wrap:wrap;gap:8px"><h3 class="cbp-h2" style="font-size:24px;margin:0">Partidos de la semana</h3><div style="display:flex;gap:8px;flex-wrap:wrap">' +
    (ctx.esCoord ? '<button type="button" class="cbp-btn sm pri" data-g="semana" style="color:#FF9F00">Proponer semana sin conflictos</button><button type="button" class="cbp-btn sm" data-g="pistas">Asignar pistas</button>' : '') +
    '<button type="button" class="cbp-btn sm" data-g="fbpa">Para cargar en la FBPA</button><button type="button" class="cbp-btn sm" data-g="excel">Excel Patronato</button></div></div>';
  return herramientas + (ctx.esCoord ? '' : '<p class="cbp-nota" style="margin:0">Puedes proponer cambios y dar el OK como entrenador. Validar, pedir cambios a otros clubes y cargar en la FBPA lo hace el grupo «Coordinación deportiva».</p>') +
    '<section class="cbp-box" style="padding:8px 12px"><div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Hora</th><th>Partido</th><th>Instalación</th><th>Pista</th><th>Estado</th><th></th></tr></thead><tbody>' + ps.map(fila).join('') + '</tbody></table></div></section>';
}

/* ---------- Acciones ---------- */
export function enlazarGestion(root: HTMLElement, ctx: IGCtx, al: IAlerta[]): void {
  const pid = (s: string): IPartidoG | undefined => ctx.semana.filter((p: IPartidoG) => p.id === +s)[0];
  root.querySelectorAll('[data-ac]').forEach((b: Element) => b.addEventListener('click', () => {
    const [a, id] = (b.getAttribute('data-ac') || '').split('|'); const p = pid(id); if (!p) { return; }
    const g = ctx.gest[p.id];
    const fin = (msg: string): () => void => () => { aviso(msg); ctx.recargar(); };
    const err = (x: Error): void => aviso('Error: ' + x.message, true);
    if (a === 'proponer') { modalPropuesta(ctx, p, al); }
    else if (a === 'validar') { guardarGest(ctx, p, { EstadoG: 'Validado' }, 'Validado').then(fin('Propuesta validada')).catch(err); }
    else if (a === 'ok') { guardarGest(ctx, p, { OKEntrenador: nombreCorto(ctx.usuario) }, 'OK del entrenador').then(fin('OK registrado')).catch(err); }
    else if (a === 'descartar') { guardarGest(ctx, p, { EstadoG: 'Oficial', FechaPropuesta: null }, 'Propuesta descartada').then(fin('Propuesta descartada')).catch(err); }
    else if (a === 'mensaje') { modalMensaje(ctx, p); }
    else if (a === 'respuesta') { modalRespuesta(ctx, p); }
    else if (a === 'aceptarcp') { guardarGest(ctx, p, { EstadoG: 'Aceptado' }, 'Aceptada la hora del rival').then(fin('Hora del rival aceptada')).catch(err); }
    else if (a === 'cargado') { guardarGest(ctx, p, { EstadoG: 'Cargado en FBPA' }, 'Cargado en la intranet FBPA').then(fin('Marcado como cargado')).catch(err); }
    else if (a === 'historial' && g) { abrirModal('Historial · ' + p.equipo, '<pre style="white-space:pre-wrap;font:inherit;font-size:14px;margin:0">' + esc(g.historial || 'Sin movimientos') + '</pre>' + (g.motivo ? '<p class="cbp-nota"><b>Justificación:</b> ' + esc(g.motivo) + '</p>' : ''), 560); }
  }));
  root.querySelectorAll('[data-pista]').forEach((s: Element) => s.addEventListener('change', () => {
    const p = pid((s as HTMLElement).getAttribute('data-pista') || ''); if (!p) { return; }
    guardarGest(ctx, p, { Pista: (s as HTMLSelectElement).value }, 'Pista ' + ((s as HTMLSelectElement).value || 'sin asignar')).then(() => { aviso('Pista guardada'); ctx.recargar(); }).catch((x: Error) => aviso('Error: ' + x.message, true));
  }));
  root.querySelectorAll('[data-g]').forEach((b: Element) => b.addEventListener('click', () => {
    const k = b.getAttribute('data-g');
    if (k === 'excel') { excelPatronato(ctx); } else if (k === 'fbpa') { listaFbpa(ctx); } else if (k === 'pistas') { autoPistas(ctx); } else if (k === 'semana') { proponerSemana(ctx, al); }
  }));
  root.querySelectorAll('[data-sol]').forEach((b: Element) => b.addEventListener('click', (ev: Event) => {
    ev.stopPropagation(); const p = pid(b.getAttribute('data-sol') || ''); if (p) { modalPropuesta(ctx, p, al); }
  }));
}

/** Justificación por defecto a partir de la alerta */
function justificacion(p: IPartidoG, al: IAlerta[]): string {
  const a = al.filter((x: IAlerta) => x.partidos.indexOf(p.id) >= 0 && x.grav === 'critica')[0] || al.filter((x: IAlerta) => x.partidos.indexOf(p.id) >= 0)[0];
  if (!a) { return ''; }
  if (/Entrenador/.test(a.titulo)) { return 'Ese día el mismo entrenador dirige otro partido de nuestro club y no llega a tiempo.'; }
  if (/Jugadores|convocatoria/i.test(a.titulo)) { return 'Varios jugadores juegan también en otro equipo de nuestro club que tiene partido a una hora muy cercana.'; }
  if (a.tipo === 'saturacion' || a.tipo === 'reserva') { return 'No tenemos pista disponible a esa hora en nuestra instalación.'; }
  if (a.tipo === 'franja') { return 'Ajustar el partido a la franja horaria oficial de la FBPA.'; }
  return a.titulo;
}

function modalPropuesta(ctx: IGCtx, p: IPartidoG, al: IAlerta[]): void {
  const g = ctx.gest[p.id];
  const info = ctx.eqInfo(p);
  const asign: { [id: number]: string } = {}; Object.keys(ctx.gest).forEach((k: string) => { asign[+k] = ctx.gest[+k].pista; });
  const sug: IHueco[] = sugerir(p, ctx.semana, ctx.reglas, ctx.pistas, ctx.margen, ctx.jugs, info.pref, info.alt, asign);
  const mias = al.filter((a: IAlerta) => a.partidos.indexOf(p.id) >= 0);
  const mods = ctx.pistas.filter((k: IPista) => info.pref.concat(info.alt).some((c: string) => norm(c) === norm(k.clave)));
  const html = '<p style="margin:0"><b>' + esc(p.equipo) + '</b> ' + (p.local ? 'vs ' : 'en casa de ') + esc(p.rival) + '<br><span class="cbp-nota">Oficial FBPA: ' + fTxt(p.oficial) + (p.local ? '' : ' · ' + esc(p.lugar)) + '</span></p>' +
    (mias.length ? '<div class="cbp-odd" style="margin:0;background:#FFF6F6">' + mias.map((a: IAlerta) => '• ' + esc(a.titulo)).join('\n') + '</div>' : '') +
    '<h4 style="margin:6px 0 0;font-size:13px;letter-spacing:.06em;text-transform:uppercase">Huecos sin conflicto' + (p.local ? '' : ' (para pedir al club local)') + '</h4>' +
    (sug.length ? '<div style="display:flex;flex-direction:column;gap:6px">' + sug.map((h: IHueco, i: number) => '<label class="cbp-check" style="border:1px solid #E6E3DC;border-radius:12px;padding:10px 12px;font-weight:500"><input type="radio" name="hueco" value="' + i + '"' + (i === 0 ? ' checked' : '') + '><span><b>' + fTxt(h.fecha) + '</b>' + (h.pista ? ' · ' + esc((ctx.pistas.filter((k: IPista) => k.clave === h.pista)[0] || { modulo: h.pista } as IPista).inst + ' ' + ((ctx.pistas.filter((k: IPista) => k.clave === h.pista)[0] || { modulo: '' } as IPista).modulo)) : '') + '<br><small class="cbp-nota">' + esc(h.nota) + '</small></span></label>').join('') + '</div>'
      : '<p class="cbp-nota" style="margin:0">El motor no encuentra huecos libres con las reservas y reglas actuales. Puedes poner una hora a mano.</p>') +
    '<label class="cbp-check" style="margin-top:4px"><input type="radio" name="hueco" value="manual"' + (sug.length ? '' : ' checked') + '>Otra hora</label>' +
    '<div class="cbp-form" style="grid-template-columns:repeat(' + (p.local ? 3 : 2) + ',minmax(0,1fr))"><label class="cbp-field">Día<input type="date" name="dia" value="' + fechaInput(g && g.fechaProp ? g.fechaProp : p.fecha) + '"></label>' +
    '<label class="cbp-field">Hora<input type="time" name="hora" step="900" value="' + hm(minDia(g && g.fechaProp ? g.fechaProp : p.fecha)) + '"></label>' +
    (p.local ? '<label class="cbp-field">Pista<select name="pista">' + mods.map((k: IPista) => '<option value="' + esc(k.clave) + '"' + (g && g.pista === k.clave ? ' selected' : '') + '>' + esc(k.inst + ' ' + k.modulo) + '</option>').join('') + '</select></label>' : '') + '</div>' +
    '<label class="cbp-field">Justificación para el otro club<textarea name="motivo" rows="2">' + esc(g && g.motivo ? g.motivo : justificacion(p, al)) + '</textarea></label>' +
    (ctx.esCoord ? '<label class="cbp-check"><input type="checkbox" name="validar" checked>Validar directamente (soy coordinación)</label>' : '') +
    '<div class="cbp-btns"><button type="button" class="cbp-btn" data-c="1">Cancelar</button><button type="button" class="cbp-btn pri" data-ok="1">Guardar propuesta</button></div>';
  const body = abrirModal('Proponer cambio', html, 640);
  (body.querySelector('[data-c]') as HTMLElement).addEventListener('click', cerrarModal);
  (body.querySelector('[data-ok]') as HTMLElement).addEventListener('click', () => {
    const sel = (body.querySelector('[name="hueco"]:checked') as HTMLInputElement).value;
    let fecha: Date; let pista = '';
    if (sel !== 'manual') { const h = sug[+sel]; fecha = h.fecha; pista = h.pista; }
    else {
      const d = (body.querySelector('[name="dia"]') as HTMLInputElement).value; const t = (body.querySelector('[name="hora"]') as HTMLInputElement).value;
      if (!d || !t) { aviso('Indica día y hora', true); return; }
      fecha = new Date(d + 'T' + t + ':00'); const ps = body.querySelector('[name="pista"]') as HTMLSelectElement; pista = ps ? ps.value : '';
    }
    const validar = !!(body.querySelector('[name="validar"]') as HTMLInputElement || { checked: false }).checked;
    const motivo = (body.querySelector('[name="motivo"]') as HTMLTextAreaElement).value.trim();
    const cambios: { [k: string]: any } = { FechaPropuesta: fecha.toISOString(), EstadoG: validar ? 'Validado' : 'Propuesto', Motivo: motivo, OKEntrenador: '' }; // eslint-disable-line @typescript-eslint/no-explicit-any
    if (p.local && pista) { cambios.Pista = pista; }
    guardarGest(ctx, p, cambios, (validar ? 'Propuesto y validado: ' : 'Propuesto: ') + fTxt(fecha)).then(() => { cerrarModal(); aviso('Propuesta guardada'); ctx.recargar(); }).catch((x: Error) => aviso('Error: ' + x.message, true));
  });
}

/* ---------- Mensaje al club rival ---------- */
function textoMensaje(ctx: IGCtx, p: IPartidoG, club?: IClub): string {
  const g = ctx.gest[p.id]; const nueva = g && g.fechaProp ? g.fechaProp : p.fecha;
  const saludo = 'Hola' + (club && club.contacto ? ' ' + club.contacto.split(' ')[0] : '') + ', os escribimos desde la coordinación deportiva del Gijón CBP (C.B. Pumarín).';
  const partido = 'Partido ' + (p.catPDM || p.equipo) + (p.jornada ? ', jornada ' + p.jornada : '') + ': ' + (p.local ? 'Gijón CBP – ' + p.rival : p.rival + ' – Gijón CBP') + ', fijado el ' + fTxt(p.oficial) + (p.local ? '' : ' en ' + p.lugar) + '.';
  const pide = p.local
    ? 'Os proponemos jugarlo el ' + fTxt(nueva) + ' en ' + (g && g.pista ? (ctx.pistas.filter((k: IPista) => k.clave === g.pista)[0] || { inst: p.lugar } as IPista).inst : p.lugar) + '. ¿Nos dais vuestra conformidad para cargarlo en la intranet de la FBPA?'
    : 'Os pedimos, si es posible, pasarlo al ' + fTxt(nueva) + '. Si os encaja, ¿podéis cambiarlo en la intranet de la FBPA?';
  return saludo + '\n\n' + partido + '\n' + pide + (g && g.motivo ? '\nMotivo: ' + g.motivo : '') + '\n\nMuchas gracias y un saludo,\nCoordinación deportiva · Gijón CBP';
}
function modalMensaje(ctx: IGCtx, p: IPartidoG): void {
  const club = clubDe(ctx.clubes, p.rival);
  const txt = textoMensaje(ctx, p, club);
  const tel = club && club.tel ? club.tel.replace(/[^\d+]/g, '') : '';
  const wa = tel ? 'https://wa.me/' + (tel.charAt(0) === '+' ? tel.slice(1) : (tel.length === 9 ? '34' + tel : tel)) + '?text=' + encodeURIComponent(txt) : '';
  const asunto = 'Gijón CBP · Cambio de horario ' + (p.catPDM || p.equipo) + ' – ' + p.rival;
  const mail = club && club.email ? 'mailto:' + club.email + '?subject=' + encodeURIComponent(asunto) + '&body=' + encodeURIComponent(txt) : '';
  const html = '<p style="margin:0"><b>' + esc(club ? club.nombre : p.rival) + '</b>' + (club && club.contacto ? ' · ' + esc(club.contacto) : '') + '<br><span class="cbp-nota">' + (club ? (club.email || 'sin correo') + ' · ' + (club.tel || 'sin teléfono') : 'Club no identificado: añade sus datos en la pestaña «Clubes»') + '</span></p>' +
    '<label class="cbp-field">Mensaje<textarea name="txt" rows="11">' + esc(txt) + '</textarea></label>' +
    '<div class="cbp-btns" style="justify-content:flex-start">' +
    (wa ? '<a class="cbp-btn pri" data-env="wa" href="' + esc(wa) + '" target="_blank" rel="noopener" style="color:#FF9F00">Abrir en WhatsApp</a>' : '') +
    (mail ? '<a class="cbp-btn' + (wa ? '' : ' pri') + '" data-env="mail" href="' + esc(mail) + '">Abrir correo</a>' : '') +
    '<button type="button" class="cbp-btn" data-env="copiar">Copiar texto</button></div>' +
    '<p class="cbp-nota" style="margin:0">Al enviarlo, el partido pasa a «Pedido al rival». Cuando contesten, registra su respuesta.</p>';
  const body = abrirModal((p.local ? 'Comunicar a ' : 'Pedir a ') + (club ? club.nombre : p.rival), html, 620);
  const marcar = (): void => { if (ctx.gest[p.id] && ctx.gest[p.id].estado === 'Pedido al rival') { return; } guardarGest(ctx, p, { EstadoG: 'Pedido al rival' }, 'Petición enviada a ' + (club ? club.nombre : p.rival)).then(() => ctx.recargar()); };
  body.querySelectorAll('[data-env]').forEach((b: Element) => b.addEventListener('click', (ev: Event) => {
    const t = (body.querySelector('[name="txt"]') as HTMLTextAreaElement).value;
    const k = b.getAttribute('data-env');
    if (k === 'copiar') { ev.preventDefault(); navigator.clipboard.writeText(t).then(() => aviso('Texto copiado')); marcar(); return; }
    if (k === 'wa' && tel) { (b as HTMLAnchorElement).href = wa.split('?')[0] + '?text=' + encodeURIComponent(t); }
    if (k === 'mail' && club) { (b as HTMLAnchorElement).href = 'mailto:' + club.email + '?subject=' + encodeURIComponent(asunto) + '&body=' + encodeURIComponent(t); }
    marcar();
  }));
}

function modalRespuesta(ctx: IGCtx, p: IPartidoG): void {
  formModal('Respuesta de ' + p.rival, [
    { n: 'r', label: 'Respuesta', tipo: 'select', opciones: ['Aceptan', 'Proponen otra hora', 'No aceptan'], valor: 'Aceptan', full: true },
    { n: 'dia', label: 'Su propuesta: día', tipo: 'date', valor: fechaInput(p.fecha) }, { n: 'hora', label: 'Hora', tipo: 'text', valor: hm(minDia(p.fecha)), ph: '12:00' },
    { n: 'nota', label: 'Nota', tipo: 'area', ph: 'Lo que han contestado' }
  ], 'Guardar respuesta', (v: { [k: string]: any }) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    if (v.r === 'Aceptan') { return guardarGest(ctx, p, { EstadoG: 'Aceptado', Respuesta: v.nota || 'Aceptan' }, 'El rival acepta').then(() => { aviso('Aceptado: queda cargarlo en la FBPA'); ctx.recargar(); }); }
    if (v.r === 'No aceptan') { return guardarGest(ctx, p, { EstadoG: 'Rechazado', Respuesta: v.nota || 'No aceptan', FechaPropuesta: null }, 'El rival no acepta').then(() => { aviso('Rechazado: se vuelve al horario oficial'); ctx.recargar(); }); }
    const f = new Date(v.dia + 'T' + (/^\d{1,2}:\d{2}$/.test(v.hora) ? ('0' + v.hora).slice(-5) : '12:00') + ':00');
    return guardarGest(ctx, p, { EstadoG: 'Contrapropuesta', FechaPropuesta: f.toISOString(), Respuesta: 'Proponen ' + fTxt(f) + (v.nota ? '. ' + v.nota : '') }, 'Contrapropuesta del rival: ' + fTxt(f))
      .then(() => { aviso('Contrapropuesta registrada: revisa las alertas con su hora'); ctx.recargar(); });
  }, 560);
}

/* ---------- Pistas automáticas ---------- */
function autoPistas(ctx: IGCtx): void {
  const asign: { [id: number]: string } = {}; Object.keys(ctx.gest).forEach((k: string) => { if (ctx.gest[+k].pista) { asign[+k] = ctx.gest[+k].pista; } });
  const res = asignarPistas(ctx.semana, ctx.pistas, ctx.margen, ctx.eqInfo, asign);
  const nuevos = ctx.semana.filter((p: IPartidoG) => res[p.id] && !asign[p.id]);
  if (!nuevos.length) { aviso('Todos los partidos en casa ya tienen pista'); return; }
  nuevos.reduce((pr: Promise<void>, p: IPartidoG) => pr.then(() => guardarGest(ctx, p, { Pista: res[p.id] }, 'Pista asignada automáticamente: ' + res[p.id])), Promise.resolve())
    .then(() => { aviso(nuevos.length + ' pistas asignadas'); ctx.recargar(); }).catch((x: Error) => aviso('Error: ' + x.message, true));
}

/* ---------- Proponer la semana completa ---------- */
function proponerSemana(ctx: IGCtx, al: IAlerta[]): void {
  const criticas = al.filter((a: IAlerta) => a.grav === 'critica');
  if (!criticas.length) { aviso('No hay conflictos que resolver esta semana'); return; }
  const trabajo: IPartidoG[] = ctx.semana.map((p: IPartidoG) => ({ ...p }));
  const asign: { [id: number]: string } = {}; Object.keys(ctx.gest).forEach((k: string) => { asign[+k] = ctx.gest[+k].pista; });
  const propuestas: { p: IPartidoG; h: IHueco; motivo: string }[] = [];
  const sinSol: string[] = [];
  criticas.forEach((a: IAlerta) => {
    const ya = a.partidos.some((id: number) => propuestas.some((x: { p: IPartidoG }) => x.p.id === id)); if (ya) { return; }
    const cands = a.partidos.map((id: number) => trabajo.filter((p: IPartidoG) => p.id === id)[0]).filter((p: IPartidoG | undefined) => !!p && !(ctx.gest[p.id] && ctx.gest[p.id].estado !== 'Oficial' && ctx.gest[p.id].estado !== 'Rechazado')) as IPartidoG[];
    cands.sort((x: IPartidoG, y: IPartidoG) => (y.local ? 1 : 0) - (x.local ? 1 : 0));
    for (const p of cands) {
      const info = ctx.eqInfo(p);
      const h = sugerir(p, trabajo, ctx.reglas, ctx.pistas, ctx.margen, ctx.jugs, info.pref, info.alt, asign)[0];
      if (h) { propuestas.push({ p, h, motivo: justificacion(p, al) }); p.fecha = h.fecha; if (h.pista) { asign[p.id] = h.pista; } return; }
    }
    sinSol.push(a.titulo);
  });
  const resto = analizar(trabajo, ctx.reglas, ctx.pistas, ctx.margen, ctx.jugs).filter((a: IAlerta) => a.grav === 'critica').length;
  const html = (propuestas.length ? '<table class="cbp-tabla"><thead><tr><th>Partido</th><th>FBPA</th><th>Propuesta</th></tr></thead><tbody>' + propuestas.map((x: { p: IPartidoG; h: IHueco }) =>
    '<tr><td><b>' + esc(x.p.equipo) + '</b><br><small>' + (x.p.local ? 'en casa' : 'fuera, a pedir al rival') + '</small></td><td>' + fTxt(x.p.oficial) + '</td><td><b>' + fTxt(x.h.fecha) + '</b>' + (x.h.pista ? '<br><small>' + esc(x.h.pista) + '</small>' : '') + '</td></tr>').join('') + '</tbody></table>' : '') +
    (sinSol.length ? '<p class="cbp-nota" style="color:#B00020">Sin solución automática: ' + esc(sinSol.join('; ')) + '</p>' : '') +
    '<p class="cbp-nota">Con estos cambios quedarían <b>' + resto + '</b> conflicto' + (resto === 1 ? '' : 's') + '. Se guardan como «Propuesto» para que los entrenadores den el OK y coordinación valide.</p>' +
    '<div class="cbp-btns"><button type="button" class="cbp-btn" data-c="1">Cancelar</button>' + (propuestas.length ? '<button type="button" class="cbp-btn pri" data-ok="1">Guardar ' + propuestas.length + ' propuesta' + (propuestas.length === 1 ? '' : 's') + '</button>' : '') + '</div>';
  const body = abrirModal('Propuesta de la semana', html, 700);
  (body.querySelector('[data-c]') as HTMLElement).addEventListener('click', cerrarModal);
  const ok = body.querySelector('[data-ok]') as HTMLElement;
  if (ok) {
    ok.addEventListener('click', () => {
      propuestas.reduce((pr: Promise<void>, x: { p: IPartidoG; h: IHueco; motivo: string }) => pr.then(() => {
        const orig = ctx.semana.filter((q: IPartidoG) => q.id === x.p.id)[0];
        const c: { [k: string]: any } = { FechaPropuesta: x.h.fecha.toISOString(), EstadoG: 'Propuesto', Motivo: x.motivo, OKEntrenador: '' }; // eslint-disable-line @typescript-eslint/no-explicit-any
        if (x.h.pista) { c.Pista = x.h.pista; }
        return guardarGest(ctx, orig, c, 'Propuesta automática: ' + fTxt(x.h.fecha));
      }), Promise.resolve()).then(() => { cerrarModal(); aviso('Propuestas guardadas'); ctx.recargar(); }).catch((x: Error) => aviso('Error: ' + x.message, true));
    });
  }
}

/* ---------- Para cargar en la FBPA ---------- */
function listaFbpa(ctx: IGCtx): void {
  const ps = ctx.semana.filter((p: IPartidoG) => { const g = ctx.gest[p.id]; return g && (g.estado === 'Validado' || g.estado === 'Aceptado' || g.estado === 'Pedido al rival' || g.estado === 'Cargado en FBPA') && g.fechaProp; });
  const html = ps.length ? '<p class="cbp-nota" style="margin:0">Cambios para pasar a la intranet de la federación. La sincronización nocturna comprueba si la FBPA ya los refleja.</p><table class="cbp-tabla"><thead><tr><th>Partido</th><th>Ahora en FBPA</th><th>Cambiar a</th><th>Estado</th></tr></thead><tbody>' +
    ps.map((p: IPartidoG) => { const g = ctx.gest[p.id]; const ya = g.fechaProp && Math.abs(g.fechaProp.getTime() - p.oficial.getTime()) < 60000;
      return '<tr><td><b>' + esc(p.equipo) + '</b><br><small>' + (p.local ? 'vs ' : 'en ') + esc(p.rival) + (p.jornada ? ' · J' + esc(p.jornada) : '') + '</small></td><td>' + fTxt(p.oficial) + '</td><td><b>' + fTxt(g.fechaProp as Date) + '</b>' + (g.pista ? '<br><small>' + esc(g.pista) + '</small>' : '') + '</td><td>' + (ya ? '<span class="cbp-fpill ok">Ya en FBPA</span>' : pillEstado(g.estado)) + (g.estado === 'Pedido al rival' ? '<br><small>esperando al rival</small>' : '') + '</td></tr>'; }).join('') + '</tbody></table>'
    : '<p class="cbp-nota">No hay cambios validados pendientes de cargar esta semana.</p>';
  abrirModal('Para cargar en la FBPA', html, 760);
}

/* ---------- Excel del Patronato ---------- */
export function excelPatronato(ctx: IGCtx): void {
  const casa = ctx.semana.filter((p: IPartidoG) => p.local).sort((a: IPartidoG, b: IPartidoG) => a.fecha.getTime() - b.fecha.getTime());
  const asign: { [id: number]: string } = {}; Object.keys(ctx.gest).forEach((k: string) => { if (ctx.gest[+k].pista) { asign[+k] = ctx.gest[+k].pista; } });
  const pistasAuto = asignarPistas(ctx.semana, ctx.pistas, ctx.margen, ctx.eqInfo, asign);
  const filas = casa.map((p: IPartidoG) => {
    const d = p.fecha; const g = ctx.gest[p.id];
    const clave = (g && g.pista) || pistasAuto[p.id] || '';
    const pi = ctx.pistas.filter((k: IPista) => k.clave === clave)[0];
    const club = clubDe(ctx.clubes, p.rival);
    const pend = g && (g.estado === 'Propuesto' || g.estado === 'Pedido al rival' || g.estado === 'Contrapropuesta');
    return { p, pend, sinPista: !(g && g.pista), celdas: [('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + String(d.getFullYear()).slice(-2), 'Bcesto', 'CB Pumarín',
      club && club.corto ? club.corto : titulo(club ? club.nombre : p.rival), p.catPDM || p.equipo, pi ? pi.pdm || pi.clave : p.lugar, DIAS_C[d.getDay()], hm(minDia(d)) + 'h'] };
  });
  const html = '<p class="cbp-nota" style="margin:0">Partidos en casa de la semana con el horario vigente (los cambios validados ya aplicados). Revisa los marcados antes de enviarlo.</p>' +
    '<div class="cbp-scroll"><table class="cbp-tabla"><tbody>' + filas.map((f: { p: IPartidoG; pend: boolean; sinPista: boolean; celdas: string[] }) => '<tr' + (f.pend ? ' style="background:#FFF6E5"' : '') + '>' + f.celdas.map((c: string) => '<td>' + esc(c) + '</td>').join('') +
      '<td>' + (f.pend ? '<span class="cbp-fpill amb">Sin confirmar</span>' : '') + (f.sinPista ? ' <span class="cbp-fpill">Pista sugerida</span>' : '') + '</td></tr>').join('') + '</tbody></table></div>' +
    '<div class="cbp-btns"><button type="button" class="cbp-btn" data-copia="1">Copiar tabla</button><button type="button" class="cbp-btn pri" data-xlsx="1">Descargar Excel</button></div>';
  const body = abrirModal('Excel para el Patronato', html, 900);
  const lunes = ctx.semana.length ? ctx.semana[0].fecha : new Date();
  (body.querySelector('[data-xlsx]') as HTMLElement).addEventListener('click', () => {
    const blob = libroXlsx('Partidos', [], filas.map((f: { celdas: string[] }) => f.celdas), [10, 8, 13, 28, 24, 16, 6, 8]);
    descargar('Patronato-GijonCBP-' + fechaInput(lunes) + '.xlsx', blob, blob.type);
  });
  (body.querySelector('[data-copia]') as HTMLElement).addEventListener('click', () => {
    navigator.clipboard.writeText(filas.map((f: { celdas: string[] }) => f.celdas.join('\t')).join('\n')).then(() => aviso('Tabla copiada: pégala en el Excel del Patronato'));
  });
}

/* ---------- Clubes ---------- */
export function vClubes(ctx: IGCtx, rivales: string[]): string {
  const juegan = (k: IClub): boolean => rivales.some((r: string) => clubDe(ctx.clubes, r) === k);
  const lista = ctx.clubes.slice().sort((a: IClub, b: IClub) => (juegan(b) ? 1 : 0) - (juegan(a) ? 1 : 0) || a.nombre.localeCompare(b.nombre));
  const sinClub = rivales.filter((r: string) => !clubDe(ctx.clubes, r));
  return '<p class="cbp-nota" style="margin:0">Clubes de la FBPA con los equipos con los que jugamos. La FBPA no publica correos ni teléfonos: añádelos una vez y se usarán en las peticiones de cambio. El «nombre corto» es el que aparece en el Excel del Patronato.</p>' +
    (sinClub.length ? '<p class="cbp-nota" style="color:#B36F00;margin:0">Equipos rivales sin club identificado: ' + esc(sinClub.slice(0, 12).join(', ')) + (sinClub.length > 12 ? '…' : '') + '</p>' : '') +
    '<section class="cbp-box" style="padding:8px 12px"><div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Club</th><th>Nombre corto</th><th>Contacto</th><th>Correo</th><th>Teléfono</th></tr></thead><tbody>' +
    lista.map((k: IClub) => '<tr class="clic" data-club="' + k.id + '"' + (juegan(k) ? '' : ' style="opacity:.6"') + '><td><b>' + esc(k.nombre) + '</b>' + (juegan(k) && !k.email && !k.tel ? ' <span class="cbp-falta">sin contacto</span>' : '') + '</td><td>' + esc(k.corto) + '</td><td>' + esc(k.contacto) + '</td><td>' + esc(k.email) + '</td><td>' + esc(k.tel) + '</td></tr>').join('') +
    '</tbody></table></div></section>';
}
export function enlazarClubes(root: HTMLElement, ctx: IGCtx): void {
  root.querySelectorAll('[data-club]').forEach((b: Element) => b.addEventListener('click', () => {
    const k = ctx.clubes.filter((x: IClub) => x.id === +(b.getAttribute('data-club') || 0))[0]; if (!k) { return; }
    formModal(k.nombre, [
      { n: 'NombreCorto', label: 'Nombre corto (Patronato)', valor: k.corto, ph: 'Ej.: Centro Asturiano Oviedo' }, { n: 'Contacto', label: 'Persona de contacto', valor: k.contacto },
      { n: 'Email', label: 'Correo', tipo: 'email', valor: k.email }, { n: 'Telefono', label: 'Teléfono / WhatsApp', valor: k.tel, ph: '600 000 000' },
      { n: 'EquiposManual', label: 'Equipos de este club (corrección manual)', tipo: 'area', valor: k.manualTxt, ayuda: '(uno por línea, tal como aparecen en la FBPA; tienen prioridad sobre la detección automática)' }
    ], 'Guardar', (v: { [x: string]: any }) => spPost(ctx.c, porTitulo(ctx.w, L_CLU) + '/items(' + k.id + ')', v, true).then(() => { aviso('Club guardado'); ctx.recargar(); }), 520, // eslint-disable-line @typescript-eslint/no-explicit-any
      '<p class="cbp-nota full" style="margin:0">' + esc(k.direccion) + (k.url ? ' · <a href="https://www.fbpa.es' + esc(k.url) + '" target="_blank" rel="noopener">Ver en la FBPA</a>' : '') + '<br>Equipos detectados: ' + esc(k.equipos.join(' · ') || 'ninguno') + '</p>');
  }));
}
