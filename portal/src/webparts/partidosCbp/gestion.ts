import { SPHttpClient } from '@microsoft/sp-http';
import { esc, MESES } from '../../common/cbp';
import { spPost, abrirModal, cerrarModal, aviso, descargar, fechaInput } from '../../common/ui';
import { porTitulo, getJ, formModal } from '../../common/sp2';
import { libroXlsx } from '../../common/xlsx';
import { IPartido, IRegla, IPista, IAlerta, IJugV, IHueco, IPlazo, sugerir, asignarPistas, analizar, norm, minDia, hm, DIA_TXT, enPlazo, limitePlazo } from './motor';

export const ESTADOS = ['Oficial', 'Propuesto', 'Validado', 'Pedido al rival', 'Aceptado', 'Contrapropuesta', 'Rechazado', 'Cargado en FBPA'];
const L_GES = 'Partidos · Gestión'; const L_CLU = 'Partidos · Clubes';

export interface IPartidoG extends IPartido { oficial: Date; comp: string; jornada: string; catPDM: string; lugarFbpa?: string; }
export interface IGest { id: number; pid: number; fechaProp?: Date; pista: string; estado: string; motivo: string; respuesta: string; ok: string; historial: string; fijo: boolean; petHora?: Date; petNota: string; petDe: string; }
export interface IClub { id: number; nombre: string; corto: string; email: string; tel: string; contacto: string; url: string; direccion: string; equipos: string[]; manual: string[]; manualTxt: string; }
export interface IEqInfo { pref: string[]; alt: string[]; }

export interface IGCtx {
  c: SPHttpClient; w: string; usuario: string; esCoord: boolean; margen: number; plazo: IPlazo;
  semana: IPartidoG[]; gest: { [pid: number]: IGest }; clubes: IClub[]; pistas: IPista[]; reglas: IRegla[]; jugs: IJugV[];
  eqInfo: (p: IPartido) => IEqInfo; recargar: () => void;
}

/* ---------- Carga ---------- */
export function cargarGestion(c: SPHttpClient, w: string): Promise<{ [pid: number]: IGest }> {
  return getJ(c, porTitulo(w, L_GES) + '/items?$select=Id,PartidoId,FechaPropuesta,Pista,EstadoG,Motivo,Respuesta,OKEntrenador,Historial,Fijo,PeticionHora,PeticionNota,PeticionDe&$top=3000').then((j: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    const m: { [pid: number]: IGest } = {};
    (j.value || []).forEach((i: any) => { m[i.PartidoId] = { id: i.Id, pid: i.PartidoId, fechaProp: i.FechaPropuesta ? new Date(i.FechaPropuesta) : undefined, pista: i.Pista || '', estado: i.EstadoG || 'Oficial', motivo: i.Motivo || '', respuesta: i.Respuesta || '', ok: i.OKEntrenador || '', historial: i.Historial || '', fijo: !!i.Fijo, petHora: i.PeticionHora ? new Date(i.PeticionHora) : undefined, petNota: i.PeticionNota || '', petDe: i.PeticionDe || '' }; }); // eslint-disable-line @typescript-eslint/no-explicit-any
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

/** Partidos que el motor no debe mover: candado de coordinación o cambio ya en trámite con el rival o la FBPA */
export const EN_TRAMITE = ['Validado', 'Pedido al rival', 'Aceptado', 'Contrapropuesta', 'Cargado en FBPA'];
export function esFijo(g?: IGest): boolean { return !!g && (g.fijo || EN_TRAMITE.indexOf(g.estado) >= 0); }
function fCorta(d: Date): string { return DIA_TXT[d.getDay()].slice(0, 3) + ' ' + d.getDate() + ', ' + hm(minDia(d)); }
export function plazoTxt(ctx: IGCtx, p: IPartidoG): string {
  if (!p.local) { return ''; }
  const lim = limitePlazo({ ...p, fecha: p.oficial }, ctx.plazo);
  return enPlazo({ ...p, fecha: p.oficial }, ctx.plazo) ? '<span class="cbp-plazo ok" title="Dentro del plazo ordinario: el local fija la hora sin conformidad del rival">En plazo · hasta ' + fCorta(lim) + '</span>'
    : '<span class="cbp-plazo" title="Pasado el plazo ordinario: el cambio necesita la conformidad del rival">Fuera de plazo</span>';
}

/* ---------- Tabla de partidos de la semana ---------- */
export function vPartidos(ctx: IGCtx, al: IAlerta[]): string {
  const ps = ctx.semana.slice().sort((a: IPartidoG, b: IPartidoG) => a.fecha.getTime() - b.fecha.getTime());
  const grav: { [id: number]: string } = {}; const tit: { [id: number]: string[] } = {};
  al.filter((a: IAlerta) => a.grav !== 'info').forEach((a: IAlerta) => a.partidos.forEach((id: number) => { if (!grav[id] || a.grav === 'critica') { grav[id] = a.grav; } (tit[id] = tit[id] || []).push(a.titulo); }));
  const opcPistas = (sel: string, lugar: string): string => {
    const insts: { [k: string]: IPista[] } = {}; ctx.pistas.forEach((k: IPista) => { (insts[k.inst] = insts[k.inst] || []).push(k); });
    return '<option value="">' + esc(lugar || 'Sin instalación') + ' (FBPA)</option>' + Object.keys(insts).map((inst: string) => '<optgroup label="' + esc(inst) + '">' +
      insts[inst].map((k: IPista) => '<option value="' + esc(k.clave) + '"' + (sel === k.clave ? ' selected' : '') + '>' + esc(inst + (k.modulo ? ' · ' + k.modulo : '')) + '</option>').join('') + '</optgroup>').join('');
  };
  const fila = (p: IPartidoG): string => {
    const g = ctx.gest[p.id]; const est = g ? g.estado : 'Oficial'; const fijo = !!(g && g.fijo);
    const cambiado = p.fecha.getTime() !== p.oficial.getTime();
    const enFbpa = g && g.fechaProp && Math.abs(g.fechaProp.getTime() - p.oficial.getTime()) < 60000;
    const pk = g && g.pista ? ctx.pistas.filter((k: IPista) => k.clave === g.pista)[0] : undefined;
    const lugar = p.local ? (ctx.esCoord && !fijo ? '<select class="cbp-select" data-pista="' + p.id + '" style="min-height:34px;font-size:13px;max-width:240px">' + opcPistas(g ? g.pista : '', p.lugarFbpa || p.lugar) + '</select>'
      : esc(pk ? pk.inst + (pk.modulo ? ' · ' + pk.modulo : '') : p.lugar)) : '<small>Fuera · ' + esc(p.lugar) + '</small>';
    const acc: string[] = [];
    const b = (a: string, t: string, pri?: boolean): string => '<button type="button" class="cbp-btn sm' + (pri ? ' pri' : '') + '" data-ac="' + a + '|' + p.id + '"' + (pri ? ' style="color:#FF9F00"' : '') + '>' + t + '</button>';
    const plazoOk = p.local && enPlazo({ ...p, fecha: p.oficial }, ctx.plazo);
    if (!fijo) {
      if (est === 'Oficial' || est === 'Rechazado') { if (ctx.esCoord) { acc.push(b('proponer', 'Proponer cambio', grav[p.id] === 'critica')); } acc.push(b('peticion', g && g.petHora ? 'Cambiar petición' : 'Pedir otra hora', !ctx.esCoord && grav[p.id] === 'critica')); }
      if (est === 'Propuesto') { if (!g || !g.ok) { acc.push(b('ok', 'Dar OK', !ctx.esCoord)); } if (ctx.esCoord) { acc.push(b('validar', 'Validar', true)); acc.push(b('proponer', 'Editar')); acc.push(b('descartar', 'Descartar')); } }
      if (est === 'Validado' && ctx.esCoord) {
        if (plazoOk) { acc.push(b('cargado', 'Cargado en FBPA', true)); acc.push(b('mensaje', 'Avisar al rival')); }
        else { acc.push(b('mensaje', p.local ? 'Pedir conformidad al rival' : 'Pedir al rival', true)); acc.push(b('cargado', 'Cargado en FBPA')); }
        acc.push(b('volver', 'Volver a propuesta')); acc.push(b('descartar', 'Descartar'));
      }
      if (est === 'Pedido al rival' && ctx.esCoord) { acc.push(b('respuesta', 'Respuesta del rival', true)); acc.push(b('mensaje', 'Reenviar')); acc.push(b('descartar', 'Descartar')); }
      if (est === 'Contrapropuesta' && ctx.esCoord) { acc.push(b('aceptarcp', 'Aceptar su hora', true)); acc.push(b('proponer', 'Proponer otra')); acc.push(b('descartar', 'Descartar')); }
      if (est === 'Aceptado' && ctx.esCoord) { acc.push(b('cargado', 'Cargado en FBPA', true)); acc.push(b('descartar', 'Descartar')); }
    }
    if (est !== 'Oficial') { acc.push(b('historial', 'Historial')); }
    const candado = ctx.esCoord ? '<button type="button" class="cbp-lock' + (fijo ? ' on' : '') + '" data-ac="fijar|' + p.id + '" title="' + (fijo ? 'Inamovible: no se puede proponer ni mover. Pulsa para soltarlo' : 'Marcar como inamovible') + '" aria-pressed="' + fijo + '">' + (fijo ? '🔒' : '🔓') + '</button>' : (fijo ? '<span title="Inamovible">🔒</span>' : '');
    const icono = grav[p.id] ? '<span class="cbp-al-ico ' + grav[p.id] + '" title="' + esc((tit[p.id] || []).join('\n')) + '">' + (grav[p.id] === 'critica' ? '!' : '·') + '</span>' : '<span class="cbp-al-ico ok" title="Sin conflictos con el horario vigente">✓</span>';
    return '<tr' + (grav[p.id] === 'critica' ? ' class="con"' : '') + (fijo ? ' data-fijo="1"' : '') + '><td>' + icono + '</td><td><b>' + esc(DIA_TXT[p.fecha.getDay()].slice(0, 3)) + ' ' + hm(minDia(p.fecha)) + '</b>' + (cambiado ? '<br><small style="text-decoration:line-through;color:#5c5c5c">' + hm(minDia(p.oficial)) + (p.oficial.toDateString() !== p.fecha.toDateString() ? ' ' + DIA_TXT[p.oficial.getDay()].slice(0, 3) : '') + ' FBPA</small>' : '') + '</td>' +
      '<td><div style="display:flex;gap:6px;align-items:flex-start">' + candado + '<div><b>' + esc(p.equipo) + '</b><br><small style="color:#5c5c5c">' + (p.local ? 'vs ' : 'en ') + esc(p.rival) + '</small>' +
      (p.local ? '<br>' + plazoTxt(ctx, p) : '') + (g && g.petHora && (est === 'Oficial' || est === 'Rechazado') ? '<br><small class="cbp-pet">Petición: ' + esc(fCorta(g.petHora)) + (g.petDe ? ' · ' + esc(g.petDe) : '') + (g.petNota ? ' · ' + esc(g.petNota) : '') + '</small>' : '') + (fijo ? '<br><small style="font-weight:700">Inamovible</small>' : '') + '</div></div></td>' +
      '<td>' + lugar + '</td>' +
      '<td>' + pillEstado(est) + (est === 'Propuesto' ? '<br><small style="color:' + (g && g.ok ? '#1B5E20' : '#B36F00') + '">' + (g && g.ok ? '✓ OK de ' + esc(g.ok) : 'Falta el OK del entrenador') + '</small>' : g && g.ok ? '<br><small style="color:#1B5E20">✓ ' + esc(g.ok) + '</small>' : '') + (est === 'Contrapropuesta' && g ? '<br><small>' + esc(g.respuesta) + '</small>' : '') +
      (est === 'Cargado en FBPA' && !enFbpa ? '<br><small style="color:#B36F00">Esperando a que la FBPA lo publique</small>' : '') + '</td>' +
      '<td><div style="display:flex;gap:6px;flex-wrap:wrap">' + acc.join('') + '</div></td></tr>';
  };
  const nProp = ps.filter((p: IPartidoG) => ctx.gest[p.id] && ctx.gest[p.id].estado === 'Propuesto').length;
  const herramientas = '<div class="cbp-box-h" style="flex-wrap:wrap;gap:8px"><h3 class="cbp-h2" style="font-size:24px;margin:0">Partidos de la semana</h3><div style="display:flex;gap:8px;flex-wrap:wrap">' +
    (ctx.esCoord && nProp ? '<button type="button" class="cbp-btn sm pri" data-g="validartodo" style="color:#FF9F00">Validar propuestas (' + nProp + ')</button>' : '') +
    '<button type="button" class="cbp-btn sm" data-g="resumen">Resumen para entrenadores</button><button type="button" class="cbp-btn sm" data-g="fbpa">Para cargar en la FBPA</button><button type="button" class="cbp-btn sm" data-g="excel">Excel Patronato</button></div></div>';
  return herramientas + '<p class="cbp-nota" style="margin:0">' + (ctx.esCoord ? 'Revisa partido a partido: propón cambios, marca inamovibles (🔒) o deja todo como está. El icono de la izquierda solo se pone en verde cuando el horario vigente, incluida la propuesta guardada, no tiene conflictos. Flujo: propuesta → OK del entrenador → validación → carga en la FBPA (o petición al rival si está fuera de plazo o jugamos fuera).'
    : 'Como entrenador puedes pedir otra hora para tus partidos en cualquier momento y dar el OK a las propuestas de coordinación. Validar, pedir cambios a otros clubes y cargar en la FBPA lo hace coordinación deportiva.') + '</p>' +
    '<section class="cbp-box" style="padding:8px 12px"><div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th></th><th>Hora</th><th>Partido</th><th>Instalación y pista</th><th>Estado</th><th></th></tr></thead><tbody>' + ps.map(fila).join('') + '</tbody></table></div></section>';
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
    else if (a === 'volver') { guardarGest(ctx, p, { EstadoG: 'Propuesto' }, 'Devuelto a propuesta').then(fin('Vuelve a estar como propuesta')).catch(err); }
    else if (a === 'fijar') { guardarGest(ctx, p, { Fijo: !(g && g.fijo) }, g && g.fijo ? 'Ya no es inamovible' : 'Marcado como inamovible').then(fin(g && g.fijo ? 'Partido liberado' : 'Partido inamovible')).catch(err); }
    else if (a === 'peticion') { modalPeticion(ctx, p); }
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
    if (k === 'excel') { excelPatronato(ctx); } else if (k === 'fbpa') { listaFbpa(ctx); } else if (k === 'resumen') { resumenEntrenadores(ctx); }
    else if (k === 'validartodo') {
      const ps = ctx.semana.filter((p: IPartidoG) => ctx.gest[p.id] && ctx.gest[p.id].estado === 'Propuesto');
      const sinOk = ps.filter((p: IPartidoG) => !ctx.gest[p.id].ok).length;
      const caja = abrirModal('Validar propuestas', '<p style="margin:0">Se validarán <b>' + ps.length + ' propuesta' + (ps.length === 1 ? '' : 's') + '</b>' + (sinOk ? ', ' + sinOk + ' todavía sin el OK del entrenador' : '') + '.</p><div class="cbp-btns"><button type="button" class="cbp-btn" data-c="1">Cancelar</button><button type="button" class="cbp-btn pri" data-si="1">Validar</button></div>', 480);
      (caja.querySelector('[data-c]') as HTMLElement).addEventListener('click', cerrarModal);
      (caja.querySelector('[data-si]') as HTMLElement).addEventListener('click', () => {
        ps.reduce((pr: Promise<void>, p: IPartidoG) => pr.then(() => guardarGest(ctx, p, { EstadoG: 'Validado' }, 'Validado')), Promise.resolve()).then(() => { cerrarModal(); aviso(ps.length + ' propuestas validadas'); ctx.recargar(); }).catch((x: Error) => aviso('Error: ' + x.message, true));
      });
    }
  }));
  root.querySelectorAll('[data-sol]').forEach((b: Element) => b.addEventListener('click', (ev: Event) => {
    ev.stopPropagation(); const p = pid(b.getAttribute('data-sol') || ''); if (p) { modalPropuesta(ctx, p, al); }
  }));
}

/** Justificación por defecto a partir de la alerta */
export function justificacion(p: IPartidoG, al: IAlerta[]): string {
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

/* ---------- Petición de hora del entrenador ---------- */
function modalPeticion(ctx: IGCtx, p: IPartidoG): void {
  const g = ctx.gest[p.id]; const f = g && g.petHora ? g.petHora : p.fecha;
  formModal('Pedir otra hora · ' + p.equipo, [
    { n: 'dia', label: 'Día', tipo: 'date', valor: fechaInput(f), req: true }, { n: 'hora', label: 'Hora', tipo: 'text', valor: hm(minDia(f)), ph: '12:00', req: true },
    { n: 'nota', label: 'Motivo', tipo: 'area', valor: g ? g.petNota : '', ph: 'Ej.: el entrenador trabaja hasta las 13:00 ese sábado' }
  ], 'Guardar petición', (v: { [k: string]: any }) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    const h = /^\d{1,2}:\d{2}$/.test(v.hora) ? ('0' + v.hora).slice(-5) : '';
    if (!h) { return Promise.reject(new Error('Hora no válida (usa 12:00)')); }
    const fecha = new Date(v.dia + 'T' + h + ':00');
    return guardarGest(ctx, p, { PeticionHora: fecha.toISOString(), PeticionNota: v.nota, PeticionDe: nombreCorto(ctx.usuario) }, 'Petición del entrenador: ' + fTxt(fecha)).then(() => { aviso('Petición guardada: coordinación la tendrá en cuenta al resolver la semana'); ctx.recargar(); });
  }, 520, g && g.petHora ? '<button type="button" class="cbp-btn full" data-del-pet="1" style="justify-self:start">Retirar petición</button>' : '');
  const del = document.querySelector('[data-del-pet]') as HTMLElement;
  if (del) { del.addEventListener('click', () => { guardarGest(ctx, p, { PeticionHora: null, PeticionNota: '', PeticionDe: '' }, 'Petición retirada').then(() => { cerrarModal(); aviso('Petición retirada'); ctx.recargar(); }); }); }
}

/* ---------- Mensaje al club rival ---------- */
function textoMensaje(ctx: IGCtx, p: IPartidoG, club?: IClub): string {
  const g = ctx.gest[p.id]; const nueva = g && g.fechaProp ? g.fechaProp : p.fecha;
  const saludo = 'Hola' + (club && club.contacto ? ' ' + club.contacto.split(' ')[0] : '') + ', os escribimos desde la coordinación deportiva del Gijón CBP (C.B. Pumarín).';
  const partido = 'Partido ' + (p.catPDM || p.equipo) + (p.jornada ? ', jornada ' + p.jornada : '') + ': ' + (p.local ? 'Gijón CBP – ' + p.rival : p.rival + ' – Gijón CBP') + ', fijado el ' + fTxt(p.oficial) + (p.local ? '' : ' en ' + p.lugar) + '.';
  const plazoOk = p.local && enPlazo({ ...p, fecha: p.oficial }, ctx.plazo);
  const pide = plazoOk ? 'Os informamos de que, dentro del plazo ordinario, lo hemos fijado el ' + fTxt(nueva) + ' en ' + (g && g.pista ? (ctx.pistas.filter((k: IPista) => k.clave === g.pista)[0] || { inst: p.lugar } as IPista).inst : p.lugar) + '. Lo cargamos en la intranet de la FBPA.'
    : p.local
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
    '<p class="cbp-nota" style="margin:0">' + (p.local && enPlazo({ ...p, fecha: p.oficial }, ctx.plazo) ? 'Está dentro del plazo ordinario: es solo un aviso, no hace falta su conformidad. Recuerda marcarlo como cargado en la FBPA.' : 'Al enviarlo, el partido pasa a «Pedido al rival». Cuando contesten, registra su respuesta.') + '</p>';
  const body = abrirModal((p.local ? 'Comunicar a ' : 'Pedir a ') + (club ? club.nombre : p.rival), html, 620);
  const marcar = (): void => { if (p.local && enPlazo({ ...p, fecha: p.oficial }, ctx.plazo)) { return; } if (ctx.gest[p.id] && ctx.gest[p.id].estado === 'Pedido al rival') { return; } guardarGest(ctx, p, { EstadoG: 'Pedido al rival' }, 'Petición enviada a ' + (club ? club.nombre : p.rival)).then(() => ctx.recargar()); };
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

/* ---------- Para cargar en la FBPA ---------- */
function listaFbpa(ctx: IGCtx): void {
  const ps = ctx.semana.filter((p: IPartidoG) => { const g = ctx.gest[p.id]; return g && (g.estado === 'Validado' || g.estado === 'Aceptado' || g.estado === 'Pedido al rival' || g.estado === 'Cargado en FBPA') && g.fechaProp; })
    .sort((a: IPartidoG, b: IPartidoG) => a.fecha.getTime() - b.fecha.getTime());
  const dd = (d: Date): string => ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear();
  const copia = (v: string): string => '<span class="cbp-cp"><code>' + esc(v) + '</code><button type="button" class="cbp-cp-b" data-cp="' + esc(v) + '" title="Copiar" aria-label="Copiar ' + esc(v) + '">⧉</button></span>';
  const campo = (p: IPartidoG, g: IGest): string => { const k = g.pista ? ctx.pistas.filter((x: IPista) => x.clave === g.pista)[0] : undefined; return k ? k.alias || k.inst : p.lugarFbpa || p.lugar; };
  const html = ps.length ? '<p class="cbp-nota" style="margin:0">Para la «Solicitud de cambio de partido» de la intranet: fecha nueva, hora nueva y campo nuevo, cada uno con su botón de copiar. La sincronización nocturna comprueba si la FBPA ya lo publica; cuando lo hace, el partido queda con el nuevo horario oficial y sus alertas se recalculan solas.</p><div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Partido</th><th>Original</th><th>Fecha nueva</th><th>Hora nueva</th><th>Campo nuevo</th><th>Estado</th></tr></thead><tbody>' +
    ps.map((p: IPartidoG) => { const g = ctx.gest[p.id]; const f = g.fechaProp as Date; const ya = Math.abs(f.getTime() - p.oficial.getTime()) < 60000;
      return '<tr><td><b>' + esc(p.equipo) + '</b><br><small>' + (p.local ? 'vs ' : 'en ') + esc(p.rival) + (p.jornada ? ' · J' + esc(p.jornada) : '') + '</small></td><td><small>' + dd(p.oficial) + ' ' + hm(minDia(p.oficial)) + '</small></td>' +
        '<td>' + copia(dd(f)) + '</td><td>' + copia(hm(minDia(f))) + '</td><td>' + (p.local ? copia(campo(p, g)) : '<small>' + esc(p.lugar) + '</small>') + '</td>' +
        '<td>' + (ya ? '<span class="cbp-fpill ok">Ya en FBPA</span>' : pillEstado(g.estado)) + (g.estado === 'Pedido al rival' ? '<br><small>esperando al rival</small>' : '') + '</td></tr>'; }).join('') + '</tbody></table></div>'
    : '<p class="cbp-nota">No hay cambios validados pendientes de cargar esta semana.</p>';
  const body = abrirModal('Para cargar en la FBPA', html, 860);
  body.querySelectorAll('[data-cp]').forEach((b: Element) => b.addEventListener('click', () => {
    const v = b.getAttribute('data-cp') || '';
    navigator.clipboard.writeText(v).then(() => { b.textContent = '✓'; b.classList.add('ok'); window.setTimeout(() => { b.textContent = '⧉'; b.classList.remove('ok'); }, 1500); }, () => aviso('No se pudo copiar', true));
  }));
}

/* ---------- Resumen para los entrenadores afectados ---------- */
function resumenEntrenadores(ctx: IGCtx): void {
  const estTxt: { [e: string]: string } = { 'Propuesto': 'propuesta, falta tu OK', 'Validado': 'validado por coordinación', 'Pedido al rival': 'pendiente de que acepte el rival', 'Contrapropuesta': 'el rival propone otra hora', 'Aceptado': 'aceptado por el rival', 'Cargado en FBPA': 'cargado en la FBPA' };
  const ps = ctx.semana.filter((p: IPartidoG) => { const g = ctx.gest[p.id]; return g && estTxt[g.estado] && g.fechaProp; }).sort((a: IPartidoG, b: IPartidoG) => a.equipo.localeCompare(b.equipo) || a.fecha.getTime() - b.fecha.getTime());
  if (!ps.length) { aviso('No hay cambios de horario esta semana'); return; }
  const lineas = ps.map((p: IPartidoG) => { const g = ctx.gest[p.id]; const f = g.fechaProp as Date; const k = g.pista ? ctx.pistas.filter((x: IPista) => x.clave === g.pista)[0] : undefined;
    return '• ' + p.equipo + ' (' + (p.local ? 'vs ' : 'en ') + p.rival + '): ' + fTxt(p.oficial) + ' → ' + fTxt(f) + (k ? ' en ' + k.inst + (k.modulo ? ' ' + k.modulo : '') : '') + ' · ' + estTxt[g.estado] + '.'; });
  const lun = new Date(ps[0].fecha); lun.setDate(lun.getDate() - ((lun.getDay() + 6) % 7));
  const txt = '🏀 Cambios de horario · semana del ' + lun.getDate() + ' de ' + MESES[lun.getMonth()].toLowerCase() + '\n\n' + lineas.join('\n') + '\n\nSi algo no os encaja, decídnoslo o pedid otra hora desde el portal (Partidos).\nCoordinación deportiva · Gijón CBP';
  const body = abrirModal('Resumen para entrenadores', '<p class="cbp-nota" style="margin:0">Cambios de esta semana con su estado, listos para el grupo de entrenadores.</p><label class="cbp-field">Mensaje<textarea name="txt" rows="12">' + esc(txt) + '</textarea></label>' +
    '<div class="cbp-btns" style="justify-content:flex-start"><button type="button" class="cbp-btn pri" data-copiar="1" style="color:#FF9F00">Copiar</button><a class="cbp-btn" data-wa="1" target="_blank" rel="noopener" href="https://wa.me/?text=' + encodeURIComponent(txt) + '">Abrir en WhatsApp</a></div>', 640);
  const ta = body.querySelector('[name="txt"]') as HTMLTextAreaElement;
  (body.querySelector('[data-copiar]') as HTMLElement).addEventListener('click', () => { navigator.clipboard.writeText(ta.value).then(() => aviso('Resumen copiado')); });
  (body.querySelector('[data-wa]') as HTMLAnchorElement).addEventListener('click', (e: Event) => { (e.currentTarget as HTMLAnchorElement).href = 'https://wa.me/?text=' + encodeURIComponent(ta.value); });
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
