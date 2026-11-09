import { SPHttpClient } from '@microsoft/sp-http';
import { esc } from './cbp';
import { spPost, webPath, aviso } from './ui';
import { porTitulo, getJ, subirArchivo, formModal, fCorta, relativo, diasHasta, iniciales, isoDia } from './sp2';
import { COLORES, fondoColor } from './estilos4';
import { enlazarUrls } from './estilos3';

const L_COM = 'Comisiones';
const L_TAB = 'Comisiones · Tablón';
const ESTADOS = ['Pendiente', 'En curso', 'Hecho'];
const COLABORAR = 1073741827;
const LEER = 1073741826;

interface IMiembro { id: number; nombre: string; email: string; login: string; }
interface ICom { id: number; nombre: string; objetivo: string; estado: string; color: string; grupo: number; coord: string; creada: Date; miembros: IMiembro[]; abiertas: number; }
interface IEntrada { id: number; titulo: string; cuerpo: string; tipo: string; estado: string; resp: string; fecha?: Date; autor: string; creada: Date; }
interface IDoc { nombre: string; url: string; fecha: Date; autor: string; }

export interface IComOpts { client: SPHttpClient; w: string; embebido?: boolean; }

function carpetaTab(w: string, id: number): string { return webPath(w) + '/Lists/ComisionesTablon/C' + id; }
function carpetaDoc(w: string, id: number): string { return webPath(w) + '/ComisionesDocumentos/C' + id; }
function api(w: string, rel: string): string { return w + '/_api/' + rel; }
function limpiaNombre(n: string): string { return n.replace(/\s*-\s*Gij[oó]n CBP.*$/i, ''); }

/** Punto de entrada: cuadrícula de comisiones o espacio de una comisión (#comision=ID) */
export function renderComisiones(el: HTMLElement, o: IComOpts): void {
  const hm = window.location.hash.match(/comision=(\d+)/);
  el.innerHTML = '<div class="cbp-cargando">Cargando comisiones…</div>';
  Promise.all([cargarComisiones(o), esGestor(o)]).then((r: [ICom[], boolean]) => {
    const sel = hm ? r[0].filter((c: ICom) => c.id === +hm[1])[0] : undefined;
    if (sel) { espacio(el, o, sel, r[1]); } else { cuadricula(el, o, r[0], r[1]); }
  }).catch((x: Error) => { el.innerHTML = '<div class="cbp-vacio">No se pudieron cargar las comisiones (' + esc(x.message) + ').</div>'; });
}

function esGestor(o: IComOpts): Promise<boolean> {
  return getJ(o.client, porTitulo(o.w, L_COM) + '/EffectiveBasePermissions')
    .then((j: any) => (parseInt(j.Low, 10) & 0x2000000) !== 0).catch(() => false); // eslint-disable-line @typescript-eslint/no-explicit-any
}

function cargarComisiones(o: IComOpts): Promise<ICom[]> {
  return getJ(o.client, porTitulo(o.w, L_COM) + '/items?$select=Id,Title,Objetivo,EstadoCom,ColorCom,GrupoId,Coordinacion,Created&$orderby=Created desc&$top=200')
    .then((j: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
      const coms: ICom[] = (j.value || []).map((i: any): ICom => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
        id: i.Id, nombre: i.Title || '', objetivo: i.Objetivo || '', estado: i.EstadoCom || 'Activa', color: i.ColorCom || '#090909',
        grupo: i.GrupoId || 0, coord: i.Coordinacion || '', creada: new Date(i.Created), miembros: [], abiertas: 0
      }));
      const tareas = getJ(o.client, porTitulo(o.w, L_TAB) + "/items?$select=ComisionId,EstadoE&$filter=TipoEntrada eq 'Tarea'&$top=2000")
        .then((t: any) => { (t.value || []).forEach((x: any) => { const c = coms.filter((k: ICom) => k.id === x.ComisionId)[0]; if (c && x.EstadoE !== 'Hecho') { c.abiertas++; } }); }) // eslint-disable-line @typescript-eslint/no-explicit-any
        .catch(() => undefined);
      return Promise.all([tareas].concat(coms.map((c: ICom) => cargarMiembros(o, c)))).then(() => coms);
    });
}

function cargarMiembros(o: IComOpts, c: ICom): Promise<void> {
  if (!c.grupo) { return Promise.resolve(); }
  return getJ(o.client, api(o.w, 'web/sitegroups(' + c.grupo + ')/users?$select=Id,Title,Email,LoginName'))
    .then((j: any) => { c.miembros = (j.value || []).filter((u: any) => !/SHAREPOINT\\system|app@sharepoint/i.test(u.LoginName)) // eslint-disable-line @typescript-eslint/no-explicit-any
      .map((u: any): IMiembro => ({ id: u.Id, nombre: limpiaNombre(u.Title || u.Email), email: u.Email || '', login: u.LoginName })); }) // eslint-disable-line @typescript-eslint/no-explicit-any
    .catch(() => undefined);
}

function pila(ms: IMiembro[]): string {
  return '<span class="cbp-pila">' + ms.slice(0, 5).map((m: IMiembro) => '<span class="cbp-avatar" title="' + esc(m.nombre) + '">' + esc(iniciales(m.nombre)) + '</span>').join('') +
    (ms.length > 5 ? '<span class="cbp-avatar" style="background:#E6E3DC;color:#090909">+' + (ms.length - 5) + '</span>' : '') + '</span>';
}

/* ---------------- Cuadrícula ---------------- */
function cuadricula(el: HTMLElement, o: IComOpts, coms: ICom[], gestor: boolean): void {
  const activas = coms.filter((c: ICom) => c.estado !== 'Cerrada');
  const cerradas = coms.filter((c: ICom) => c.estado === 'Cerrada');
  const tarjeta = (c: ICom): string => '<button type="button" class="cbp-com" data-com="' + c.id + '">' +
    '<div class="cbp-com-h" style="' + fondoColor(c.color) + '"><small>' + esc(c.estado === 'Activa' ? 'Comisión de trabajo' : c.estado) + '</small><h3>' + esc(c.nombre) + '</h3></div>' +
    '<div class="cbp-com-b">' + (c.objetivo ? '<p>' + esc(c.objetivo) + '</p>' : '<p style="color:#5c5c5c">Sin objetivo descrito.</p>') +
    '<div class="cbp-com-pie">' + (c.miembros.length ? pila(c.miembros) : '<span>Sin personas invitadas</span>') +
    '<span>' + (c.abiertas ? c.abiertas + ' tarea' + (c.abiertas === 1 ? '' : 's') + ' abierta' + (c.abiertas === 1 ? '' : 's') : 'Al día') + '</span></div></div></button>';
  const nueva = gestor ? '<button type="button" class="cbp-com cbp-com-nueva" data-nueva="1"><b>+</b>Crear comisión</button>' : '';
  el.innerHTML = (o.embebido ? '' : '<div class="cbp-sec-head"><h2 class="cbp-h2">Comisiones de trabajo</h2>' + (gestor ? '<button type="button" class="cbp-btn pri" data-nueva="1">+ Nueva comisión</button>' : '') + '</div>') +
    '<p class="cbp-nota" style="margin:0 0 4px">' + (gestor
      ? 'La junta crea las comisiones e invita a quien haga falta, aunque no sea de la junta. Cada persona solo ve las comisiones en las que participa.'
      : 'Estas son las comisiones en las que participas.') + '</p>' +
    (activas.length || gestor ? '<div class="cbp-coms">' + activas.map(tarjeta).join('') + nueva + '</div>' : '<div class="cbp-vacio">Todavía no participas en ninguna comisión.</div>') +
    (cerradas.length ? '<h3 style="margin:18px 0 0;font-size:14px;letter-spacing:.08em;text-transform:uppercase;color:#5c5c5c">Cerradas</h3><div class="cbp-coms" style="opacity:.7">' + cerradas.map(tarjeta).join('') + '</div>' : '');
  el.querySelectorAll('[data-com]').forEach((b: Element) => b.addEventListener('click', () => {
    history.pushState(null, '', '#comision=' + b.getAttribute('data-com')); renderComisiones(el, o); el.scrollIntoView({ block: 'start' });
  }));
  el.querySelectorAll('[data-nueva]').forEach((b: Element) => b.addEventListener('click', () => crearComision(el, o)));
}

/* ---------------- Alta de comisión ---------------- */
function selectorColor(sel: string): string {
  return '<div class="cbp-field full">Color<div class="cbp-colores">' + COLORES.map((c: string, i: number) =>
    '<label style="background:' + c + '" title="Color ' + (i + 1) + '"><input type="radio" name="color" value="' + c + '"' + (c === sel ? ' checked' : '') + '></label>').join('') + '</div></div>';
}

function crearComision(el: HTMLElement, o: IComOpts): void {
  const body = formModal('Nueva comisión', [
    { n: 'nombre', label: 'Nombre', req: true, full: true, ph: 'Ej.: Torneo de Navidad, Patrocinios, Eventos 50 aniversario' },
    { n: 'objetivo', label: 'Objetivo', tipo: 'area', ph: 'Para qué se crea y qué tiene que entregar' },
    { n: 'coord', label: 'Coordinación', full: true, ph: 'Persona que la coordina' },
    { n: 'emails', label: 'Invitar personas', tipo: 'area', ayuda: '(correos separados por comas o líneas; puedes hacerlo después)', ph: 'nombre@cbpumarin.com, otra@gmail.com' }
  ], 'Crear comisión', (v: { [k: string]: any }) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    const color = ((body.querySelector('input[name="color"]:checked') as HTMLInputElement) || { value: '#090909' }).value;
    return nuevaComision(o, v.nombre, v.objetivo, v.coord, color).then((id: number) => {
      const emails = String(v.emails || '').split(/[\s,;]+/).filter((x: string) => /@/.test(x));
      return (emails.length ? invitar(o, id, emails) : Promise.resolve('')).then((msg: string) => {
        aviso('Comisión creada' + (msg ? ' · ' + msg : ''));
        history.pushState(null, '', '#comision=' + id); renderComisiones(el, o);
      });
    });
  }, 620, selectorColor('#090909'));
}

function nuevaComision(o: IComOpts, nombre: string, objetivo: string, coord: string, color: string): Promise<number> {
  let gid = 0; let iid = 0;
  return getJ(o.client, api(o.w, "web/sitegroups/getbyname('Junta Directiva')?$select=Id")).then((g: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    const junta = g.Id;
    return spPost(o.client, api(o.w, 'web/sitegroups'), { Title: ('Comisión · ' + nombre).slice(0, 250), Description: (objetivo || '').slice(0, 500), OnlyAllowMembersViewMembership: false, AllowMembersEditMembership: false })
      .then((gr: any) => { gid = gr.Id; return propietarioGrupo(o, gid, junta).catch(() => undefined); }); // eslint-disable-line @typescript-eslint/no-explicit-any
  }).then(() => spPost(o.client, porTitulo(o.w, L_COM) + '/items', { Title: nombre, Objetivo: objetivo, Coordinacion: coord, ColorCom: color, EstadoCom: 'Activa', GrupoId: gid }))
    .then((it: any) => { iid = it.Id; // eslint-disable-line @typescript-eslint/no-explicit-any
      return spPost(o.client, porTitulo(o.w, L_COM) + '/items(' + iid + ')/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)'); })
    .then(() => spPost(o.client, porTitulo(o.w, L_COM) + '/items(' + iid + ')/roleassignments/addroleassignment(principalid=' + gid + ',roledefid=' + LEER + ')'))
    .then(() => carpeta(o, L_TAB, webPath(o.w) + '/Lists/ComisionesTablon', 'C' + iid, gid))
    .then(() => carpeta(o, 'Comisiones · Documentos', webPath(o.w) + '/ComisionesDocumentos', 'C' + iid, gid))
    .then(() => iid);
}

/** Crea una carpeta como elemento de la lista y le da permisos propios para el grupo de la comisión */
function carpeta(o: IComOpts, lista: string, raiz: string, nombre: string, gid: number): Promise<void> {
  return spPost(o.client, porTitulo(o.w, lista) + '/AddValidateUpdateItemUsingPath', {
    listItemCreateInfo: { FolderPath: { DecodedUrl: raiz }, UnderlyingObjectType: 1, LeafName: { DecodedUrl: nombre } }, formValues: [{ FieldName: 'Title', FieldValue: nombre }], bNewDocumentUpdate: false
  }).then((r: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    const v = (r && r.value) || []; const e = v.filter((x: any) => x.HasException)[0]; // eslint-disable-line @typescript-eslint/no-explicit-any
    if (e) { throw new Error(e.ErrorMessage); }
    const id = v.filter((x: any) => x.FieldName === 'Id')[0].FieldValue; // eslint-disable-line @typescript-eslint/no-explicit-any
    return spPost(o.client, porTitulo(o.w, lista) + '/items(' + id + ')/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)')
      .then(() => spPost(o.client, porTitulo(o.w, lista) + '/items(' + id + ')/roleassignments/addroleassignment(principalid=' + gid + ',roledefid=' + COLABORAR + ')'));
  }).then(() => undefined);
}

/** La junta pasa a ser propietaria del grupo (así cualquier miembro de la junta puede invitar o quitar personas) */
function propietarioGrupo(o: IComOpts, gid: number, owner: number): Promise<void> {
  const xml = '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="GijonCBP" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">' +
    '<Actions><SetProperty Id="10" ObjectPathId="5" Name="Owner"><Parameter ObjectPathId="7" /></SetProperty><Method Name="Update" Id="11" ObjectPathId="5" /></Actions>' +
    '<ObjectPaths><StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="2" ParentId="1" Name="Web" /><Property Id="3" ParentId="2" Name="SiteGroups" />' +
    '<Method Id="5" ParentId="3" Name="GetById"><Parameters><Parameter Type="Number">' + gid + '</Parameter></Parameters></Method>' +
    '<Method Id="7" ParentId="3" Name="GetById"><Parameters><Parameter Type="Number">' + owner + '</Parameter></Parameters></Method></ObjectPaths></Request>';
  return o.client.post(o.w + '/_vti_bin/client.svc/ProcessQuery', SPHttpClient.configurations.v1, { body: xml, headers: { 'Content-Type': 'text/xml' } })
    .then((r: any) => r.text()).then((t: string) => { if (/"ErrorInfo":\{/.test(t)) { throw new Error('No se pudo asignar la junta como propietaria'); } }); // eslint-disable-line @typescript-eslint/no-explicit-any
}

/* ---------------- Invitaciones ---------------- */
function invitar(o: IComOpts, comId: number, emails: string[]): Promise<string> {
  return getJ(o.client, porTitulo(o.w, L_COM) + '/items(' + comId + ')?$select=Title,GrupoId').then((c: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    return getJ(o.client, api(o.w, "web/sitegroups/getbyname('Participantes en comisiones')?$select=Id")).catch(() => ({ Id: 0 })).then((p: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
      const pagina = window.location.origin + webPath(o.w) + '/SitePages/Comisiones.aspx#comision=' + comId;
      let ok = 0; const fuera: string[] = [];
      const uno = (email: string): Promise<void> => spPost(o.client, api(o.w, 'web/ensureuser'), { logonName: email })
        .then((u: any) => Promise.all([ // eslint-disable-line @typescript-eslint/no-explicit-any
          spPost(o.client, api(o.w, 'web/sitegroups(' + c.GrupoId + ')/users'), { LoginName: u.LoginName }),
          p.Id ? spPost(o.client, api(o.w, 'web/sitegroups(' + p.Id + ')/users'), { LoginName: u.LoginName }).catch(() => undefined) : Promise.resolve()
        ]).then(() => {
          ok++;
          return spPost(o.client, api(o.w, 'SP.Utilities.Utility.SendEmail'), { properties: { To: [u.Email || email], Subject: 'Te han invitado a la comisión «' + c.Title + '» · Gijón CBP',
            Body: '<p>Hola:</p><p>La junta directiva del Gijón CBP te ha invitado a participar en la comisión <b>' + esc(c.Title) + '</b>.</p><p>Entra en el espacio de trabajo aquí: <a href="' + pagina + '">' + pagina + '</a></p><p>Nacimos en Pumarín. Crecemos con Gijón.</p>' } }).catch(() => undefined);
        }))
        .catch(() => compartirExterno(o, pagina, email, c.GrupoId, c.Title).then((b: boolean) => { if (b) { ok++; } else { fuera.push(email); } }));
      return emails.reduce((pr: Promise<void>, e: string) => pr.then(() => uno(e)), Promise.resolve()).then(() =>
        ok + ' invitada' + (ok === 1 ? '' : 's') + (fuera.length ? '. No se pudo invitar a: ' + fuera.join(', ') + ' (pide al administrador que la dé de alta como invitada)' : ''));
    });
  });
}

/** Personas sin cuenta del club: se comparte la página con el grupo de la comisión (invitación por correo de Microsoft) */
function compartirExterno(o: IComOpts, pagina: string, email: string, gid: number, nombre: string): Promise<boolean> {
  const ppl = JSON.stringify([{ Key: email, DisplayText: email, IsResolved: true, Description: email, EntityType: '', EntityData: { Email: email }, ProviderName: '', ProviderDisplayName: '' }]);
  return spPost(o.client, api(o.w, 'SP.Web.ShareObject'), {
    url: pagina.split('#')[0], peoplePickerInput: ppl, roleValue: 'group:' + gid, groupId: 0, propagateAcl: false, sendEmail: true,
    includeAnonymousLinkInEmail: false, emailSubject: 'Invitación a la comisión «' + nombre + '» · Gijón CBP',
    emailBody: 'La junta directiva del Gijón CBP te invita a participar en la comisión «' + nombre + '».', useSimplifiedRoles: true
  }).then((r: any) => !!r && (r.StatusCode === 0 || r.StatusCode === undefined)).catch(() => false); // eslint-disable-line @typescript-eslint/no-explicit-any
}

/* ---------------- Espacio de la comisión ---------------- */
let subTab = 'tablon';

function espacio(el: HTMLElement, o: IComOpts, c: ICom, gestor: boolean): void {
  el.innerHTML = '<div class="cbp-cargando">Abriendo ' + esc(c.nombre) + '…</div>';
  Promise.all([
    getJ(o.client, porTitulo(o.w, L_TAB) + '/items?$select=Id,Title,Cuerpo,TipoEntrada,EstadoE,ResponsableE,FechaE,Created,Author/Title&$expand=Author&$filter=ComisionId eq ' + c.id + '&$orderby=Created desc&$top=500').catch(() => ({ value: [] })),
    getJ(o.client, api(o.w, "web/GetFolderByServerRelativeUrl('" + encodeURIComponent(carpetaDoc(o.w, c.id)) + "')/Files?$select=Name,ServerRelativeUrl,TimeLastModified&$expand=ModifiedBy&$orderby=TimeLastModified desc")).catch(() => ({ value: [] }))
  ]).then((r: any[]) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    const ents: IEntrada[] = (r[0].value || []).map((i: any): IEntrada => ({ id: i.Id, titulo: i.Title || '', cuerpo: i.Cuerpo || '', tipo: i.TipoEntrada || 'Nota', estado: i.EstadoE || 'Pendiente', // eslint-disable-line @typescript-eslint/no-explicit-any
      resp: i.ResponsableE || '', fecha: i.FechaE ? new Date(i.FechaE) : undefined, autor: limpiaNombre(i.Author ? i.Author.Title : ''), creada: new Date(i.Created) }));
    const docs: IDoc[] = (r[1].value || []).map((f: any): IDoc => ({ nombre: f.Name, url: f.ServerRelativeUrl, fecha: new Date(f.TimeLastModified), autor: limpiaNombre(f.ModifiedBy ? f.ModifiedBy.Title : '') })); // eslint-disable-line @typescript-eslint/no-explicit-any
    pintarEspacio(el, o, c, gestor, ents, docs);
  });
}

function pintarEspacio(el: HTMLElement, o: IComOpts, c: ICom, gestor: boolean, ents: IEntrada[], docs: IDoc[]): void {
  const recargar = (): void => espacio(el, o, c, gestor);
  const tareas = ents.filter((e: IEntrada) => e.tipo === 'Tarea');
  const posts = ents.filter((e: IEntrada) => e.tipo !== 'Tarea');
  const hero = '<section class="cbp-chero" style="' + fondoColor(c.color) + '"><div><small>Comisión de trabajo · ' + esc(c.estado) + '</small><h2>' + esc(c.nombre) + '</h2>' +
    (c.objetivo ? '<p>' + esc(c.objetivo) + '</p>' : '') + (c.coord ? '<p style="margin-top:8px;font-weight:700">Coordina: ' + esc(c.coord) + '</p>' : '') + '</div>' +
    '<div class="cbp-btns">' + (gestor ? '<button type="button" class="cbp-btn" data-inv="1">+ Invitar personas</button><button type="button" class="cbp-btn" data-edc="1">Editar</button>' : '') + '</div></section>';

  const post = (e: IEntrada): string => '<article class="cbp-post"><span class="cbp-avatar">' + esc(iniciales(e.autor || '?')) + '</span><div>' +
    '<div class="cbp-post-h"><b>' + esc(e.autor) + '</b><span>' + relativo(e.creada) + '</span><span class="cbp-fpill' + (e.tipo === 'Acta' ? ' amb' : e.tipo === 'Idea' ? ' ok' : '') + '">' + esc(e.tipo) + '</span></div>' +
    (e.titulo ? '<h4>' + esc(e.titulo) + '</h4>' : '') + (e.cuerpo ? '<p>' + enlazarUrls(esc(e.cuerpo)) + '</p>' : '') +
    '<div class="cbp-post-acc"><button type="button" data-borrar="' + e.id + '">Eliminar</button></div></div></article>';

  const composer = '<form class="cbp-compose" data-comp="1"><input name="t" placeholder="Título (opcional)" style="font:inherit;font-size:15px;font-weight:700;border:0;outline:none;padding:4px">' +
    '<textarea name="c" placeholder="Escribe una nota, un acta o una idea para la comisión…"></textarea>' +
    '<div class="cbp-compose-b"><div class="cbp-tipo">' + ['Nota', 'Acta', 'Idea'].map((t: string, i: number) => '<label><input type="radio" name="tipo" value="' + t + '"' + (i === 0 ? ' checked' : '') + '>' + t + '</label>').join('') + '</div>' +
    '<span style="display:flex;gap:8px"><label class="cbp-btn sm" style="cursor:pointer">Adjuntar<input type="file" data-adj="1" multiple hidden></label><button type="submit" class="cbp-btn sm pri">Publicar</button></span></div></form>';

  const kcard = (e: IEntrada): string => {
    const i = ESTADOS.indexOf(e.estado); const d = e.fecha ? diasHasta(e.fecha) : 0;
    return '<div class="cbp-kcard"><h4 data-edt="' + e.id + '">' + esc(e.titulo) + '</h4>' +
      '<div class="det">' + (e.resp ? '<span class="cbp-avatar" style="width:24px;height:24px;font-size:10px">' + esc(iniciales(e.resp)) + '</span>' + esc(e.resp) : '<span>Sin responsable</span>') +
      (e.fecha ? '<span class="cbp-fpill' + (e.estado !== 'Hecho' && d < 0 ? ' mal' : d <= 3 && e.estado !== 'Hecho' ? ' amb' : '') + '">' + fCorta(e.fecha) + '</span>' : '') + '</div>' +
      '<div class="cbp-kmov"><button type="button" data-mov="' + e.id + '|-1"' + (i <= 0 ? ' disabled style="visibility:hidden"' : '') + '>← ' + (ESTADOS[i - 1] || '') + '</button>' +
      '<button type="button" data-mov="' + e.id + '|1"' + (i >= 2 ? ' disabled style="visibility:hidden"' : '') + '>' + (ESTADOS[i + 1] || '') + ' →</button></div></div>';
  };
  const kanban = '<div class="cbp-box-h"><span class="cbp-nota">' + tareas.filter((t: IEntrada) => t.estado !== 'Hecho').length + ' abiertas</span><button type="button" class="cbp-btn sm pri" data-ntarea="1">+ Nueva tarea</button></div>' +
    '<div class="cbp-kanban">' + ESTADOS.map((s: string) => { const xs = tareas.filter((t: IEntrada) => t.estado === s); return '<div class="cbp-kcol"><div class="cbp-kcol-h">' + s + '<span>' + xs.length + '</span></div>' + xs.map(kcard).join('') + '</div>'; }).join('') + '</div>';

  const izq = '<div class="cbp-of-tabs" role="tablist">' +
    [['tablon', 'Tablón', posts.length], ['tareas', 'Tareas', tareas.filter((t: IEntrada) => t.estado !== 'Hecho').length]].map((t: any[]) => // eslint-disable-line @typescript-eslint/no-explicit-any
      '<button type="button" role="tab" class="cbp-of-tab" data-st="' + t[0] + '" aria-selected="' + (subTab === t[0]) + '">' + t[1] + (t[2] ? '<i>' + t[2] + '</i>' : '') + '</button>').join('') + '</div>' +
    (subTab === 'tareas' ? kanban : composer + '<div class="cbp-feed">' + (posts.length ? posts.map(post).join('') : '<div class="cbp-vacio">Aún no hay nada en el tablón. Empieza con el objetivo o la primera reunión.</div>') + '</div>');

  const miembros = '<section class="cbp-box"><div class="cbp-box-h"><h3>Personas</h3>' + (gestor ? '<button type="button" class="lnk" data-inv="1">+ Invitar</button>' : '<small>' + c.miembros.length + '</small>') + '</div>' +
    (c.miembros.length ? '<div class="cbp-miembros">' + c.miembros.map((m: IMiembro) => '<div class="cbp-miembro"><span class="cbp-avatar">' + esc(iniciales(m.nombre)) + '</span><span>' + esc(m.nombre) + '<small>' + esc(m.email) + '</small></span>' +
      (gestor ? '<button type="button" class="cbp-quitar" data-quitar="' + m.id + '" title="Quitar de la comisión" aria-label="Quitar a ' + esc(m.nombre) + '">×</button>' : '<span></span>') + '</div>').join('') + '</div>'
      : '<p class="cbp-nota">Solo la junta tiene acceso por ahora.' + (gestor ? ' Invita a las personas que vayan a trabajar en ella.' : '') + '</p>') +
    '<p class="cbp-nota" style="font-size:12px">La junta directiva tiene acceso a todas las comisiones.</p></section>';

  const ext = (n: string): string => (n.split('.').pop() || '').toLowerCase().slice(0, 4);
  const documentos = '<section class="cbp-box"><div class="cbp-box-h"><h3>Documentos</h3><label class="lnk" style="font-size:13px;font-weight:700;color:#B36F00;cursor:pointer">+ Subir<input type="file" data-subir="1" multiple hidden></label></div>' +
    (docs.length ? '<div class="cbp-docs">' + docs.slice(0, 12).map((d: IDoc) => '<a class="cbp-doc" href="' + esc(d.url) + '?web=1" target="_blank" rel="noopener"><span class="cbp-doc-ico ' + ext(d.nombre) + '">' + esc(ext(d.nombre)) + '</span><span><b>' + esc(d.nombre) + '</b><small>' + relativo(d.fecha) + (d.autor ? ' · ' + esc(d.autor) : '') + '</small></span><span></span></a>').join('') + '</div>'
      : '<p class="cbp-nota">Arrastra aquí o pulsa «Subir» para compartir archivos con la comisión.</p>') + '</section>';

  el.innerHTML = '<div><button type="button" class="cbp-volver" data-volver="1">← Todas las comisiones</button></div>' + hero +
    '<div class="cbp-of-grid"><div class="cbp-col">' + izq + '</div><div class="cbp-col">' + miembros + documentos + '</div></div>';

  const q = (s: string): NodeListOf<HTMLElement> => el.querySelectorAll(s) as NodeListOf<HTMLElement>;
  (el.querySelector('[data-volver]') as HTMLElement).addEventListener('click', () => { history.pushState(null, '', window.location.pathname + (o.embebido ? '#comisiones' : '')); renderComisiones(el, o); });
  q('[data-st]').forEach((b: HTMLElement) => b.addEventListener('click', () => { subTab = b.getAttribute('data-st') || 'tablon'; pintarEspacio(el, o, c, gestor, ents, docs); }));
  q('[data-inv]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
    formModal('Invitar a ' + c.nombre, [{ n: 'e', label: 'Correos electrónicos', tipo: 'area', req: true, ayuda: '(separados por comas o líneas)', ph: 'nombre@cbpumarin.com, persona@gmail.com' }], 'Invitar',
      (v: { [k: string]: any }) => invitar(o, c.id, String(v.e).split(/[\s,;]+/).filter((x: string) => /@/.test(x))).then((m: string) => { aviso(m); return cargarMiembros(o, c).then(() => pintarEspacio(el, o, c, gestor, ents, docs)); }), 520, // eslint-disable-line @typescript-eslint/no-explicit-any
      '<p class="cbp-nota full">Las personas del club entran directamente y reciben un correo con el enlace. Las de fuera reciben una invitación de Microsoft para acceder como invitadas.</p>');
  }));
  q('[data-quitar]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
    const m = c.miembros.filter((x: IMiembro) => x.id === +(b.getAttribute('data-quitar') || 0))[0];
    if (!m || !window.confirm('¿Quitar a ' + m.nombre + ' de la comisión?')) { return; }
    spPost(o.client, api(o.w, 'web/sitegroups(' + c.grupo + ')/users/removebyid(' + m.id + ')')).then(() => { c.miembros = c.miembros.filter((x: IMiembro) => x.id !== m.id); aviso(m.nombre + ' ya no está en la comisión'); pintarEspacio(el, o, c, gestor, ents, docs); })
      .catch((x: Error) => aviso('Error: ' + x.message, true));
  }));
  q('[data-edc]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
    const body = formModal('Editar comisión', [
      { n: 'nombre', label: 'Nombre', req: true, full: true, valor: c.nombre }, { n: 'objetivo', label: 'Objetivo', tipo: 'area', valor: c.objetivo },
      { n: 'coord', label: 'Coordinación', valor: c.coord }, { n: 'estado', label: 'Estado', tipo: 'select', opciones: ['Activa', 'En pausa', 'Cerrada'], valor: c.estado }
    ], 'Guardar', (v: { [k: string]: any }) => { // eslint-disable-line @typescript-eslint/no-explicit-any
      const color = ((body.querySelector('input[name="color"]:checked') as HTMLInputElement) || { value: c.color }).value;
      return spPost(o.client, porTitulo(o.w, L_COM) + '/items(' + c.id + ')', { Title: v.nombre, Objetivo: v.objetivo, Coordinacion: v.coord, EstadoCom: v.estado, ColorCom: color }, true)
        .then(() => { c.nombre = v.nombre; c.objetivo = v.objetivo; c.coord = v.coord; c.estado = v.estado; c.color = color; aviso('Comisión actualizada'); pintarEspacio(el, o, c, gestor, ents, docs); });
    }, 620, selectorColor(c.color));
  }));

  // Publicar en el tablón
  const comp = el.querySelector('[data-comp]') as HTMLFormElement;
  if (comp) {
    const adj = comp.querySelector('[data-adj]') as HTMLInputElement;
    adj.addEventListener('change', () => subir(o, c, adj.files, recargar));
    comp.addEventListener('submit', (ev: Event) => {
      ev.preventDefault();
      const t = (comp.querySelector('[name="t"]') as HTMLInputElement).value.trim(); const cu = (comp.querySelector('[name="c"]') as HTMLTextAreaElement).value.trim();
      const tipo = ((comp.querySelector('[name="tipo"]:checked') as HTMLInputElement) || { value: 'Nota' }).value;
      if (!t && !cu) { return; }
      crearEntrada(o, c.id, { Title: t || cu.split('\n')[0].slice(0, 120), Cuerpo: t ? cu : cu.split('\n').slice(1).join('\n').trim() || (cu.length > 120 ? cu : ''), TipoEntrada: tipo })
        .then(() => { aviso('Publicado en el tablón'); recargar(); }).catch((x: Error) => aviso('Error: ' + x.message, true));
    });
  }
  q('[data-borrar]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
    if (!window.confirm('¿Eliminar esta publicación?')) { return; }
    spPost(o.client, porTitulo(o.w, L_TAB) + '/items(' + b.getAttribute('data-borrar') + ')/recycle').then(recargar).catch((x: Error) => aviso('Error: ' + x.message, true));
  }));
  // Tareas
  const formTarea = (e?: IEntrada): void => {
    formModal(e ? 'Editar tarea' : 'Nueva tarea', [
      { n: 'Title', label: 'Tarea', req: true, full: true, valor: e ? e.titulo : '' }, { n: 'Cuerpo', label: 'Detalle', tipo: 'area', valor: e ? e.cuerpo : '' },
      { n: 'ResponsableE', label: 'Responsable', valor: e ? e.resp : '', ph: 'Nombre' }, { n: 'FechaE', label: 'Fecha límite', tipo: 'date', valor: e && e.fecha ? e.fecha.toISOString().slice(0, 10) : '' },
      { n: 'EstadoE', label: 'Estado', tipo: 'select', opciones: ESTADOS, valor: e ? e.estado : 'Pendiente' }
    ], e ? 'Guardar' : 'Crear tarea', (v: { [k: string]: any }) => { // eslint-disable-line @typescript-eslint/no-explicit-any
      const f = isoDia(v.FechaE);
      const pr = e ? spPost(o.client, porTitulo(o.w, L_TAB) + '/items(' + e.id + ')', { Title: v.Title, Cuerpo: v.Cuerpo, ResponsableE: v.ResponsableE, EstadoE: v.EstadoE, FechaE: f }, true)
        : crearEntrada(o, c.id, { Title: v.Title, Cuerpo: v.Cuerpo, ResponsableE: v.ResponsableE, EstadoE: v.EstadoE, TipoEntrada: 'Tarea' }, f);
      return pr.then(() => { subTab = 'tareas'; aviso(e ? 'Tarea guardada' : 'Tarea creada'); recargar(); });
    }, 560);
  };
  q('[data-ntarea]').forEach((b: HTMLElement) => b.addEventListener('click', () => formTarea()));
  q('[data-edt]').forEach((b: HTMLElement) => b.addEventListener('click', () => formTarea(tareas.filter((t: IEntrada) => t.id === +(b.getAttribute('data-edt') || 0))[0])));
  q('[data-mov]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
    const p = (b.getAttribute('data-mov') || '').split('|'); const t = tareas.filter((x: IEntrada) => x.id === +p[0])[0]; if (!t) { return; }
    const n = ESTADOS[ESTADOS.indexOf(t.estado) + +p[1]]; if (!n) { return; }
    spPost(o.client, porTitulo(o.w, L_TAB) + '/items(' + t.id + ')', { EstadoE: n }, true).then(() => { t.estado = n; pintarEspacio(el, o, c, gestor, ents, docs); }).catch((x: Error) => aviso('Error: ' + x.message, true));
  }));
  // Documentos
  const sb = el.querySelector('[data-subir]') as HTMLInputElement;
  if (sb) { sb.addEventListener('change', () => subir(o, c, sb.files, recargar)); }
  const caja = sb ? sb.closest('.cbp-box') as HTMLElement : null;
  if (caja) {
    caja.addEventListener('dragover', (ev: DragEvent) => { ev.preventDefault(); caja.style.outline = '2px dashed #FF9F00'; });
    caja.addEventListener('dragleave', () => { caja.style.outline = ''; });
    caja.addEventListener('drop', (ev: DragEvent) => { ev.preventDefault(); caja.style.outline = ''; if (ev.dataTransfer) { subir(o, c, ev.dataTransfer.files, recargar); } });
  }
}

function crearEntrada(o: IComOpts, comId: number, campos: { [k: string]: string }, fecha?: string | null): Promise<void> {
  const fv = Object.keys(campos).filter((k: string) => campos[k] !== undefined && campos[k] !== '').map((k: string) => ({ FieldName: k, FieldValue: String(campos[k]) }));
  fv.push({ FieldName: 'ComisionId', FieldValue: String(comId) });
  return spPost(o.client, porTitulo(o.w, L_TAB) + '/AddValidateUpdateItemUsingPath', {
    listItemCreateInfo: { FolderPath: { DecodedUrl: carpetaTab(o.w, comId) }, UnderlyingObjectType: 0 }, formValues: fv, bNewDocumentUpdate: false
  }).then((r: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
    const vals = (r && r.value) || [];
    const err = vals.filter((x: any) => x.HasException)[0]; // eslint-disable-line @typescript-eslint/no-explicit-any
    if (err) { throw new Error(err.ErrorMessage || 'No se pudo guardar'); }
    const id = vals.filter((x: any) => x.FieldName === 'Id')[0]; // eslint-disable-line @typescript-eslint/no-explicit-any
    if (fecha && id) { return spPost(o.client, porTitulo(o.w, L_TAB) + '/items(' + id.FieldValue + ')', { FechaE: fecha }, true).then(() => undefined); }
    return undefined;
  });
}

function subir(o: IComOpts, c: ICom, files: FileList | null, despues: () => void): void {
  if (!files || !files.length) { return; }
  const arr: File[] = []; for (let i = 0; i < files.length; i++) { arr.push(files[i]); }
  aviso('Subiendo ' + arr.length + ' archivo' + (arr.length === 1 ? '' : 's') + '…');
  arr.reduce((p: Promise<void>, f: File) => p.then(() => subirArchivo(o.client, o.w, carpetaDoc(o.w, c.id), f, { ComisionId: c.id })), Promise.resolve())
    .then(() => { aviso('Archivos subidos'); despues(); }).catch((x: Error) => aviso('Error al subir: ' + x.message, true));
}

