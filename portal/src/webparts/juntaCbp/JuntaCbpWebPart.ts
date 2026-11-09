import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { cargarEstilos, esc, MESES } from '../../common/cbp';
import { chips } from '../../common/render';
import { cargarEstilosUi, spPost, aviso } from '../../common/ui';
import { cargarEstilos3, enlazarUrls } from '../../common/estilos3';
import { cargarEstilos4 } from '../../common/estilos4';
import { porTitulo, getJ, subirArchivo, formModal, fCorta, fLarga, hhmm, euros, relativo, cuenta, diasHasta, iniciales, isoDia, isoDT, localDT } from '../../common/sp2';
import { renderComisiones } from '../../common/comisiones';

export interface IJuntaCbpWebPartProps { titulo: string; }

const L = { reu: 'Junta · Reuniones', tar: 'Junta · Acuerdos y tareas', pat: 'Junta · Patrocinios y convenios', doc: 'Junta · Documentos' };
const EST_T = ['Pendiente', 'En curso', 'Hecho'];
const EST_P = ['En negociación', 'Activo', 'Renovar', 'Finalizado'];
const TIPOS_P = ['Patrocinio', 'Colaboración', 'Aportación voluntaria', 'Convenio institucional'];
const TIPOS_D = ['Acta', 'Presupuesto', 'Contrato o convenio', 'Subvención', 'Estatutos y reglamentos', 'Otro'];

interface IReu { id: number; titulo: string; fecha?: Date; lugar: string; orden: string; estado: string; }
interface ITar { id: number; titulo: string; resp: string; fecha?: Date; estado: string; origen: string; }
interface IPat { id: number; titulo: string; tipo: string; importe: number; vigencia?: Date; estado: string; contacto: string; }
interface IDocJ { id: number; nombre: string; titulo: string; url: string; tipo: string; fecha?: Date; resumen: string; mod: Date; autor: string; }

type Tab = 'resumen' | 'reuniones' | 'tareas' | 'patrocinios' | 'documentos' | 'comisiones';
const TABS: [Tab, string][] = [['resumen', 'Resumen'], ['reuniones', 'Reuniones'], ['tareas', 'Acuerdos y tareas'], ['patrocinios', 'Patrocinios'], ['documentos', 'Documentos'], ['comisiones', 'Comisiones']];

const d = (v: any): Date | undefined => (v ? new Date(v) : undefined); // eslint-disable-line @typescript-eslint/no-explicit-any

export default class JuntaCbpWebPart extends BaseClientSideWebPart<IJuntaCbpWebPartProps> {
  private estado: 'cargando' | 'ok' | 'denegado' = 'cargando';
  private tab: Tab = 'resumen';
  private reus: IReu[] = []; private tars: ITar[] = []; private pats: IPat[] = []; private docs: IDocJ[] = [];
  private carpetaDocs: string = '';
  private nComisiones: number = 0;
  private filtroDoc: string = 'Todos';
  private verHechas: boolean = false;
  private onHash = (): void => { this.leerHash(); this.render(); };

  protected onInit(): Promise<void> {
    cargarEstilos(); cargarEstilosUi(); cargarEstilos3(); cargarEstilos4();
    this.leerHash();
    window.addEventListener('hashchange', this.onHash);
    return this.cargar();
  }
  protected onDispose(): void { window.removeEventListener('hashchange', this.onHash); }

  private leerHash(): void {
    const h = window.location.hash.replace('#', '');
    if (/^comision=/.test(h)) { this.tab = 'comisiones'; return; }
    const t = TABS.filter((x: [Tab, string]) => x[0] === h)[0];
    if (t) { this.tab = t[0]; }
  }

  private cargar(): Promise<void> {
    const c = this.context.spHttpClient; const w = this.context.pageContext.web.absoluteUrl;
    return getJ(c, porTitulo(w, L.reu) + '/items?$select=Id,Title,FechaReunion,LugarReunion,OrdenDia,EstadoReunion&$orderby=FechaReunion desc&$top=200')
      .then((r: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
        this.reus = (r.value || []).map((i: any): IReu => ({ id: i.Id, titulo: i.Title || '', fecha: d(i.FechaReunion), lugar: i.LugarReunion || '', orden: i.OrdenDia || '', estado: i.EstadoReunion || 'Convocada' })); // eslint-disable-line @typescript-eslint/no-explicit-any
        return Promise.all([
          getJ(c, porTitulo(w, L.tar) + '/items?$select=Id,Title,Responsable,FechaLimiteJ,EstadoTarea,ReunionOrigen&$top=1000'),
          getJ(c, porTitulo(w, L.pat) + '/items?$select=Id,Title,TipoAcuerdo,Importe,Vigencia,EstadoAcuerdo,ContactoAcuerdo&$top=500'),
          getJ(c, porTitulo(w, L.doc) + '/items?$select=Id,Title,FileLeafRef,FileRef,TipoDoc,FechaDoc,Resumen,Modified,Editor/Title,FSObjType&$expand=Editor&$orderby=Modified desc&$top=500'),
          getJ(c, porTitulo(w, L.doc) + '/RootFolder?$select=ServerRelativeUrl'),
          getJ(c, porTitulo(w, 'Comisiones') + "/items?$select=Id&$filter=EstadoCom ne 'Cerrada'&$top=500").catch(() => ({ value: [] }))
        ]);
      }).then((r: any[]) => { // eslint-disable-line @typescript-eslint/no-explicit-any
        this.tars = (r[0].value || []).map((i: any): ITar => ({ id: i.Id, titulo: i.Title || '', resp: i.Responsable || '', fecha: d(i.FechaLimiteJ), estado: i.EstadoTarea || 'Pendiente', origen: i.ReunionOrigen || '' })); // eslint-disable-line @typescript-eslint/no-explicit-any
        this.pats = (r[1].value || []).map((i: any): IPat => ({ id: i.Id, titulo: i.Title || '', tipo: i.TipoAcuerdo || '', importe: i.Importe || 0, vigencia: d(i.Vigencia), estado: i.EstadoAcuerdo || 'En negociación', contacto: i.ContactoAcuerdo || '' })); // eslint-disable-line @typescript-eslint/no-explicit-any
        this.docs = (r[2].value || []).filter((i: any) => i.FSObjType !== 1).map((i: any): IDocJ => ({ id: i.Id, nombre: i.FileLeafRef || '', titulo: i.Title || '', url: i.FileRef || '', tipo: i.TipoDoc || 'Otro', // eslint-disable-line @typescript-eslint/no-explicit-any
          fecha: d(i.FechaDoc), resumen: i.Resumen || '', mod: new Date(i.Modified), autor: i.Editor ? String(i.Editor.Title).replace(/\s*-\s*Gij[oó]n CBP.*$/i, '') : '' }));
        this.carpetaDocs = r[3].ServerRelativeUrl;
        this.nComisiones = (r[4].value || []).length;
        this.estado = 'ok';
      }).catch(() => { this.estado = 'denegado'; });
  }

  private recargar(): void { this.cargar().then(() => this.render()); }

  public render(): void {
    if (this.estado === 'cargando') { this.domElement.innerHTML = '<div class="cbp"><div class="cbp-cargando">Abriendo la oficina de la junta…</div></div>'; return; }
    if (this.estado === 'denegado') {
      this.domElement.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:24px"><section class="cbp-restr"><span class="cbp-candado">Acceso restringido</span><h2>Oficina de la junta</h2>' +
        '<p>Este espacio es solo para los miembros de la junta directiva. Si participas en una comisión de trabajo, entra desde <a href="' + esc(this.context.pageContext.web.absoluteUrl + '/SitePages/Comisiones.aspx') + '" style="color:#FF9F00">Comisiones</a>.</p></section></div></div>';
      return;
    }
    const ahora = new Date();
    const prox = this.reus.filter((r: IReu) => r.fecha && r.fecha >= new Date(ahora.getTime() - 3 * 3600000) && r.estado !== 'Celebrada').sort((a: IReu, b: IReu) => (a.fecha as Date).getTime() - (b.fecha as Date).getTime())[0];
    const abiertas = this.tars.filter((t: ITar) => t.estado !== 'Hecho');
    const vencidas = abiertas.filter((t: ITar) => t.fecha && diasHasta(t.fecha) < 0);
    const activos = this.pats.filter((p: IPat) => p.estado === 'Activo' || p.estado === 'Renovar');
    const total = activos.reduce((s: number, p: IPat) => s + p.importe, 0);
    const nombre = (this.context.pageContext.user.displayName || '').replace(/\s*-\s*Gij[oó]n CBP.*$/i, '').split(' ')[0];

    const kpi = (tab: Tab, n: string, et: string, sub: string, alerta?: boolean): string => '<button type="button" class="cbp-kpi' + (alerta ? ' alerta' : '') + '" data-tab="' + tab + '"><b>' + n + '</b><span>' + et + '</span><small>' + sub + '</small></button>';
    const hero = '<section class="cbp-of-hero"><div class="cbp-of-top"><div><span class="cbp-over">Junta directiva · Gijón CBP</span><h2>' + esc(this.properties.titulo || 'Oficina de la junta') + '</h2>' +
      '<p>' + (nombre ? 'Hola, ' + esc(nombre) + '. ' : '') + 'Hoy es ' + fLarga(ahora) + '.</p></div><span class="cbp-candado">Solo junta directiva</span></div>' +
      '<div class="cbp-kpis">' +
      kpi('reuniones', prox && prox.fecha ? String(Math.max(0, diasHasta(prox.fecha))) : '—', prox && prox.fecha ? (diasHasta(prox.fecha) === 1 ? 'día para la reunión' : 'días para la reunión') : 'Próxima reunión', prox && prox.fecha ? fCorta(prox.fecha) + ' · ' + hhmm(prox.fecha) : 'Sin convocar') +
      kpi('tareas', String(abiertas.length), 'Tareas abiertas', vencidas.length ? vencidas.length + ' vencida' + (vencidas.length === 1 ? '' : 's') : 'Ninguna vencida', vencidas.length > 0) +
      kpi('patrocinios', total ? euros(total) : String(activos.length), total ? 'En acuerdos activos' : 'Acuerdos activos', activos.length + ' patrocinio' + (activos.length === 1 ? '' : 's') + ' y convenios') +
      kpi('comisiones', String(this.nComisiones), 'Comisiones activas', 'Espacios de trabajo') + '</div></section>';

    const cuenta2 = (t: Tab): number => t === 'tareas' ? abiertas.length : t === 'reuniones' ? this.reus.filter((r: IReu) => r.estado === 'Convocada').length : t === 'documentos' ? this.docs.length : t === 'comisiones' ? this.nComisiones : t === 'patrocinios' ? this.pats.filter((p: IPat) => p.estado !== 'Finalizado').length : 0;
    const tabs = '<nav class="cbp-of-tabs" role="tablist" aria-label="Secciones de la oficina">' + TABS.map((t: [Tab, string]) => { const n = cuenta2(t[0]);
      return '<button type="button" role="tab" class="cbp-of-tab" data-tab="' + t[0] + '" aria-selected="' + (this.tab === t[0]) + '">' + t[1] + (n ? '<i>' + n + '</i>' : '') + '</button>'; }).join('') + '</nav>';

    let cuerpo = '';
    if (this.tab === 'resumen') { cuerpo = this.vResumen(prox, abiertas, vencidas); }
    else if (this.tab === 'reuniones') { cuerpo = this.vReuniones(); }
    else if (this.tab === 'tareas') { cuerpo = this.vTareas(); }
    else if (this.tab === 'patrocinios') { cuerpo = this.vPatrocinios(); }
    else if (this.tab === 'documentos') { cuerpo = this.vDocumentos(); }
    else { cuerpo = '<div data-coms="1" style="display:flex;flex-direction:column;gap:16px"></div>'; }

    this.domElement.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:20px;gap:18px">' + hero + tabs + cuerpo + '</div></div>';
    this.enlazar();
    if (this.tab === 'comisiones') {
      const box = this.domElement.querySelector('[data-coms]') as HTMLElement;
      renderComisiones(box, { client: this.context.spHttpClient, w: this.context.pageContext.web.absoluteUrl, embebido: true });
    }
  }

  /* ---------- Vistas ---------- */
  private reuCard(r: IReu, grande: boolean): string {
    const f = r.fecha;
    const acuerdos = this.tars.filter((t: ITar) => t.origen && t.origen === r.titulo);
    return '<div class="cbp-reu"><div class="cbp-reu-f">' + (f ? '<small>' + ['DOM', 'LUN', 'MAR', 'MIÉ', 'JUE', 'VIE', 'SÁB'][f.getDay()] + '</small><b>' + f.getDate() + '</b><span>' + MESES[f.getMonth()] + '</span>' : '<b>?</b>') + '</div>' +
      '<div><h4>' + esc(r.titulo) + '</h4><div class="meta">' + (f ? '<span>' + hhmm(f) + ' h</span>' : '') + (r.lugar ? '<span>' + esc(r.lugar) + '</span>' : '') +
      '<span class="cbp-fpill' + (r.estado === 'Celebrada' ? ' ok' : r.estado === 'Aplazada' ? ' mal' : ' amb') + '">' + esc(r.estado) + (r.estado === 'Convocada' && f ? ' · ' + cuenta(f) : '') + '</span></div>' +
      (grande && r.orden ? '<div class="cbp-odd"><b style="display:block;margin-bottom:4px;font-size:12px;letter-spacing:.08em;text-transform:uppercase">Orden del día</b>' + esc(r.orden) + '</div>' : '') +
      (grande && acuerdos.length ? '<div class="cbp-odd" style="background:#fff;border:1px solid #E6E3DC"><b style="display:block;margin-bottom:4px;font-size:12px;letter-spacing:.08em;text-transform:uppercase">Acuerdos</b>' +
        acuerdos.map((t: ITar) => (t.estado === 'Hecho' ? '✓ ' : '• ') + esc(t.titulo) + (t.resp ? ' — ' + esc(t.resp) : '')).join('\n') + '</div>' : '') +
      (grande ? '<div class="cbp-btns" style="justify-content:flex-start;margin-top:12px">' +
        '<button type="button" class="cbp-btn sm" data-reu-ed="' + r.id + '">Editar</button>' +
        '<button type="button" class="cbp-btn sm" data-reu-conv="' + r.id + '">Copiar convocatoria</button>' +
        '<button type="button" class="cbp-btn sm" data-reu-acu="' + r.id + '">+ Acuerdo</button>' +
        (r.estado !== 'Celebrada' ? '<button type="button" class="cbp-btn sm pri" data-reu-cel="' + r.id + '">Marcar celebrada</button>' : '') + '</div>' : '') + '</div></div>';
  }

  private miniTarea(t: ITar): string {
    const dd = t.fecha ? diasHasta(t.fecha) : 99;
    return '<div class="cbp-mini-it"><span><b data-tar-ed="' + t.id + '" style="cursor:pointer">' + esc(t.titulo) + '</b><small>' + esc(t.resp || 'Sin responsable') + (t.origen ? ' · ' + esc(t.origen) : '') + '</small></span>' +
      (t.fecha ? '<span class="cbp-fpill' + (dd < 0 ? ' mal' : dd <= 7 ? ' amb' : '') + '">' + (dd < 0 ? 'Venció ' : '') + fCorta(t.fecha) + '</span>' : '<span class="cbp-fpill">Sin fecha</span>') + '</div>';
  }

  private ext(n: string): string { return (n.split('.').pop() || '').toLowerCase().slice(0, 4); }
  private docFila(x: IDocJ): string {
    return '<a class="cbp-doc" href="' + esc(x.url) + '?web=1" target="_blank" rel="noopener"><span class="cbp-doc-ico ' + this.ext(x.nombre) + '">' + esc(this.ext(x.nombre)) + '</span>' +
      '<span><b>' + esc(x.titulo || x.nombre) + '</b><small>' + esc(x.tipo) + (x.fecha ? ' · ' + fCorta(x.fecha) + ' ' + x.fecha.getFullYear() : '') + ' · ' + relativo(x.mod) + (x.autor ? ' por ' + esc(x.autor) : '') + '</small></span><span></span></a>';
  }

  private vResumen(prox: IReu | undefined, abiertas: ITar[], vencidas: ITar[]): string {
    const urg = abiertas.slice().sort((a: ITar, b: ITar) => (a.fecha ? a.fecha.getTime() : 9e15) - (b.fecha ? b.fecha.getTime() : 9e15)).slice(0, 6);
    const renov = this.pats.filter((p: IPat) => p.estado !== 'Finalizado' && p.vigencia && diasHasta(p.vigencia) <= 90).sort((a: IPat, b: IPat) => (a.vigencia as Date).getTime() - (b.vigencia as Date).getTime()).slice(0, 5);
    const neg = this.pats.filter((p: IPat) => p.estado === 'En negociación').length;
    return '<div class="cbp-of-grid"><div class="cbp-col">' +
      '<section class="cbp-box"><div class="cbp-box-h"><h3>Próxima reunión</h3><button type="button" class="lnk" data-reu-new="1">+ Convocar</button></div>' +
      (prox ? this.reuCard(prox, true) : '<p class="cbp-nota">No hay ninguna reunión convocada. Convoca la siguiente y prepara el orden del día aquí.</p>') + '</section>' +
      '<section class="cbp-box"><div class="cbp-box-h"><h3>Acuerdos pendientes</h3><button type="button" class="lnk" data-tab="tareas">Ver tablero</button></div>' +
      (urg.length ? '<div class="cbp-mini">' + urg.map((t: ITar) => this.miniTarea(t)).join('') + '</div>' : '<p class="cbp-nota">Todo al día. 👏</p>') +
      (vencidas.length ? '<p class="cbp-nota" style="color:#B00020;font-weight:700">' + vencidas.length + ' acuerdo' + (vencidas.length === 1 ? '' : 's') + ' fuera de plazo.</p>' : '') + '</section>' +
      '</div><div class="cbp-col">' +
      '<section class="cbp-box"><div class="cbp-box-h"><h3>Patrocinios</h3><button type="button" class="lnk" data-tab="patrocinios">Abrir</button></div>' +
      (renov.length ? '<div class="cbp-mini">' + renov.map((p: IPat) => { const dd = diasHasta(p.vigencia as Date);
        return '<div class="cbp-mini-it"><span><b>' + esc(p.titulo) + '</b><small>' + esc(p.tipo) + (p.importe ? ' · ' + euros(p.importe) : '') + '</small></span><span class="cbp-fpill' + (dd < 0 ? ' mal' : ' amb') + '">' + (dd < 0 ? 'Venció ' : 'Vence ') + fCorta(p.vigencia as Date) + '</span></div>'; }).join('') + '</div>'
        : '<p class="cbp-nota">Ningún acuerdo vence en los próximos 90 días.</p>') +
      (neg ? '<p class="cbp-nota">' + neg + ' en negociación.</p>' : '') + '</section>' +
      '<section class="cbp-box"><div class="cbp-box-h"><h3>Últimos documentos</h3><button type="button" class="lnk" data-tab="documentos">Todos</button></div>' +
      (this.docs.length ? '<div class="cbp-docs">' + this.docs.slice(0, 5).map((x: IDocJ) => this.docFila(x)).join('') + '</div>' : '<p class="cbp-nota">Sube actas, presupuestos y convenios para tenerlos a mano.</p>') + '</section>' +
      '</div></div>';
  }

  private vReuniones(): string {
    const fut = this.reus.filter((r: IReu) => r.estado !== 'Celebrada').sort((a: IReu, b: IReu) => (a.fecha ? a.fecha.getTime() : 0) - (b.fecha ? b.fecha.getTime() : 0));
    const pas = this.reus.filter((r: IReu) => r.estado === 'Celebrada');
    return '<div class="cbp-box-h"><span class="cbp-nota">Convoca, prepara el orden del día y registra los acuerdos de cada reunión.</span><button type="button" class="cbp-btn pri" data-reu-new="1">+ Convocar reunión</button></div>' +
      (fut.length ? fut.map((r: IReu) => '<section class="cbp-box">' + this.reuCard(r, true) + '</section>').join('') : '<div class="cbp-vacio">No hay reuniones convocadas.</div>') +
      (pas.length ? '<h3 class="cbp-h2" style="font-size:24px;margin-top:8px">Celebradas</h3>' + pas.map((r: IReu) => '<section class="cbp-box">' + this.reuCard(r, true) + '</section>').join('') : '');
  }

  private vTareas(): string {
    const cols = this.verHechas ? EST_T : EST_T.slice(0, 2);
    const card = (t: ITar): string => {
      const i = EST_T.indexOf(t.estado); const dd = t.fecha ? diasHasta(t.fecha) : 99;
      return '<div class="cbp-kcard"><h4 data-tar-ed="' + t.id + '">' + esc(t.titulo) + '</h4><div class="det">' +
        (t.resp ? '<span class="cbp-avatar" style="width:24px;height:24px;font-size:10px">' + esc(iniciales(t.resp)) + '</span>' + esc(t.resp) : 'Sin responsable') +
        (t.fecha ? '<span class="cbp-fpill' + (t.estado !== 'Hecho' && dd < 0 ? ' mal' : t.estado !== 'Hecho' && dd <= 7 ? ' amb' : '') + '">' + fCorta(t.fecha) + '</span>' : '') + '</div>' +
        (t.origen ? '<div class="det" style="font-size:12px">De: ' + esc(t.origen) + '</div>' : '') +
        '<div class="cbp-kmov"><button type="button" data-tar-mov="' + t.id + '|-1"' + (i <= 0 ? ' style="visibility:hidden"' : '') + '>← ' + (EST_T[i - 1] || '') + '</button>' +
        '<button type="button" data-tar-mov="' + t.id + '|1"' + (i >= 2 ? ' style="visibility:hidden"' : '') + '>' + (EST_T[i + 1] || '') + ' →</button></div></div>';
    };
    return '<div class="cbp-box-h"><label class="cbp-check"><input type="checkbox" data-hechas="1"' + (this.verHechas ? ' checked' : '') + '>Mostrar hechas (' + this.tars.filter((t: ITar) => t.estado === 'Hecho').length + ')</label>' +
      '<button type="button" class="cbp-btn pri" data-tar-new="1">+ Nuevo acuerdo o tarea</button></div>' +
      '<div class="cbp-kanban" style="--cols:' + cols.length + '">' + cols.map((s: string) => {
        const xs = this.tars.filter((t: ITar) => t.estado === s).sort((a: ITar, b: ITar) => (a.fecha ? a.fecha.getTime() : 9e15) - (b.fecha ? b.fecha.getTime() : 9e15));
        return '<div class="cbp-kcol"><div class="cbp-kcol-h">' + s + '<span>' + xs.length + '</span></div>' + (xs.length ? xs.map(card).join('') : '<p class="cbp-nota" style="text-align:center;padding:12px">Nada por aquí</p>') + '</div>';
      }).join('') + '</div>';
  }

  private vPatrocinios(): string {
    const card = (p: IPat): string => { const dd = p.vigencia ? diasHasta(p.vigencia) : 999; const i = EST_P.indexOf(p.estado);
      return '<div class="cbp-kcard"><h4 data-pat-ed="' + p.id + '">' + esc(p.titulo) + '</h4>' + (p.importe ? '<span class="imp">' + euros(p.importe) + '</span>' : '') +
        '<div class="det"><span class="cbp-fpill">' + esc(p.tipo || 'Acuerdo') + '</span>' + (p.vigencia ? '<span class="cbp-fpill' + (p.estado !== 'Finalizado' && dd < 0 ? ' mal' : p.estado !== 'Finalizado' && dd <= 60 ? ' amb' : '') + '">Hasta ' + fCorta(p.vigencia) + ' ' + p.vigencia.getFullYear() + '</span>' : '') + '</div>' +
        (p.contacto ? '<div class="det">' + enlazarUrls(esc(p.contacto)) + '</div>' : '') +
        '<div class="cbp-kmov"><button type="button" data-pat-mov="' + p.id + '|-1"' + (i <= 0 ? ' style="visibility:hidden"' : '') + '>←</button><button type="button" data-pat-mov="' + p.id + '|1"' + (i >= EST_P.length - 1 ? ' style="visibility:hidden"' : '') + '>' + (EST_P[i + 1] || '') + ' →</button></div></div>'; };
    return '<div class="cbp-box-h"><span class="cbp-nota">Seguimiento de patrocinadores, colaboraciones y convenios. Los importes suman en «Activo» y «Renovar».</span><button type="button" class="cbp-btn pri" data-pat-new="1">+ Nuevo acuerdo</button></div>' +
      '<div class="cbp-kanban" style="--cols:4">' + EST_P.map((s: string) => { const xs = this.pats.filter((p: IPat) => p.estado === s); const sum = xs.reduce((a: number, p: IPat) => a + p.importe, 0);
        return '<div class="cbp-kcol"><div class="cbp-kcol-h">' + s + (sum ? '<small>' + euros(sum) + '</small>' : '<span>' + xs.length + '</span>') + '</div>' + xs.map(card).join('') + '</div>'; }).join('') + '</div>';
  }

  private vDocumentos(): string {
    const tipos = ['Todos'].concat(TIPOS_D.filter((t: string) => this.docs.some((x: IDocJ) => x.tipo === t)));
    const lista = this.docs.filter((x: IDocJ) => this.filtroDoc === 'Todos' || x.tipo === this.filtroDoc);
    return '<div class="cbp-box-h"><div class="cbp-chips">' + chips(tipos, this.filtroDoc, 'fdoc') + '</div><button type="button" class="cbp-btn pri" data-doc-new="1">+ Subir documento</button></div>' +
      '<section class="cbp-box" data-drop="1">' + (lista.length ? '<div class="cbp-docs">' + lista.map((x: IDocJ) => this.docFila(x)).join('') + '</div>' : '<div class="cbp-vacio">No hay documentos' + (this.filtroDoc !== 'Todos' ? ' de este tipo' : '') + '. Arrastra archivos aquí para subirlos.</div>') + '</section>';
  }

  /* ---------- Acciones ---------- */
  private enlazar(): void {
    const c = this.context.spHttpClient; const w = this.context.pageContext.web.absoluteUrl;
    const q = (s: string): NodeListOf<HTMLElement> => this.domElement.querySelectorAll(s) as NodeListOf<HTMLElement>;
    const num = (b: HTMLElement, a: string): number => +((b.getAttribute(a) || '').split('|')[0]);
    q('[data-tab]').forEach((b: HTMLElement) => b.addEventListener('click', () => { this.tab = b.getAttribute('data-tab') as Tab; history.replaceState(null, '', '#' + this.tab); this.render(); }));

    // Reuniones
    const formReu = (r?: IReu): void => { formModal(r ? 'Editar reunión' : 'Convocar reunión', [
      { n: 'Title', label: 'Reunión', req: true, full: true, valor: r ? r.titulo : 'Junta ordinaria · ' + MESES[new Date().getMonth()].toLowerCase() },
      { n: 'FechaReunion', label: 'Fecha y hora', tipo: 'datetime', req: true, valor: r && r.fecha ? localDT(r.fecha) : '' },
      { n: 'LugarReunion', label: 'Lugar', valor: r ? r.lugar : '', ph: 'Sede, pabellón o videollamada' },
      { n: 'OrdenDia', label: 'Orden del día', tipo: 'area', valor: r ? r.orden : '1. Lectura y aprobación del acta anterior\n2. \n3. Ruegos y preguntas' },
      { n: 'EstadoReunion', label: 'Estado', tipo: 'select', opciones: ['Convocada', 'Celebrada', 'Aplazada'], valor: r ? r.estado : 'Convocada' }
    ], r ? 'Guardar' : 'Convocar', (v: { [k: string]: any }) => { v.FechaReunion = isoDT(v.FechaReunion); // eslint-disable-line @typescript-eslint/no-explicit-any
      return spPost(c, porTitulo(w, L.reu) + '/items' + (r ? '(' + r.id + ')' : ''), v, !!r).then(() => { aviso(r ? 'Reunión guardada' : 'Reunión convocada'); this.recargar(); }); }); };
    q('[data-reu-new]').forEach((b: HTMLElement) => b.addEventListener('click', () => formReu()));
    q('[data-reu-ed]').forEach((b: HTMLElement) => b.addEventListener('click', () => formReu(this.reus.filter((r: IReu) => r.id === num(b, 'data-reu-ed'))[0])));
    q('[data-reu-cel]').forEach((b: HTMLElement) => b.addEventListener('click', () => spPost(c, porTitulo(w, L.reu) + '/items(' + num(b, 'data-reu-cel') + ')', { EstadoReunion: 'Celebrada' }, true).then(() => { aviso('Reunión marcada como celebrada'); this.recargar(); })));
    q('[data-reu-conv]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
      const r = this.reus.filter((x: IReu) => x.id === num(b, 'data-reu-conv'))[0]; if (!r) { return; }
      const txt = '*' + r.titulo.toUpperCase() + '*\n📅 ' + (r.fecha ? fLarga(r.fecha) + ' a las ' + hhmm(r.fecha) : 'Fecha por confirmar') + (r.lugar ? '\n📍 ' + r.lugar : '') + '\n\n*Orden del día*\n' + r.orden + '\n\nJunta directiva · Gijón CBP';
      navigator.clipboard.writeText(txt).then(() => aviso('Convocatoria copiada: pégala en WhatsApp o en un correo')).catch(() => aviso('No se pudo copiar', true));
    }));
    q('[data-reu-acu]').forEach((b: HTMLElement) => b.addEventListener('click', () => { const r = this.reus.filter((x: IReu) => x.id === num(b, 'data-reu-acu'))[0]; formTar(undefined, r ? r.titulo : ''); }));

    // Tareas
    const formTar = (t?: ITar, origen?: string): void => { formModal(t ? 'Editar acuerdo' : 'Nuevo acuerdo o tarea', [
      { n: 'Title', label: 'Acuerdo o tarea', req: true, full: true, valor: t ? t.titulo : '' },
      { n: 'Responsable', label: 'Responsable', valor: t ? t.resp : '' }, { n: 'FechaLimiteJ', label: 'Fecha límite', tipo: 'date', valor: t && t.fecha ? t.fecha.toISOString().slice(0, 10) : '' },
      { n: 'EstadoTarea', label: 'Estado', tipo: 'select', opciones: EST_T, valor: t ? t.estado : 'Pendiente' },
      { n: 'ReunionOrigen', label: 'Reunión de origen', valor: t ? t.origen : (origen || ''), ph: 'Opcional' }
    ], t ? 'Guardar' : 'Crear', (v: { [k: string]: any }) => { v.FechaLimiteJ = isoDia(v.FechaLimiteJ); // eslint-disable-line @typescript-eslint/no-explicit-any
      return spPost(c, porTitulo(w, L.tar) + '/items' + (t ? '(' + t.id + ')' : ''), v, !!t).then(() => { aviso('Guardado'); this.recargar(); }); },
      560, t ? '<button type="button" class="cbp-btn full" data-tar-del="' + t.id + '" style="color:#B00020;border-color:#B00020;justify-self:start">Eliminar</button>' : '');
      const del = document.querySelector('[data-tar-del]') as HTMLElement;
      if (del && t) { del.addEventListener('click', () => { if (!window.confirm('¿Eliminar «' + t.titulo + '»?')) { return; } spPost(c, porTitulo(w, L.tar) + '/items(' + t.id + ')/recycle').then(() => { (document.querySelector('#cbp-modal .cbp-x') as HTMLElement).click(); aviso('Eliminado'); this.recargar(); }); }); }
    };
    q('[data-tar-new]').forEach((b: HTMLElement) => b.addEventListener('click', () => formTar()));
    q('[data-tar-ed]').forEach((b: HTMLElement) => b.addEventListener('click', () => formTar(this.tars.filter((t: ITar) => t.id === num(b, 'data-tar-ed'))[0])));
    q('[data-tar-mov]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
      const p = (b.getAttribute('data-tar-mov') || '').split('|'); const t = this.tars.filter((x: ITar) => x.id === +p[0])[0]; const n = t ? EST_T[EST_T.indexOf(t.estado) + +p[1]] : '';
      if (!t || !n) { return; }
      spPost(c, porTitulo(w, L.tar) + '/items(' + t.id + ')', { EstadoTarea: n }, true).then(() => { t.estado = n; this.render(); }).catch((x: Error) => aviso('Error: ' + x.message, true));
    }));
    const hh = this.domElement.querySelector('[data-hechas]') as HTMLInputElement;
    if (hh) { hh.addEventListener('change', () => { this.verHechas = hh.checked; this.render(); }); }

    // Patrocinios
    const formPat = (p?: IPat): void => { formModal(p ? 'Editar acuerdo' : 'Nuevo patrocinio o convenio', [
      { n: 'Title', label: 'Entidad', req: true, full: true, valor: p ? p.titulo : '', ph: 'Empresa o institución' },
      { n: 'TipoAcuerdo', label: 'Tipo', tipo: 'select', opciones: TIPOS_P, valor: p ? p.tipo : 'Patrocinio' },
      { n: 'EstadoAcuerdo', label: 'Estado', tipo: 'select', opciones: EST_P, valor: p ? p.estado : 'En negociación' },
      { n: 'Importe', label: 'Importe anual (€)', tipo: 'number', valor: p ? p.importe : '' },
      { n: 'Vigencia', label: 'Vigente hasta', tipo: 'date', valor: p && p.vigencia ? p.vigencia.toISOString().slice(0, 10) : '' },
      { n: 'ContactoAcuerdo', label: 'Contacto', full: true, valor: p ? p.contacto : '', ph: 'Nombre, teléfono o correo' }
    ], p ? 'Guardar' : 'Crear', (v: { [k: string]: any }) => { v.Vigencia = isoDia(v.Vigencia); // eslint-disable-line @typescript-eslint/no-explicit-any
      return spPost(c, porTitulo(w, L.pat) + '/items' + (p ? '(' + p.id + ')' : ''), v, !!p).then(() => { aviso('Guardado'); this.recargar(); }); }); };
    q('[data-pat-new]').forEach((b: HTMLElement) => b.addEventListener('click', () => formPat()));
    q('[data-pat-ed]').forEach((b: HTMLElement) => b.addEventListener('click', () => formPat(this.pats.filter((p: IPat) => p.id === num(b, 'data-pat-ed'))[0])));
    q('[data-pat-mov]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
      const p = (b.getAttribute('data-pat-mov') || '').split('|'); const x = this.pats.filter((y: IPat) => y.id === +p[0])[0]; const n = x ? EST_P[EST_P.indexOf(x.estado) + +p[1]] : '';
      if (!x || !n) { return; }
      spPost(c, porTitulo(w, L.pat) + '/items(' + x.id + ')', { EstadoAcuerdo: n }, true).then(() => { x.estado = n; this.render(); }).catch((e: Error) => aviso('Error: ' + e.message, true));
    }));

    // Documentos
    q('[data-fdoc]').forEach((b: HTMLElement) => b.addEventListener('click', () => { this.filtroDoc = b.getAttribute('data-fdoc') || 'Todos'; this.render(); }));
    const subirDocs = (files: File[]): void => {
      const body = formModal('Subir ' + (files.length === 1 ? files[0].name : files.length + ' documentos'), [
        { n: 'TipoDoc', label: 'Tipo', tipo: 'select', opciones: TIPOS_D, valor: 'Acta' }, { n: 'FechaDoc', label: 'Fecha del documento', tipo: 'date', valor: new Date().toISOString().slice(0, 10) },
        { n: 'Resumen', label: 'Resumen', tipo: 'area', ph: 'Opcional: de qué trata, acuerdos clave…' }
      ], 'Subir', (v: { [k: string]: any }) => { // eslint-disable-line @typescript-eslint/no-explicit-any
        const campos = { TipoDoc: v.TipoDoc, FechaDoc: isoDia(v.FechaDoc), Resumen: v.Resumen };
        return files.reduce((pr: Promise<void>, f: File) => pr.then(() => subirArchivo(c, w, this.carpetaDocs, f, campos)), Promise.resolve()).then(() => { aviso('Documentos subidos'); this.filtroDoc = 'Todos'; this.recargar(); });
      }, 540, files.length ? '' : '<label class="cbp-field full">Archivos<input type="file" multiple data-files="1" required></label>');
      const inp = body.querySelector('[data-files]') as HTMLInputElement;
      if (inp) { inp.addEventListener('change', () => { files = []; for (let i = 0; i < (inp.files || []).length; i++) { files.push((inp.files as FileList)[i]); } }); }
    };
    q('[data-doc-new]').forEach((b: HTMLElement) => b.addEventListener('click', () => subirDocs([])));
    const drop = this.domElement.querySelector('[data-drop]') as HTMLElement;
    if (drop) {
      drop.addEventListener('dragover', (ev: DragEvent) => { ev.preventDefault(); drop.style.outline = '2px dashed #FF9F00'; });
      drop.addEventListener('dragleave', () => { drop.style.outline = ''; });
      drop.addEventListener('drop', (ev: DragEvent) => { ev.preventDefault(); drop.style.outline = ''; const fs: File[] = []; const l = ev.dataTransfer ? ev.dataTransfer.files : null; if (l) { for (let i = 0; i < l.length; i++) { fs.push(l[i]); } } if (fs.length) { subirDocs(fs); } });
    }
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: 'Oficina de la junta' }, groups: [{ groupName: 'General', groupFields: [PropertyPaneTextField('titulo', { label: 'Título' })] }] }] };
  }
}
