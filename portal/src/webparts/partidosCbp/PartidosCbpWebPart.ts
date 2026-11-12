import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneSlider } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { cargarEstilos, cargarEventos, IEvento, esc, MESES } from '../../common/cbp';
import { cargarEstilosUi, spPost, aviso } from '../../common/ui';
import { cargarEstilos3 } from '../../common/estilos3';
import { cargarEstilos4 } from '../../common/estilos4';
import { porTitulo, getJ, formModal } from '../../common/sp2';
import { IPartidoG, IGest, IClub, IGCtx, cargarGestion, cargarClubes, esCoordinador, guardarGest, vPartidos, enlazarGestion, vClubes, enlazarClubes } from './gestion';
import { IPartido, IRegla, IPista, IAlerta, IJugV, IPlazo, cargarFranjas, analizar, leerDisponibilidad, instalacionDe, minDia, hm, durTxt, DIA_TXT, PLAZO_DEF } from './motor';
import { IEscEstado, nuevoEstado, vEscenarios, enlazarEscenarios, vista, CSS_ESC } from './escenarios';

export interface IPartidosCbpWebPartProps { titulo: string; margen: number; plazoDia: string; plazoHora: string; }
const DIAS_PLAZO = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];

interface IEq { id: number; nombre: string; codigo: string; franja: string; dur: number; pref: string; alt: string; catPDM: string; }
type Tab = 'semana' | 'temporada' | 'clubes' | 'reglas';
const L = { reglas: 'Partidos · Conflictos', pistas: 'Partidos · Pistas', franjas: 'Partidos · Franjas' };
const DIAS_CORTOS = ['dom', 'lun', 'mar', 'mié', 'jue', 'vie', 'sáb'];
interface IFranjaRow { id: number; grupo: string; dias: string; desde: string; hasta: string; tipo: string; cond: string; habitual: boolean; }
const COND = ['Cualquiera', 'Local', 'Visitante'];
const MOTIVOS = ['Entrenador', 'Jugadores', 'Otro'];
const GRUPOS = ['Júnior y sénior', 'Minibasket', 'Cadete e infantil 1ª', 'Resto de categorías', 'Nacional', 'Sin franja'];
const DIA = 86400000;

const CSS5 = `
.cbp-pt-sem{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.cbp-pt-sem button{font:inherit;width:40px;height:40px;border-radius:10px;border:1px solid #333;background:#1a1a1a;color:#fff;cursor:pointer;font-size:18px}
.cbp-pt-sem button:hover{border-color:#FF9F00}
.cbp-pt-sem b{font-size:16px;color:#fff;min-width:210px;text-align:center}
.cbp-alerta{display:grid;grid-template-columns:6px minmax(0,1fr);gap:14px;background:#fff;border:1px solid #E6E3DC;border-radius:14px;padding:14px 16px 14px 0;overflow:hidden;cursor:pointer}
.cbp-alerta:before{content:"";border-radius:0 6px 6px 0;background:#9a9a9a}
.cbp-alerta.critica:before{background:#B00020}.cbp-alerta.aviso:before{background:#FF9F00}.cbp-alerta.info:before{background:#1F4E9A}
.cbp-alerta h4{margin:0 0 4px;font-size:15px}
.cbp-alerta p{margin:0;font-size:13.5px;color:#333;line-height:1.45}
.cbp-alerta .acc{margin-top:6px;font-size:13px;font-weight:700;color:#090909}
.cbp-alerta .acc:before{content:"→ ";color:#B36F00}
.cbp-alerta.sel{outline:2px solid #090909}
.cbp-tl{background:#fff;border:1px solid #E6E3DC;border-radius:18px;padding:16px 18px;display:flex;flex-direction:column;gap:6px;overflow-x:auto}
.cbp-tl h3{margin:0 0 6px;font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:21px;text-transform:uppercase}
.cbp-tl-esc{display:grid;grid-template-columns:150px minmax(560px,1fr);font-size:11px;color:#5c5c5c;font-weight:700}
.cbp-tl-esc div{position:relative;height:16px}
.cbp-tl-esc span{position:absolute;transform:translateX(-50%)}
.cbp-tl-fila{display:grid;grid-template-columns:150px minmax(560px,1fr);align-items:stretch;border-top:1px solid #EFECE6}
.cbp-tl-fila>b{font-size:12.5px;padding:8px 8px 8px 0;line-height:1.25}
.cbp-tl-fila>b small{display:block;font-weight:600;color:#5c5c5c;font-size:11px}
.cbp-tl-carril{position:relative;background-image:repeating-linear-gradient(90deg,#F1EEE8 0,#F1EEE8 1px,transparent 1px,transparent calc(100%/14))}
.cbp-tl-res{position:absolute;top:0;bottom:0;background:rgba(255,159,0,.10);border-left:2px solid rgba(255,159,0,.5);border-right:2px solid rgba(255,159,0,.5)}
.cbp-tl-p{position:absolute;height:26px;border-radius:7px;background:#090909;color:#fff;font-size:11.5px;font-weight:700;padding:0 7px;display:flex;align-items:center;gap:5px;white-space:nowrap;overflow:hidden;cursor:pointer;border:0;font-family:inherit}

.cbp-tl-p{flex-direction:column;align-items:flex-start;justify-content:center;gap:0;line-height:1.15;height:40px!important;padding:0 7px}
.cbp-tl-p b{font-size:11.5px}.cbp-tl-p span{font-size:11px;font-weight:600;opacity:.9;overflow:hidden;text-overflow:ellipsis;max-width:100%}
.cbp-al-det{background:#fff;border:1px solid #E6E3DC;border-radius:14px;padding:10px 14px}
.cbp-al-det summary{cursor:pointer;font-size:14px;padding:4px 0}
.cbp-al-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:10px;margin-top:10px}
.cbp-al-grid .cbp-alerta{padding:10px 12px 10px 0}.cbp-al-grid .cbp-alerta p{font-size:13px}
.cbp-esc-banda{background:#090909;color:#fff;border-radius:12px;padding:10px 14px;font-size:14px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.cbp-esc-banda b{color:#FF9F00}.cbp-esc-banda .cbp-btn{background:#fff}
.cbp-tl-p.fuera{background:#fff;color:#090909;border:1.5px dashed #090909}
.cbp-tl-p.critica{background:#B00020;color:#fff;border:0}.cbp-tl-p.aviso{box-shadow:inset 0 0 0 2px #FF9F00}
.cbp-tl-p.fuera.critica{background:#FDE7E9;color:#B00020;border:1.5px dashed #B00020}
.cbp-tl-p.hl{outline:3px solid #FF9F00;outline-offset:1px;z-index:2}
.cbp-tl-ley{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:#5c5c5c;margin-top:6px}
.cbp-tl-ley i{display:inline-block;width:14px;height:10px;border-radius:3px;margin-right:5px;vertical-align:middle}
.cbp-tabla.pt td,.cbp-tabla.pt th{vertical-align:top}
.cbp-tabla.pt tr.clic{cursor:pointer}.cbp-tabla.pt tr.clic:hover td{background:#F6F5F1}
.cbp-num-c{font-weight:800;text-align:center}
.cbp-num-c.r{color:#B00020}.cbp-num-c.a{color:#B36F00}
`;

function lunes(d: Date): Date { const x = new Date(d.getFullYear(), d.getMonth(), d.getDate()); const w = (x.getDay() + 6) % 7; x.setDate(x.getDate() - w); return x; }

export default class PartidosCbpWebPart extends BaseClientSideWebPart<IPartidosCbpWebPartProps> {
  private cargado: boolean = false;
  private error: string = '';
  private tab: Tab = 'semana';
  private semana: Date = lunes(new Date(Date.now() + DIA));
  private partidos: IPartidoG[] = [];
  private gest: { [pid: number]: IGest } = {}; private clubes: IClub[] = []; private esCoord: boolean = false;
  private eqs: IEq[] = []; private reglas: IRegla[] = []; private pistas: IPista[] = [];
  private sel: number = -1;
  private jugs: IJugV[] = [];
  private est: IEscEstado = nuevoEstado();
  private alAbiertas: boolean = false;
  private franjas: IFranjaRow[] = [];

  protected onInit(): Promise<void> {
    cargarEstilos(); cargarEstilosUi(); cargarEstilos3(); cargarEstilos4();
    if (!document.getElementById('cbp-css5')) { const s = document.createElement('style'); s.id = 'cbp-css5'; s.textContent = CSS5 + CSS_ESC; document.head.appendChild(s); }
    const h = window.location.hash.match(/semana=(\d{4}-\d{2}-\d{2})/); if (h) { this.semana = lunes(new Date(h[1] + 'T12:00:00')); }
    return this.cargar();
  }

  private get margen(): number { return this.properties.margen === undefined ? 0 : this.properties.margen; }
  private get plazo(): IPlazo {
    const d = DIAS_PLAZO.indexOf((this.properties.plazoDia || '').toLowerCase()); const h = (this.properties.plazoHora || '').match(/^(\d{1,2}):(\d{2})$/);
    return { dia: d >= 0 ? d : PLAZO_DEF.dia, hora: h ? +h[1] * 60 + +h[2] : PLAZO_DEF.hora };
  }

  private cargar(): Promise<void> {
    const c = this.context.spHttpClient; const w = this.context.pageContext.web.absoluteUrl;
    return Promise.all([
      cargarEventos(c, w),
      getJ(c, porTitulo(w, 'Equipos') + '/items?$select=Id,Title,Codigo,GrupoFranja,Duracion,PistasPref,PistasAlt,CategoriaPDM&$top=200'),
      getJ(c, porTitulo(w, L.reglas) + '/items?$select=Id,Title,EquipoA,CondA,EquipoB,CondB,Minutos,Motivo&$top=500'),
      getJ(c, porTitulo(w, L.pistas) + '/items?$select=Id,Title,Instalacion,Modulo,AliasFBPA,Disponibilidad,Controlar,Orden,NombrePDM,Permitidos&$orderby=Orden&$top=200'),
      getJ(c, porTitulo(w, 'Jugadores') + '/items?$select=Title,EquipoJ,Activo,EquiposHabitual,EquiposOcasional&$top=3000').catch(() => ({ value: [] })),
      cargarGestion(c, w), cargarClubes(c, w), esCoordinador(c, w),
      getJ(c, porTitulo(w, L.franjas) + '/items?$select=Id,Title,Dias,Desde,Hasta,TipoF,Condicion,Habitual&$top=500').catch(() => ({ value: [] }))
    ]).then((r: any[]) => { // eslint-disable-line @typescript-eslint/no-explicit-any
      this.franjas = (r[8].value || []).map((i: any): IFranjaRow => ({ id: i.Id, grupo: i.Title || '', dias: i.Dias || '', desde: i.Desde || '', hasta: i.Hasta || '', tipo: i.TipoF || 'oficial', cond: i.Condicion || '', habitual: i.Habitual !== false })); // eslint-disable-line @typescript-eslint/no-explicit-any
      cargarFranjas(this.franjas.filter((x: IFranjaRow) => /^\d{1,2}:\d{2}$/.test(x.desde) && /^\d{1,2}:\d{2}$/.test(x.hasta)).map((x: IFranjaRow) => ({ grupo: x.grupo, f: { id: x.id, dias: x.dias.toLowerCase().split(/[,\s/]+/).map((d: string) => DIAS_CORTOS.indexOf(d.slice(0, 3).replace('mie', 'mié').replace('sab', 'sáb'))).filter((n: number) => n >= 0),
        desde: ('0' + x.desde).slice(-5), hasta: ('0' + x.hasta).slice(-5), tipo: (x.tipo as 'oficial' | 'conformidad' | 'condicion'), cond: x.cond, pref: x.habitual } })));
      this.eqs = (r[1].value || []).map((i: any): IEq => ({ id: i.Id, nombre: i.Title, codigo: (i.Codigo || '').toUpperCase(), franja: i.GrupoFranja || 'Sin franja', dur: i.Duracion || 90, pref: i.PistasPref || '', alt: i.PistasAlt || '', catPDM: i.CategoriaPDM || '' })); // eslint-disable-line @typescript-eslint/no-explicit-any
      this.reglas = (r[2].value || []).map((i: any): IRegla => ({ id: i.Id, titulo: i.Title || '', a: (i.EquipoA || '').toUpperCase(), condA: i.CondA || 'Cualquiera', b: (i.EquipoB || '').toUpperCase(), condB: i.CondB || 'Cualquiera', min: i.Minutos || 0, motivo: i.Motivo || 'Entrenador' })); // eslint-disable-line @typescript-eslint/no-explicit-any
      this.pistas = (r[3].value || []).map((i: any): IPista => ({ id: i.Id, clave: i.Title, inst: i.Instalacion || i.Title, modulo: i.Modulo || '', alias: i.AliasFBPA || i.Instalacion || '', disp: leerDisponibilidad(i.Disponibilidad || ''), dispTexto: i.Disponibilidad || '', permitidos: (i.Permitidos || '').split(/\n/).map((x: string) => x.trim()).filter((x: string) => x), controlar: i.Controlar !== false, orden: i.Orden || 0, pdm: i.NombrePDM || '' })); // eslint-disable-line @typescript-eslint/no-explicit-any
      this.jugs = (r[4].value || []).filter((i: any) => i.Activo !== false && ((i.EquiposHabitual || []).length || (i.EquiposOcasional || []).length)) // eslint-disable-line @typescript-eslint/no-explicit-any
        .map((i: any): IJugV => ({ nombre: i.Title || '', equipo: i.EquipoJ || '', habitual: i.EquiposHabitual || [], ocasional: i.EquiposOcasional || [] })); // eslint-disable-line @typescript-eslint/no-explicit-any
      const porNombre: { [k: string]: IEq } = {}; this.eqs.forEach((e: IEq) => { porNombre[e.nombre] = e; });
      this.gest = r[5]; this.clubes = r[6]; this.esCoord = r[7];
      this.partidos = (r[0] as IEvento[]).filter((e: IEvento) => e.tipo === 'Partido').map((e: IEvento): IPartidoG => {
        const q = porNombre[e.equipo]; const g = this.gest[e.id];
        const vigente = g && g.fechaProp && g.estado !== 'Rechazado' && g.estado !== 'Oficial' ? g.fechaProp : e.fecha;
        const jm = e.notas.match(/Jornada (\d+)/);
        const pk = e.local && g && g.pista ? this.pistas.filter((k: IPista) => k.clave === g.pista)[0] : undefined;
        return { id: e.id, equipo: e.equipo, codigo: q ? q.codigo : '', fecha: vigente, oficial: e.fecha, local: e.local, rival: e.rival, lugar: pk ? (pk.alias || pk.inst) : e.lugar, lugarFbpa: e.lugar, sinHora: /Hora por confirmar/i.test(e.notas), dur: q ? q.dur : 90, franja: q ? q.franja : 'Sin franja',
          comp: e.notas.split(' · ')[0] || '', jornada: jm ? jm[1] : '', catPDM: q ? q.catPDM : '' };
      });
      // La FBPA ya refleja el cambio: se marca como cargado
      if (this.esCoord) {
        this.partidos.forEach((p: IPartidoG) => { const g = this.gest[p.id];
          if (g && g.fechaProp && ['Validado', 'Pedido al rival', 'Aceptado'].indexOf(g.estado) >= 0 && Math.abs(g.fechaProp.getTime() - p.oficial.getTime()) < 60000) {
            g.estado = 'Cargado en FBPA'; guardarGest(this.ctx([]), p, { EstadoG: 'Cargado en FBPA' }, 'La FBPA ya muestra el nuevo horario').catch(() => undefined);
          } });
      }
      this.cargado = true;
    }).catch((x: Error) => { this.error = x.message; this.cargado = true; });
  }

  private nombreCod(c: string): string { const e = this.eqs.filter((x: IEq) => x.codigo === c)[0]; return e ? e.nombre : c; }
  private semanaDe(d: Date): IPartidoG[] { const a = d.getTime(); const b = a + 7 * DIA; return this.partidos.filter((p: IPartidoG) => p.fecha.getTime() >= a && p.fecha.getTime() < b); }
  private ctx(semana: IPartidoG[]): IGCtx {
    const lineas = (t: string): string[] => t.split(/\n/).map((x: string) => x.trim()).filter((x: string) => x);
    return { c: this.context.spHttpClient, w: this.context.pageContext.web.absoluteUrl, usuario: this.context.pageContext.user.displayName, esCoord: this.esCoord, margen: this.margen, plazo: this.plazo,
      semana, gest: this.gest, clubes: this.clubes, pistas: this.pistas, reglas: this.reglas, jugs: this.jugs,
      eqInfo: (p: IPartido) => { const e = this.eqs.filter((x: IEq) => x.nombre === p.equipo)[0]; return { pref: e ? lineas(e.pref) : [], alt: e ? lineas(e.alt) : [] }; },
      recargar: () => { this.cargar().then(() => this.render()); } };
  }

  public render(): void {
    if (!this.cargado) { this.domElement.innerHTML = '<div class="cbp"><div class="cbp-cargando">Analizando partidos…</div></div>'; return; }
    if (this.error) { this.domElement.innerHTML = '<div class="cbp"><div class="cbp-wrap"><div class="cbp-vacio">No se pudieron cargar los datos: ' + esc(this.error) + '</div></div></div>'; return; }
    const ps = this.semanaDe(this.semana);
    const al = analizar(ps, this.reglas, this.pistas, this.margen, this.jugs);
    const fin = new Date(this.semana.getTime() + 6 * DIA);
    const rango = this.semana.getDate() + (this.semana.getMonth() !== fin.getMonth() ? ' ' + MESES[this.semana.getMonth()].toLowerCase() : '') + ' – ' + fin.getDate() + ' ' + MESES[fin.getMonth()].toLowerCase();
    const n = (g: string): number => al.filter((a: IAlerta) => a.grav === g).length;
    const casa = ps.filter((p: IPartido) => p.local).length;

    const kpi = (v: string, et: string, sub: string, cls?: string): string => '<div class="cbp-kpi' + (cls ? ' ' + cls : '') + '" style="cursor:default"><b>' + v + '</b><span>' + et + '</span><small>' + sub + '</small></div>';
    const hero = '<section class="cbp-of-hero"><div class="cbp-of-top"><div><span class="cbp-over">Coordinación deportiva</span><h2>' + esc(this.properties.titulo || 'Gestión de partidos') + '</h2>' +
      '<p>Alertas sobre el horario de la FBPA, propuestas de cambio, peticiones a otros clubes, carga en la FBPA y Excel del Patronato.' + (this.esCoord ? ' <b style="color:#FF9F00">Tienes permisos de coordinación.</b>' : '') + '</p></div>' +
      '<div class="cbp-pt-sem"><button type="button" data-sem="-1" aria-label="Semana anterior">‹</button><b>Semana ' + rango + '</b><button type="button" data-sem="1" aria-label="Semana siguiente">›</button></div></div>' +
      '<div class="cbp-kpis">' + kpi(String(ps.length), 'Partidos', casa + ' en casa · ' + (ps.length - casa) + ' fuera') +
      kpi(String(n('critica')), 'Conflictos', n('critica') ? 'Hay que resolverlos' : 'Ninguno', n('critica') ? 'alerta' : '') +
      kpi(String(n('aviso')), 'Avisos', 'Conformidades y reservas') + kpi(String(n('info')), 'Pendientes', 'Horas sin fijar') + '</div></section>';

    const tabs = '<nav class="cbp-of-tabs" role="tablist">' + ([['semana', 'Semana'], ['temporada', 'Temporada'], ['clubes', 'Clubes'], ['reglas', 'Reglas y pistas']] as [Tab, string][]).map((t: [Tab, string]) =>
      '<button type="button" role="tab" class="cbp-of-tab" data-tab="' + t[0] + '" aria-selected="' + (this.tab === t[0]) + '">' + t[1] + '</button>').join('') + '</nav>';

    let cuerpo = '';
    const gctx = this.ctx(ps);
    const clave = String(this.semana.getTime());
    if (this.est.semanaKey !== clave) { this.est = nuevoEstado(); this.est.semanaKey = clave; }
    if (this.tab === 'semana') {
      const v = vista(this.est, ps, al);
      const escSel = this.est.res && this.est.sel ? this.est.res.filter((e: { clave: string }) => e.clave === this.est.sel)[0] : undefined;
      const banda = escSel ? '<div class="cbp-esc-banda">Estás viendo el escenario <b>«' + esc(escSel.nombre) + '»</b> sobre el gantt. <button type="button" class="cbp-btn sm" data-esc="">Volver a la situación actual</button> <a href="#cbp-motor" class="cbp-btn sm">Ir al motor</a></div>' : '';
      cuerpo = banda + this.vSemana(v.ps, v.al, v.movidos) + (ps.length ? vPartidos(gctx, al) : '') + (this.esCoord && ps.length ? '<div id="cbp-motor"></div>' + vEscenarios(gctx, this.est, al) : '');
    } else if (this.tab === 'temporada') { cuerpo = this.vTemporada(); }
    else if (this.tab === 'clubes') { const riv: string[] = []; this.partidos.forEach((p: IPartidoG) => { if (riv.indexOf(p.rival) < 0) { riv.push(p.rival); } }); cuerpo = vClubes(gctx, riv); } else { cuerpo = this.vReglas(); }
    this.domElement.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:20px;gap:18px">' + hero + tabs + cuerpo + '</div></div>';
    this.enlazar(al);
    enlazarGestion(this.domElement, gctx, al); enlazarClubes(this.domElement, gctx);
    if (this.tab === 'semana' && this.esCoord) { enlazarEscenarios(this.domElement, gctx, this.est, al, () => this.render()); }
  }

  /* ---------- Semana ---------- */
  private vSemana(ps: IPartido[], al: IAlerta[], movidos?: { [id: number]: Date }): string {
    const mov = movidos || {};
    if (!ps.length) { return '<div class="cbp-vacio">No hay partidos publicados por la FBPA esta semana.</div>'; }
    const grav: { [id: number]: string } = {};
    al.forEach((a: IAlerta) => a.partidos.forEach((id: number) => { if (!grav[id] || a.grav === 'critica') { grav[id] = a.grav; } }));
    const alertas = al.length ? al.map((a: IAlerta, i: number) => '<div class="cbp-alerta ' + a.grav + (this.sel === i ? ' sel' : '') + '" data-al="' + i + '" role="button" tabindex="0"><div><h4>' + esc(a.titulo) + '</h4><p>' + esc(a.detalle) + '</p><div class="acc">' + esc(a.accion) + '</div></div></div>').join('')
      : '<div class="cbp-vacio" style="border-style:solid;background:#E8F5E9;color:#1B5E20;font-weight:700">Semana limpia: sin conflictos con las reglas configuradas.</div>';
    const hl = this.sel >= 0 && al[this.sel] ? al[this.sel].partidos : [];

    // Línea de tiempo por día
    const H0 = 8 * 60; const H1 = 22 * 60; const pos = (m: number): number => Math.max(0, Math.min(100, ((m - H0) / (H1 - H0)) * 100));
    const dias: Date[] = []; for (let i = 0; i < 7; i++) { const d = new Date(this.semana.getTime() + i * DIA); if (ps.some((p: IPartido) => p.fecha.toDateString() === d.toDateString())) { dias.push(d); } }
    const escala = '<div class="cbp-tl-esc"><i></i><div>' + [8, 10, 12, 14, 16, 18, 20, 22].map((h: number) => '<span style="left:' + pos(h * 60) + '%">' + h + ':00</span>').join('') + '</div></div>';
    const tl = dias.map((d: Date) => {
      const delDia = ps.filter((p: IPartido) => p.fecha.toDateString() === d.toDateString()).sort((a: IPartido, b: IPartido) => a.fecha.getTime() - b.fecha.getTime());
      const filas: { [k: string]: IPartido[] } = {};
      delDia.forEach((p: IPartido) => { const k = p.local ? (instalacionDe(p, this.pistas) || p.lugar || 'Pista por confirmar') : '~Fuera de casa'; (filas[k] = filas[k] || []).push(p); });
      const claves = Object.keys(filas).sort((a: string, b: string) => (a.charAt(0) === '~' ? 1 : 0) - (b.charAt(0) === '~' ? 1 : 0) || a.localeCompare(b));
      return '<div class="cbp-tl"><h3>' + DIA_TXT[d.getDay()] + ' ' + d.getDate() + ' ' + MESES[d.getMonth()].toLowerCase() + '</h3>' + escala + claves.map((k: string) => {
        const lista = filas[k];
        const carriles: number[] = []; const lane: { [id: number]: number } = {};
        lista.forEach((p: IPartido) => { const s = minDia(p.fecha); let i = carriles.findIndex((fin: number) => fin <= s); if (i < 0) { i = carriles.length; carriles.push(0); } carriles[i] = s + p.dur; lane[p.id] = i; });
        const alto = Math.max(1, carriles.length) * 44 + 6;
        const mods = this.pistas.filter((x: IPista) => x.inst === k);
        const res = mods.map((m: IPista) => m.disp.filter((v: { dia: number }) => v.dia === d.getDay()).map((v: { desde: number; hasta: number }) => '<div class="cbp-tl-res" style="left:' + pos(v.desde) + '%;width:' + (pos(v.hasta) - pos(v.desde)) + '%" title="Reserva ' + esc(m.modulo) + ' ' + hm(v.desde) + '–' + hm(v.hasta) + '"></div>').join('')).join('');
        const resTxt = mods.map((m: IPista) => m.disp.filter((v: { dia: number }) => v.dia === d.getDay()).map((v: { desde: number; hasta: number }) => (m.modulo || m.inst) + ' ' + hm(v.desde).replace(':00', '') + '–' + hm(v.hasta).replace(':00', '')).join(', ')).filter((x: string) => x);
        const bloques = lista.map((p: IPartido) => { const s = minDia(p.fecha); const antes = mov[p.id];
          const fantasma = antes && antes.toDateString() === p.fecha.toDateString() ? '<span class="cbp-tl-p ghost' + (p.local ? '' : ' fuera') + '" style="left:' + pos(minDia(antes)) + '%;width:' + Math.max(6, pos(minDia(antes) + p.dur) - pos(minDia(antes))) + '%;top:' + (3 + lane[p.id] * 44) + 'px">' + hm(minDia(antes)) + '</span>' : '';
          return fantasma + '<button type="button" class="cbp-tl-p' + (p.local ? '' : ' fuera') + (grav[p.id] ? ' ' + grav[p.id] : '') + (antes ? ' mov' : '') + (hl.indexOf(p.id) >= 0 ? ' hl' : '') + '" data-p="' + p.id + '" style="left:' + pos(s) + '%;width:' + Math.max(6, pos(s + p.dur) - pos(s)) + '%;top:' + (3 + lane[p.id] * 44) + 'px" title="' + esc(p.equipo + ' · ' + hm(s) + ' · ' + (p.local ? 'vs ' : 'en ') + p.rival + (p.local ? '' : ' · ' + p.lugar) + (antes ? ' · antes ' + DIA_TXT[antes.getDay()].slice(0, 3) + ' ' + hm(minDia(antes)) : '')) + '"><b>' + hm(s) + '</b><span>' + esc(p.codigo || p.equipo) + '</span></button>'; }).join('');
        return '<div class="cbp-tl-fila"><b>' + esc(k.charAt(0) === '~' ? 'Fuera de casa' : k) + '<small>' + (k.charAt(0) === '~' ? lista.length + ' desplazamiento' + (lista.length === 1 ? '' : 's') : (resTxt.length ? 'Reservado: ' + resTxt.map(esc).join('<br>') : 'Sin reserva del club ese día')) + '</small></b><div class="cbp-tl-carril" style="height:' + alto + 'px">' + res + bloques + '</div></div>';
      }).join('') + '</div>';
    }).join('');
    const ley = '<div class="cbp-tl-ley">' + (Object.keys(mov).length ? '<span><i style="outline:2px dashed #FF9F00"></i>Movido por el escenario</span>' : '') + '<span><i style="background:#090909"></i>En casa</span><span><i style="border:1.5px dashed #090909"></i>Fuera</span><span><i style="background:#B00020"></i>Conflicto</span><span><i style="box-shadow:inset 0 0 0 2px #FF9F00"></i>Aviso</span><span><i style="background:rgba(255,159,0,.2)"></i>Pista reservada</span></div>';
    const nC = al.filter((a: IAlerta) => a.grav === 'critica').length; const nA = al.filter((a: IAlerta) => a.grav === 'aviso').length; const nI = al.length - nC - nA;
    return '<div class="cbp-col" style="gap:12px">' + tl + ley + '</div>' +
      '<details class="cbp-al-det"' + (this.alAbiertas ? ' open' : '') + '><summary><b>Detalle de alertas</b> · ' + (al.length ? nC + ' conflicto' + (nC === 1 ? '' : 's') + ', ' + nA + ' aviso' + (nA === 1 ? '' : 's') + (nI ? ', ' + nI + ' pendiente' + (nI === 1 ? '' : 's') : '') : 'ninguna') + '</summary>' + (al.length ? '<div class="cbp-al-grid">' + alertas + '</div>' : alertas) + '</details>';
  }

  /* ---------- Temporada ---------- */
  private vTemporada(): string {
    const hoy = lunes(new Date());
    const semanas: Date[] = []; for (let i = 0; i < 30; i++) { const d = new Date(hoy.getTime() + i * 7 * DIA); if (this.semanaDe(d).length) { semanas.push(d); } }
    if (!semanas.length) { return '<div class="cbp-vacio">No hay partidos futuros publicados.</div>'; }
    return '<p class="cbp-nota" style="margin:0">Todas las semanas con partidos publicados por la FBPA a partir de hoy. Pulsa una fila para revisarla.</p><section class="cbp-box"><div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Semana</th><th>Partidos</th><th>En casa</th><th>Conflictos</th><th>Avisos</th><th>Principal problema</th></tr></thead><tbody>' +
      semanas.map((d: Date) => { const ps = this.semanaDe(d); const al = analizar(ps, this.reglas, this.pistas, this.margen, this.jugs); const c = al.filter((a: IAlerta) => a.grav === 'critica'); const a = al.filter((x: IAlerta) => x.grav === 'aviso');
        const f = new Date(d.getTime() + 6 * DIA);
        return '<tr class="clic" data-ir="' + d.getTime() + '"><td><b>' + d.getDate() + ' ' + MESES[d.getMonth()].toLowerCase() + ' – ' + f.getDate() + ' ' + MESES[f.getMonth()].toLowerCase() + '</b></td><td class="cbp-num-c">' + ps.length + '</td><td class="cbp-num-c">' + ps.filter((p: IPartido) => p.local).length + '</td>' +
          '<td class="cbp-num-c' + (c.length ? ' r' : '') + '">' + c.length + '</td><td class="cbp-num-c' + (a.length ? ' a' : '') + '">' + a.length + '</td><td>' + esc(c[0] ? c[0].titulo : a[0] ? a[0].titulo : '—') + '</td></tr>'; }).join('') +
      '</tbody></table></div></section>';
  }

  /* ---------- Reglas y pistas ---------- */
  private vReglas(): string {
    const cond = (c: string): string => c === 'Cualquiera' ? '' : ' <span class="cbp-fpill">' + (c === 'Local' ? 'en casa' : 'fuera') + '</span>';
    const reglas = this.reglas.slice().sort((a: IRegla, b: IRegla) => a.a.localeCompare(b.a) || a.b.localeCompare(b.b));
    const tReglas = '<section class="cbp-box"><div class="cbp-box-h"><h3>Conflictos entre equipos</h3><button type="button" class="cbp-btn sm pri" data-rnew="1">+ Nueva regla</button></div>' +
      '<p class="cbp-nota" style="margin:0">Parejas de equipos que no pueden jugar demasiado seguidas el mismo día, normalmente por compartir entrenador. La separación se mide entre las horas de inicio. Las subidas de jugadores se calculan solas a partir de las plantillas (' + this.jugs.length + ' jugador' + (this.jugs.length === 1 ? '' : 'es') + ' con vínculos en otros equipos).</p>' +
      '<div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Equipo A</th><th>Equipo B</th><th>Separación</th><th>Motivo</th><th>Nota</th></tr></thead><tbody>' +
      reglas.map((r: IRegla) => '<tr class="clic" data-red="' + r.id + '"><td><b>' + esc(this.nombreCod(r.a)) + '</b>' + cond(r.condA) + '</td><td><b>' + esc(this.nombreCod(r.b)) + '</b>' + cond(r.condB) + '</td><td>' + durTxt(r.min) + '</td><td>' + esc(r.motivo) + '</td><td>' + esc(/↔/.test(r.titulo) ? '' : r.titulo) + '</td></tr>').join('') +
      '</tbody></table></div></section>';
    const tPistas = '<section class="cbp-box"><div class="cbp-box-h"><h3>Pistas y reservas</h3><button type="button" class="cbp-btn sm" data-pnew="1">+ Añadir pista</button></div>' +
      '<p class="cbp-nota" style="margin:0">Módulos que el club tiene reservados cada fin de semana. Los partidos en casa se cuentan contra estos módulos.</p>' +
      '<div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Instalación</th><th>Módulo</th><th>Reserva</th><th>Solo para</th></tr></thead><tbody>' +
      this.pistas.map((p: IPista) => '<tr class="clic" data-ped="' + p.id + '"><td><b>' + esc(p.inst) + '</b></td><td>' + esc(p.modulo) + '</td><td style="white-space:pre-line">' + esc(p.dispTexto || '—') + '</td><td>' + esc((p.permitidos || []).join(', ') || 'Todos') + '</td></tr>').join('') +
      '</tbody></table></div></section>';
    const tEq = '<section class="cbp-box"><div class="cbp-box-h"><h3>Equipos</h3></div><p class="cbp-nota" style="margin:0">Código, franja horaria de la FBPA y duración estimada de cada partido (para calcular la ocupación de pista).</p>' +
      '<div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Código</th><th>Equipo</th><th>Franja FBPA</th><th>Duración</th><th>Pistas preferentes</th></tr></thead><tbody>' +
      this.eqs.filter((e: IEq) => e.codigo).sort((a: IEq, b: IEq) => a.codigo.localeCompare(b.codigo)).map((e: IEq) => '<tr class="clic" data-eed="' + e.id + '"><td><b>' + esc(e.codigo) + '</b></td><td>' + esc(e.nombre) + '</td><td>' + esc(e.franja) + '</td><td>' + e.dur + ' min</td><td style="white-space:pre-line">' + esc(e.pref) + '</td></tr>').join('') +
      '</tbody></table></div></section>';
    const grupos: string[] = []; this.franjas.forEach((x: IFranjaRow) => { if (grupos.indexOf(x.grupo) < 0) { grupos.push(x.grupo); } });
    const fr = '<section class="cbp-box"><div class="cbp-box-h"><h3>Franjas FBPA</h3>' + (this.esCoord ? '<button type="button" class="cbp-btn sm" data-fnew="1">+ Añadir franja</button>' : '') + '</div>' +
      '<p class="cbp-nota" style="margin:0">Horas de inicio por grupo de equipos. El motor de partidos solo propone horas de franjas <b>oficiales</b>, y empieza por las <b>habituales</b>; las secundarias solo si no hay otra salida. Pulsa una fila para cambiarla. Los equipos sin grupo (p. ej. Sénior Femenino A, competición nacional) no los mueve el motor.</p>' +
      '<div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Grupo</th><th>Días</th><th>Horas de inicio</th><th>Tipo</th><th>Uso</th></tr></thead><tbody>' +
      grupos.map((g: string) => this.franjas.filter((x: IFranjaRow) => x.grupo === g).map((x: IFranjaRow, i: number) => '<tr class="clic" data-fed="' + x.id + '"><td>' + (i === 0 ? '<b>' + esc(g) + '</b>' : '') + '</td><td>' + esc(x.dias) + '</td><td>' + esc(x.desde === x.hasta ? x.desde : x.desde + '–' + x.hasta) + '</td><td>' +
        (x.tipo === 'oficial' ? 'Oficial' : x.tipo === 'conformidad' ? 'Con conformidad del rival' : 'Condicionada') + (x.cond ? '<br><small>' + esc(x.cond) + '</small>' : '') + '</td><td>' + (x.tipo !== 'oficial' ? '—' : x.habitual ? '<span class="cbp-plazo ok">Habitual</span>' : '<span class="cbp-plazo">Secundaria</span>') + '</td></tr>').join('')).join('') +
      '</tbody></table></div></section>';
    const pl = this.plazo;
    const tPlazo = '<section class="cbp-box"><div class="cbp-box-h"><h3>Plazo ordinario y normativa</h3></div>' +
      '<p class="cbp-nota" style="margin:0">El equipo local puede fijar o ajustar la hora de sus partidos sin notificar al rival hasta el <b>' + DIAS_PLAZO[pl.dia] + ' a las ' + hm(pl.hora) + ' de la semana anterior</b> al partido. Pasado ese plazo, el cambio necesita la conformidad del rival. El plazo se cambia en la configuración de esta sección.</p>' +
      '<p class="cbp-nota" style="margin:0">El Reglamento General 2026/27 (art. 155 B) fija además las 20:00 del lunes anterior a la jornada como límite para pedir cambios sin cuota; las condiciones exactas de cada competición están en sus normas específicas. <a href="https://www.fbpa.es/descargar?seccion=documentos&id=432&delegacion=1" target="_blank" rel="noopener">Reglamento General y de Competiciones FBPA 2026/27</a></p></section>';
    return tPlazo + tReglas + '<div class="cbp-of-grid" style="grid-template-columns:minmax(0,1fr) minmax(0,1fr)"><div class="cbp-col">' + tPistas + fr + '</div><div class="cbp-col">' + tEq + '</div></div>';
  }

  /* ---------- Eventos ---------- */
  private enlazar(al: IAlerta[]): void {
    const c = this.context.spHttpClient; const w = this.context.pageContext.web.absoluteUrl;
    const q = (s: string): NodeListOf<HTMLElement> => this.domElement.querySelectorAll(s) as NodeListOf<HTMLElement>;
    const recargar = (): void => { this.cargar().then(() => this.render()); };
    q('[data-tab]').forEach((b: HTMLElement) => b.addEventListener('click', () => { this.tab = b.getAttribute('data-tab') as Tab; this.sel = -1; this.render(); }));
    q('[data-sem]').forEach((b: HTMLElement) => b.addEventListener('click', () => { this.semana = new Date(this.semana.getTime() + +(b.getAttribute('data-sem') || 0) * 7 * DIA + 3 * 3600000); this.semana = lunes(this.semana); this.sel = -1; this.tab = this.tab === 'reglas' ? 'semana' : this.tab; this.render(); }));
    q('[data-ir]').forEach((b: HTMLElement) => b.addEventListener('click', () => { this.semana = lunes(new Date(+(b.getAttribute('data-ir') || 0) + 3 * 3600000)); this.tab = 'semana'; this.sel = -1; this.render(); window.scrollTo(0, 0); }));
    q('.cbp-al-det').forEach((d: HTMLElement) => d.addEventListener('toggle', () => { this.alAbiertas = (d as HTMLDetailsElement).open; }));
    q('[data-al]').forEach((b: HTMLElement) => { const f = (): void => { const i = +(b.getAttribute('data-al') || 0); this.sel = this.sel === i ? -1 : i; this.render(); }; b.addEventListener('click', f); b.addEventListener('keydown', (e: KeyboardEvent) => { if (e.key === 'Enter') { f(); } }); });
    q('[data-p]').forEach((b: HTMLElement) => b.addEventListener('click', () => { const id = +(b.getAttribute('data-p') || 0); const i = al.findIndex((a: IAlerta) => a.partidos.indexOf(id) >= 0); this.sel = i; this.render(); }));

    const opcEq = this.eqs.filter((e: IEq) => e.codigo).sort((a: IEq, b: IEq) => a.codigo.localeCompare(b.codigo)).map((e: IEq) => e.codigo + ' · ' + e.nombre);
    const cod = (s: string): string => (s || '').split(' · ')[0];
    const formRegla = (r?: IRegla): void => {
      formModal(r ? 'Editar regla' : 'Nueva regla de conflicto', [
        { n: 'a', label: 'Equipo A', tipo: 'select', opciones: opcEq, valor: r ? opcEq.filter((o: string) => cod(o) === r.a)[0] : '' },
        { n: 'condA', label: 'Cuando A juega', tipo: 'select', opciones: COND, valor: r ? r.condA : 'Cualquiera' },
        { n: 'b', label: 'Equipo B', tipo: 'select', opciones: opcEq, valor: r ? opcEq.filter((o: string) => cod(o) === r.b)[0] : '' },
        { n: 'condB', label: 'Cuando B juega', tipo: 'select', opciones: COND, valor: r ? r.condB : 'Cualquiera' },
        { n: 'horas', label: 'Separación mínima (horas)', tipo: 'number', req: true, valor: r ? r.min / 60 : 2 },
        { n: 'motivo', label: 'Motivo', tipo: 'select', opciones: MOTIVOS, valor: r ? r.motivo : 'Entrenador' },
        { n: 'nota', label: 'Nota', full: true, valor: r && !/↔/.test(r.titulo) ? r.titulo : '', ph: 'Ej.: Cuesta entrena a los dos equipos' }
      ], r ? 'Guardar' : 'Crear regla', (v: { [k: string]: any }) => { // eslint-disable-line @typescript-eslint/no-explicit-any
        const body = { Title: v.nota || (cod(v.a) + ' ↔ ' + cod(v.b)), EquipoA: cod(v.a), CondA: v.condA, EquipoB: cod(v.b), CondB: v.condB, Minutos: Math.round((v.horas || 0) * 60), Motivo: v.motivo };
        return spPost(c, porTitulo(w, L.reglas) + '/items' + (r ? '(' + r.id + ')' : ''), body, !!r).then(() => { aviso('Regla guardada'); recargar(); });
      }, 600, r ? '<button type="button" class="cbp-btn full" data-del-r="1" style="color:#B00020;border-color:#B00020;justify-self:start">Eliminar regla</button>' : '');
      const del = document.querySelector('[data-del-r]') as HTMLElement;
      if (del && r) { del.addEventListener('click', () => { if (!window.confirm('¿Eliminar esta regla?')) { return; } spPost(c, porTitulo(w, L.reglas) + '/items(' + r.id + ')/recycle').then(() => { (document.querySelector('#cbp-modal .cbp-x') as HTMLElement).click(); aviso('Regla eliminada'); recargar(); }); }); }
    };
    q('[data-rnew]').forEach((b: HTMLElement) => b.addEventListener('click', () => formRegla()));
    q('[data-red]').forEach((b: HTMLElement) => b.addEventListener('click', () => formRegla(this.reglas.filter((r: IRegla) => r.id === +(b.getAttribute('data-red') || 0))[0])));

    const formPista = (p?: IPista): void => {
      formModal(p ? 'Editar ' + p.clave : 'Nueva pista', [
        { n: 'Instalacion', label: 'Instalación', req: true, valor: p ? p.inst : '', ph: 'Como aparece en la FBPA' },
        { n: 'Modulo', label: 'Módulo', valor: p ? p.modulo : '', ph: 'M1, M2, Central…' },
        { n: 'Disponibilidad', label: 'Reserva', tipo: 'area', valor: p ? p.dispTexto : '', ayuda: '(una línea por franja: «Sábado 09:00-14:00»)' },
        { n: 'Permitidos', label: 'Solo para estos equipos', tipo: 'area', valor: p ? (p.permitidos || []).join('\n') : '', ayuda: '(una línea por categoría, equipo o código: «Junior», «Sénior», «SFA». Vacío = todos)' },
        { n: 'Controlar', label: 'Controlar capacidad', tipo: 'check', valor: p ? p.controlar : true }
      ], 'Guardar', (v: { [k: string]: any }) => { // eslint-disable-line @typescript-eslint/no-explicit-any
        const body = { Title: v.Instalacion + (v.Modulo ? ' · ' + v.Modulo : ''), Instalacion: v.Instalacion, Modulo: v.Modulo, AliasFBPA: v.Instalacion, Disponibilidad: v.Disponibilidad, Controlar: v.Controlar, Permitidos: v.Permitidos, Orden: p ? p.orden : this.pistas.length + 1 };
        return spPost(c, porTitulo(w, L.pistas) + '/items' + (p ? '(' + p.id + ')' : ''), body, !!p).then(() => { aviso('Pista guardada'); recargar(); });
      }, 520);
    };
    q('[data-pnew]').forEach((b: HTMLElement) => b.addEventListener('click', () => formPista()));
    q('[data-ped]').forEach((b: HTMLElement) => b.addEventListener('click', () => formPista(this.pistas.filter((p: IPista) => p.id === +(b.getAttribute('data-ped') || 0))[0])));
    const formFranja = (x?: IFranjaRow): void => {
      if (!this.esCoord) { return; }
      const gr: string[] = GRUPOS.filter((g: string) => g !== 'Nacional' && g !== 'Sin franja');
      formModal(x ? 'Editar franja' : 'Nueva franja', [
        { n: 'Title', label: 'Grupo', tipo: 'select', opciones: gr, valor: x ? x.grupo : gr[0] },
        { n: 'Dias', label: 'Días', req: true, valor: x ? x.dias : 'sáb', ph: 'sáb · dom · mié, jue, vie' },
        { n: 'Desde', label: 'Primera hora de inicio', req: true, valor: x ? x.desde : '09:30', ph: '09:30' }, { n: 'Hasta', label: 'Última hora de inicio', req: true, valor: x ? x.hasta : '12:30', ph: '12:30' },
        { n: 'TipoF', label: 'Tipo', tipo: 'select', opciones: ['oficial', 'conformidad', 'condicion'], valor: x ? x.tipo : 'oficial' },
        { n: 'Habitual', label: 'Habitual (el motor la usa primero)', tipo: 'check', valor: x ? x.habitual : true },
        { n: 'Condicion', label: 'Condición', tipo: 'area', full: true, valor: x ? x.cond : '' }
      ], 'Guardar', (v: { [k: string]: any }) => { // eslint-disable-line @typescript-eslint/no-explicit-any
        if (!/^\d{1,2}:\d{2}$/.test(v.Desde) || !/^\d{1,2}:\d{2}$/.test(v.Hasta)) { return Promise.reject(new Error('Horas en formato 09:30')); }
        return spPost(c, porTitulo(w, L.franjas) + '/items' + (x ? '(' + x.id + ')' : ''), v, !!x).then(() => { aviso('Franja guardada'); recargar(); });
      }, 560, x ? '<button type="button" class="cbp-btn full" data-del-f="1" style="color:#B00020;border-color:#B00020;justify-self:start">Eliminar franja</button>' : '');
      const del = document.querySelector('[data-del-f]') as HTMLElement;
      if (del && x) { del.addEventListener('click', () => { spPost(c, porTitulo(w, L.franjas) + '/items(' + x.id + ')/recycle').then(() => { (document.querySelector('#cbp-modal .cbp-x') as HTMLElement).click(); aviso('Franja eliminada'); recargar(); }); }); }
    };
    q('[data-fnew]').forEach((b: HTMLElement) => b.addEventListener('click', () => formFranja()));
    q('[data-fed]').forEach((b: HTMLElement) => b.addEventListener('click', () => formFranja(this.franjas.filter((x: IFranjaRow) => x.id === +(b.getAttribute('data-fed') || 0))[0])));

    q('[data-eed]').forEach((b: HTMLElement) => b.addEventListener('click', () => {
      const e = this.eqs.filter((x: IEq) => x.id === +(b.getAttribute('data-eed') || 0))[0]; if (!e) { return; }
      formModal(e.nombre, [
        { n: 'Codigo', label: 'Código', req: true, valor: e.codigo }, { n: 'GrupoFranja', label: 'Franja FBPA', tipo: 'select', opciones: GRUPOS, valor: e.franja },
        { n: 'Duracion', label: 'Duración del partido (min)', tipo: 'number', valor: e.dur },
        { n: 'PistasPref', label: 'Pistas preferentes', tipo: 'area', valor: e.pref, ayuda: '(una por línea)' }, { n: 'PistasAlt', label: 'Pistas alternativas', tipo: 'area', valor: e.alt }
      ], 'Guardar', (v: { [k: string]: any }) => { v.Codigo = String(v.Codigo).toUpperCase(); return spPost(c, porTitulo(w, 'Equipos') + '/items(' + e.id + ')', v, true).then(() => { aviso('Equipo guardado'); recargar(); }); }, 560); // eslint-disable-line @typescript-eslint/no-explicit-any
    }));
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: 'Gestión de partidos' }, groups: [{ groupName: 'General', groupFields: [
      PropertyPaneTextField('titulo', { label: 'Título' }),
      PropertyPaneSlider('margen', { label: 'Margen entre partidos en la misma pista (min)', min: 0, max: 60, step: 5 }),
      PropertyPaneTextField('plazoDia', { label: 'Plazo ordinario: día de la semana anterior', description: 'Por defecto viernes' }),
      PropertyPaneTextField('plazoHora', { label: 'Plazo ordinario: hora', description: 'Por defecto 11:00' })
    ] }] }] };
  }
}
