import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneSlider } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { cargarEstilos, cargarEventos, IEvento, esc, MESES } from '../../common/cbp';
import { cargarEstilosUi, spPost, aviso } from '../../common/ui';
import { cargarEstilos3 } from '../../common/estilos3';
import { cargarEstilos4 } from '../../common/estilos4';
import { porTitulo, getJ, formModal } from '../../common/sp2';
import { IPartidoG, IGest, IClub, IGCtx, cargarGestion, cargarClubes, esCoordinador, guardarGest, vPartidos, enlazarGestion, vClubes, enlazarClubes } from './gestion';
import { IPartido, IRegla, IPista, IAlerta, IJugV, analizar, leerDisponibilidad, instalacionDe, minDia, hm, durTxt, DIA_TXT, FRANJAS } from './motor';

export interface IPartidosCbpWebPartProps { titulo: string; margen: number; }

interface IEq { id: number; nombre: string; codigo: string; franja: string; dur: number; pref: string; alt: string; catPDM: string; }
type Tab = 'semana' | 'temporada' | 'clubes' | 'reglas';
const L = { reglas: 'Partidos · Conflictos', pistas: 'Partidos · Pistas' };
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

  protected onInit(): Promise<void> {
    cargarEstilos(); cargarEstilosUi(); cargarEstilos3(); cargarEstilos4();
    if (!document.getElementById('cbp-css5')) { const s = document.createElement('style'); s.id = 'cbp-css5'; s.textContent = CSS5; document.head.appendChild(s); }
    const h = window.location.hash.match(/semana=(\d{4}-\d{2}-\d{2})/); if (h) { this.semana = lunes(new Date(h[1] + 'T12:00:00')); }
    return this.cargar();
  }

  private get margen(): number { return this.properties.margen === undefined ? 30 : this.properties.margen; }

  private cargar(): Promise<void> {
    const c = this.context.spHttpClient; const w = this.context.pageContext.web.absoluteUrl;
    return Promise.all([
      cargarEventos(c, w),
      getJ(c, porTitulo(w, 'Equipos') + '/items?$select=Id,Title,Codigo,GrupoFranja,Duracion,PistasPref,PistasAlt,CategoriaPDM&$top=200'),
      getJ(c, porTitulo(w, L.reglas) + '/items?$select=Id,Title,EquipoA,CondA,EquipoB,CondB,Minutos,Motivo&$top=500'),
      getJ(c, porTitulo(w, L.pistas) + '/items?$select=Id,Title,Instalacion,Modulo,AliasFBPA,Disponibilidad,Controlar,Orden,NombrePDM&$orderby=Orden&$top=200'),
      getJ(c, porTitulo(w, 'Jugadores') + '/items?$select=Title,EquipoJ,Activo,EquiposHabitual,EquiposOcasional&$top=3000').catch(() => ({ value: [] })),
      cargarGestion(c, w), cargarClubes(c, w), esCoordinador(c, w)
    ]).then((r: any[]) => { // eslint-disable-line @typescript-eslint/no-explicit-any
      this.eqs = (r[1].value || []).map((i: any): IEq => ({ id: i.Id, nombre: i.Title, codigo: (i.Codigo || '').toUpperCase(), franja: i.GrupoFranja || 'Sin franja', dur: i.Duracion || 90, pref: i.PistasPref || '', alt: i.PistasAlt || '', catPDM: i.CategoriaPDM || '' })); // eslint-disable-line @typescript-eslint/no-explicit-any
      this.reglas = (r[2].value || []).map((i: any): IRegla => ({ id: i.Id, titulo: i.Title || '', a: (i.EquipoA || '').toUpperCase(), condA: i.CondA || 'Cualquiera', b: (i.EquipoB || '').toUpperCase(), condB: i.CondB || 'Cualquiera', min: i.Minutos || 0, motivo: i.Motivo || 'Entrenador' })); // eslint-disable-line @typescript-eslint/no-explicit-any
      this.pistas = (r[3].value || []).map((i: any): IPista => ({ id: i.Id, clave: i.Title, inst: i.Instalacion || i.Title, modulo: i.Modulo || '', alias: i.AliasFBPA || i.Instalacion || '', disp: leerDisponibilidad(i.Disponibilidad || ''), dispTexto: i.Disponibilidad || '', controlar: i.Controlar !== false, orden: i.Orden || 0, pdm: i.NombrePDM || '' })); // eslint-disable-line @typescript-eslint/no-explicit-any
      this.jugs = (r[4].value || []).filter((i: any) => i.Activo !== false && ((i.EquiposHabitual || []).length || (i.EquiposOcasional || []).length)) // eslint-disable-line @typescript-eslint/no-explicit-any
        .map((i: any): IJugV => ({ nombre: i.Title || '', equipo: i.EquipoJ || '', habitual: i.EquiposHabitual || [], ocasional: i.EquiposOcasional || [] })); // eslint-disable-line @typescript-eslint/no-explicit-any
      const porNombre: { [k: string]: IEq } = {}; this.eqs.forEach((e: IEq) => { porNombre[e.nombre] = e; });
      this.gest = r[5]; this.clubes = r[6]; this.esCoord = r[7];
      this.partidos = (r[0] as IEvento[]).filter((e: IEvento) => e.tipo === 'Partido').map((e: IEvento): IPartidoG => {
        const q = porNombre[e.equipo]; const g = this.gest[e.id];
        const vigente = g && g.fechaProp && g.estado !== 'Rechazado' && g.estado !== 'Oficial' ? g.fechaProp : e.fecha;
        const jm = e.notas.match(/Jornada (\d+)/);
        return { id: e.id, equipo: e.equipo, codigo: q ? q.codigo : '', fecha: vigente, oficial: e.fecha, local: e.local, rival: e.rival, lugar: e.lugar, sinHora: /Hora por confirmar/i.test(e.notas), dur: q ? q.dur : 90, franja: q ? q.franja : 'Sin franja',
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
    return { c: this.context.spHttpClient, w: this.context.pageContext.web.absoluteUrl, usuario: this.context.pageContext.user.displayName, esCoord: this.esCoord, margen: this.margen,
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
    if (this.tab === 'semana') { cuerpo = this.vSemana(ps, al) + (ps.length ? vPartidos(gctx, al) : ''); } else if (this.tab === 'temporada') { cuerpo = this.vTemporada(); }
    else if (this.tab === 'clubes') { const riv: string[] = []; this.partidos.forEach((p: IPartidoG) => { if (riv.indexOf(p.rival) < 0) { riv.push(p.rival); } }); cuerpo = vClubes(gctx, riv); } else { cuerpo = this.vReglas(); }
    this.domElement.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:20px;gap:18px">' + hero + tabs + cuerpo + '</div></div>';
    this.enlazar(al);
    enlazarGestion(this.domElement, gctx, al); enlazarClubes(this.domElement, gctx);
  }

  /* ---------- Semana ---------- */
  private vSemana(ps: IPartido[], al: IAlerta[]): string {
    if (!ps.length) { return '<div class="cbp-vacio">No hay partidos publicados por la FBPA esta semana.</div>'; }
    const grav: { [id: number]: string } = {};
    al.forEach((a: IAlerta) => a.partidos.forEach((id: number) => { if (!grav[id] || a.grav === 'critica') { grav[id] = a.grav; } }));
    const alertas = al.length ? al.map((a: IAlerta, i: number) => '<div class="cbp-alerta ' + a.grav + (this.sel === i ? ' sel' : '') + '" data-al="' + i + '" role="button" tabindex="0"><div><h4>' + esc(a.titulo) + '</h4><p>' + esc(a.detalle) + '</p><div class="acc">' + esc(a.accion) + '</div>' + (a.grav !== 'info' ? '<button type="button" class="cbp-btn sm" data-sol="' + (ps.filter((p: IPartido) => a.partidos.indexOf(p.id) >= 0 && p.local)[0] || ps.filter((p: IPartido) => a.partidos.indexOf(p.id) >= 0)[0] || { id: 0 }).id + '" style="margin-top:8px">Proponer solución</button>' : '') + '</div></div>').join('')
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
        const alto = Math.max(1, carriles.length) * 30 + 6;
        const mods = this.pistas.filter((x: IPista) => x.inst === k);
        const res = mods.map((m: IPista) => m.disp.filter((v: { dia: number }) => v.dia === d.getDay()).map((v: { desde: number; hasta: number }) => '<div class="cbp-tl-res" style="left:' + pos(v.desde) + '%;width:' + (pos(v.hasta) - pos(v.desde)) + '%" title="Reserva ' + esc(m.modulo) + ' ' + hm(v.desde) + '–' + hm(v.hasta) + '"></div>').join('')).join('');
        const nMods = mods.filter((m: IPista) => m.disp.some((v: { dia: number }) => v.dia === d.getDay())).length;
        const bloques = lista.map((p: IPartido) => { const s = minDia(p.fecha);
          return '<button type="button" class="cbp-tl-p' + (p.local ? '' : ' fuera') + (grav[p.id] ? ' ' + grav[p.id] : '') + (hl.indexOf(p.id) >= 0 ? ' hl' : '') + '" data-p="' + p.id + '" style="left:' + pos(s) + '%;width:' + Math.max(6, pos(s + p.dur) - pos(s)) + '%;top:' + (3 + lane[p.id] * 30) + 'px" title="' + esc(p.equipo + ' · ' + hm(s) + ' · ' + (p.local ? 'vs ' : 'en ') + p.rival + (p.local ? '' : ' · ' + p.lugar)) + '">' + hm(s) + ' ' + esc(p.codigo || p.equipo) + '</button>'; }).join('');
        return '<div class="cbp-tl-fila"><b>' + esc(k.charAt(0) === '~' ? 'Fuera de casa' : k) + '<small>' + (k.charAt(0) === '~' ? lista.length + ' desplazamiento' + (lista.length === 1 ? '' : 's') : nMods + ' módulo' + (nMods === 1 ? '' : 's') + ' reservado' + (nMods === 1 ? '' : 's')) + '</small></b><div class="cbp-tl-carril" style="height:' + alto + 'px">' + res + bloques + '</div></div>';
      }).join('') + '</div>';
    }).join('');
    const ley = '<div class="cbp-tl-ley"><span><i style="background:#090909"></i>En casa</span><span><i style="border:1.5px dashed #090909"></i>Fuera</span><span><i style="background:#B00020"></i>Conflicto</span><span><i style="box-shadow:inset 0 0 0 2px #FF9F00"></i>Aviso</span><span><i style="background:rgba(255,159,0,.2)"></i>Pista reservada</span></div>';
    return '<div class="cbp-of-grid" style="grid-template-columns:minmax(0,1fr) minmax(0,1.6fr)"><div class="cbp-col"><div class="cbp-box-h"><h3 class="cbp-h2" style="font-size:24px;margin:0">Alertas</h3><span class="cbp-nota">' + al.length + '</span></div>' + alertas + '</div>' +
      '<div class="cbp-col">' + tl + ley + '</div></div>';
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
      '<div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Instalación</th><th>Módulo</th><th>Reserva</th><th>Control</th></tr></thead><tbody>' +
      this.pistas.map((p: IPista) => '<tr class="clic" data-ped="' + p.id + '"><td><b>' + esc(p.inst) + '</b></td><td>' + esc(p.modulo) + '</td><td style="white-space:pre-line">' + esc(p.dispTexto || '—') + '</td><td>' + (p.controlar ? 'Sí' : 'No') + '</td></tr>').join('') +
      '</tbody></table></div></section>';
    const tEq = '<section class="cbp-box"><div class="cbp-box-h"><h3>Equipos</h3></div><p class="cbp-nota" style="margin:0">Código, franja horaria de la FBPA y duración estimada de cada partido (para calcular la ocupación de pista).</p>' +
      '<div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Código</th><th>Equipo</th><th>Franja FBPA</th><th>Duración</th><th>Pistas preferentes</th></tr></thead><tbody>' +
      this.eqs.filter((e: IEq) => e.codigo).sort((a: IEq, b: IEq) => a.codigo.localeCompare(b.codigo)).map((e: IEq) => '<tr class="clic" data-eed="' + e.id + '"><td><b>' + esc(e.codigo) + '</b></td><td>' + esc(e.nombre) + '</td><td>' + esc(e.franja) + '</td><td>' + e.dur + ' min</td><td style="white-space:pre-line">' + esc(e.pref) + '</td></tr>').join('') +
      '</tbody></table></div></section>';
    const fr = '<section class="cbp-box"><div class="cbp-box-h"><h3>Franjas FBPA 2026/27</h3></div><p class="cbp-nota" style="margin:0">Horas de inicio oficiales por grupo (texto de bandas facilitado por el club). La Sénior Femenina A es competición nacional y no se valida aquí.</p><div class="cbp-scroll"><table class="cbp-tabla pt"><thead><tr><th>Grupo</th><th>Oficial</th><th>Con conformidad o condición</th></tr></thead><tbody>' +
      Object.keys(FRANJAS).map((g: string) => { const t = (tp: string[]): string => FRANJAS[g].filter((f: { tipo: string }) => tp.indexOf(f.tipo) >= 0).map((f: { dias: number[]; desde: string; hasta: string }) => f.dias.map((x: number) => DIA_TXT[x].slice(0, 3)).join('/') + ' ' + (f.desde === f.hasta ? f.desde : f.desde + '–' + f.hasta)).join(' · ');
        return '<tr><td><b>' + esc(g) + '</b></td><td>' + t(['oficial']) + '</td><td>' + t(['conformidad', 'condicion']) + '</td></tr>'; }).join('') + '</tbody></table></div></section>';
    return tReglas + '<div class="cbp-of-grid" style="grid-template-columns:minmax(0,1fr) minmax(0,1fr)"><div class="cbp-col">' + tPistas + fr + '</div><div class="cbp-col">' + tEq + '</div></div>';
  }

  /* ---------- Eventos ---------- */
  private enlazar(al: IAlerta[]): void {
    const c = this.context.spHttpClient; const w = this.context.pageContext.web.absoluteUrl;
    const q = (s: string): NodeListOf<HTMLElement> => this.domElement.querySelectorAll(s) as NodeListOf<HTMLElement>;
    const recargar = (): void => { this.cargar().then(() => this.render()); };
    q('[data-tab]').forEach((b: HTMLElement) => b.addEventListener('click', () => { this.tab = b.getAttribute('data-tab') as Tab; this.sel = -1; this.render(); }));
    q('[data-sem]').forEach((b: HTMLElement) => b.addEventListener('click', () => { this.semana = new Date(this.semana.getTime() + +(b.getAttribute('data-sem') || 0) * 7 * DIA + 3 * 3600000); this.semana = lunes(this.semana); this.sel = -1; this.tab = this.tab === 'reglas' ? 'semana' : this.tab; this.render(); }));
    q('[data-ir]').forEach((b: HTMLElement) => b.addEventListener('click', () => { this.semana = lunes(new Date(+(b.getAttribute('data-ir') || 0) + 3 * 3600000)); this.tab = 'semana'; this.sel = -1; this.render(); window.scrollTo(0, 0); }));
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
        { n: 'Controlar', label: 'Controlar capacidad', tipo: 'check', valor: p ? p.controlar : true }
      ], 'Guardar', (v: { [k: string]: any }) => { // eslint-disable-line @typescript-eslint/no-explicit-any
        const body = { Title: v.Instalacion + (v.Modulo ? ' · ' + v.Modulo : ''), Instalacion: v.Instalacion, Modulo: v.Modulo, AliasFBPA: v.Instalacion, Disponibilidad: v.Disponibilidad, Controlar: v.Controlar, Orden: p ? p.orden : this.pistas.length + 1 };
        return spPost(c, porTitulo(w, L.pistas) + '/items' + (p ? '(' + p.id + ')' : ''), body, !!p).then(() => { aviso('Pista guardada'); recargar(); });
      }, 520);
    };
    q('[data-pnew]').forEach((b: HTMLElement) => b.addEventListener('click', () => formPista()));
    q('[data-ped]').forEach((b: HTMLElement) => b.addEventListener('click', () => formPista(this.pistas.filter((p: IPista) => p.id === +(b.getAttribute('data-ped') || 0))[0])));

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
      PropertyPaneSlider('margen', { label: 'Margen entre partidos en la misma pista (min)', min: 0, max: 60, step: 5 })
    ] }] }] };
  }
}
