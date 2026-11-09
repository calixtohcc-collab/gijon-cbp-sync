import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { cargarEstilos, cargarEventos, cargarEquipos, IEvento, IEquipo, esc, icon, CATEGORIAS, MESES_LARGO, MESES, DIAS_LARGO, claveDia, mismoDia, hora, categoriaDe } from '../../common/cbp';
import { filaEvento, chips } from '../../common/render';
import { cargarEstilosUi, descargar, generarIcs } from '../../common/ui';
import { mostrarEvento, mostrarDia, enlazarEventos, nuevoEvento } from '../../common/evento';

export interface ICalendarioCbpWebPartProps { titulo: string; }
type Vista = 'proximos' | 'semana' | 'mes' | 'resultados';

interface IBloque { inicio: Date; fin: Date; titulo: string; sub: string; cls: string; ev?: IEvento; }

const DIAMAP: { [k: string]: number } = { lunes: 1, martes: 2, 'miércoles': 3, miercoles: 3, jueves: 4, viernes: 5, 'sábado': 6, sabado: 6, domingo: 0 };
const H0 = 8; const H1 = 23; const PX = 44;

export default class CalendarioCbpWebPart extends BaseClientSideWebPart<ICalendarioCbpWebPartProps> {
  private eventos: IEvento[] = [];
  private equipos: IEquipo[] = [];
  private cargado: boolean = false;
  private vista: Vista = 'proximos';
  private cat: string = 'Todos';
  private equipo: string = '';
  private mes: Date = new Date();
  private lunes: Date = new Date();
  private verEntrenos: boolean = true;

  protected onInit(): Promise<void> {
    cargarEstilos(); cargarEstilosUi();
    this.mes = new Date(); this.mes.setDate(1);
    this.lunes = this.inicioSemana(new Date());
    return this.recargar();
  }

  private recargar(): Promise<void> {
    const c = this.context.spHttpClient; const w = this.context.pageContext.web.absoluteUrl;
    return Promise.all([cargarEventos(c, w).catch(() => []), cargarEquipos(c, w).catch(() => [])])
      .then((r: [IEvento[], IEquipo[]]) => { this.eventos = r[0]; this.equipos = r[1]; this.cargado = true; });
  }

  private inicioSemana(d: Date): Date {
    const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    x.setDate(x.getDate() - ((x.getDay() + 6) % 7));
    return x;
  }

  private pasaFiltro(equipo: string): boolean {
    if (this.equipo) { return equipo === this.equipo || equipo === 'Todo el club'; }
    const c = categoriaDe(equipo);
    return this.cat === 'Todos' || c === this.cat || c === 'Club';
  }

  public render(): void {
    if (!this.cargado) { this.domElement.innerHTML = '<div class="cbp"><div class="cbp-cargando">Cargando calendario…</div></div>'; return; }
    const w = this.context.pageContext.web.absoluteUrl;
    const lista = this.eventos.filter((e: IEvento) => this.pasaFiltro(e.equipo));
    const ahora = new Date();
    const cats = ['Todos'].concat(CATEGORIAS.filter((c: string) => this.eventos.some((e: IEvento) => e.categoria === c) || this.equipos.some((q: IEquipo) => q.categoria === c)));
    const equipos: string[] = this.equipos.map((q: IEquipo) => q.nombre);
    const equiposCat = equipos.filter((n: string) => this.cat === 'Todos' || categoriaDe(n) === this.cat);

    let cuerpo = '';
    if (this.vista === 'proximos') {
      const desde = new Date(ahora.getTime() - 2 * 3600 * 1000);
      const prox = lista.filter((e: IEvento) => e.fecha >= desde).slice(0, 60);
      cuerpo = prox.length ? this.agrupado(prox, w, false) : '<div class="cbp-vacio">No hay próximos eventos con este filtro.</div>';
    } else if (this.vista === 'resultados') {
      const res = lista.filter((e: IEvento) => e.pf !== undefined).reverse();
      const v = res.filter((e: IEvento) => (e.pf as number) > (e.pa as number)).length;
      cuerpo = res.length ? '<p style="margin:0;font-weight:600">' + res.length + ' partidos jugados · ' + v + ' victorias · ' + (res.length - v) + ' derrotas</p>' + this.agrupado(res, w, true)
        : '<div class="cbp-vacio">Todavía no hay resultados con este filtro.</div>';
    } else if (this.vista === 'semana') {
      cuerpo = this.semanaHtml(lista, ahora);
    } else {
      cuerpo = this.mesHtml(lista, ahora);
    }

    this.domElement.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:24px;gap:18px">' +
      '<div class="cbp-sec-head"><h2 class="cbp-h2">' + esc(this.properties.titulo || 'Calendario') + '</h2>' +
      '<div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">' +
      '<div class="cbp-tabs" role="tablist" aria-label="Vista">' +
      this.tab('proximos', 'Próximos') + this.tab('semana', 'Semana') + this.tab('mes', 'Mes') + this.tab('resultados', 'Resultados') + '</div>' +
      '<button type="button" class="cbp-btn pri" data-nuevo="1">+ Nuevo evento</button></div></div>' +
      '<div class="cbp-toolbar"><div class="cbp-chips" role="group" aria-label="Categoría">' + chips(cats, this.equipo ? '' : this.cat, 'cat') + '</div>' +
      '<div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap"><label style="display:flex;gap:8px;align-items:center;font-size:13px;font-weight:600">Equipo <select class="cbp-select" data-equipo="1"><option value="">Todos los equipos</option>' +
      equiposCat.map((n: string) => '<option value="' + esc(n) + '"' + (n === this.equipo ? ' selected' : '') + '>' + esc(n) + '</option>').join('') + '</select></label>' +
      '<button type="button" class="cbp-btn sm" data-ics="1" title="Descarga los partidos y eventos del filtro actual para importarlos en Outlook, Google Calendar o el móvil">Descargar .ics</button>' +
      '<a class="cbp-btn sm pri" href="https://calixtohcc-collab.github.io/gijon-cbp-sync/" target="_blank" rel="noopener" title="Página para suscribirse al calendario de cada equipo desde el móvil. Se actualiza sola cada día con la FBPA. Se puede compartir con las familias." style="color:#FF9F00">Calendario en el móvil</a></div></div>' +
      cuerpo + '</div></div>';

    this.enlazar(lista);
  }

  private tab(v: Vista, t: string): string {
    return '<button type="button" role="tab" class="cbp-tab" data-vista="' + v + '" aria-selected="' + (this.vista === v ? 'true' : 'false') + '">' + t + '</button>';
  }

  private agrupado(evs: IEvento[], w: string, res: boolean): string {
    let html = '<div class="cbp-card cbp-lista" style="overflow:hidden">'; let k = '';
    evs.forEach((e: IEvento) => {
      const kk = claveDia(e.fecha);
      if (kk !== k) { k = kk; html += '<div class="cbp-diahead">' + DIAS_LARGO[e.fecha.getDay()] + ' ' + e.fecha.getDate() + ' de ' + MESES_LARGO[e.fecha.getMonth()] + '</div>'; }
      html += filaEvento(e, w, res);
    });
    return html + '</div>';
  }

  /* ---------- Semana ---------- */
  private entrenosSemana(lunes: Date): IBloque[] {
    const out: IBloque[] = [];
    const ini = new Date(2026, 8, 1); const fin = new Date(2027, 5, 15);
    this.equipos.forEach((q: IEquipo) => {
      if (!this.pasaFiltro(q.nombre)) { return; }
      q.horario.split(/\n/).forEach((l: string) => {
        const m = l.trim().match(/^(\S+)\s+(\d{1,2}):(\d{2})\s*[–-]\s*(\d{1,2}):(\d{2})(.*)$/);
        if (!m) { return; }
        const dow = DIAMAP[m[1].toLowerCase()]; if (dow === undefined) { return; }
        const d = new Date(lunes); d.setDate(lunes.getDate() + ((dow + 6) % 7));
        if (d < ini || d > fin) { return; }
        const i = new Date(d.getFullYear(), d.getMonth(), d.getDate(), +m[2], +m[3]);
        const f = new Date(d.getFullYear(), d.getMonth(), d.getDate(), +m[4], +m[5]);
        const lugar = (m[6].split('·')[1] || '').replace(/\(.*\)/, '').trim();
        out.push({ inicio: i, fin: f, titulo: q.nombre, sub: hora(i) + ' · ' + lugar, cls: 'ent' });
      });
    });
    return out;
  }

  private conEntrenos(): boolean { return !!this.equipo || this.cat !== 'Todos'; }

  private semanaHtml(lista: IEvento[], hoy: Date): string {
    const lunes = this.lunes; const dom = new Date(lunes); dom.setDate(lunes.getDate() + 7);
    const bloques: IBloque[] = lista.filter((e: IEvento) => e.fecha >= lunes && e.fecha < dom).map((e: IEvento): IBloque => {
      const dur = e.tipo === 'Partido' ? 90 : 60;
      const cls = e.tipo === 'Partido' ? (e.local ? 'casa' : 'fuera') : e.tipo === 'Torneo' ? 'tor' : 'otro';
      return { inicio: e.fecha, fin: new Date(e.fecha.getTime() + dur * 60000), titulo: e.tipo === 'Partido' ? e.equipo : e.titulo,
        sub: hora(e.fecha) + (e.tipo === 'Partido' ? ' · vs ' + e.rival : ''), cls, ev: e };
    }).concat(this.verEntrenos && this.conEntrenos() ? this.entrenosSemana(lunes) : []);

    let cols = '';
    for (let i = 0; i < 7; i++) {
      const d = new Date(lunes); d.setDate(lunes.getDate() + i);
      const delDia = bloques.filter((b: IBloque) => mismoDia(b.inicio, d)).sort((a: IBloque, b: IBloque) => a.inicio.getTime() - b.inicio.getTime());
      // Reparto en carriles por grupos de solape
      const carril: number[] = []; const grupo: number[] = []; const nGrupo: number[] = [];
      let g = -1; let finGrupo = 0; let finCarril: number[] = [];
      delDia.forEach((b: IBloque, j: number) => {
        if (b.inicio.getTime() >= finGrupo) { g++; finCarril = []; nGrupo[g] = 0; }
        let c = 0; while (finCarril[c] !== undefined && finCarril[c] > b.inicio.getTime()) { c++; }
        carril[j] = c; finCarril[c] = b.fin.getTime(); grupo[j] = g;
        nGrupo[g] = Math.max(nGrupo[g], finCarril.length); finGrupo = Math.max(finGrupo, b.fin.getTime());
      });
      cols += '<div class="cbp-sem-col" style="height:' + ((H1 - H0) * PX) + 'px">' + delDia.map((b: IBloque, j: number) => {
        const top = Math.max(0, (b.inicio.getHours() + b.inicio.getMinutes() / 60 - H0) * PX);
        const alto = Math.max(22, (b.fin.getTime() - b.inicio.getTime()) / 3600000 * PX - 2);
        const ancho = 100 / Math.max(1, nGrupo[grupo[j]]);
        return '<button type="button" class="cbp-blk ' + b.cls + '"' + (b.ev ? ' data-evid="' + b.ev.id + '"' : ' tabindex="-1"') +
          ' style="top:' + top + 'px;height:' + alto + 'px;left:calc(' + (carril[j] * ancho) + '% + 2px);width:calc(' + ancho + '% - 4px)" title="' + esc(b.titulo + ' · ' + b.sub) + '">' +
          '<b>' + esc(b.titulo) + '</b><span>' + esc(b.sub) + '</span></button>';
      }).join('') + '</div>';
    }
    let cab = '<div class="cbp-sem-h"></div>';
    for (let i = 0; i < 7; i++) {
      const d = new Date(lunes); d.setDate(lunes.getDate() + i);
      cab += '<div class="cbp-sem-h' + (mismoDia(d, hoy) ? ' hoy' : '') + '">' + ['LUN', 'MAR', 'MIÉ', 'JUE', 'VIE', 'SÁB', 'DOM'][i] + ' ' + d.getDate() + '</div>';
    }
    let horas = '<div class="cbp-sem-horas" style="height:' + ((H1 - H0) * PX) + 'px">';
    for (let h = H0 + 1; h < H1; h++) { horas += '<span style="top:' + ((h - H0) * PX) + 'px">' + h + ':00</span>'; }
    horas += '</div>';
    const fin = new Date(lunes); fin.setDate(lunes.getDate() + 6);
    const rango = lunes.getDate() + ' ' + MESES[lunes.getMonth()].toLowerCase() + ' – ' + fin.getDate() + ' ' + MESES[fin.getMonth()].toLowerCase();
    return '<div class="cbp-toolbar"><div class="cbp-mes-head">' +
      '<button type="button" class="cbp-nav" data-sem="-1" aria-label="Semana anterior">' + icon('flechaIzq', 18) + '</button>' +
      '<h3>' + rango + '</h3>' +
      '<button type="button" class="cbp-nav" data-sem="1" aria-label="Semana siguiente">' + icon('flechaDer', 18) + '</button>' +
      '<button type="button" class="cbp-btn sm" data-sem="0">Hoy</button></div>' +
      '<div class="cbp-leyenda"><label style="display:inline-flex;gap:6px;align-items:center;font-weight:600;color:#090909"><input type="checkbox" data-entrenos="1"' + (this.verEntrenos ? ' checked' : '') + '> Mostrar entrenamientos habituales</label>' +
      '<span><i style="background:#FF9F00"></i>Partido en casa</span><span><i style="background:#fff;border:1.5px solid #FF9F00"></i>Fuera</span><span><i style="background:#ECEAE4"></i>Entrenamiento</span></div></div>' +
      '<div class="cbp-scroll"><div class="cbp-sem">' + cab + horas + cols + '</div></div>' +
      '<p style="margin:0;font-size:12px;color:#5c5c5c">' + (this.conEntrenos() ? 'Los entrenamientos se muestran según el horario habitual de cada equipo (lista Equipos). ' : 'Elige una categoría o un equipo para ver también sus entrenamientos. ') + 'Pulsa un partido para ver su ficha.</p>';
  }

  /* ---------- Mes ---------- */
  private mesHtml(lista: IEvento[], hoy: Date): string {
    const y = this.mes.getFullYear(); const m = this.mes.getMonth();
    const primero = new Date(y, m, 1);
    const offset = (primero.getDay() + 6) % 7;
    const inicio = new Date(y, m, 1 - offset);
    const porDia: { [k: string]: IEvento[] } = {};
    lista.forEach((e: IEvento) => { const k = claveDia(e.fecha); (porDia[k] = porDia[k] || []).push(e); });
    let celdas = ['LUN', 'MAR', 'MIÉ', 'JUE', 'VIE', 'SÁB', 'DOM'].map((d: string) => '<div class="cbp-dsem">' + d + '</div>').join('');
    for (let i = 0; i < 42; i++) {
      const d = new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate() + i);
      if (i >= 35 && d.getMonth() !== m) { break; }
      const evs = porDia[claveDia(d)] || [];
      const cls = 'cbp-dia' + (d.getMonth() !== m ? ' fuera' : '') + (mismoDia(d, hoy) ? ' hoy' : '');
      const max = 3;
      celdas += '<div class="' + cls + '" data-dia="' + d.getTime() + '" role="button" tabindex="0" aria-label="' + d.getDate() + ' de ' + MESES_LARGO[d.getMonth()] + ', ' + evs.length + ' eventos"><span class="cbp-num">' + d.getDate() + '</span>' +
        evs.slice(0, max).map((e: IEvento) => {
          const c = e.tipo === 'Torneo' ? ' torneo' : (e.tipo === 'Partido' && e.local ? ' casa' : '');
          const t = hora(e.fecha) + ' ' + (e.tipo === 'Partido' ? e.equipo + ' vs ' + e.rival : e.titulo);
          return '<span class="cbp-ev' + c + '" title="' + esc(t) + '" style="display:block">' + esc(t) + '</span>';
        }).join('') + (evs.length > max ? '<span class="cbp-mas">+' + (evs.length - max) + ' más</span>' : '') + '</div>';
    }
    return '<div class="cbp-toolbar"><div class="cbp-mes-head">' +
      '<button type="button" class="cbp-nav" data-mes="-1" aria-label="Mes anterior">' + icon('flechaIzq', 18) + '</button>' +
      '<h3>' + MESES_LARGO[m] + ' ' + y + '</h3>' +
      '<button type="button" class="cbp-nav" data-mes="1" aria-label="Mes siguiente">' + icon('flechaDer', 18) + '</button></div>' +
      '<div class="cbp-leyenda"><span><i style="background:#FF9F00"></i>En casa</span><span><i style="background:#CFCABF"></i>Fuera / otros</span><span><i style="background:#090909"></i>Torneo</span><span>Pulsa un día para ver todos sus eventos</span></div></div>' +
      '<div class="cbp-mes">' + celdas + '</div>';
  }

  private enlazar(lista: IEvento[]): void {
    const w = this.context.pageContext.web.absoluteUrl;
    const q = (s: string): NodeListOf<Element> => this.domElement.querySelectorAll(s);
    const on = (s: string, fn: (el: HTMLElement) => void): void => {
      const els = q(s);
      for (let i = 0; i < els.length; i++) { els[i].addEventListener('click', (ev: Event) => fn(ev.currentTarget as HTMLElement)); }
    };
    on('[data-vista]', (el: HTMLElement) => { this.vista = el.getAttribute('data-vista') as Vista; this.render(); });
    on('[data-cat]', (el: HTMLElement) => { this.cat = el.getAttribute('data-cat') || 'Todos'; this.equipo = ''; this.render(); });
    on('[data-mes]', (el: HTMLElement) => { const n = parseInt(el.getAttribute('data-mes') || '0', 10); this.mes = new Date(this.mes.getFullYear(), this.mes.getMonth() + n, 1); this.render(); });
    on('[data-sem]', (el: HTMLElement) => {
      const n = parseInt(el.getAttribute('data-sem') || '0', 10);
      if (n === 0) { this.lunes = this.inicioSemana(new Date()); } else { const d = new Date(this.lunes); d.setDate(d.getDate() + 7 * n); this.lunes = d; }
      this.render();
    });
    on('[data-nuevo]', () => nuevoEvento(this.context.spHttpClient, w, undefined, this.equipo, () => { this.recargar().then(() => this.render()); }));
    on('[data-ics]', () => {
      const nombre = 'Gijón CBP – ' + (this.equipo || (this.cat === 'Todos' ? 'Todo el club' : this.cat));
      descargar((this.equipo || this.cat).toLowerCase().replace(/[^a-z0-9]+/g, '-') + '.ics', generarIcs(lista, nombre), 'text/calendar');
    });
    const dias = q('[data-dia]');
    for (let i = 0; i < dias.length; i++) {
      const abrir = (ev: Event): void => {
        const d = new Date(parseInt((ev.currentTarget as HTMLElement).getAttribute('data-dia') || '0', 10));
        mostrarDia(d, lista.filter((e: IEvento) => mismoDia(e.fecha, d)), w);
      };
      dias[i].addEventListener('click', abrir);
      dias[i].addEventListener('keydown', (ev: Event) => { const k = (ev as KeyboardEvent).key; if (k === 'Enter' || k === ' ') { ev.preventDefault(); abrir(ev); } });
    }
    enlazarEventos(this.domElement, lista, w);
    const chk = this.domElement.querySelector('[data-entrenos]') as HTMLInputElement;
    if (chk) { chk.addEventListener('change', () => { this.verEntrenos = chk.checked; this.render(); }); }
    const sel = this.domElement.querySelector('[data-equipo]') as HTMLSelectElement;
    if (sel) { sel.addEventListener('change', () => { this.equipo = sel.value; this.render(); }); }
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: 'Calendario Gijón CBP' }, groups: [{ groupName: 'General', groupFields: [
      PropertyPaneTextField('titulo', { label: 'Título' })
    ] }] }] };
  }
}
