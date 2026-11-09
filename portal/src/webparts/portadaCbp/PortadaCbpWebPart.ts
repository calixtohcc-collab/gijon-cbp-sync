import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { cargarEstilos, cargarEventos, cargarAvisos, cargarEquipos, IEvento, IAviso, IEquipo, esc, icon, CATEGORIAS } from '../../common/cbp';
import { filaEvento, filaAviso, chips } from '../../common/render';
import { cargarEstilosUi } from '../../common/ui';
import { enlazarEventos } from '../../common/evento';

export interface IPortadaCbpWebPartProps { mostrarCabecera: boolean; }

export default class PortadaCbpWebPart extends BaseClientSideWebPart<IPortadaCbpWebPartProps> {
  private eventos: IEvento[] = [];
  private avisos: IAviso[] = [];
  private equipos: IEquipo[] = [];
  private filtro: string = 'Todos';
  private cargado: boolean = false;
  private error: string = '';

  protected onInit(): Promise<void> {
    cargarEstilos(); cargarEstilosUi();
    const c = this.context.spHttpClient; const w = this.context.pageContext.web.absoluteUrl;
    return Promise.all([
      cargarEventos(c, w).catch(() => []),
      cargarAvisos(c, w).catch(() => []),
      cargarEquipos(c, w).catch(() => [])
    ]).then((r: [IEvento[], IAviso[], IEquipo[]]) => {
      this.eventos = r[0]; this.avisos = r[1].filter((a: IAviso) => !a.limite || a.limite.getTime() >= new Date().setHours(0, 0, 0, 0)); this.equipos = r[2]; this.cargado = true;
    }).catch((e: Error) => { this.error = e.message; this.cargado = true; });
  }

  public render(): void {
    const w = this.context.pageContext.web.absoluteUrl;
    if (!this.cargado) { this.domElement.innerHTML = '<div class="cbp"><div class="cbp-cargando">Cargando…</div></div>'; return; }
    const ahora = new Date();
    const desde = new Date(ahora.getTime() - 2 * 3600 * 1000);
    const en7 = new Date(ahora.getTime() + 7 * 24 * 3600 * 1000);
    const proximos = this.eventos.filter((e: IEvento) => e.fecha >= desde);
    const filtrados = proximos.filter((e: IEvento) => this.filtro === 'Todos' || e.categoria === this.filtro || e.categoria === 'Club').slice(0, 8);
    const semana = proximos.filter((e: IEvento) => e.tipo === 'Partido' && e.fecha <= en7).length;
    const jugados = this.eventos.filter((e: IEvento) => e.pf !== undefined && e.fecha < ahora);
    const victorias = jugados.filter((e: IEvento) => (e.pf as number) > (e.pa as number)).length;
    const ultimos = jugados.slice(-6).reverse();
    const cats = ['Todos'].concat(CATEGORIAS.filter((c: string) => proximos.some((e: IEvento) => e.categoria === c)));

    const logo = this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, '') + '/SiteAssets/escudo-gijon-cbp-ambar.png';
    const m = ahora.getMonth(); const d = ahora.getDate();
    let fase = 1;
    if (m === 7 || (m === 8 && d < 19)) { fase = 0; } else if (m === 11 && d >= 20) { fase = 2; } else if (m === 5 || m === 6) { fase = 3; }
    const fases = [['Pretemporada', 1], ['Competición', 6], ['Torneo de Navidad', 1.4], ['Fases finales y campus', 1.6]];

    const hero = this.properties.mostrarCabecera === false ? '' :
      '<section class="cbp-hero cbp-dark"><div class="cbp-hero-in">' +
      '<div class="cbp-hero-txt"><span class="cbp-over">Gijón CBP · Portal interno · Temporada 2026–27</span>' +
      '<h1 class="cbp-titular">Nacimos en Pumarín.<br><em>Crecemos con Gijón.</em></h1>' +
      '<p>Partidos, entrenamientos, equipos y materiales para el cuerpo técnico y la junta directiva.</p>' +
      '<div class="cbp-stats">' +
      '<div class="cbp-stat"><b>' + (this.equipos.length || '–') + '</b><span>equipos y grupos</span></div>' +
      '<div class="cbp-stat"><b>' + semana + '</b><span>partidos en 7 días</span></div>' +
      '<div class="cbp-stat"><b>' + victorias + '–' + (jugados.length - victorias) + '</b><span>victorias–derrotas</span></div>' +
      '</div>' +
      '<div class="cbp-fases" role="img" aria-label="Fase actual: ' + fases[fase][0] + '">' +
      fases.map((f: (string | number)[], i: number) => '<div class="cbp-fase' + (i === fase ? ' on' : '') + '" style="flex:' + f[1] + ' 1 0"><i></i><span>' + f[0] + '</span></div>').join('') +
      '</div></div>' +
      '<div class="cbp-hero-logo"><img src="' + esc(logo) + '" alt="Escudo de Gijón CBP"></div>' +
      '</div></section>';

    const accesos: string[][] = [
      ['calendario', 'Calendario', 'Partidos y eventos', w + '/SitePages/Temporada.aspx'],
      ['equipos', 'Equipos', 'Horarios y pistas', w + '/SitePages/Equipos.aspx'],
      ['materiales', 'Materiales', 'Ejercicios y documentos', w + '/SitePages/Materiales.aspx'],
      ['avisos', 'Avisos', 'Comunicados de la junta', w + '/SitePages/Avisos.aspx'],
      ['herramientas', 'Herramientas', 'Pizarra, convocatorias…', w + '/SitePages/Herramientas.aspx'],
      ['junta', 'Junta directiva', 'Acceso restringido', w + '/SitePages/Junta.aspx']
    ];
    const accesosHtml = '<nav class="cbp-accesos" aria-label="Accesos rápidos">' + accesos.map((a: string[]) =>
      '<a class="cbp-acceso" href="' + esc(a[3]) + '"><span class="cbp-acceso-ico">' + icon(a[0], 22) + '</span><span><b>' + a[1] + '</b><small>' + a[2] + '</small></span></a>').join('') + '</nav>';

    const proxHtml = filtrados.length ? '<div class="cbp-card cbp-lista">' + filtrados.map((e: IEvento) => filaEvento(e, w)).join('') + '</div>'
      : '<div class="cbp-vacio">No hay próximos eventos para esta categoría.</div>';
    const avisosHtml = this.avisos.length ? '<div class="cbp-card">' + this.avisos.slice(0, 5).map((a: IAviso) => filaAviso(a, w)).join('') + '</div>'
      : '<div class="cbp-vacio">Sin avisos.</div>';
    const resHtml = ultimos.length ? '<div class="cbp-card cbp-lista">' + ultimos.map((e: IEvento) => filaEvento(e, w, true)).join('') + '</div>'
      : '<div class="cbp-vacio">Todavía no hay resultados.</div>';

    this.domElement.innerHTML = '<div class="cbp">' + hero +
      '<div class="cbp-wrap">' + accesosHtml +
      '<div class="cbp-dos"><section class="cbp-main cbp-sec"><div class="cbp-sec-head"><h2 class="cbp-h2">Próximos partidos</h2><a class="cbp-link" href="' + esc(w + '/SitePages/Temporada.aspx') + '">Calendario completo</a></div>' +
      '<div class="cbp-chips" role="group" aria-label="Filtrar por categoría">' + chips(cats, this.filtro, 'cat') + '</div>' + proxHtml + '</section>' +
      '<aside class="cbp-side cbp-sec"><div class="cbp-sec-head"><h2 class="cbp-h2">Avisos</h2><a class="cbp-link" href="' + esc(w + '/SitePages/Avisos.aspx') + '">Ver todos</a></div>' + avisosHtml + '</aside></div>' +
      '<section class="cbp-sec"><div class="cbp-sec-head"><h2 class="cbp-h2">Últimos resultados</h2></div>' + resHtml + '</section>' +
      (this.error ? '<div class="cbp-vacio">No se pudieron cargar algunos datos.</div>' : '') +
      '</div></div>';

    enlazarEventos(this.domElement, this.eventos, w);
    const botones = this.domElement.querySelectorAll('[data-cat]');
    for (let i = 0; i < botones.length; i++) {
      botones[i].addEventListener('click', (ev: Event) => {
        this.filtro = (ev.currentTarget as HTMLElement).getAttribute('data-cat') || 'Todos';
        this.render();
      });
    }
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: 'Portada Gijón CBP' }, groups: [{ groupName: 'Diseño', groupFields: [
      PropertyPaneToggle('mostrarCabecera', { label: 'Mostrar cabecera negra con escudo', onText: 'Sí', offText: 'No' })
    ] }] }] };
  }
}
