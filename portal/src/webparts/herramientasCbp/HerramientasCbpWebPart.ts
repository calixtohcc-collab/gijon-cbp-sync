import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { cargarEstilos, cargarEventos, IEvento, icon } from '../../common/cbp';
import { cargarEstilosUi } from '../../common/ui';
import { ICtx, IJugador, cargarJugadores } from './datos';
import { renderConvocatorias, preseleccionarPartido } from './convocatorias';
import { renderAsistencia } from './asistencia';
import { renderPizarra } from './pizarra';
import { renderSesiones } from './sesiones';
import { renderPlantillas } from './plantillas';

export interface IHerramientasCbpWebPartProps { inicial: string; }

const TABS: { k: string; t: string; ico: string; d: string }[] = [
  { k: 'convocatorias', t: 'Convocatorias', ico: 'equipos', d: 'Elige quién juega cada partido y compártelo por WhatsApp.' },
  { k: 'asistencia', t: 'Asistencia', ico: 'calendario', d: 'Pasa lista en cada entrenamiento y consulta los porcentajes.' },
  { k: 'pizarra', t: 'Pizarra táctica', ico: 'herramientas', d: 'Dibuja jugadas y ejercicios y guárdalos en Materiales.' },
  { k: 'sesiones', t: 'Sesiones', ico: 'materiales', d: 'Planifica entrenamientos por bloques con ejercicios de la biblioteca.' },
  { k: 'plantillas', t: 'Plantillas', ico: 'junta', d: 'Jugadores y jugadoras de cada equipo.' }
];

export default class HerramientasCbpWebPart extends BaseClientSideWebPart<IHerramientasCbpWebPartProps> {
  private tab: string = 'convocatorias';
  private eventos: IEvento[] = [];
  private jugadores: IJugador[] = [];
  private equipo: string = '';
  private cargado: boolean = false;

  protected onInit(): Promise<void> {
    cargarEstilos(); cargarEstilosUi();
    this.tab = this.properties.inicial || 'convocatorias';
    const h = window.location.hash.replace('#', '');
    const m = h.match(/^(\w+)(?:=(\d+))?$/);
    if (m && TABS.some((t: { k: string }) => t.k === m[1])) { this.tab = m[1]; if (m[1] === 'convocatorias' && m[2]) { preseleccionarPartido(+m[2]); } }
    try { this.equipo = window.localStorage.getItem('cbp-equipo') || ''; } catch (e) { this.equipo = ''; }
    const c = this.context.spHttpClient; const w = this.context.pageContext.web.absoluteUrl;
    return Promise.all([cargarEventos(c, w).catch(() => []), cargarJugadores(c, w).catch(() => [])])
      .then((r: [IEvento[], IJugador[]]) => { this.eventos = r[0]; this.jugadores = r[1]; this.cargado = true; });
  }

  private ctx(): ICtx {
    return {
      client: this.context.spHttpClient, w: this.context.pageContext.web.absoluteUrl, eventos: this.eventos, jugadores: this.jugadores, equipo: this.equipo,
      setEquipo: (e: string) => { this.equipo = e; try { window.localStorage.setItem('cbp-equipo', e); } catch (x) { /* sin almacenamiento */ } },
      recargarJugadores: () => cargarJugadores(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl).then((j: IJugador[]) => { this.jugadores = j; }),
      irA: (t: string) => { this.tab = t; this.render(); }
    };
  }

  public render(): void {
    if (!this.cargado) { this.domElement.innerHTML = '<div class="cbp"><div class="cbp-cargando">Cargando herramientas…</div></div>'; return; }
    const t = TABS.filter((x: { k: string }) => x.k === this.tab)[0] || TABS[0];
    this.domElement.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:24px;gap:18px">' +
      '<div class="cbp-sec-head"><h2 class="cbp-h2">Herramientas</h2></div>' +
      '<div class="cbp-tooltabs" role="tablist" aria-label="Herramientas">' + TABS.map((x: { k: string; t: string; ico: string }) =>
        '<button type="button" role="tab" class="cbp-tooltab" data-tab="' + x.k + '" aria-selected="' + (x.k === t.k ? 'true' : 'false') + '">' + icon(x.ico, 18) + x.t + '</button>').join('') + '</div>' +
      '<p style="margin:0;color:#5c5c5c">' + t.d + '</p>' +
      '<section class="cbp-card cbp-panel" role="tabpanel" data-panel="1"></section></div></div>';
    const tabs = this.domElement.querySelectorAll('[data-tab]');
    for (let i = 0; i < tabs.length; i++) {
      tabs[i].addEventListener('click', (e: Event) => {
        this.tab = (e.currentTarget as HTMLElement).getAttribute('data-tab') || 'convocatorias';
        try { window.history.replaceState(null, '', '#' + this.tab); } catch (x) { /* nada */ }
        this.render();
      });
    }
    const panel = this.domElement.querySelector('[data-panel]') as HTMLElement;
    const ctx = this.ctx();
    if (t.k === 'convocatorias') { renderConvocatorias(panel, ctx); }
    else if (t.k === 'asistencia') { renderAsistencia(panel, ctx); }
    else if (t.k === 'pizarra') { renderPizarra(panel, ctx); }
    else if (t.k === 'sesiones') { renderSesiones(panel, ctx); }
    else { renderPlantillas(panel, ctx); }
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: 'Herramientas Gijón CBP' }, groups: [{ groupName: 'General', groupFields: [
      PropertyPaneDropdown('inicial', { label: 'Herramienta inicial', options: TABS.map((x: { k: string; t: string }) => ({ key: x.k, text: x.t })) })
    ] }] }] };
  }
}
