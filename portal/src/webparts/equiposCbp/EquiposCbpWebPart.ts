import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { cargarEstilos, cargarEquipos, cargarEventos, IEquipo, IEvento, esc, CATEGORIAS, MESES, DIAS, hora } from '../../common/cbp';
import { chips } from '../../common/render';
import { cargarEstilosUi } from '../../common/ui';
import { cargarEstilos3 } from '../../common/estilos3';
import { renderFicha } from './ficha';

export interface IEquiposCbpWebPartProps { titulo: string; }

const ABR: { [k: string]: string } = { lunes: 'LUN', martes: 'MAR', 'miércoles': 'MIÉ', miercoles: 'MIÉ', jueves: 'JUE', viernes: 'VIE', 'sábado': 'SÁB', domingo: 'DOM' };

export default class EquiposCbpWebPart extends BaseClientSideWebPart<IEquiposCbpWebPartProps> {
  private equipos: IEquipo[] = [];
  private eventos: IEvento[] = [];
  private cargado: boolean = false;
  private cat: string = 'Todas';
  private rama: string = 'Todas';

  private onHash = (): void => { this.render(); };

  protected onInit(): Promise<void> {
    cargarEstilos(); cargarEstilosUi(); cargarEstilos3();
    window.addEventListener('hashchange', this.onHash);
    const c = this.context.spHttpClient; const w = this.context.pageContext.web.absoluteUrl;
    return Promise.all([cargarEquipos(c, w).catch(() => []), cargarEventos(c, w).catch(() => [])])
      .then((r: [IEquipo[], IEvento[]]) => { this.equipos = r[0]; this.eventos = r[1]; this.cargado = true; });
  }

  public render(): void {
    if (!this.cargado) { this.domElement.innerHTML = '<div class="cbp"><div class="cbp-cargando">Cargando equipos…</div></div>'; return; }
    const w = this.context.pageContext.web.absoluteUrl;
    const hm = window.location.hash.match(/^#equipo=(\d+)/);
    const sel = hm ? this.equipos.filter((x: IEquipo) => x.id === +hm[1])[0] : undefined;
    if (sel) {
      renderFicha(this.domElement, {
        client: this.context.spHttpClient, w, equipo: sel, eventos: this.eventos,
        volver: () => { history.pushState(null, '', window.location.pathname + window.location.search); this.render(); window.scrollTo(0, 0); },
        alCambiarEquipo: (n: IEquipo) => { this.equipos = this.equipos.map((x: IEquipo) => (x.id === n.id ? n : x)); this.render(); }
      });
      return;
    }
    const ahora = new Date();
    const cats = ['Todas'].concat(CATEGORIAS.filter((c: string) => this.equipos.some((e: IEquipo) => e.categoria === c)));
    const ramas = ['Todas', 'Femenino', 'Masculino', 'Mixto'].filter((r: string) => r === 'Todas' || this.equipos.some((e: IEquipo) => e.rama === r));
    const lista = this.equipos.filter((e: IEquipo) => (this.cat === 'Todas' || e.categoria === this.cat) && (this.rama === 'Todas' || e.rama === this.rama));

    const tarjetas = lista.map((e: IEquipo) => {
      const sesiones = e.horario.split(/\n/).filter((l: string) => l.trim()).map((l: string) => {
        const m = l.match(/^(\S+)\s+(.*)$/);
        const dia = m ? (ABR[m[1].toLowerCase()] || '') : '';
        const resto = dia && m ? m[2] : l;
        return '<li><b>' + (dia || '·') + '</b><span>' + esc(resto) + '</span></li>';
      }).join('');
      const prox = this.eventos.filter((v: IEvento) => v.equipo === e.nombre && v.fecha >= ahora)[0];
      const proxHtml = prox ? '<div class="cbp-prox"><small>Próximo</small><a href="' + esc(w + '/Lists/Calendario/DispForm.aspx?ID=' + prox.id) + '" style="text-decoration:none;font-weight:700">' +
        DIAS[prox.fecha.getDay()].toLowerCase() + ' ' + prox.fecha.getDate() + ' ' + MESES[prox.fecha.getMonth()].toLowerCase() + ' · ' + hora(prox.fecha) + ' · ' +
        esc(prox.tipo === 'Partido' ? (prox.local ? 'vs ' : 'en ') + prox.rival : prox.titulo) + '</a></div>' : '';
      return '<article class="cbp-card cbp-equipo"><div class="cbp-equipo-head"><span class="cbp-over">' + esc(e.categoria) + ' · ' + esc(e.rama) + '</span>' +
        '<h3><a href="#equipo=' + e.id + '" data-ficha="' + e.id + '" style="text-decoration:none;color:inherit">' + esc(e.nombre) + '</a></h3></div>' +
        '<div class="cbp-equipo-body"><ul class="cbp-sesiones" aria-label="Horario de entrenamiento">' + sesiones + '</ul>' +
        (e.competicion ? '<span class="cbp-comp">' + esc(e.competicion) + '</span>' : '') + proxHtml +
        '<a class="cbp-btn sm" href="#equipo=' + e.id + '" data-ficha="' + e.id + '" style="align-self:flex-start;margin-top:auto">Ver ficha y plantilla →</a></div></article>';
    }).join('');

    this.domElement.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:24px;gap:20px">' +
      '<div class="cbp-sec-head"><h2 class="cbp-h2">' + esc(this.properties.titulo || 'Equipos') + '</h2><span style="font-size:14px;color:#5c5c5c">' + lista.length + ' de ' + this.equipos.length + '</span></div>' +
      '<div class="cbp-toolbar"><div class="cbp-chips" role="group" aria-label="Categoría">' + chips(cats, this.cat, 'cat') + '</div>' +
      '<div class="cbp-chips" role="group" aria-label="Rama">' + chips(ramas, this.rama, 'rama') + '</div></div>' +
      (lista.length ? '<div class="cbp-equipos">' + tarjetas + '</div>' : '<div class="cbp-vacio">No hay equipos con este filtro.</div>') +
      '</div></div>';

    const bind = (sel: string, fn: (v: string) => void): void => {
      const els = this.domElement.querySelectorAll(sel);
      for (let i = 0; i < els.length; i++) { els[i].addEventListener('click', (ev: Event) => fn((ev.currentTarget as HTMLElement).getAttribute(sel.slice(1, -1)) || '')); }
    };
    bind('[data-cat]', (v: string) => { this.cat = v; this.render(); });
    bind('[data-rama]', (v: string) => { this.rama = v; this.render(); });
    const fichas = this.domElement.querySelectorAll('[data-ficha]');
    for (let i = 0; i < fichas.length; i++) {
      fichas[i].addEventListener('click', (ev: Event) => { ev.preventDefault(); history.pushState(null, '', '#equipo=' + (ev.currentTarget as HTMLElement).getAttribute('data-ficha')); this.render(); window.scrollTo(0, 0); });
    }
  }

  protected onDispose(): void { window.removeEventListener('hashchange', this.onHash); }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: 'Equipos Gijón CBP' }, groups: [{ groupName: 'General', groupFields: [
      PropertyPaneTextField('titulo', { label: 'Título' })
    ] }] }] };
  }
}
