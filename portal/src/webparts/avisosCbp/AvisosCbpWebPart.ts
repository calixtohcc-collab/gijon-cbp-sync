import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { cargarEstilos, esc, MESES_LARGO } from '../../common/cbp';
import { chips } from '../../common/render';
import { cargarEstilosUi, spGet, spPost, listUrl, abrirModal, cerrarModal, aviso, fechaInput } from '../../common/ui';
import { cargarEstilos3, enlazarUrls } from '../../common/estilos3';

export interface IAvisosCbpWebPartProps { titulo: string; }

interface IAv { id: number; titulo: string; detalle: string; limite?: Date; importante: boolean; creado: Date; autor: string; auto: boolean; }

const FILTROS = ['Vigentes', 'Importantes', 'Caducados', 'Todos'];
const DIA = 86400000;

function inicioDia(d: Date): number { return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime(); }
function fechaLarga(d: Date): string { return d.getDate() + ' de ' + MESES_LARGO[d.getMonth()]; }
function haceTiempo(d: Date): string {
  const n = Math.round((inicioDia(new Date()) - inicioDia(d)) / DIA);
  if (n <= 0) { return 'hoy'; } if (n === 1) { return 'ayer'; } if (n < 30) { return 'hace ' + n + ' días'; }
  return 'el ' + fechaLarga(d);
}

export default class AvisosCbpWebPart extends BaseClientSideWebPart<IAvisosCbpWebPartProps> {
  private avisos: IAv[] = [];
  private cargado: boolean = false;
  private filtro: string = 'Vigentes';

  protected onInit(): Promise<void> {
    cargarEstilos(); cargarEstilosUi(); cargarEstilos3();
    return this.cargar();
  }

  private cargar(): Promise<void> {
    const w = this.context.pageContext.web.absoluteUrl;
    return spGet(this.context.spHttpClient, listUrl(w, 'Lists/Avisos') + '/items?$select=Id,Title,Detalle,FechaLimite,Prioridad,Created,Author/Title&$expand=Author&$orderby=Created desc&$top=300')
      .then((j: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
        this.avisos = (j.value || []).map((i: any): IAv => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
          id: i.Id, titulo: i.Title || '', detalle: i.Detalle || '', limite: i.FechaLimite ? new Date(i.FechaLimite) : undefined,
          importante: i.Prioridad === 'Importante', creado: new Date(i.Created), autor: i.Author ? i.Author.Title || '' : '',
          auto: /^Cambio:/.test(i.Title || '')
        }));
        this.cargado = true;
      }).catch(() => { this.cargado = true; });
  }

  private caducado(a: IAv): boolean { return !!a.limite && inicioDia(a.limite) < inicioDia(new Date()); }

  private tarjeta(a: IAv): string {
    const cad = this.caducado(a);
    let top = '';
    if (a.importante && !cad) { top += '<span class="cbp-pill imp">Importante</span>'; }
    if (a.auto) { top += '<span class="cbp-pill auto">Automático · FBPA</span>'; }
    if (a.limite) {
      const d = Math.round((inicioDia(a.limite) - inicioDia(new Date())) / DIA);
      if (cad) { top += '<span class="cbp-pill">Caducó el ' + fechaLarga(a.limite) + '</span>'; }
      else if (d === 0) { top += '<span class="cbp-pill urg">Vence hoy</span>'; }
      else if (d <= 3) { top += '<span class="cbp-pill urg">Quedan ' + d + ' día' + (d === 1 ? '' : 's') + '</span>'; }
      else { top += '<span class="cbp-pill">Hasta el ' + fechaLarga(a.limite) + '</span>'; }
    }
    const autor = a.auto ? 'Sincronización FBPA' : (a.autor || '');
    return '<article class="cbp-av' + (a.importante && !cad ? ' imp' : '') + (cad ? ' cad' : '') + '" id="aviso-' + a.id + '">' +
      (top ? '<div class="cbp-av-top">' + top + '</div>' : '') +
      '<h3>' + esc(a.titulo) + '</h3>' + (a.detalle ? '<p>' + enlazarUrls(esc(a.detalle)) + '</p>' : '') +
      '<div class="cbp-av-pie"><span>Publicado ' + haceTiempo(a.creado) + (autor ? ' · ' + esc(autor) : '') + '</span>' +
      '<span class="cbp-av-acc"><button type="button" data-ed="' + a.id + '">Editar</button><button type="button" data-del="' + a.id + '">Eliminar</button></span></div></article>';
  }

  public render(): void {
    if (!this.cargado) { this.domElement.innerHTML = '<div class="cbp"><div class="cbp-cargando">Cargando avisos…</div></div>'; return; }
    const vig = this.avisos.filter((a: IAv) => !this.caducado(a));
    const imp = vig.filter((a: IAv) => a.importante);
    let lista: IAv[];
    if (this.filtro === 'Vigentes') { lista = vig; }
    else if (this.filtro === 'Importantes') { lista = imp; }
    else if (this.filtro === 'Caducados') { lista = this.avisos.filter((a: IAv) => this.caducado(a)); }
    else { lista = this.avisos; }
    // Importantes primero; luego por fecha límite más cercana; luego más recientes
    lista = lista.slice().sort((a: IAv, b: IAv) => {
      const ca = this.caducado(a); const cb = this.caducado(b);
      if (ca !== cb) { return ca ? 1 : -1; }
      if (a.importante !== b.importante) { return a.importante ? -1 : 1; }
      const ta = a.limite && !ca ? a.limite.getTime() : 9e15; const tb = b.limite && !cb ? b.limite.getTime() : 9e15;
      return ta - tb || b.creado.getTime() - a.creado.getTime();
    });
    const pronto = vig.filter((a: IAv) => a.limite && (inicioDia(a.limite) - inicioDia(new Date())) / DIA <= 7).length;

    const dest = '<section class="cbp-av-dest"><span class="num">' + vig.length + '</span><div><h3>Aviso' + (vig.length === 1 ? '' : 's') + ' en vigor</h3>' +
      '<p>' + (imp.length ? imp.length + ' importante' + (imp.length === 1 ? '' : 's') : 'Ninguno marcado como importante') +
      (pronto ? ' · ' + pronto + ' vence' + (pronto === 1 ? '' : 'n') + ' en los próximos 7 días' : '') +
      '. Los cambios de horario que publica la FBPA aparecen aquí solos.</p></div></section>';

    this.domElement.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:24px;gap:18px">' +
      '<div class="cbp-sec-head"><h2 class="cbp-h2">' + esc(this.properties.titulo || 'Avisos') + '</h2>' +
      '<button type="button" class="cbp-btn pri" data-nuevo="1">+ Nuevo aviso</button></div>' + dest +
      '<div class="cbp-toolbar"><div class="cbp-chips" role="group" aria-label="Filtro">' + chips(FILTROS, this.filtro, 'f') + '</div></div>' +
      (lista.length ? '<div class="cbp-av-grid">' + lista.map((a: IAv) => this.tarjeta(a)).join('') + '</div>'
        : '<div class="cbp-vacio">' + (this.filtro === 'Caducados' ? 'No hay avisos caducados.' : 'No hay avisos en vigor.') + '</div>') +
      '</div></div>';

    const q = (s: string): NodeListOf<HTMLElement> => this.domElement.querySelectorAll(s) as NodeListOf<HTMLElement>;
    q('[data-f]').forEach((b: HTMLElement) => b.addEventListener('click', () => { this.filtro = b.getAttribute('data-f') || 'Vigentes'; this.render(); }));
    (this.domElement.querySelector('[data-nuevo]') as HTMLElement).addEventListener('click', () => this.formulario());
    q('[data-ed]').forEach((b: HTMLElement) => b.addEventListener('click', () => this.formulario(this.avisos.filter((a: IAv) => a.id === +(b.getAttribute('data-ed') || 0))[0])));
    q('[data-del]').forEach((b: HTMLElement) => b.addEventListener('click', () => this.eliminar(+(b.getAttribute('data-del') || 0))));

    const hm = window.location.hash.match(/^#aviso=(\d+)/);
    if (hm) {
      const el = this.domElement.querySelector('#aviso-' + hm[1]) as HTMLElement;
      if (!el && this.filtro !== 'Todos') { this.filtro = 'Todos'; this.render(); return; }
      if (el) { el.classList.add('foco'); setTimeout(() => el.scrollIntoView({ behavior: 'smooth', block: 'center' }), 100); }
    }
  }

  private formulario(a?: IAv): void {
    const c = abrirModal(a ? 'Editar aviso' : 'Nuevo aviso',
      '<form class="cbp-form" data-fm="1">' +
      '<label class="cbp-field full">Título<input name="t" required maxlength="255" value="' + esc(a ? a.titulo : '') + '" placeholder="Ej.: Cambio de horario del Cadete A el sábado"></label>' +
      '<label class="cbp-field full">Detalle<textarea name="d" rows="5" placeholder="Qué pasa, a quién afecta y qué hay que hacer. Puedes pegar enlaces.">' + esc(a ? a.detalle : '') + '</textarea></label>' +
      '<label class="cbp-field">Visible hasta <span style="font-weight:400;color:#5c5c5c">(opcional)</span><input type="date" name="l" value="' + (a && a.limite ? fechaInput(a.limite) : '') + '"></label>' +
      '<label class="cbp-check" style="align-self:end;min-height:44px"><input type="checkbox" name="i"' + (a && a.importante ? ' checked' : '') + '>Marcar como importante</label>' +
      '<div class="cbp-btns full"><button type="button" class="cbp-btn" data-c="1">Cancelar</button><button type="submit" class="cbp-btn pri">' + (a ? 'Guardar' : 'Publicar aviso') + '</button></div></form>', 580);
    (c.querySelector('[data-c]') as HTMLElement).addEventListener('click', cerrarModal);
    const f = c.querySelector('[data-fm]') as HTMLFormElement;
    f.addEventListener('submit', (ev: Event) => {
      ev.preventDefault();
      const v = (n: string): HTMLInputElement => f.querySelector('[name="' + n + '"]') as HTMLInputElement;
      const lim = v('l').value;
      const body = { Title: v('t').value.trim(), Detalle: v('d').value.trim(), Prioridad: v('i').checked ? 'Importante' : 'Normal', FechaLimite: lim ? lim + 'T12:00:00Z' : null };
      const w = this.context.pageContext.web.absoluteUrl;
      spPost(this.context.spHttpClient, listUrl(w, 'Lists/Avisos') + '/items' + (a ? '(' + a.id + ')' : ''), body, !!a)
        .then(() => { cerrarModal(); aviso(a ? 'Aviso actualizado' : 'Aviso publicado'); return this.cargar(); })
        .then(() => { if (!a) { this.filtro = 'Vigentes'; } this.render(); })
        .catch((x: Error) => aviso('Error: ' + x.message, true));
    });
  }

  private eliminar(id: number): void {
    const a = this.avisos.filter((x: IAv) => x.id === id)[0];
    if (!a || !window.confirm('¿Eliminar el aviso «' + a.titulo + '»? Irá a la papelera del sitio.')) { return; }
    const w = this.context.pageContext.web.absoluteUrl;
    spPost(this.context.spHttpClient, listUrl(w, 'Lists/Avisos') + '/items(' + id + ')/recycle')
      .then(() => { aviso('Aviso eliminado'); return this.cargar(); }).then(() => this.render())
      .catch((x: Error) => aviso('Error: ' + x.message, true));
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: 'Avisos Gijón CBP' }, groups: [{ groupName: 'General', groupFields: [
      PropertyPaneTextField('titulo', { label: 'Título' })
    ] }] }] };
  }
}
