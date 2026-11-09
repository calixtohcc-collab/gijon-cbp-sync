import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { cargarEstilos, esc, MESES } from '../../common/cbp';
import { chips } from '../../common/render';
import { cargarEstilosUi, spGet, spPost, listUrl, webPath, abrirModal, cerrarModal, aviso } from '../../common/ui';

export interface IMaterialesCbpWebPartProps { titulo: string; }

interface IMat { id: number; titulo: string; archivo: string; ruta: string; tipo: string; cat: string; desc: string; mod: Date; autor: string; ext: string; }

const TIPOS = ['Ejercicio', 'Sistema', 'Planificación', 'Documento', 'Vídeo'];
const CATS = ['Escuela', 'Benjamín', 'Alevín', 'Infantil', 'Cadete', 'Junior', 'Sénior', 'Madres y padres', 'Todo el club'];
const IMG = ['png', 'jpg', 'jpeg', 'gif', 'webp'];
const PREV = ['pdf', 'docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls'];

export default class MaterialesCbpWebPart extends BaseClientSideWebPart<IMaterialesCbpWebPartProps> {
  private mats: IMat[] = [];
  private cargado: boolean = false;
  private tipo: string = 'Todos';
  private cat: string = 'Todas';
  private q: string = '';

  protected onInit(): Promise<void> { cargarEstilos(); cargarEstilosUi(); return this.recargar(); }

  private recargar(): Promise<void> {
    const w = this.context.pageContext.web.absoluteUrl;
    return spGet(this.context.spHttpClient, listUrl(w, 'Materiales') + '/items?$select=Id,Title,FileLeafRef,FileRef,TipoMaterial,Categoria,Descripcion,Modified,Editor/Title,FSObjType&$expand=Editor&$orderby=Modified desc&$top=1000')
      .then((j: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
        this.mats = (j.value || []).filter((i: any) => i.FSObjType === 0 || i.FSObjType === '0').map((i: any): IMat => { // eslint-disable-line @typescript-eslint/no-explicit-any
          const f: string = i.FileLeafRef || '';
          return { id: i.Id, titulo: i.Title || f.replace(/\.[^.]+$/, ''), archivo: f, ruta: i.FileRef, tipo: i.TipoMaterial || '', cat: i.Categoria || '',
            desc: i.Descripcion || '', mod: new Date(i.Modified), autor: (i.Editor && i.Editor.Title) || '', ext: (f.split('.').pop() || '').toLowerCase() };
        });
        this.cargado = true;
      }).catch(() => { this.cargado = true; });
  }

  private preview(m: IMat): string {
    const origin = window.location.protocol + '//' + window.location.host;
    if (IMG.indexOf(m.ext) > -1) { return '<img loading="lazy" alt="" src="' + esc(m.ruta) + '">'; }
    if (PREV.indexOf(m.ext) > -1) {
      return '<img loading="lazy" alt="" src="' + esc(this.context.pageContext.web.absoluteUrl + '/_layouts/15/getpreview.ashx?path=' + encodeURIComponent(origin + m.ruta) + '&resolution=1') + '" onerror="this.style.display=\'none\'">' +
        '<span class="cbp-mat-ext" style="position:absolute;z-index:-1">' + esc(m.ext.toUpperCase()) + '</span>';
    }
    return '<span class="cbp-mat-ext">' + esc((m.ext || 'DOC').toUpperCase()) + '</span>';
  }

  public render(): void {
    if (!this.cargado) { this.domElement.innerHTML = '<div class="cbp"><div class="cbp-cargando">Cargando materiales…</div></div>'; return; }
    const q = this.q.toLowerCase();
    const lista = this.mats.filter((m: IMat) => (this.tipo === 'Todos' || m.tipo === this.tipo) && (this.cat === 'Todas' || m.cat === this.cat || m.cat === 'Todo el club') &&
      (!q || (m.titulo + ' ' + m.desc + ' ' + m.archivo).toLowerCase().indexOf(q) > -1));
    const tarjetas = lista.map((m: IMat) => {
      const href = /\.(docx?|pptx?|xlsx?)$/i.test(m.archivo) ? m.ruta + '?web=1' : m.ruta;
      return '<a class="cbp-card cbp-mat" href="' + esc(href) + '" target="_blank" rel="noopener">' +
        '<div class="cbp-mat-prev">' + this.preview(m) + (m.tipo ? '<span class="cbp-mat-tipo">' + esc(m.tipo) + '</span>' : '') + '</div>' +
        '<div class="cbp-mat-body"><b>' + esc(m.titulo) + '</b>' + (m.desc ? '<p>' + esc(m.desc) + '</p>' : '') +
        '<div class="cbp-mat-meta"><span>' + esc(m.cat || 'Sin categoría') + '</span><span>' + m.mod.getDate() + ' ' + MESES[m.mod.getMonth()].toLowerCase() + '</span></div></div></a>';
    }).join('');
    this.domElement.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:24px;gap:18px">' +
      '<div class="cbp-sec-head"><h2 class="cbp-h2">' + esc(this.properties.titulo || 'Materiales') + '</h2>' +
      '<div style="display:flex;gap:10px;flex-wrap:wrap"><a class="cbp-btn" href="' + esc(this.context.pageContext.web.absoluteUrl + '/Materiales') + '">Abrir biblioteca</a>' +
      '<button type="button" class="cbp-btn pri" data-subir="1">+ Subir material</button></div></div>' +
      '<div class="cbp-toolbar"><input type="search" class="cbp-search" placeholder="Buscar por título, descripción o archivo…" aria-label="Buscar materiales" value="' + esc(this.q) + '" data-q="1">' +
      '<span style="font-size:14px;color:#5c5c5c">' + lista.length + ' de ' + this.mats.length + ' materiales</span></div>' +
      '<div class="cbp-chips" role="group" aria-label="Tipo">' + chips(['Todos'].concat(TIPOS), this.tipo, 'tipo') + '</div>' +
      '<div class="cbp-chips" role="group" aria-label="Categoría">' + chips(['Todas'].concat(CATS.filter((c: string) => c !== 'Todo el club')), this.cat, 'cat') + '</div>' +
      (lista.length ? '<div class="cbp-mats">' + tarjetas + '</div>' : '<div class="cbp-vacio">' + (this.mats.length ? 'Ningún material coincide con la búsqueda.' : 'Todavía no hay materiales. Pulsa «Subir material» para añadir el primero.') + '</div>') +
      '</div></div>';
    const bind = (sel: string, attr: string, fn: (v: string) => void): void => {
      const els = this.domElement.querySelectorAll(sel);
      for (let i = 0; i < els.length; i++) { els[i].addEventListener('click', (ev: Event) => fn((ev.currentTarget as HTMLElement).getAttribute(attr) || '')); }
    };
    bind('[data-tipo]', 'data-tipo', (v: string) => { this.tipo = v; this.render(); });
    bind('[data-cat]', 'data-cat', (v: string) => { this.cat = v; this.render(); });
    bind('[data-subir]', 'data-subir', () => this.subir());
    const inp = this.domElement.querySelector('[data-q]') as HTMLInputElement;
    let t: number = 0;
    inp.addEventListener('input', () => { window.clearTimeout(t); t = window.setTimeout(() => { this.q = inp.value; this.render(); const n = this.domElement.querySelector('[data-q]') as HTMLInputElement; n.focus(); n.setSelectionRange(n.value.length, n.value.length); }, 250); });
  }

  private subir(): void {
    const html = '<form class="cbp-form" novalidate>' +
      '<div class="full"><label class="cbp-drop" data-drop="1"><input type="file" name="file" hidden><span data-fname="1">Arrastra aquí el archivo o pulsa para elegirlo</span></label></div>' +
      '<label class="cbp-field full">Título<input name="titulo" required placeholder="Ej.: Ataque contra zona 2-3"></label>' +
      '<label class="cbp-field">Tipo<select name="tipo">' + TIPOS.map((x: string) => '<option>' + x + '</option>').join('') + '</select></label>' +
      '<label class="cbp-field">Categoría<select name="cat">' + CATS.map((x: string) => '<option>' + esc(x) + '</option>').join('') + '</select></label>' +
      '<label class="cbp-field full">Descripción<textarea name="desc" placeholder="Objetivo, material necesario, consignas…"></textarea></label>' +
      '<div class="cbp-btns full"><button type="button" class="cbp-btn" data-a="cancelar">Cancelar</button><button type="submit" class="cbp-btn pri">Subir</button></div></form>';
    const body = abrirModal('Subir material', html, 600);
    const form = body.querySelector('form') as HTMLFormElement;
    const fileInp = form.elements.namedItem('file') as HTMLInputElement;
    const drop = body.querySelector('[data-drop]') as HTMLElement;
    const fname = body.querySelector('[data-fname]') as HTMLElement;
    let archivo: File | undefined;
    const elegir = (f: File | undefined): void => {
      archivo = f; if (!f) { return; }
      fname.textContent = f.name + ' · ' + Math.round(f.size / 1024) + ' KB';
      const t = form.elements.namedItem('titulo') as HTMLInputElement;
      if (!t.value) { t.value = f.name.replace(/\.[^.]+$/, '').replace(/[_-]+/g, ' '); }
      const tipo = form.elements.namedItem('tipo') as HTMLSelectElement;
      if (/\.(mp4|mov|avi|webm)$/i.test(f.name)) { tipo.value = 'Vídeo'; }
    };
    fileInp.addEventListener('change', () => elegir(fileInp.files ? fileInp.files[0] : undefined));
    drop.addEventListener('dragover', (e: DragEvent) => { e.preventDefault(); drop.classList.add('on'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('on'));
    drop.addEventListener('drop', (e: DragEvent) => { e.preventDefault(); drop.classList.remove('on'); elegir(e.dataTransfer && e.dataTransfer.files[0] ? e.dataTransfer.files[0] : undefined); });
    (body.querySelector('[data-a=cancelar]') as HTMLElement).addEventListener('click', cerrarModal);
    form.addEventListener('submit', (ev: Event) => {
      ev.preventDefault();
      if (!archivo) { aviso('Elige un archivo', true); return; }
      const val = (n: string): string => ((form.elements.namedItem(n) as HTMLInputElement).value || '').trim();
      const btn = form.querySelector('[type=submit]') as HTMLButtonElement; btn.disabled = true; btn.textContent = 'Subiendo…';
      this.subirArchivo(archivo, { Title: val('titulo'), TipoMaterial: val('tipo'), Categoria: val('cat'), Descripcion: val('desc') })
        .then(() => { cerrarModal(); aviso('Material subido'); return this.recargar(); })
        .then(() => this.render())
        .catch((e: Error) => { btn.disabled = false; btn.textContent = 'Subir'; aviso('No se pudo subir: ' + e.message, true); });
    });
  }

  /** Sube un archivo a Materiales y rellena sus metadatos. También lo usa la pizarra. */
  public subirArchivo(f: Blob & { name?: string }, campos: { [k: string]: string }, nombre?: string): Promise<void> {
    return subirAMateriales(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl, f, campos, nombre);
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: 'Materiales Gijón CBP' }, groups: [{ groupName: 'General', groupFields: [PropertyPaneTextField('titulo', { label: 'Título' })] }] }] };
  }
}

export function subirAMateriales(client: SPHttpClient, w: string, f: Blob & { name?: string }, campos: { [k: string]: string }, nombre?: string): Promise<void> {
  const base = (nombre || f.name || 'archivo').replace(/[#%*:<>?/\\|"]/g, '-');
  const unico = base.replace(/(\.[^.]+)$/, '-' + Date.now().toString(36) + '$1');
  const carpeta = webPath(w) + '/Materiales';
  return client.post(w + "/_api/web/GetFolderByServerRelativeUrl('" + encodeURIComponent(carpeta) + "')/Files/add(url='" + encodeURIComponent(unico.replace(/'/g, "''")) + "',overwrite=false)?$expand=ListItemAllFields",
    SPHttpClient.configurations.v1, { body: f, headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' } })
    .then((r: SPHttpClientResponse) => { if (!r.ok) { return r.text().then((t: string) => { throw new Error('HTTP ' + r.status + ' ' + t.slice(0, 160)); }); } return r.json(); })
    .then((j: any) => spPost(client, listUrl(w, 'Materiales') + '/items(' + j.ListItemAllFields.Id + ')', campos, true)) // eslint-disable-line @typescript-eslint/no-explicit-any
    .then(() => undefined);
}
