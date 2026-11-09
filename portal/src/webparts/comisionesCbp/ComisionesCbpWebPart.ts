import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { cargarEstilos } from '../../common/cbp';
import { cargarEstilosUi } from '../../common/ui';
import { cargarEstilos3 } from '../../common/estilos3';
import { cargarEstilos4 } from '../../common/estilos4';
import { renderComisiones } from '../../common/comisiones';

export interface IComisionesCbpWebPartProps { titulo: string; }

export default class ComisionesCbpWebPart extends BaseClientSideWebPart<IComisionesCbpWebPartProps> {
  private onHash = (): void => { this.render(); };

  protected onInit(): Promise<void> {
    cargarEstilos(); cargarEstilosUi(); cargarEstilos3(); cargarEstilos4();
    window.addEventListener('hashchange', this.onHash);
    return Promise.resolve();
  }
  protected onDispose(): void { window.removeEventListener('hashchange', this.onHash); }

  public render(): void {
    this.domElement.innerHTML = '<div class="cbp"><div class="cbp-wrap" style="padding-top:24px;gap:18px" data-box="1"></div></div>';
    renderComisiones(this.domElement.querySelector('[data-box]') as HTMLElement, { client: this.context.spHttpClient, w: this.context.pageContext.web.absoluteUrl });
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: 'Comisiones de trabajo' }, groups: [{ groupName: 'General', groupFields: [PropertyPaneTextField('titulo', { label: 'Título' })] }] }] };
  }
}
