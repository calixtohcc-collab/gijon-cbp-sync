import { IEvento, IAviso, esc, MESES, DIAS, hora, icon } from './cbp';

export function filaEvento(e: IEvento, webUrl: string, conResultado?: boolean): string {
  const href = webUrl + '/Lists/Calendario/DispForm.aspx?ID=' + e.id;
  const titulo = e.tipo === 'Partido' ? 'vs ' + e.rival : e.titulo;
  let tag: string;
  if (e.tipo === 'Partido') { tag = e.local ? '<span class="cbp-tag casa">EN CASA</span>' : '<span class="cbp-tag">FUERA</span>'; }
  else if (e.tipo === 'Torneo') { tag = '<span class="cbp-tag torneo">TORNEO</span>'; }
  else { tag = '<span class="cbp-tag otro">' + esc(e.tipo.toUpperCase()) + '</span>'; }
  let derecha = tag;
  if (conResultado && e.pf !== undefined && e.pa !== undefined) {
    const v = (e.pf as number) > (e.pa as number);
    derecha = '<div class="cbp-res"><span class="cbp-marcador">' + e.pf + '–' + e.pa + '</span><span class="cbp-wl ' + (v ? 'v' : 'd') + '" aria-label="' + (v ? 'Victoria' : 'Derrota') + '">' + (v ? 'V' : 'D') + '</span></div>';
  }
  const lugar = e.lugar ? e.lugar : 'Lugar por confirmar';
  return '<a class="cbp-partido" data-evid="' + e.id + '" href="' + esc(href) + '">' +
    '<div class="cbp-fecha" aria-hidden="true"><small>' + DIAS[e.fecha.getDay()] + '</small><b>' + e.fecha.getDate() + '</b><span>' + MESES[e.fecha.getMonth()] + '</span></div>' +
    '<div class="cbp-pinfo"><span class="cbp-eq">' + esc(e.equipo) + '</span><span class="cbp-riv">' + esc(titulo) + '</span>' +
    '<span class="cbp-meta"><span>' + icon('reloj', 14) + hora(e.fecha) + '</span><span>' + icon('pin', 14) + esc(lugar) + '</span></span></div>' +
    derecha + '</a>';
}

export function filaAviso(a: IAviso, webUrl: string): string {
  const href = webUrl + '/SitePages/Avisos.aspx#aviso=' + a.id;
  let top = '';
  if (a.importante) { top += '<span class="cbp-badge">IMPORTANTE</span>'; }
  if (a.limite) { top += '<span class="cbp-limite">Hasta el ' + a.limite.getDate() + ' ' + MESES[a.limite.getMonth()].toLowerCase() + '</span>'; }
  return '<article class="cbp-aviso">' + (top ? '<div class="cbp-aviso-top">' + top + '</div>' : '') +
    '<h3><a href="' + esc(href) + '" style="text-decoration:none">' + esc(a.titulo) + '</a></h3>' +
    (a.detalle ? '<p>' + esc(a.detalle) + '</p>' : '') + '</article>';
}

export function chips(opciones: string[], activa: string, attr: string): string {
  return opciones.map((o: string) => '<button type="button" class="cbp-chip" data-' + attr + '="' + esc(o) + '" aria-pressed="' + (o === activa ? 'true' : 'false') + '">' + esc(o) + '</button>').join('');
}
