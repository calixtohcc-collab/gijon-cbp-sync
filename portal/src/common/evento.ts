import { SPHttpClient } from '@microsoft/sp-http';
import { IEvento, esc, DIAS_LARGO, MESES_LARGO, hora } from './cbp';
import { abrirModal, cerrarModal, descargar, generarIcs, mapa, spPost, listUrl, opcionesEquipo, fechaInput, aviso } from './ui';
import { filaEvento } from './render';

/** Ficha de un evento en ventana modal. */
export function mostrarEvento(e: IEvento, webUrl: string): void {
  const titulo = e.tipo === 'Partido' ? e.equipo : e.titulo;
  const comp = (e.notas.split(' · ')[0] || '').replace(/Fuente: FBPA/, '');
  let res = '';
  if (e.pf !== undefined && e.pa !== undefined) {
    const v = (e.pf as number) > (e.pa as number);
    res = '<div class="cbp-bigres"><b>' + e.pf + '–' + e.pa + '</b><span style="font-weight:700;letter-spacing:.06em">' + (v ? 'VICTORIA' : 'DERROTA') + '</span></div>';
  }
  const html = res +
    (e.tipo === 'Partido' ? '<div style="font-size:20px;font-weight:700">' + (e.local ? 'vs ' : 'en casa de ') + esc(e.rival) + '</div>' : '') +
    '<dl class="cbp-dl">' +
    '<dt>Equipo</dt><dd>' + esc(e.equipo) + '</dd>' +
    '<dt>Tipo</dt><dd>' + esc(e.tipo) + (e.tipo === 'Partido' ? (e.local ? ' · en casa' : ' · fuera') : '') + '</dd>' +
    '<dt>Fecha</dt><dd>' + DIAS_LARGO[e.fecha.getDay()] + ' ' + e.fecha.getDate() + ' de ' + MESES_LARGO[e.fecha.getMonth()] + ' de ' + e.fecha.getFullYear() + ' · ' + hora(e.fecha) + '</dd>' +
    '<dt>Lugar</dt><dd>' + (e.lugar ? esc(e.lugar) + ' · <a href="' + esc(mapa(e.lugar)) + '" target="_blank" rel="noopener">Cómo llegar</a>' : 'Por confirmar') + '</dd>' +
    (comp && e.tipo === 'Partido' ? '<dt>Competición</dt><dd>' + esc(comp) + '</dd>' : '') +
    (e.notas && e.tipo !== 'Partido' ? '<dt>Notas</dt><dd>' + esc(e.notas) + '</dd>' : '') +
    '</dl>' +
    '<div class="cbp-btns">' +
    (e.tipo === 'Partido' ? '<a class="cbp-btn" href="' + esc(webUrl + '/SitePages/Herramientas.aspx#convocatorias=' + e.id) + '">Preparar convocatoria</a>' : '') +
    '<button type="button" class="cbp-btn" data-a="ics">Añadir a mi calendario</button>' +
    '<a class="cbp-btn" href="' + esc(webUrl + '/Lists/Calendario/EditForm.aspx?ID=' + e.id) + '" target="_blank" rel="noopener">Editar</a>' +
    '<button type="button" class="cbp-btn pri" data-a="cerrar">Cerrar</button></div>';
  const body = abrirModal(titulo, html, 560);
  (body.querySelector('[data-a=ics]') as HTMLElement).addEventListener('click', () => descargar('partido-' + e.id + '.ics', generarIcs([e], 'Gijón CBP'), 'text/calendar'));
  (body.querySelector('[data-a=cerrar]') as HTMLElement).addEventListener('click', cerrarModal);
}

/** Lista de eventos de un día en ventana modal. */
export function mostrarDia(d: Date, evs: IEvento[], webUrl: string): void {
  const body = abrirModal(DIAS_LARGO[d.getDay()] + ' ' + d.getDate() + ' de ' + MESES_LARGO[d.getMonth()],
    evs.length ? '<div class="cbp-card cbp-lista" style="overflow:hidden">' + evs.map((e: IEvento) => filaEvento(e, webUrl, true)).join('') + '</div>' : '<div class="cbp-vacio">Sin eventos este día.</div>', 640);
  enlazarEventos(body, evs, webUrl);
}

/** Intercepta los clics en filas de evento para abrir la ficha. */
export function enlazarEventos(root: HTMLElement, evs: IEvento[], webUrl: string): void {
  const filas = root.querySelectorAll('[data-evid]');
  for (let i = 0; i < filas.length; i++) {
    filas[i].addEventListener('click', (ev: Event) => {
      const id = parseInt((ev.currentTarget as HTMLElement).getAttribute('data-evid') || '0', 10);
      const e = evs.filter((x: IEvento) => x.id === id)[0];
      if (e) { ev.preventDefault(); mostrarEvento(e, webUrl); }
    });
  }
}

/** Formulario de nuevo evento. Llama a alGuardar cuando se crea. */
export function nuevoEvento(client: SPHttpClient, webUrl: string, fecha: Date | undefined, equipo: string, alGuardar: () => void): void {
  const f = fecha || new Date();
  const html = '<form class="cbp-form" novalidate>' +
    '<label class="cbp-field">Tipo<select name="tipo"><option>Entrenamiento</option><option>Partido</option><option>Torneo</option><option>Reunión</option><option>Hito</option></select></label>' +
    '<label class="cbp-field">Equipo<select name="equipo">' + opcionesEquipo(equipo || 'Todo el club', true) + '</select></label>' +
    '<label class="cbp-field">Fecha<input type="date" name="fecha" required value="' + fechaInput(f) + '"></label>' +
    '<label class="cbp-field">Hora<input type="time" name="hora" required value="18:00"></label>' +
    '<label class="cbp-field full" data-solo="no-partido">Título<input name="titulo" placeholder="Ej.: Entrenamiento extra, reunión de padres…"></label>' +
    '<label class="cbp-field" data-solo="partido" hidden>Rival<input name="rival"></label>' +
    '<label class="cbp-field" data-solo="partido" hidden>Juega<select name="local"><option value="1">En casa</option><option value="0">Fuera</option></select></label>' +
    '<label class="cbp-field full">Lugar<input name="lugar" placeholder="Pabellón o pista"></label>' +
    '<label class="cbp-field full">Notas<textarea name="notas"></textarea></label>' +
    '<div class="cbp-btns full"><button type="button" class="cbp-btn" data-a="cancelar">Cancelar</button><button type="submit" class="cbp-btn pri">Guardar evento</button></div></form>';
  const body = abrirModal('Nuevo evento', html, 620);
  const form = body.querySelector('form') as HTMLFormElement;
  const val = (n: string): string => ((form.elements.namedItem(n) as HTMLInputElement).value || '').trim();
  const tipoSel = form.elements.namedItem('tipo') as HTMLSelectElement;
  const toggle = (): void => {
    const p = tipoSel.value === 'Partido';
    const els = form.querySelectorAll('[data-solo]');
    for (let i = 0; i < els.length; i++) { (els[i] as HTMLElement).hidden = (els[i].getAttribute('data-solo') === 'partido') !== p; }
  };
  tipoSel.addEventListener('change', toggle);
  (body.querySelector('[data-a=cancelar]') as HTMLElement).addEventListener('click', cerrarModal);
  form.addEventListener('submit', (ev: Event) => {
    ev.preventDefault();
    const tipo = val('tipo'); const eq = val('equipo'); const fs = val('fecha'); const hs = val('hora');
    if (!fs || !hs) { aviso('Indica fecha y hora', true); return; }
    const p = fs.split('-'); const h = hs.split(':');
    const d = new Date(+p[0], +p[1] - 1, +p[2], +h[0], +h[1]);
    const rival = val('rival'); const local = val('local') === '1';
    if (tipo === 'Partido' && !rival) { aviso('Indica el rival', true); return; }
    const titulo = tipo === 'Partido' ? eq + ' · ' + rival + ' (' + (local ? 'local' : 'visitante') + ')' : (val('titulo') || tipo + ' · ' + eq);
    const notas = (tipo === 'Partido' ? (local ? '· Local' : '· Visitante') + ' · ' : '') + val('notas');
    const btn = form.querySelector('[type=submit]') as HTMLButtonElement; btn.disabled = true;
    spPost(client, listUrl(webUrl, 'Lists/Calendario') + '/items', { Title: titulo.slice(0, 255), Fecha: d.toISOString(), Equipo: eq, TipoEvento: tipo, Rival: rival, Lugar: val('lugar'), Notas: notas })
      .then(() => { cerrarModal(); aviso('Evento guardado'); alGuardar(); })
      .catch((e: Error) => { btn.disabled = false; aviso('No se pudo guardar: ' + e.message, true); });
  });
}
