// Hoja semanal por equipo: el fin de semana de un equipo listo para leer en el móvil y reenviar a las familias.
// Pensada para correo (HTML con estilos en línea y tablas, que es lo que respetan Outlook y Gmail).

const PORTAL = 'https://gijoncbp.sharepoint.com/sites/PortalCBP';
const DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const C = { negro: '#090909', ambar: '#FF9F00', gris: '#5c5c5c', linea: '#E6E3DC', fondo: '#F6F5F1', rojo: '#B00020', verde: '#1B5E20' };

const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const madrid = (iso) => { // fecha y hora en Gijón
  const p = new Intl.DateTimeFormat('es-ES', { timeZone: 'Europe/Madrid', hourCycle: 'h23', year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit', weekday: 'long' })
    .formatToParts(new Date(iso)).reduce((a, x) => ((a[x.type] = x.value), a), {});
  return { dia: +p.day, mes: +p.month - 1, anio: +p.year, h: +p.hour, mi: +p.minute, semana: p.weekday };
};
const hm = (h, mi) => String(h).padStart(2, '0') + ':' + String(mi).padStart(2, '0');
const fechaLarga = (f) => f.semana + ' ' + f.dia + ' de ' + MESES[f.mes];
const mapa = (lugar) => 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(lugar + ', Asturias');
const whatsapp = (txt) => 'https://api.whatsapp.com/send?text=' + encodeURIComponent(txt);
const titulo = (s) => String(s || '').toLowerCase().replace(/(^|[\s(."-])([a-záéíóúñ])/g, (m, a, b) => a + b.toUpperCase()).replace(/ (De|Del|La|Las|Los|Y|En)(?= )/g, (m, w) => ' ' + w.toLowerCase())
  .replace(/\b(C\.b|A\.c\.d|C\.d|C\.d\.b|I\.e\.s|Cd|Cb|Cp|Ies|Col|Pol|Pab)\b\.?/g, (m) => /^(Col|Pol|Pab)/.test(m) ? m : m.toUpperCase());

/**
 * datos = {
 *   equipo: { nombre, entrenadores, competicion },
 *   jugadores: número en plantilla,
 *   partidos: [{ id, fecha (ISO), rival, lugar, local, competicion, jornada,
 *               estado ('Oficial'|'Propuesto'|…), propuesta (ISO|null),
 *               convocatoria: { citacion, convocados } | null,
 *               viaje: { minutos, estimado } | null }],      // los del fin de semana
 *   siguientes: [{ fecha, rival, local, lugar }],           // las dos semanas después
 *   finde: { desde ISO, hasta ISO }
 * }
 */
export function hojaEquipo(d) {
  const ini = madrid(d.finde.desde); const fin = madrid(d.finde.hasta);
  const rango = ini.mes === fin.mes ? ini.dia + '–' + fin.dia + ' de ' + MESES[fin.mes] : ini.dia + ' de ' + MESES[ini.mes] + ' – ' + fin.dia + ' de ' + MESES[fin.mes];
  const pendientes = [];

  const tarjetas = d.partidos.map((p) => {
    const f = madrid(p.propuesta && p.estado !== 'Oficial' ? p.propuesta : p.fecha);
    const rival = titulo(p.rival); const lugar = titulo(p.lugar || 'Pabellón por confirmar');
    const conv = p.convocatoria;
    if (!conv) { pendientes.push('Hacer la convocatoria del ' + f.semana); }
    if (!p.lugar || /por confirmar/i.test(p.lugar)) { pendientes.push('Confirmar el pabellón del ' + f.semana); }
    // Plan de viaje: llegar 1 h antes del inicio
    let viaje = '';
    if (!p.local && p.viaje && p.viaje.minutos) {
      const salida = f.h * 60 + f.mi - 60 - p.viaje.minutos;
      viaje = fila('🚗', 'Salida desde Gijón', '<b>' + hm(Math.floor(salida / 60), salida % 60) + '</b> <span style="color:' + C.gris + '">· trayecto ' + (p.viaje.estimado ? '≈ ' : '') + p.viaje.minutos + ' min, para llegar 1 h antes</span>');
    }
    const enTramite = p.estado && p.estado !== 'Oficial' && p.estado !== 'Cargado en FBPA';
    return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid ' + C.linea + ';border-radius:14px;background:#fff;margin:0 0 14px">' +
      '<tr><td style="padding:18px 20px 6px">' +
      '<div style="font:700 12px/1 Arial,sans-serif;letter-spacing:.08em;text-transform:uppercase;color:' + (p.local ? C.verde : C.gris) + '">' + (p.local ? 'En casa' : 'Fuera') + ' · ' + esc(p.competicion) + (p.jornada ? ' · Jornada ' + esc(p.jornada) : '') + '</div>' +
      '<div style="font:800 26px/1.15 Impact,\'Arial Black\',sans-serif;color:' + C.negro + ';margin:8px 0 2px;text-transform:uppercase">' + esc(f.semana) + ' ' + f.dia + ' · ' + hm(f.h, f.mi) + '</div>' +
      '<div style="font:700 17px/1.3 Arial,sans-serif;color:' + C.negro + '">' + (p.local ? 'vs ' : 'en casa de ') + esc(rival) + '</div>' +
      (enTramite ? '<div style="margin-top:8px;font:13px/1.4 Arial,sans-serif;color:#8a5600;background:#FFF3DF;border-radius:8px;padding:6px 10px">Cambio de horario en trámite (' + esc(p.estado.toLowerCase()) + '). Horario oficial: ' + esc(fechaLarga(madrid(p.fecha))) + ' · ' + hm(madrid(p.fecha).h, madrid(p.fecha).mi) + '.</div>' : '') +
      '</td></tr><tr><td style="padding:6px 20px 16px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
      fila('📍', 'Lugar', esc(lugar) + (p.lugar && !/por confirmar/i.test(p.lugar) ? ' · <a href="' + esc(mapa(p.lugar)) + '" style="color:' + C.negro + ';font-weight:700">Cómo llegar</a>' : '')) +
      viaje +
      fila('⏰', 'Citación', conv && conv.citacion ? '<b>' + esc(conv.citacion) + '</b>' : '<span style="color:' + C.rojo + ';font-weight:700">Sin convocatoria todavía</span>') +
      fila('👥', 'Convocatoria', conv ? conv.convocados + ' de ' + d.jugadores + ' jugadoras/es' : '<a href="' + esc(PORTAL + '/SitePages/Herramientas.aspx#convocatorias=' + p.id + '&eq=' + encodeURIComponent(d.equipo.nombre)) + '" style="color:' + C.negro + ';font-weight:700">Hacerla en el portal →</a>') +
      '</table></td></tr></table>';
  }).join('');

  const sinPartido = !d.partidos.length ? '<p style="font:15px/1.5 Arial,sans-serif;color:' + C.gris + ';margin:0 0 14px">Este fin de semana no hay partido.</p>' : '';
  const siguientes = d.siguientes.length ? '<div style="font:700 12px/1 Arial,sans-serif;letter-spacing:.08em;text-transform:uppercase;color:' + C.gris + ';margin:22px 0 8px">Lo siguiente</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font:14px/1.4 Arial,sans-serif;color:' + C.negro + '">' +
    d.siguientes.map((s) => { const f = madrid(s.fecha); return '<tr><td style="padding:6px 0;border-top:1px solid ' + C.linea + ';width:150px;white-space:nowrap;color:' + C.gris + '">' + f.semana.slice(0, 3) + ' ' + f.dia + ' ' + MESES[f.mes].slice(0, 3) + ' · ' + hm(f.h, f.mi) + '</td><td style="padding:6px 0;border-top:1px solid ' + C.linea + '">' + (s.local ? 'vs ' : 'en ') + esc(titulo(s.rival)) + '</td></tr>'; }).join('') + '</table>' : '';

  // Mensaje para reenviar al grupo de familias (sin datos internos)
  const msgFamilias = '🏀 *' + d.equipo.nombre + '* · fin de semana ' + rango + '\n\n' + (d.partidos.length ? d.partidos.map((p) => {
    const f = madrid(p.propuesta && p.estado !== 'Oficial' ? p.propuesta : p.fecha);
    const salida = !p.local && p.viaje && p.viaje.minutos ? (() => { const m = f.h * 60 + f.mi - 60 - p.viaje.minutos; return '\n🚗 Salida desde Gijón: ' + hm(Math.floor(m / 60), m % 60); })() : '';
    return '📅 ' + fechaLarga(f) + ' · ' + hm(f.h, f.mi) + '\n' + (p.local ? '🏠 vs ' : '✈️ en casa de ') + titulo(p.rival) + '\n📍 ' + titulo(p.lugar || 'por confirmar') + (p.lugar ? ' · ' + mapa(p.lugar) : '') + salida +
      (p.convocatoria && p.convocatoria.citacion ? '\n⏰ Citación: ' + p.convocatoria.citacion : '\n⏰ Citación: la confirmamos con la convocatoria');
  }).join('\n\n') : 'Este fin de semana no hay partido.') + '\n\n¡Vamos Gijón CBP!';

  const html = '<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + esc(d.equipo.nombre) + ' · fin de semana</title></head>' +
    '<body style="margin:0;padding:0;background:' + C.fondo + '">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.fondo + '"><tr><td align="center" style="padding:18px 12px">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px">' +
    // Cabecera
    '<tr><td style="background:' + C.negro + ';border-radius:16px 16px 0 0;padding:22px 22px 20px">' +
    '<div style="font:700 11px/1 Arial,sans-serif;letter-spacing:.14em;color:' + C.ambar + ';text-transform:uppercase">Gijón CBP · Hoja del fin de semana</div>' +
    '<div style="font:800 30px/1.1 Impact,\'Arial Black\',sans-serif;color:#fff;margin-top:10px;text-transform:uppercase">' + esc(d.equipo.nombre) + '</div>' +
    '<div style="font:15px/1.4 Arial,sans-serif;color:#ddd;margin-top:6px">' + esc(rango) + (d.equipo.entrenadores ? ' · ' + esc(d.equipo.entrenadores) : '') + '</div>' +
    '</td></tr>' +
    '<tr><td style="background:' + C.fondo + ';padding:18px 0 4px">' +
    (pendientes.length ? '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 14px"><tr><td style="background:#FFF3DF;border-left:4px solid ' + C.ambar + ';border-radius:8px;padding:10px 14px;font:14px/1.5 Arial,sans-serif;color:' + C.negro + '"><b>Pendiente:</b> ' + esc(pendientes.join(' · ')) + '</td></tr></table>' : '') +
    tarjetas + sinPartido +
    // Botones
    '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:6px 0 4px"><tr>' +
    '<td style="padding:0 8px 8px 0"><a href="' + esc(whatsapp(msgFamilias)) + '" style="display:inline-block;background:' + C.negro + ';color:' + C.ambar + ';font:700 15px/1 Arial,sans-serif;text-decoration:none;padding:13px 18px;border-radius:999px">Reenviar a las familias por WhatsApp</a></td>' +
    '<td style="padding:0 0 8px"><a href="' + esc(PORTAL + '/SitePages/Equipos.aspx') + '" style="display:inline-block;border:1px solid ' + C.negro + ';color:' + C.negro + ';font:700 15px/1 Arial,sans-serif;text-decoration:none;padding:12px 18px;border-radius:999px">Abrir el portal</a></td>' +
    '</tr></table>' +
    siguientes +
    '</td></tr>' +
    '<tr><td style="padding:22px 4px 8px;font:12px/1.5 Arial,sans-serif;color:' + C.gris + '">Te llega porque figuras como contacto de ' + esc(d.equipo.nombre) + ' en el portal. Horarios de la FBPA actualizados cada noche; si la federación cambia algo de este fin de semana recibirás otro aviso. Si algo no cuadra, avisa a coordinación deportiva.</td></tr>' +
    '</table></td></tr></table></body></html>';

  return { asunto: d.equipo.nombre + ' · fin de semana ' + rango + (pendientes.length ? ' · falta: ' + pendientes[0].toLowerCase() : ''), html, familias: msgFamilias, pendientes };

  function fila(ico, etiqueta, valor) {
    return '<tr><td style="padding:7px 0;width:28px;vertical-align:top;font-size:16px">' + ico + '</td><td style="padding:7px 10px 7px 0;width:110px;vertical-align:top;font:13px/1.4 Arial,sans-serif;color:' + C.gris + '">' + etiqueta + '</td><td style="padding:7px 0;vertical-align:top;font:15px/1.4 Arial,sans-serif;color:' + C.negro + '">' + valor + '</td></tr>';
  }
}
