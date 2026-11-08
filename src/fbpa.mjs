// Lectura de la página de club de la FBPA (sin dependencias, funciona en Node y en el navegador).

const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', aacute: 'á', eacute: 'é', iacute: 'í', oacute: 'ó', uacute: 'ú', Aacute: 'Á', Eacute: 'É', Iacute: 'Í', Oacute: 'Ó', Uacute: 'Ú', ntilde: 'ñ', Ntilde: 'Ñ', uuml: 'ü', Uuml: 'Ü', ordf: 'ª', ordm: 'º', iexcl: '¡', iquest: '¿', ccedil: 'ç', acute: '´', middot: '·' };
export function texto(html) {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCharCode(parseInt(n, 16)))
    .replace(/&([a-z]+);/gi, (m, n) => (ENT[n] !== undefined ? ENT[n] : m))
    .replace(/\s+/g, ' ')
    .trim();
}
export function norm(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[ºª°]/g, '').replace(/\s+/g, ' ').trim();
}

// Correspondencia competición (inicio del encabezado normalizado) + nombre FBPA -> equipo del club
export const MAPA = [
  ['CTO. ESPANA 1', 'ALLER ALVAREZ GIJON CBP', 'Sénior Femenino A'],
  ['SEGUNDA FBPA FEM', 'GIJON CBP "B"', 'Sénior Femenino B'],
  ['SEGUNDA FBPA MAS', 'LA SANTINA GIJON CBP', 'Sénior Masculino A'],
  ['JUNIOR FEM 1', 'I-LANZA GIJON CBP "A"', 'Junior Femenino A'],
  ['JUNIOR MAS 1', 'I-LANZA GIJON CBP "A"', 'Junior Masculino A'],
  ['JUNIOR MAS 2', 'I-LANZA GIJON CBP "B"', 'Junior Masculino B'],
  ['CADETE FEM 1', 'I-LANZA GIJON CBP "A"', 'Cadete Femenino A'],
  ['CADETE FEM 2', 'GIJON CBP B', 'Cadete Femenino B'],
  ['CADETE MAS 1', 'I-LANZA GIJON CBP "A"', 'Cadete Masculino A'],
  ['CADETE MAS 2', 'GIJON CBP B', 'Cadete Masculino B'],
  ['CADETE MAS 2', 'I-LANZA GIJON CBP "C"', 'Cadete Masculino C'],
  ['INFANTIL FEM 1', 'I-LANZA GIJON CBP "A"', 'Infantil Femenino A'],
  ['INFANTIL MAS 1', 'I-LANZA GIJON CBP "A"', 'Infantil Masculino A'],
  ['INFANTIL MAS 2', 'GIJON CBP B', 'Infantil Masculino B'],
  ['INFANTIL MAS 2', 'C.B. ESTUDIANTES - GIJON CBP', 'Infantil Masculino C'],
  ['LIGA ALEVIN FBPA FEM - 1 FASE - 1', 'GRUPO DAGO GIJON CBP', 'Alevín Femenino A'],
  ['LIGA ALEVIN FBPA FEM - 1 FASE - 2', 'FISIOTERAPIA FDEZ. BLANCO GIJON CBP', 'Alevín Femenino B'],
  ['LIGA ALEVIN FBPA MAS - 1 FASE - 1', 'I-LANZA GIJON CBP "A"', 'Alevín Masculino A'],
  ['LIGA ALEVIN FBPA MAS - 1 FASE - 2', 'I-LANZA GIJON CBP "B"-ESTUDIANTES N 1', 'Alevín Masculino B'],
  ['LIGA ALEVIN FBPA MAS - 1 FASE - 2', 'I-LANZA GIJON CBP "C"', 'Alevín Masculino C'],
  ['LIGA BENJAMIN FBPA FEM', 'I-LANZA GIJON CBP "A"', 'Benjamín Femenino'],
  ['LIGA BENJAMIN FBPA MAS - 1 FASE - 1', 'I-LANZA GIJON CBP "A"', 'Benjamín Masculino A'],
  ['LIGA BENJAMIN FBPA MAS - 1 FASE - 2', 'I-LANZA GIJON CBP "B"-ESTUDIANTES N 1', 'Benjamín Masculino B'],
  ['LIGA BENJAMIN FBPA MAS - 1 FASE - 2', 'I-LANZA GIJON CBP "C"', 'Benjamín Masculino C'],
  ['LIGA DE MADRES', 'C.B. PUMARIN', 'Equipo de Madres'],
  ['LIGA DE PADRES', 'C.B. PUMARIN', 'Equipo de Padres']
];
const nn = (s) => norm(s).replace(/N\s?1\b/, 'N 1').replace(/N1\b/, 'N 1');

/**
 * Devuelve los partidos del club: { equipo, competicion, jornada, local, rival, campo, fecha:{y,m,d,h,mi}|null, pf, pa }
 */
export function leerPartidos(html) {
  const ini = html.indexOf('id="nav-calendario"');
  if (ini < 0) { throw new Error('No se encontró la pestaña de calendario en la página de la FBPA'); }
  const resto = html.slice(ini);
  const finRel = resto.search(/id="nav-(?!calendario)[a-z-]+"/);
  const sec = finRel > 0 ? resto.slice(0, finRel) : resto;

  // Títulos de cada bloque de competición
  const titulos = {};
  const reA = /<a[^>]*href="#(accordionCalendario[^"]+)"[^>]*>([\s\S]*?)<\/a>/g;
  let m;
  while ((m = reA.exec(html))) { titulos[m[1]] = texto(m[2]); }

  // Marcadores en orden: apertura de bloque, "Jornada N", tabla
  const marcas = [];
  const reC = /<div[^>]*id="(accordionCalendario[^"]*)"[^>]*>/g;
  while ((m = reC.exec(sec))) { if (titulos[m[1]] && / - /.test(titulos[m[1]])) { marcas.push({ p: m.index, t: 'c', v: m[1] }); } }
  const reJ = /Jornada\s+(\d+)/g;
  while ((m = reJ.exec(sec))) { marcas.push({ p: m.index, t: 'j', v: m[1] }); }
  const reT = /<table[\s\S]*?<\/table>/g;
  while ((m = reT.exec(sec))) { marcas.push({ p: m.index, t: 't', v: m[0] }); }
  marcas.sort((a, b) => a.p - b.p);

  const out = [];
  let comp = ''; let jor = '';
  for (const k of marcas) {
    if (k.t === 'c') { comp = titulos[k.v]; jor = ''; continue; }
    if (k.t === 'j') { jor = k.v; continue; }
    if (!comp) { continue; }
    const H = norm(comp);
    const filas = k.v.match(/<tr[\s\S]*?<\/tr>/g) || [];
    for (const f of filas) {
      const tds = (f.match(/<td[\s\S]*?<\/td>/g) || []).map((x) => texto(x));
      if (tds.length < 6) { continue; }
      const [loc, pl, pv, vis, fec, campo] = tds;
      for (const [hp, team, equipo] of MAPA) {
        if (!H.startsWith(hp)) { continue; }
        const T = nn(team);
        let local = null;
        if (nn(loc) === T) { local = true; } else if (nn(vis) === T) { local = false; }
        if (local === null) { continue; }
        const rival = local ? vis : loc;
        if (/DESCANSA/i.test(rival)) { continue; }
        const fm = fec.match(/(\d{2})\/(\d{2})\/(\d{4})\s*(\d{1,2}):(\d{2})?/) || fec.match(/(\d{2})\/(\d{2})\/(\d{4})/);
        const fecha = fm ? { y: +fm[3], m: +fm[2], d: +fm[1], h: fm[4] ? +fm[4] : 12, mi: fm[5] ? +fm[5] : 0, sinHora: !fm[4] } : null;
        const pf = local ? pl : pv; const pa = local ? pv : pl;
        out.push({ equipo, competicion: comp, jornada: jor, local, rival, campo: /^-+$/.test(campo) ? '' : campo, fecha,
          pf: /^\d+$/.test(pf) ? +pf : null, pa: /^\d+$/.test(pa) ? +pa : null });
      }
    }
  }
  return out;
}
