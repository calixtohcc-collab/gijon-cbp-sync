/* Detección de conflictos sobre el horario oficial de la FBPA (fase 1) */

export interface IPartido { id: number; equipo: string; codigo: string; fecha: Date; local: boolean; rival: string; lugar: string; sinHora: boolean; dur: number; franja: string; }
export interface IRegla { id: number; titulo: string; a: string; condA: string; b: string; condB: string; min: number; motivo: string; }
export interface IVentana { dia: number; desde: number; hasta: number; }
export interface IPista { id: number; clave: string; inst: string; modulo: string; alias: string; disp: IVentana[]; dispTexto: string; controlar: boolean; orden: number; pdm?: string; permitidos?: string[]; }
/** ¿Puede este partido usar este módulo? (vacío = cualquiera; si no, categorías, nombres o códigos de equipo) */
export function puedeUsar(p: IPartido, k: IPista): boolean { const ps = k.permitidos || []; if (!ps.length) { return true; } const eq = norm(p.equipo); return ps.some((t: string) => { const n = norm(t); return !!n && (eq.indexOf(n) === 0 || norm(p.codigo) === n); }); }
export interface IJugV { nombre: string; equipo: string; habitual: string[]; ocasional: string[]; }
export type Gravedad = 'critica' | 'aviso' | 'info';
export interface IAlerta { tipo: 'conflicto' | 'saturacion' | 'reserva' | 'franja' | 'hora'; grav: Gravedad; partidos: number[]; titulo: string; detalle: string; accion: string; }

const DIAS_N: { [k: string]: number } = { domingo: 0, lunes: 1, martes: 2, miercoles: 3, jueves: 4, viernes: 5, sabado: 6 };
export const DIA_TXT = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];

export function norm(s: string): string { return (s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim(); }
export function minDia(d: Date): number { return d.getHours() * 60 + d.getMinutes(); }
export function hm(m: number): string { const h = Math.floor(m / 60); const x = m % 60; return ('0' + h).slice(-2) + ':' + ('0' + x).slice(-2); }
export function durTxt(m: number): string { const h = Math.floor(m / 60); const x = m % 60; return (h ? h + ' h' : '') + (x ? (h ? ' ' : '') + x + ' min' : (h ? '' : '0 min')); }
function mismoDia(a: Date, b: Date): boolean { return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate(); }

/** «Sábado 09:00-14:00» por línea → ventanas */
export function leerDisponibilidad(txt: string): IVentana[] {
  const out: IVentana[] = [];
  (txt || '').split(/\n|;/).forEach((l: string) => {
    const m = norm(l).match(/^([A-Z]+)\s+(\d{1,2}):(\d{2})\s*[-–a]\s*(\d{1,2}):(\d{2})/);
    if (!m) { return; }
    const d = DIAS_N[m[1].toLowerCase()];
    if (d === undefined) { return; }
    out.push({ dia: d, desde: +m[2] * 60 + +m[3], hasta: +m[4] * 60 + +m[5] });
  });
  return out;
}

/* ---------- Franjas horarias FBPA 2026/27 (horas de inicio) ---------- */
export interface IFranja { id?: number; dias: number[]; desde: string; hasta: string; tipo: 'oficial' | 'conformidad' | 'condicion'; cond?: string; pref?: boolean; }
const F = (dias: number[], desde: string, hasta: string, tipo: IFranja['tipo'], cond?: string): IFranja => ({ dias, desde, hasta, tipo, cond, pref: true });
export const FRANJAS: { [g: string]: IFranja[] } = {
  'Júnior y sénior': [F([4, 5], '18:30', '21:00', 'conformidad'), F([6], '16:00', '16:30', 'oficial'), F([6], '18:00', '18:30', 'oficial'), F([6], '20:00', '20:30', 'oficial'),
    F([0], '10:00', '12:30', 'oficial'), F([0], '16:00', '16:30', 'oficial'), F([0], '18:00', '18:30', 'oficial'), F([0], '20:00', '20:00', 'oficial')],
  'Minibasket': [F([2, 3, 4, 5], '16:30', '19:30', 'conformidad'), F([6], '09:30', '13:30', 'condicion', 'conformidad del rival si coincide con jornada de liga de Juegos Deportivos'), F([0], '09:30', '13:30', 'oficial')],
  'Cadete e infantil 1ª': [F([3, 4, 5], '17:30', '20:00', 'conformidad'), F([6], '09:30', '12:30', 'oficial'), F([6], '16:00', '16:30', 'oficial'), F([6], '18:00', '18:30', 'oficial'), F([6], '20:00', '20:30', 'oficial'),
    F([0], '09:30', '12:30', 'oficial'), F([0], '16:00', '16:30', 'oficial'), F([0], '18:00', '18:30', 'oficial'), F([0], '20:00', '20:00', 'condicion', 'a partir de la primera jornada de enero y sin desplazamiento subvencionado DGD')],
  'Resto de categorías': [F([1, 2], '16:30', '20:00', 'condicion', 'conformidad del rival; no se designa árbitro'), F([3, 4, 5], '16:30', '20:00', 'conformidad'), F([6], '09:30', '14:00', 'oficial'),
    F([0], '09:30', '12:30', 'condicion', 'conformidad del rival y, si pide transporte, del coordinador de zona visitante')]
};
const aMin = (s: string): number => +s.slice(0, 2) * 60 + +s.slice(3, 5);
/** Sustituye las franjas por las configuradas en la lista «Partidos · Franjas» */
export function cargarFranjas(fs: { grupo: string; f: IFranja }[]): void {
  if (!fs.length) { return; }
  Object.keys(FRANJAS).forEach((k: string) => { delete FRANJAS[k]; });
  fs.forEach((x) => { (FRANJAS[x.grupo] = FRANJAS[x.grupo] || []).push(x.f); });
}

function franjaDe(p: IPartido): { tipo: 'oficial' | 'conformidad' | 'condicion' | 'fuera' | 'n/a'; cond?: string; opciones: string; pref?: boolean } {
  const fr = FRANJAS[p.franja];
  if (!fr) { return { tipo: 'n/a', opciones: '' }; }
  const d = p.fecha.getDay(); const t = minDia(p.fecha);
  const del = fr.filter((f: IFranja) => f.dias.indexOf(d) >= 0);
  const opciones = del.filter((f: IFranja) => f.tipo === 'oficial').map((f: IFranja) => f.desde === f.hasta ? f.desde : f.desde + '–' + f.hasta).join(', ');
  const en = del.filter((f: IFranja) => t >= aMin(f.desde) && t <= aMin(f.hasta));
  const ofs = en.filter((f: IFranja) => f.tipo === 'oficial'); if (ofs.length) { return { tipo: 'oficial', opciones, pref: ofs.some((f: IFranja) => f.pref !== false) }; }
  const co = en.filter((f: IFranja) => f.tipo === 'condicion')[0]; if (co) { return { tipo: 'condicion', cond: co.cond, opciones }; }
  const cf = en.filter((f: IFranja) => f.tipo === 'conformidad')[0]; if (cf) { return { tipo: 'conformidad', opciones }; }
  return { tipo: 'fuera', opciones };
}

/* ---------- Instalaciones ---------- */
export function instalacionDe(p: IPartido, pistas: IPista[]): string {
  const n = norm(p.lugar);
  const x = pistas.filter((k: IPista) => norm(k.alias) === n || norm(k.inst) === n)[0];
  return x ? x.inst : '';
}

function quien(p: IPartido): string { return p.equipo + ' (' + (p.local ? 'en casa' : 'fuera') + ', ' + hm(minDia(p.fecha)) + ')'; }
function cumple(c: string, p: IPartido): boolean { return c === 'Cualquiera' || (c === 'Local' && p.local) || (c === 'Visitante' && !p.local); }

/** Calcula todas las alertas de un conjunto de partidos (normalmente una semana) */
export function analizar(ps: IPartido[], reglas: IRegla[], pistas: IPista[], margen: number, jugs?: IJugV[]): IAlerta[] {
  const out: IAlerta[] = [];
  if (jugs && jugs.length) { conflictosJugadores(ps, pistas, jugs).forEach((a: IAlerta) => out.push(a)); }

  // 1. Conflictos de entrenadores y jugadores
  const vistos: { [k: string]: IAlerta & { _min: number } } = {};
  reglas.forEach((r: IRegla) => {
    const as = ps.filter((p: IPartido) => p.codigo === r.a && !p.sinHora); const bs = ps.filter((p: IPartido) => p.codigo === r.b && !p.sinHora);
    as.forEach((pa: IPartido) => bs.forEach((pb: IPartido) => {
      if (pa.id === pb.id || !mismoDia(pa.fecha, pb.fecha)) { return; }
      if (!cumple(r.condA, pa) || !cumple(r.condB, pb)) { return; }
      const delta = Math.abs(pa.fecha.getTime() - pb.fecha.getTime()) / 60000;
      if (delta >= r.min) { return; }
      const k = [pa.id, pb.id].sort().join('-');
      if (vistos[k] && vistos[k]._min >= r.min) { return; }
      const [pri, seg] = pa.fecha <= pb.fecha ? [pa, pb] : [pb, pa];
      const falta = r.min - delta;
      const mover = !seg.local && pri.local ? pri : seg.local ? seg : pri.local ? pri : null;
      let accion: string;
      if (mover === seg) { accion = 'Retrasar ' + seg.equipo + ' a las ' + hm(minDia(pri.fecha) + r.min) + ' o más tarde, o adelantar ' + pri.equipo + '.'; }
      else if (mover === pri) { accion = 'Adelantar ' + pri.equipo + ' a las ' + hm(Math.max(0, minDia(seg.fecha) - r.min)) + ' o antes, o retrasar ' + seg.equipo + '.'; }
      else { accion = 'Los dos juegan fuera: hay que pedir cambio de hora a uno de los clubes rivales (faltan ' + durTxt(Math.round(falta)) + ').'; }
      vistos[k] = { _min: r.min, tipo: 'conflicto', grav: 'critica', partidos: [pri.id, seg.id],
        titulo: (r.motivo === 'Jugadores' ? 'Jugadores compartidos' : r.motivo === 'Entrenador' ? 'Entrenador compartido' : 'Conflicto') + ': ' + pri.equipo + ' y ' + seg.equipo,
        detalle: quien(pri) + ' y ' + quien(seg) + '. Necesitan ' + durTxt(r.min) + ' entre inicios y hay ' + durTxt(Math.round(delta)) + '.' + (r.titulo && !/↔/.test(r.titulo) ? ' ' + r.titulo + '.' : ''),
        accion };
    }));
  });
  Object.keys(vistos).forEach((k: string) => { const a = vistos[k]; out.push({ tipo: a.tipo, grav: a.grav, partidos: a.partidos, titulo: a.titulo, detalle: a.detalle, accion: a.accion }); });

  // 2. Capacidad de pistas y horario reservado (solo partidos en casa)
  const casa = ps.filter((p: IPartido) => p.local && !p.sinHora);
  const insts: { [k: string]: IPista[] } = {};
  pistas.filter((k: IPista) => k.controlar).forEach((k: IPista) => { (insts[k.inst] = insts[k.inst] || []).push(k); });
  Object.keys(insts).forEach((inst: string) => {
    const mods = insts[inst];
    const delInst = casa.filter((p: IPartido) => instalacionDe(p, mods) === inst);
    const porDia: { [k: string]: IPartido[] } = {};
    delInst.forEach((p: IPartido) => { const k = p.fecha.toDateString(); (porDia[k] = porDia[k] || []).push(p); });
    Object.keys(porDia).forEach((k: string) => {
      const lista = porDia[k].slice().sort((x: IPartido, y: IPartido) => x.fecha.getTime() - y.fecha.getTime());
      const dia = lista[0].fecha.getDay();
      const ventanas = (m: IPista): IVentana[] => m.disp.filter((v: IVentana) => v.dia === dia);
      // Fuera del horario reservado
      lista.forEach((p: IPartido) => {
        const s = minDia(p.fecha); const e = s + p.dur;
        const cabe = mods.some((m: IPista) => puedeUsar(p, m) && ventanas(m).some((v: IVentana) => s >= v.desde && e <= v.hasta));
        if (!cabe) {
          const res = mods.map((m: IPista) => ventanas(m).map((v: IVentana) => m.modulo + ' ' + hm(v.desde) + '–' + hm(v.hasta)).join(', ')).filter((x: string) => x).join(' · ');
          out.push({ tipo: 'reserva', grav: 'aviso', partidos: [p.id], titulo: 'Fuera del horario reservado en ' + inst,
            detalle: quien(p) + ' termina hacia las ' + hm(e) + '. ' + (res ? 'Reserva del ' + DIA_TXT[dia] + ': ' + res + '.' : 'No hay pista reservada ese día.'),
            accion: res ? 'Encajarlo en la reserva o pedir ampliación al Patronato.' : 'Pedir pista al Patronato para ese día.' });
        }
      });
      // Saturación: más partidos a la vez que módulos reservados
      lista.forEach((p: IPartido, i: number) => {
        const s = minDia(p.fecha);
        const simult = lista.filter((q: IPartido) => { const qs = minDia(q.fecha); return qs <= s && s < qs + q.dur + margen; });
        if (simult.length < 2 || lista.slice(0, i).some((q: IPartido) => minDia(q.fecha) === s)) { return; }
        const libres = mods.filter((m: IPista) => ventanas(m).some((v: IVentana) => s >= v.desde && s < v.hasta));
        const cap = Math.max(libres.length, 0);
        // Reparto de módulos respetando qué equipos pueden usar cada uno (los más restringidos primero)
        const ocup: { [k: string]: boolean } = {};
        const sinSitio = simult.slice().sort((x: IPartido, y: IPartido) => libres.filter((m: IPista) => puedeUsar(x, m)).length - libres.filter((m: IPista) => puedeUsar(y, m)).length)
          .filter((q: IPartido) => { const m = libres.filter((k: IPista) => !ocup[k.clave] && puedeUsar(q, k))[0]; if (m) { ocup[m.clave] = true; return false; } return true; });
        // Sin ningún módulo reservado a esa hora ya lo avisa «Fuera del horario reservado»: aquí solo se mira la capacidad reservada
        if (cap > 0 && sinSitio.length > 0 && simult.length > 1) {
          const ya = out.filter((a: IAlerta) => a.tipo === 'saturacion' && simult.every((q: IPartido) => a.partidos.indexOf(q.id) >= 0))[0];
          if (ya) { return; }
          out.push({ tipo: 'saturacion', grav: 'critica', partidos: simult.map((q: IPartido) => q.id), titulo: inst + ' saturado el ' + DIA_TXT[dia] + ' a las ' + hm(s),
            detalle: simult.length + ' partidos a la vez (' + simult.map((q: IPartido) => q.equipo + ' ' + hm(minDia(q.fecha)) + '–' + hm(minDia(q.fecha) + q.dur)).join(', ') + ') y el club solo tiene reservado' + (cap === 1 ? ' ' : 's ') + cap + ' módulo' + (cap === 1 ? '' : 's') + ' a esa hora (' + libres.map((m: IPista) => m.modulo + ((m.permitidos || []).length ? ' solo ' + (m.permitidos || []).join('/') : '')).join(', ') + ').' + (margen ? ' Se cuentan ' + margen + ' min de margen entre partidos.' : ''),
            accion: 'Escalonar las horas o mover alguno a otra instalación (p. ej. Gijón Sur).' });
        }
      });
    });
  });

  // 3. Franjas horarias de la FBPA (partidos en casa)
  casa.forEach((p: IPartido) => {
    const f = franjaDe(p);
    if (f.tipo === 'fuera') {
      out.push({ tipo: 'franja', grav: 'critica', partidos: [p.id], titulo: 'Fuera de franja FBPA: ' + p.equipo,
        detalle: quien(p) + ' no está en ninguna franja de «' + p.franja + '» para el ' + DIA_TXT[p.fecha.getDay()] + '.', accion: f.opciones ? 'Horas oficiales ese día: ' + f.opciones + '.' : 'Ese día no tiene franja oficial para esta categoría.' });
    } else if (f.tipo === 'conformidad' || f.tipo === 'condicion') {
      out.push({ tipo: 'franja', grav: 'aviso', partidos: [p.id], titulo: (f.tipo === 'conformidad' ? 'Requiere conformidad del rival: ' : 'Franja condicionada: ') + p.equipo,
        detalle: quien(p) + (f.cond ? ' · Condición: ' + f.cond + '.' : ' · Hora válida solo con acuerdo del rival.'), accion: f.opciones ? 'Horas oficiales ese día: ' + f.opciones + '.' : 'Confirmar con el rival por escrito.' });
    }
  });

  // 4. Hora por confirmar
  ps.filter((p: IPartido) => p.sinHora).forEach((p: IPartido) => out.push({ tipo: 'hora', grav: 'info', partidos: [p.id], titulo: 'Hora sin fijar: ' + p.equipo,
    detalle: 'La FBPA aún no publica hora para el partido ' + (p.local ? 'contra ' : 'en casa de ') + p.rival + '.', accion: p.local ? 'Fijar la hora y cargarla en la intranet.' : 'Pendiente del club rival.' }));

  const orden: { [g: string]: number } = { critica: 0, aviso: 1, info: 2 };
  return out.sort((a: IAlerta, b: IAlerta) => orden[a.grav] - orden[b.grav]);
}

/** Jugadores que juegan en varios equipos: el primer partido tiene que terminar con margen para llegar al segundo */
const MARGEN_MISMA = 30; const MARGEN_OTRA = 60;
function conflictosJugadores(ps: IPartido[], pistas: IPista[], jugs: IJugV[]): IAlerta[] {
  const out: IAlerta[] = [];
  const conHora = ps.filter((p: IPartido) => !p.sinHora).sort((a: IPartido, b: IPartido) => a.fecha.getTime() - b.fecha.getTime());
  const nombreCorto = (n: string): string => n.split(' ')[0] + (n.split(' ')[1] ? ' ' + n.split(' ')[1].charAt(0) + '.' : '');
  conHora.forEach((pa: IPartido, i: number) => conHora.slice(i + 1).forEach((pb: IPartido) => {
    if (pa.equipo === pb.equipo || !mismoDia(pa.fecha, pb.fecha)) { return; }
    const misma = pa.local && pb.local && instalacionDe(pa, pistas) !== '' && instalacionDe(pa, pistas) === instalacionDe(pb, pistas);
    const necesita = pa.dur + (misma ? MARGEN_MISMA : MARGEN_OTRA);
    const hay = (pb.fecha.getTime() - pa.fecha.getTime()) / 60000;
    if (hay >= necesita) { return; }
    const hab: string[] = []; const oca: string[] = [];
    jugs.forEach((j: IJugV) => {
      const h = [j.equipo].concat(j.habitual); const todo = h.concat(j.ocasional);
      if (h.indexOf(pa.equipo) >= 0 && h.indexOf(pb.equipo) >= 0) { hab.push(j.nombre); }
      else if (todo.indexOf(pa.equipo) >= 0 && todo.indexOf(pb.equipo) >= 0) { oca.push(j.nombre); }
    });
    if (!hab.length && !oca.length) { return; }
    const lista = (xs: string[]): string => xs.slice(0, 5).map(nombreCorto).join(', ') + (xs.length > 5 ? ' y ' + (xs.length - 5) + ' más' : '');
    const mover = pb.local ? 'Retrasar ' + pb.equipo + ' a las ' + hm(minDia(pa.fecha) + necesita) + ' o más tarde' : pa.local ? 'Adelantar ' + pa.equipo + ' a las ' + hm(Math.max(0, minDia(pb.fecha) - necesita)) + ' o antes' : 'Pedir cambio de hora a uno de los rivales';
    out.push({ tipo: 'conflicto', grav: hab.length ? 'critica' : 'aviso', partidos: [pa.id, pb.id],
      titulo: (hab.length ? 'Jugadores en dos equipos: ' : 'Posible convocatoria doble: ') + pa.equipo + ' y ' + pb.equipo,
      detalle: quien(pa) + ' y ' + quien(pb) + '. ' + (hab.length ? hab.length + ' juega' + (hab.length === 1 ? '' : 'n') + ' habitualmente en los dos (' + lista(hab) + ')' : '') +
        (hab.length && oca.length ? '; ' : '') + (oca.length ? oca.length + ' convocatoria' + (oca.length === 1 ? '' : 's') + ' ocasional' + (oca.length === 1 ? '' : 'es') + ' (' + lista(oca) + ')' : '') +
        '. Hace falta ' + durTxt(necesita) + ' entre inicios (' + durTxt(pa.dur) + ' de partido + ' + (misma ? MARGEN_MISMA + ' min en el mismo pabellón' : MARGEN_OTRA + ' min de desplazamiento') + ') y hay ' + durTxt(Math.round(hay)) + '.',
      accion: mover + (oca.length && !hab.length ? ', o confirmar que esta semana no se les convoca.' : '.') });
  }));
  return out;
}

/* ---------- Propuestas: huecos válidos para mover un partido ---------- */
export interface IHueco { fecha: Date; pista: string; tipoFranja: string; puntos: number; nota: string; }

function franjasDelDia(p: IPartido, dia: number, soloPref?: boolean): number[] {
  const fr = FRANJAS[p.franja];
  const out: number[] = [];
  if (!fr) { for (let t = 10 * 60; t <= 20 * 60; t += 30) { out.push(t); } return out; }
  fr.filter((f: IFranja) => f.dias.indexOf(dia) >= 0 && f.tipo === 'oficial' && (!soloPref || f.pref !== false)).forEach((f: IFranja) => { for (let t = aMin(f.desde); t <= aMin(f.hasta); t += 15) { if (out.indexOf(t) < 0) { out.push(t); } } });
  return out;
}

/**
 * Busca horas alternativas para un partido sin crear conflictos nuevos.
 * En casa: dentro de las reservas de sus pistas preferentes o alternativas. Fuera: dentro de la franja oficial (para pedirlo al rival).
 */
export function sugerir(p: IPartido, semana: IPartido[], reglas: IRegla[], pistas: IPista[], margen: number, jugs: IJugV[], prefs: string[], alts: string[], asign: { [id: number]: string }): IHueco[] {
  const base = new Date(p.fecha.getFullYear(), p.fecha.getMonth(), p.fecha.getDate());
  const dias: Date[] = [base];
  const lun = new Date(base); lun.setDate(base.getDate() - ((base.getDay() + 6) % 7));
  [new Date(lun.getFullYear(), lun.getMonth(), lun.getDate() + 5), new Date(lun.getFullYear(), lun.getMonth(), lun.getDate() + 6)].forEach((d: Date) => { if (d.toDateString() !== base.toDateString()) { dias.push(d); } });
  const otros = semana.filter((q: IPartido) => q.id !== p.id);
  const vistos: { [k: string]: boolean } = {};
  const out: IHueco[] = [];
  const valido = (cand: IPartido): { ok: boolean; oca: boolean } => {
    const delDia = otros.filter((q: IPartido) => q.fecha.toDateString() === cand.fecha.toDateString());
    const al = analizar([cand].concat(delDia), reglas, pistas, margen, jugs).filter((a: IAlerta) => a.tipo === 'conflicto' && a.partidos.indexOf(cand.id) >= 0);
    return { ok: !al.some((a: IAlerta) => a.grav === 'critica'), oca: al.some((a: IAlerta) => a.grav === 'aviso') };
  };
  const original = minDia(p.fecha);
  dias.forEach((d: Date) => {
    if (p.local) {
      const claves = prefs.concat(alts.filter((a: string) => prefs.indexOf(a) < 0));
      claves.forEach((clave: string) => {
        const pi = pistas.filter((k: IPista) => norm(k.clave) === norm(clave))[0]; if (!pi || !puedeUsar(p, pi)) { return; }
        pi.disp.filter((v: IVentana) => v.dia === d.getDay()).forEach((v: IVentana) => {
          for (let t = v.desde; t + p.dur <= v.hasta; t += 15) {
            const f = new Date(d.getFullYear(), d.getMonth(), d.getDate(), Math.floor(t / 60), t % 60);
            const cand: IPartido = { ...p, fecha: f, lugar: pi.alias };
            const fr = franjaDe(cand); if (fr.tipo !== 'oficial' && fr.tipo !== 'n/a') { continue; } // solo franja oficial FBPA
            const ocupada = otros.some((q: IPartido) => q.local && asign[q.id] && norm(asign[q.id]) === norm(pi.clave) && q.fecha.toDateString() === f.toDateString() &&
              t < minDia(q.fecha) + q.dur + margen && minDia(q.fecha) < t + p.dur + margen);
            if (ocupada) { continue; }
            const v2 = valido(cand); if (!v2.ok) { continue; }
            const k = f.getTime() + '|' + pi.clave; if (vistos[k]) { continue; } vistos[k] = true;
            let pts = 100 - Math.min(35, Math.abs(t - original) / 15 * 1.5) - (d.toDateString() === base.toDateString() ? 0 : 40) - (prefs.indexOf(clave) >= 0 ? 0 : 15) - (fr.pref === false ? 15 : 0) - (v2.oca ? 10 : 0);
            pts = Math.round(pts);
            out.push({ fecha: f, pista: pi.clave, tipoFranja: fr.tipo, puntos: pts,
              nota: (prefs.indexOf(clave) >= 0 ? 'Pista habitual' : 'Pista alternativa') + (fr.pref === false ? ' · franja secundaria' : '') + (v2.oca ? ' · coincide con convocatorias ocasionales' : '') });
          }
        });
      });
    } else {
      const habit = franjasDelDia(p, d.getDay(), true);
      franjasDelDia(p, d.getDay()).forEach((t: number) => {
        const f = new Date(d.getFullYear(), d.getMonth(), d.getDate(), Math.floor(t / 60), t % 60);
        const k = String(f.getTime()); if (vistos[k] || f.getTime() === p.fecha.getTime()) { return; } vistos[k] = true;
        const cand: IPartido = { ...p, fecha: f };
        const v2 = valido(cand); if (!v2.ok) { return; }
        const sec = habit.indexOf(t) < 0;
        out.push({ fecha: f, pista: '', tipoFranja: sec ? 'secundaria' : 'oficial', puntos: Math.round(100 - Math.min(35, Math.abs(t - original) / 15 * 1.5) - (d.toDateString() === base.toDateString() ? 0 : 40) - (sec ? 15 : 0) - (v2.oca ? 10 : 0)),
          nota: (sec ? 'Franja secundaria' : 'Hora oficial de la franja') + (v2.oca ? ' · coincide con convocatorias ocasionales' : '') + ' · a pedir al club local' });
      });
    }
  });
  // Variedad: como mucho una opción por día, hora en punto y pista
  const usados: { [k: string]: boolean } = {};
  return out.sort((a: IHueco, b: IHueco) => b.puntos - a.puntos).filter((h: IHueco) => { const k = h.fecha.toDateString() + '|' + h.fecha.getHours() + '|' + h.pista; if (usados[k]) { return false; } usados[k] = true; return true; }).slice(0, 6);
}

/** Asigna módulo a los partidos en casa que no lo tienen: primero pistas habituales, luego alternativas, sin solapes */
export function asignarPistas(semana: IPartido[], pistas: IPista[], margen: number, prefsDe: (p: IPartido) => { pref: string[]; alt: string[] }, asign: { [id: number]: string }): { [id: number]: string } {
  const res: { [id: number]: string } = { ...asign };
  const casa = semana.filter((p: IPartido) => p.local && !p.sinHora).sort((a: IPartido, b: IPartido) => a.fecha.getTime() - b.fecha.getTime());
  casa.forEach((p: IPartido) => {
    if (res[p.id]) { return; }
    const s = minDia(p.fecha); const pr = prefsDe(p);
    const inst = norm(p.lugar);
    const cands = pr.pref.concat(pr.alt).map((c: string) => pistas.filter((k: IPista) => norm(k.clave) === norm(c))[0]).filter((k: IPista | undefined) => !!k) as IPista[];
    const mismos = pistas.filter((k: IPista) => norm(k.alias) === inst || norm(k.inst) === inst);
    const lista = cands.filter((k: IPista) => norm(k.inst) === inst || norm(k.alias) === inst).concat(mismos.filter((k: IPista) => cands.indexOf(k) < 0)).filter((k: IPista) => puedeUsar(p, k));
    const libre = (k: IPista): boolean => !casa.some((q: IPartido) => q.id !== p.id && res[q.id] === k.clave && q.fecha.toDateString() === p.fecha.toDateString() && s < minDia(q.fecha) + q.dur + margen && minDia(q.fecha) < s + p.dur + margen);
    const enReserva = (k: IPista): boolean => k.disp.some((v: IVentana) => v.dia === p.fecha.getDay() && s >= v.desde && s + p.dur <= v.hasta);
    const elegida = lista.filter((k: IPista) => libre(k) && enReserva(k))[0] || lista.filter(libre)[0] || lista[0];
    if (elegida) { res[p.id] = elegida.clave; }
  });
  return res;
}

/* =====================================================================
   Plazo ordinario: el local fija o ajusta la hora sin conformidad del rival
   hasta el viernes a las 11:00 de la semana anterior (configurable).
   ===================================================================== */
export interface IPlazo { dia: number; hora: number; } // dia: 5 = viernes; hora en minutos
export const PLAZO_DEF: IPlazo = { dia: 5, hora: 11 * 60 };
export function limitePlazo(p: IPartido, pl?: IPlazo): Date {
  const x = pl || PLAZO_DEF;
  const d = new Date(p.fecha.getFullYear(), p.fecha.getMonth(), p.fecha.getDate());
  const lun = new Date(d); lun.setDate(d.getDate() - ((d.getDay() + 6) % 7)); // lunes de la semana del partido
  const lim = new Date(lun); lim.setDate(lun.getDate() - 7 + ((x.dia + 6) % 7)); // ese día de la semana anterior
  lim.setHours(Math.floor(x.hora / 60), x.hora % 60, 0, 0);
  return lim;
}
export function enPlazo(p: IPartido, pl?: IPlazo, ahora?: Date): boolean { return p.local && (ahora || new Date()).getTime() < limitePlazo(p, pl).getTime(); }

/* =====================================================================
   Motor de escenarios: resuelve la semana completa a la vez
   ===================================================================== */
export interface IPeticion { fecha: Date; nota: string; de: string; }
export interface ICand { fecha: Date; lugar: string; pista: string; tipoFranja: string; pdm?: boolean; }
export interface ICambio { id: number; desde: Date; hasta: Date; lugar: string; pista: string; pideRival: boolean; enPlazo: boolean; cambiaDia: boolean; peticion: boolean; tipoFranja: string; pdm: boolean; }
export interface IEscenario { clave: string; nombre: string; desc: string; cambios: ICambio[]; alertas: IAlerta[]; criticas: number; avisos: number; holguraMin: number; coste: number; partidos: IPartido[]; }
export interface IResolverOpts {
  reglas: IRegla[]; pistas: IPista[]; margen: number; jugs: IJugV[];
  prefsDe: (p: IPartido) => { pref: string[]; alt: string[] };
  fijos: { [id: number]: boolean };            // inamovibles o ya en trámite
  forzados: { [id: number]: ICand };           // coordinación ha elegido esta hora: se respeta
  peticiones: { [id: number]: IPeticion };
  plazo?: IPlazo; ahora?: Date;
}

/** Parejas de partidos que necesitan separación (reglas de entrenadores y jugadores habituales) */
interface IPar { a: number; b: number; req: number; }
function pares(ps: IPartido[], reglas: IRegla[], pistas: IPista[], jugs: IJugV[]): IPar[] {
  const out: IPar[] = []; const k: { [s: string]: IPar } = {};
  const add = (a: IPartido, b: IPartido, req: number): void => { const s = Math.min(a.id, b.id) + '-' + Math.max(a.id, b.id); if (!k[s] || k[s].req < req) { k[s] = { a: a.id, b: b.id, req }; } };
  reglas.forEach((r: IRegla) => ps.filter((p: IPartido) => p.codigo === r.a && cumple(r.condA, p)).forEach((pa: IPartido) => ps.filter((q: IPartido) => q.codigo === r.b && q.id !== pa.id && cumple(r.condB, q)).forEach((pb: IPartido) => add(pa, pb, r.min))));
  ps.forEach((pa: IPartido, i: number) => ps.slice(i + 1).forEach((pb: IPartido) => {
    if (pa.equipo === pb.equipo) { return; }
    if (jugs.some((j: IJugV) => { const h = [j.equipo].concat(j.habitual); return h.indexOf(pa.equipo) >= 0 && h.indexOf(pb.equipo) >= 0; })) { add(pa, pb, Math.max(pa.dur, pb.dur) + MARGEN_OTRA); }
  }));
  Object.keys(k).forEach((s: string) => out.push(k[s]));
  return out;
}

/** Horas posibles para un partido: en casa, dentro de las reservas de sus pistas; fuera, horas oficiales de la franja (a pedir al rival) */
function candidatos(p: IPartido, pistas: IPista[], pref: string[], alt: string[], soloMismoDia: boolean): ICand[] {
  if (!FRANJAS[p.franja]) { return []; } // sin franjas cargadas (p. ej. competición nacional): el motor no lo mueve, solo a mano
  const base = new Date(p.fecha.getFullYear(), p.fecha.getMonth(), p.fecha.getDate());
  const dias: Date[] = [base];
  if (!soloMismoDia) {
    const lun = new Date(base); lun.setDate(base.getDate() - ((base.getDay() + 6) % 7));
    [5, 6].forEach((o: number) => { const d = new Date(lun.getFullYear(), lun.getMonth(), lun.getDate() + o); if (d.toDateString() !== base.toDateString()) { dias.push(d); } });
  }
  const out: ICand[] = []; const visto: { [k: string]: boolean } = {};
  dias.forEach((d: Date) => {
    if (p.local) {
      const claves = pref.concat(alt.filter((a: string) => pref.indexOf(a) < 0));
      // 1) Módulos reservados de sus pistas habituales y alternativas, solo en franja oficial
      claves.forEach((clave: string) => {
        const pi = pistas.filter((k: IPista) => norm(k.clave) === norm(clave))[0]; if (!pi || !puedeUsar(p, pi)) { return; }
        pi.disp.filter((v: IVentana) => v.dia === d.getDay()).forEach((v: IVentana) => {
          for (let t = v.desde; t + p.dur <= v.hasta; t += 15) {
            const f = new Date(d.getFullYear(), d.getMonth(), d.getDate(), Math.floor(t / 60), t % 60);
            const fr = franjaDe({ ...p, fecha: f }); if (fr.tipo !== 'oficial') { continue; }
            const k = f.getTime() + '|' + norm(pi.inst); if (visto[k]) { continue; } visto[k] = true;
            out.push({ fecha: f, lugar: pi.alias, pista: pi.clave, tipoFranja: fr.pref === false ? 'secundaria' : 'oficial' });
          }
        });
      });
      // 2) Último recurso: pedir al PDM un módulo en sus instalaciones a una hora de franja (fuera de la reserva del club), el mismo día
      if (d.toDateString() === base.toDateString()) {
        const insts: { [k: string]: IPista } = {};
        pistas.filter((k: IPista) => k.inst && puedeUsar(p, k)).forEach((k: IPista) => { if (!insts[norm(k.inst)]) { insts[norm(k.inst)] = k; } });
        Object.keys(insts).forEach((ki: string) => {
          const k0 = insts[ki];
          franjasDelDia(p, d.getDay(), true).forEach((t: number) => {
            const f = new Date(d.getFullYear(), d.getMonth(), d.getDate(), Math.floor(t / 60), t % 60);
            const kk = f.getTime() + '|' + ki; if (visto[kk]) { return; } visto[kk] = true;
            out.push({ fecha: f, lugar: k0.alias, pista: '', tipoFranja: 'oficial', pdm: true });
          });
        });
      }
    } else {
      const pref = franjasDelDia(p, d.getDay(), true);
      franjasDelDia(p, d.getDay()).forEach((t: number) => {
        const f = new Date(d.getFullYear(), d.getMonth(), d.getDate(), Math.floor(t / 60), t % 60);
        const k = String(f.getTime()); if (visto[k]) { return; } visto[k] = true;
        out.push({ fecha: f, lugar: p.lugar, pista: '', tipoFranja: pref.indexOf(t) >= 0 ? 'oficial' : 'secundaria' });
      });
    }
  });
  return out;
}

const ESCENARIOS: { clave: string; nombre: string; desc: string; wCambio: number; wAviso: number; wHolgura: number; soloCasa: boolean; soloMismoDia: boolean }[] = [
  { clave: 'menos', nombre: 'Menos cambios', desc: 'Resuelve los conflictos tocando el menor número de partidos, primero los de casa dentro de plazo.', wCambio: 12, wAviso: 3, wHolgura: 0.3, soloCasa: false, soloMismoDia: false },
  { clave: 'holgura', nombre: 'Más holgura', desc: 'Busca el mayor margen entre partidos que comparten entrenador o jugadores, aunque haya que mover alguno más.', wCambio: 4, wAviso: 3, wHolgura: 3, soloCasa: false, soloMismoDia: false },
  { clave: 'casa', nombre: 'Solo partidos en casa', desc: 'Solo mueve partidos en casa, el mismo día: no depende de ningún club rival.', wCambio: 10, wAviso: 3, wHolgura: 0.3, soloCasa: true, soloMismoDia: true }
];
export const NOMBRES_ESCENARIO = ESCENARIOS.map((e) => ({ clave: e.clave, nombre: e.nombre, desc: e.desc }));

/** Resuelve la semana para cada escenario. Búsqueda en haz: en cada paso mueve un partido de los que siguen en conflicto. */
export function resolver(semana: IPartido[], o: IResolverOpts): IEscenario[] {
  const ahora = o.ahora || new Date();
  const conHora = semana.filter((p: IPartido) => !p.sinHora);
  // Punto de partida: horas vigentes + lo que coordinación ha forzado
  const inicio: { [id: number]: ICand } = {};
  conHora.forEach((p: IPartido) => { inicio[p.id] = o.forzados[p.id] || { fecha: p.fecha, lugar: p.lugar, pista: '', tipoFranja: '' }; });
  const orig: { [id: number]: IPartido } = {}; conHora.forEach((p: IPartido) => { orig[p.id] = p; });
  const relacion = pares(conHora, o.reglas, o.pistas, o.jugs);

  return ESCENARIOS.map((esc) => {
    const movible = (p: IPartido): boolean => !o.fijos[p.id] && !o.forzados[p.id] && (!esc.soloCasa || p.local);
    const cands: { [id: number]: ICand[] } = {};
    conHora.filter(movible).forEach((p: IPartido) => {
      const pr = o.prefsDe(p); let cs = candidatos(p, o.pistas, pr.pref, pr.alt, esc.soloMismoDia);
      const pet = o.peticiones[p.id];
      if (pet && !cs.some((c: ICand) => c.fecha.getTime() === pet.fecha.getTime())) { cs = cs.concat([{ fecha: pet.fecha, lugar: p.lugar, pista: '', tipoFranja: '' }]); }
      cands[p.id] = cs;
    });
    const memo: { [k: string]: { crit: number; av: number; al: IAlerta[] } } = {};
    const materializa = (asig: { [id: number]: ICand }): IPartido[] => conHora.map((p: IPartido) => ({ ...p, fecha: asig[p.id].fecha, lugar: asig[p.id].lugar }));
    const evalua = (asig: { [id: number]: ICand }): { coste: number; crit: number; av: number; hol: number; minHol: number } => {
      const ps = materializa(asig);
      const porDia: { [d: string]: IPartido[] } = {};
      ps.forEach((p: IPartido) => { const k = p.fecha.toDateString(); (porDia[k] = porDia[k] || []).push(p); });
      let crit = 0; let av = 0;
      Object.keys(porDia).forEach((d: string) => {
        const key = d + '#' + porDia[d].map((p: IPartido) => p.id + '@' + p.fecha.getTime() + '@' + p.lugar).sort().join(',');
        if (!memo[key]) { const al = analizar(porDia[d], o.reglas, o.pistas, o.margen, o.jugs); memo[key] = { crit: al.filter((a: IAlerta) => a.grav === 'critica').length, av: al.filter((a: IAlerta) => a.grav === 'aviso').length, al }; }
        crit += memo[key].crit; av += memo[key].av;
      });
      let cambios = 0; let pet = 0;
      ps.forEach((p: IPartido) => {
        const p0 = orig[p.id]; const pe = o.peticiones[p.id];
        if (pe && p.fecha.getTime() === pe.fecha.getTime()) { pet++; }
        if (p.fecha.getTime() === p0.fecha.getTime() && norm(p.lugar) === norm(p0.lugar)) { return; }
        if (o.forzados[p.id]) { return; }
        let w = p.local ? (enPlazo(p0, o.plazo, ahora) ? 1 : 2.5) : 4;
        const ca = asig[p.id]; if (ca && ca.pdm) { w += 6; } if (ca && ca.tipoFranja === 'secundaria') { w += 3; }
        if (p.fecha.toDateString() !== p0.fecha.toDateString()) { w += 2; }
        w += Math.abs(p.fecha.getTime() - p0.fecha.getTime()) / 3600000 / 4;
        if (pe && p.fecha.getTime() === pe.fecha.getTime()) { w -= 0.8; }
        cambios += w;
      });
      let hol = 0; let minHol = 9999;
      const byId: { [id: number]: IPartido } = {}; ps.forEach((p: IPartido) => { byId[p.id] = p; });
      relacion.forEach((r: IPar) => {
        const a = byId[r.a]; const b = byId[r.b]; if (!a || !b || a.fecha.toDateString() !== b.fecha.toDateString()) { return; }
        const s = Math.abs(a.fecha.getTime() - b.fecha.getTime()) / 60000 - r.req;
        if (s >= 0) { hol += Math.min(s, 120) / 30; minHol = Math.min(minHol, s); }
      });
      const coste = crit * 1000 + av * esc.wAviso + cambios * esc.wCambio - hol * esc.wHolgura - pet * 6;
      return { coste, crit, av, hol, minHol };
    };
    const clave = (asig: { [id: number]: ICand }): string => Object.keys(asig).map((k: string) => k + ':' + asig[+k].fecha.getTime() + asig[+k].lugar).join('|');
    interface IEst { asig: { [id: number]: ICand }; r: { coste: number; crit: number; av: number; hol: number; minHol: number }; }
    const e0: IEst = { asig: { ...inicio }, r: evalua(inicio) };
    let haz: IEst[] = [e0]; let mejor = e0; const vistos: { [k: string]: boolean } = {}; vistos[clave(inicio)] = true;
    const ANCHO = 10; const PROF = 5;
    for (let paso = 0; paso < PROF; paso++) {
      const sig: IEst[] = [];
      haz.forEach((st: IEst) => {
        const ps = materializa(st.asig);
        const al = analizar(ps, o.reglas, o.pistas, o.margen, o.jugs);
        let obj: number[] = [];
        al.filter((a: IAlerta) => a.grav === 'critica' || (a.grav === 'aviso' && a.tipo !== 'hora')).forEach((a: IAlerta) => a.partidos.forEach((id: number) => { if (obj.indexOf(id) < 0) { obj.push(id); } }));
        if (esc.wHolgura > 1) { relacion.forEach((r: IPar) => { [r.a, r.b].forEach((id: number) => { if (obj.indexOf(id) < 0) { obj.push(id); } }); }); }
        Object.keys(o.peticiones).forEach((k: string) => { if (obj.indexOf(+k) < 0) { obj.push(+k); } });
        obj = obj.filter((id: number) => !!cands[id]);
        obj.forEach((id: number) => cands[id].forEach((c: ICand) => {
          const cur = st.asig[id]; if (cur.fecha.getTime() === c.fecha.getTime() && norm(cur.lugar) === norm(c.lugar)) { return; }
          const asig = { ...st.asig }; asig[id] = c;
          const k = clave(asig); if (vistos[k]) { return; } vistos[k] = true;
          sig.push({ asig, r: evalua(asig) });
        }));
      });
      if (!sig.length) { break; }
      sig.sort((a: IEst, b: IEst) => a.r.coste - b.r.coste);
      haz = sig.slice(0, ANCHO);
      if (haz[0].r.coste < mejor.r.coste) { mejor = haz[0]; }
    }
    const final = materializa(mejor.asig);
    const alertas = analizar(final, o.reglas, o.pistas, o.margen, o.jugs);
    const cambios: ICambio[] = final.filter((p: IPartido) => { const p0 = orig[p.id]; return p.fecha.getTime() !== p0.fecha.getTime() || norm(p.lugar) !== norm(p0.lugar); }).map((p: IPartido): ICambio => {
      const p0 = orig[p.id]; const c = mejor.asig[p.id]; const pe = o.peticiones[p.id];
      return { id: p.id, desde: p0.fecha, hasta: p.fecha, lugar: p.lugar, pista: c.pista, pideRival: !p.local || !enPlazo(p0, o.plazo, ahora) || c.tipoFranja === 'conformidad' || c.tipoFranja === 'condicion', pdm: !!c.pdm,
        enPlazo: enPlazo(p0, o.plazo, ahora), cambiaDia: p.fecha.toDateString() !== p0.fecha.toDateString(), peticion: !!pe && pe.fecha.getTime() === p.fecha.getTime(), tipoFranja: c.tipoFranja };
    }).sort((a: ICambio, b: ICambio) => a.hasta.getTime() - b.hasta.getTime());
    return { clave: esc.clave, nombre: esc.nombre, desc: esc.desc, cambios, alertas, criticas: mejor.r.crit, avisos: mejor.r.av, holguraMin: mejor.r.minHol === 9999 ? -1 : Math.round(mejor.r.minHol), coste: mejor.r.coste,
      partidos: semana.map((p: IPartido) => { const a = mejor.asig[p.id]; return a ? { ...p, fecha: a.fecha, lugar: a.lugar } : p; }) };
  });
}
