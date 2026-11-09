/* Detección de conflictos sobre el horario oficial de la FBPA (fase 1) */

export interface IPartido { id: number; equipo: string; codigo: string; fecha: Date; local: boolean; rival: string; lugar: string; sinHora: boolean; dur: number; franja: string; }
export interface IRegla { id: number; titulo: string; a: string; condA: string; b: string; condB: string; min: number; motivo: string; }
export interface IVentana { dia: number; desde: number; hasta: number; }
export interface IPista { id: number; clave: string; inst: string; modulo: string; alias: string; disp: IVentana[]; dispTexto: string; controlar: boolean; orden: number; pdm?: string; }
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
interface IFranja { dias: number[]; desde: string; hasta: string; tipo: 'oficial' | 'conformidad' | 'condicion'; cond?: string; }
const F = (dias: number[], desde: string, hasta: string, tipo: IFranja['tipo'], cond?: string): IFranja => ({ dias, desde, hasta, tipo, cond });
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

function franjaDe(p: IPartido): { tipo: 'oficial' | 'conformidad' | 'condicion' | 'fuera' | 'n/a'; cond?: string; opciones: string } {
  const fr = FRANJAS[p.franja];
  if (!fr) { return { tipo: 'n/a', opciones: '' }; }
  const d = p.fecha.getDay(); const t = minDia(p.fecha);
  const del = fr.filter((f: IFranja) => f.dias.indexOf(d) >= 0);
  const opciones = del.filter((f: IFranja) => f.tipo === 'oficial').map((f: IFranja) => f.desde === f.hasta ? f.desde : f.desde + '–' + f.hasta).join(', ');
  const en = del.filter((f: IFranja) => t >= aMin(f.desde) && t <= aMin(f.hasta));
  const of = en.filter((f: IFranja) => f.tipo === 'oficial')[0]; if (of) { return { tipo: 'oficial', opciones }; }
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
        const cabe = mods.some((m: IPista) => ventanas(m).some((v: IVentana) => s >= v.desde && e <= v.hasta));
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
        if (simult.length > cap) {
          const ya = out.filter((a: IAlerta) => a.tipo === 'saturacion' && simult.every((q: IPartido) => a.partidos.indexOf(q.id) >= 0))[0];
          if (ya) { return; }
          out.push({ tipo: 'saturacion', grav: 'critica', partidos: simult.map((q: IPartido) => q.id), titulo: inst + ' saturado el ' + DIA_TXT[dia] + ' a las ' + hm(s),
            detalle: simult.length + ' partidos a la vez (' + simult.map((q: IPartido) => q.equipo + ' ' + hm(minDia(q.fecha))).join(', ') + ') y ' + (cap ? 'solo ' + cap + ' módulo' + (cap === 1 ? '' : 's') + ' reservado' + (cap === 1 ? '' : 's') + ' (' + libres.map((m: IPista) => m.modulo).join(', ') + ')' : 'ningún módulo reservado a esa hora') + '. Se cuenta ' + margen + ' min de margen entre partidos.',
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

function franjasDelDia(p: IPartido, dia: number): number[] {
  const fr = FRANJAS[p.franja];
  const out: number[] = [];
  if (!fr) { for (let t = 10 * 60; t <= 20 * 60; t += 30) { out.push(t); } return out; }
  fr.filter((f: IFranja) => f.dias.indexOf(dia) >= 0 && f.tipo === 'oficial').forEach((f: IFranja) => { for (let t = aMin(f.desde); t <= aMin(f.hasta); t += 15) { out.push(t); } });
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
        const pi = pistas.filter((k: IPista) => norm(k.clave) === norm(clave))[0]; if (!pi) { return; }
        pi.disp.filter((v: IVentana) => v.dia === d.getDay()).forEach((v: IVentana) => {
          for (let t = v.desde; t + p.dur <= v.hasta; t += 15) {
            const f = new Date(d.getFullYear(), d.getMonth(), d.getDate(), Math.floor(t / 60), t % 60);
            const cand: IPartido = { ...p, fecha: f, lugar: pi.alias };
            const fr = franjaDe(cand); if (fr.tipo === 'fuera') { continue; }
            const ocupada = otros.some((q: IPartido) => q.local && asign[q.id] && norm(asign[q.id]) === norm(pi.clave) && q.fecha.toDateString() === f.toDateString() &&
              t < minDia(q.fecha) + q.dur + margen && minDia(q.fecha) < t + p.dur + margen);
            if (ocupada) { continue; }
            const v2 = valido(cand); if (!v2.ok) { continue; }
            const k = f.getTime() + '|' + pi.clave; if (vistos[k]) { continue; } vistos[k] = true;
            let pts = 100 - Math.min(35, Math.abs(t - original) / 15 * 1.5) - (d.toDateString() === base.toDateString() ? 0 : 40) - (prefs.indexOf(clave) >= 0 ? 0 : 15) - (fr.tipo === 'oficial' || fr.tipo === 'n/a' ? 0 : 20) - (v2.oca ? 10 : 0);
            pts = Math.round(pts);
            out.push({ fecha: f, pista: pi.clave, tipoFranja: fr.tipo, puntos: pts,
              nota: (prefs.indexOf(clave) >= 0 ? 'Pista habitual' : 'Pista alternativa') + (fr.tipo === 'conformidad' ? ' · requiere conformidad del rival' : fr.tipo === 'condicion' ? ' · franja condicionada' : '') + (v2.oca ? ' · coincide con convocatorias ocasionales' : '') });
          }
        });
      });
    } else {
      franjasDelDia(p, d.getDay()).forEach((t: number) => {
        const f = new Date(d.getFullYear(), d.getMonth(), d.getDate(), Math.floor(t / 60), t % 60);
        const k = String(f.getTime()); if (vistos[k] || f.getTime() === p.fecha.getTime()) { return; } vistos[k] = true;
        const cand: IPartido = { ...p, fecha: f };
        const v2 = valido(cand); if (!v2.ok) { return; }
        out.push({ fecha: f, pista: '', tipoFranja: 'oficial', puntos: Math.round(100 - Math.min(35, Math.abs(t - original) / 15 * 1.5) - (d.toDateString() === base.toDateString() ? 0 : 40) - (v2.oca ? 10 : 0)),
          nota: 'Hora oficial de la franja' + (v2.oca ? ' · coincide con convocatorias ocasionales' : '') + ' · a pedir al club local' });
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
    const lista = cands.filter((k: IPista) => norm(k.inst) === inst || norm(k.alias) === inst).concat(mismos.filter((k: IPista) => cands.indexOf(k) < 0));
    const libre = (k: IPista): boolean => !casa.some((q: IPartido) => q.id !== p.id && res[q.id] === k.clave && q.fecha.toDateString() === p.fecha.toDateString() && s < minDia(q.fecha) + q.dur + margen && minDia(q.fecha) < s + p.dur + margen);
    const enReserva = (k: IPista): boolean => k.disp.some((v: IVentana) => v.dia === p.fecha.getDay() && s >= v.desde && s + p.dur <= v.hasta);
    const elegida = lista.filter((k: IPista) => libre(k) && enReserva(k))[0] || lista.filter(libre)[0] || lista[0];
    if (elegida) { res[p.id] = elegida.clave; }
  });
  return res;
}
