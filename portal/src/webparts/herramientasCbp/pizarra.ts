import { esc } from '../../common/cbp';
import { abrirModal, cerrarModal, aviso, descargar } from '../../common/ui';
import { subirAMateriales } from '../materialesCbp/MaterialesCbpWebPart';
import { ICtx } from './datos';

type Herr = 'mover' | 'at' | 'df' | 'bal' | 'cono' | 'mov' | 'pase' | 'bote' | 'bloq' | 'borrar';
interface IFicha { t: 'at' | 'df' | 'bal' | 'cono'; x: number; y: number; n?: number; }
interface ILinea { t: 'mov' | 'pase' | 'bote' | 'bloq'; x1: number; y1: number; x2: number; y2: number; }
type Obj = IFicha | ILinea;
const esLinea = (o: Obj): o is ILinea => (o as ILinea).x2 !== undefined;

let modo: 'media' | 'completa' = 'media';
let objs: Obj[] = [];
let herr: Herr = 'at';

const PAD = 24;
function dims(): { w: number; h: number; s: number } { return modo === 'media' ? { w: 15 * 50 + PAD * 2, h: 14 * 50 + PAD * 2, s: 50 } : { w: 28 * 36 + PAD * 2, h: 15 * 36 + PAD * 2, s: 36 }; }

const HERRS: { k: Herr; t: string; ico: string }[] = [
  { k: 'mover', t: 'Mover', ico: '✥' }, { k: 'at', t: 'Atacante', ico: '●' }, { k: 'df', t: 'Defensor', ico: '✕' }, { k: 'bal', t: 'Balón', ico: '◍' }, { k: 'cono', t: 'Cono', ico: '▲' },
  { k: 'mov', t: 'Desplazamiento', ico: '→' }, { k: 'pase', t: 'Pase', ico: '⇢' }, { k: 'bote', t: 'Bote', ico: '⤳' }, { k: 'bloq', t: 'Bloqueo', ico: '⊣' }, { k: 'borrar', t: 'Borrar', ico: '⌫' }
];

export function renderPizarra(el: HTMLElement, ctx: ICtx): void {
  const d = dims();
  el.innerHTML = '<div class="cbp-pz-wrap">' +
    '<div class="cbp-pz-tools" role="toolbar" aria-label="Herramientas de la pizarra">' +
    HERRS.map((h: { k: Herr; t: string; ico: string }) => '<button type="button" class="cbp-pz-tool" data-h="' + h.k + '" aria-pressed="' + (herr === h.k ? 'true' : 'false') + '" title="' + h.t + '"><span aria-hidden="true">' + h.ico + '</span>' + h.t + '</button>').join('') +
    '<span class="cbp-pz-sep"></span><button type="button" class="cbp-pz-tool" data-acc="deshacer">↶ Deshacer</button><button type="button" class="cbp-pz-tool" data-acc="limpiar">Limpiar</button>' +
    '<button type="button" class="cbp-pz-tool" data-acc="modo">' + (modo === 'media' ? 'Ver pista completa' : 'Ver media pista') + '</button></div>' +
    '<canvas class="cbp-pz-canvas" width="' + d.w + '" height="' + d.h + '" aria-label="Pizarra táctica. Usa las herramientas para colocar jugadores y dibujar movimientos." role="img" style="max-width:' + (modo === 'media' ? 760 : 1100) + 'px"></canvas>' +
    '<p style="margin:0;font-size:13px;color:#5c5c5c">Pulsa en la pista para colocar fichas. Para las flechas, arrastra desde el origen hasta el destino. Con «Mover» puedes recolocar fichas.</p>' +
    '<div class="cbp-btns" style="justify-content:flex-start"><button type="button" class="cbp-btn" data-acc="png">Descargar imagen</button><button type="button" class="cbp-btn pri" data-acc="guardar">Guardar en Materiales</button></div></div>';
  const cv = el.querySelector('canvas') as HTMLCanvasElement;
  const g = cv.getContext('2d') as CanvasRenderingContext2D;
  let previa: ILinea | undefined; let arrastre: IFicha | undefined;

  const pt = (e: PointerEvent): { x: number; y: number } => { const r = cv.getBoundingClientRect(); return { x: (e.clientX - r.left) * cv.width / r.width, y: (e.clientY - r.top) * cv.height / r.height }; };
  const dibujar = (): void => { pista(g, d.s); objs.concat(previa ? [previa] : []).forEach((o: Obj) => (esLinea(o) ? linea(g, o) : ficha(g, o))); };
  const cerca = (p: { x: number; y: number }): number => {
    let best = -1; let dist = 28;
    objs.forEach((o: Obj, i: number) => {
      const dd = esLinea(o) ? distSeg(p, o) : Math.hypot(o.x - p.x, o.y - p.y);
      if (dd < dist) { dist = dd; best = i; }
    });
    return best;
  };
  cv.addEventListener('pointerdown', (e: PointerEvent) => {
    e.preventDefault(); cv.setPointerCapture(e.pointerId); const p = pt(e);
    if (herr === 'mover') { const i = cerca(p); if (i > -1 && !esLinea(objs[i])) { arrastre = objs[i] as IFicha; } return; }
    if (herr === 'borrar') { const i = cerca(p); if (i > -1) { objs.splice(i, 1); dibujar(); } return; }
    if (herr === 'mov' || herr === 'pase' || herr === 'bote' || herr === 'bloq') { previa = { t: herr, x1: p.x, y1: p.y, x2: p.x, y2: p.y }; return; }
    const n = herr === 'at' || herr === 'df' ? (objs.filter((o: Obj) => !esLinea(o) && (o as IFicha).t === herr).length % 5) + 1 : undefined;
    objs.push({ t: herr as 'at' | 'df' | 'bal' | 'cono', x: p.x, y: p.y, n }); dibujar();
  });
  cv.addEventListener('pointermove', (e: PointerEvent) => {
    const p = pt(e);
    if (arrastre) { arrastre.x = p.x; arrastre.y = p.y; dibujar(); }
    if (previa) { previa.x2 = p.x; previa.y2 = p.y; dibujar(); }
  });
  const fin = (): void => {
    if (previa && Math.hypot(previa.x2 - previa.x1, previa.y2 - previa.y1) > 12) { objs.push(previa); }
    previa = undefined; arrastre = undefined; dibujar();
  };
  cv.addEventListener('pointerup', fin); cv.addEventListener('pointercancel', fin);

  const hs = el.querySelectorAll('[data-h]');
  for (let i = 0; i < hs.length; i++) {
    hs[i].addEventListener('click', (e: Event) => {
      herr = (e.currentTarget as HTMLElement).getAttribute('data-h') as Herr;
      for (let k = 0; k < hs.length; k++) { hs[k].setAttribute('aria-pressed', hs[k] === e.currentTarget ? 'true' : 'false'); }
    });
  }
  const acc = (k: string, fn: () => void): void => { const b = el.querySelector('[data-acc=' + k + ']') as HTMLElement; if (b) { b.addEventListener('click', fn); } };
  acc('deshacer', () => { objs.pop(); dibujar(); });
  acc('limpiar', () => { if (!objs.length || window.confirm('¿Borrar toda la jugada?')) { objs = []; dibujar(); } });
  acc('modo', () => { modo = modo === 'media' ? 'completa' : 'media'; objs = []; renderPizarra(el, ctx); });
  acc('png', () => cv.toBlob((b: Blob | null) => { if (b) { descargar('jugada-gijon-cbp.png', b, 'image/png'); } }, 'image/png'));
  acc('guardar', () => guardar(cv, ctx));
  dibujar();
}

function guardar(cv: HTMLCanvasElement, ctx: ICtx): void {
  const html = '<form class="cbp-form" novalidate><label class="cbp-field full">Título<input name="titulo" required placeholder="Ej.: Salida de presión 1-2-2"></label>' +
    '<label class="cbp-field">Tipo<select name="tipo"><option>Sistema</option><option>Ejercicio</option></select></label>' +
    '<label class="cbp-field">Categoría<select name="cat">' + ['Escuela', 'Benjamín', 'Alevín', 'Infantil', 'Cadete', 'Junior', 'Sénior', 'Todo el club'].map((c: string) => '<option>' + esc(c) + '</option>').join('') + '</select></label>' +
    '<label class="cbp-field full">Descripción<textarea name="desc" placeholder="Explicación de la jugada o el ejercicio"></textarea></label>' +
    '<div class="cbp-btns full"><button type="button" class="cbp-btn" data-a="c">Cancelar</button><button type="submit" class="cbp-btn pri">Guardar</button></div></form>';
  const body = abrirModal('Guardar jugada', html, 560);
  const form = body.querySelector('form') as HTMLFormElement;
  (body.querySelector('[data-a=c]') as HTMLElement).addEventListener('click', cerrarModal);
  form.addEventListener('submit', (e: Event) => {
    e.preventDefault();
    const v = (n: string): string => ((form.elements.namedItem(n) as HTMLInputElement).value || '').trim();
    if (!v('titulo')) { aviso('Pon un título', true); return; }
    const btn = form.querySelector('[type=submit]') as HTMLButtonElement; btn.disabled = true;
    cv.toBlob((b: Blob | null) => {
      if (!b) { return; }
      subirAMateriales(ctx.client, ctx.w, b, { Title: v('titulo'), TipoMaterial: v('tipo'), Categoria: v('cat'), Descripcion: v('desc') }, v('titulo').replace(/\s+/g, '-').toLowerCase() + '.png')
        .then(() => { cerrarModal(); aviso('Jugada guardada en Materiales'); })
        .catch((er: Error) => { btn.disabled = false; aviso('Error: ' + er.message, true); });
    }, 'image/png');
  });
}

/* ---------- Dibujo ---------- */
function pista(g: CanvasRenderingContext2D, s: number): void {
  const W = g.canvas.width; const H = g.canvas.height;
  g.fillStyle = '#E9CFA0'; g.fillRect(0, 0, W, H);
  g.strokeStyle = '#090909'; g.lineWidth = 3; g.lineCap = 'round'; g.lineJoin = 'round';
  // u: distancia desde la línea de fondo; v: desde la banda izquierda
  const mitad = (map: (u: number, v: number) => [number, number]): void => {
    const poly = (pts: [number, number][], cerrar?: boolean, relleno?: string): void => {
      g.beginPath(); pts.forEach((p: [number, number], i: number) => { const q = map(p[0], p[1]); if (i) { g.lineTo(q[0], q[1]); } else { g.moveTo(q[0], q[1]); } });
      if (cerrar) { g.closePath(); } if (relleno) { g.fillStyle = relleno; g.fill(); } g.stroke();
    };
    const arco = (cu: number, cv: number, r: number, a0: number, a1: number): [number, number][] => {
      const pts: [number, number][] = []; for (let i = 0; i <= 48; i++) { const a = a0 + (a1 - a0) * i / 48; pts.push([cu + r * Math.cos(a), cv + r * Math.sin(a)]); } return pts;
    };
    const bu = 1.575; const c = 7.5;
    poly([[0, c - 2.45], [5.8, c - 2.45], [5.8, c + 2.45], [0, c + 2.45]], false, 'rgba(255,159,0,0.28)');
    poly(arco(5.8, c, 1.8, -Math.PI / 2, Math.PI / 2));
    const lim = Math.sqrt(6.75 * 6.75 - 6.6 * 6.6);
    const ang = Math.asin(6.6 / 6.75);
    poly([[0, 0.9], [bu + lim, 0.9]]); poly([[0, 15 - 0.9], [bu + lim, 15 - 0.9]]);
    poly(arco(bu, c, 6.75, -ang, ang).map((p: [number, number]) => [p[0], p[1]] as [number, number]));
    poly(arco(bu, c, 1.25, -Math.PI / 2, Math.PI / 2));
    poly([[1.2, c - 0.9], [1.2, c + 0.9]]);
    g.beginPath(); const aro = map(bu, c); g.arc(aro[0], aro[1], 0.23 * s, 0, Math.PI * 2); g.strokeStyle = '#FF9F00'; g.stroke(); g.strokeStyle = '#090909';
  };
  if (modo === 'media') {
    const map = (u: number, v: number): [number, number] => [PAD + v * s, PAD + u * s];
    g.strokeRect(PAD, PAD, 15 * s, 14 * s);
    mitad(map);
    g.beginPath(); const cc = map(14, 7.5); g.arc(cc[0], cc[1], 1.8 * s, Math.PI, 2 * Math.PI); g.stroke();
  } else {
    g.strokeRect(PAD, PAD, 28 * s, 15 * s);
    mitad((u: number, v: number): [number, number] => [PAD + u * s, PAD + v * s]);
    mitad((u: number, v: number): [number, number] => [PAD + (28 - u) * s, PAD + v * s]);
    g.beginPath(); g.moveTo(PAD + 14 * s, PAD); g.lineTo(PAD + 14 * s, PAD + 15 * s); g.stroke();
    g.beginPath(); g.arc(PAD + 14 * s, PAD + 7.5 * s, 1.8 * s, 0, Math.PI * 2); g.stroke();
  }
}

function ficha(g: CanvasRenderingContext2D, f: IFicha): void {
  g.save();
  if (f.t === 'at') {
    g.beginPath(); g.arc(f.x, f.y, 17, 0, Math.PI * 2); g.fillStyle = '#090909'; g.fill();
    g.fillStyle = '#FF9F00'; g.font = 'bold 18px Inter, Arial, sans-serif'; g.textAlign = 'center'; g.textBaseline = 'middle'; g.fillText(String(f.n || ''), f.x, f.y + 1);
  } else if (f.t === 'df') {
    g.strokeStyle = '#C0392B'; g.lineWidth = 5; g.lineCap = 'round';
    g.beginPath(); g.moveTo(f.x - 12, f.y - 12); g.lineTo(f.x + 12, f.y + 12); g.moveTo(f.x + 12, f.y - 12); g.lineTo(f.x - 12, f.y + 12); g.stroke();
    g.fillStyle = '#C0392B'; g.font = 'bold 13px Inter, Arial, sans-serif'; g.fillText(String(f.n || ''), f.x + 14, f.y - 12);
  } else if (f.t === 'bal') {
    g.beginPath(); g.arc(f.x, f.y, 10, 0, Math.PI * 2); g.fillStyle = '#E8761B'; g.fill(); g.strokeStyle = '#090909'; g.lineWidth = 1.5; g.stroke();
    g.beginPath(); g.moveTo(f.x - 10, f.y); g.lineTo(f.x + 10, f.y); g.moveTo(f.x, f.y - 10); g.lineTo(f.x, f.y + 10); g.stroke();
  } else {
    g.beginPath(); g.moveTo(f.x, f.y - 12); g.lineTo(f.x + 10, f.y + 9); g.lineTo(f.x - 10, f.y + 9); g.closePath(); g.fillStyle = '#FF9F00'; g.fill(); g.strokeStyle = '#090909'; g.lineWidth = 1.5; g.stroke();
  }
  g.restore();
}

function linea(g: CanvasRenderingContext2D, l: ILinea): void {
  g.save(); g.strokeStyle = '#090909'; g.fillStyle = '#090909'; g.lineWidth = 3; g.lineCap = 'round';
  const ang = Math.atan2(l.y2 - l.y1, l.x2 - l.x1); const len = Math.hypot(l.x2 - l.x1, l.y2 - l.y1);
  const ex = l.x2 - Math.cos(ang) * 12; const ey = l.y2 - Math.sin(ang) * 12;
  g.beginPath();
  if (l.t === 'pase') { g.setLineDash([10, 8]); g.moveTo(l.x1, l.y1); g.lineTo(ex, ey); }
  else if (l.t === 'bote') {
    const n = Math.max(2, Math.floor((len - 14) / 14)); g.moveTo(l.x1, l.y1);
    for (let i = 1; i <= n; i++) { const t = i / n; const px = l.x1 + (ex - l.x1) * t; const py = l.y1 + (ey - l.y1) * t; const o = (i % 2 ? 7 : -7); g.lineTo(px - Math.sin(ang) * o, py + Math.cos(ang) * o); }
    g.lineTo(ex, ey);
  } else { g.moveTo(l.x1, l.y1); g.lineTo(l.t === 'bloq' ? l.x2 : ex, l.t === 'bloq' ? l.y2 : ey); }
  g.stroke(); g.setLineDash([]);
  if (l.t === 'bloq') {
    g.beginPath(); g.moveTo(l.x2 - Math.sin(ang) * 14, l.y2 + Math.cos(ang) * 14); g.lineTo(l.x2 + Math.sin(ang) * 14, l.y2 - Math.cos(ang) * 14); g.lineWidth = 4; g.stroke();
  } else {
    g.beginPath(); g.moveTo(l.x2, l.y2); g.lineTo(l.x2 - Math.cos(ang - 0.45) * 16, l.y2 - Math.sin(ang - 0.45) * 16); g.lineTo(l.x2 - Math.cos(ang + 0.45) * 16, l.y2 - Math.sin(ang + 0.45) * 16); g.closePath(); g.fill();
  }
  g.restore();
}

function distSeg(p: { x: number; y: number }, l: ILinea): number {
  const dx = l.x2 - l.x1; const dy = l.y2 - l.y1; const L = dx * dx + dy * dy;
  const t = L ? Math.max(0, Math.min(1, ((p.x - l.x1) * dx + (p.y - l.y1) * dy) / L)) : 0;
  return Math.hypot(p.x - (l.x1 + t * dx), p.y - (l.y1 + t * dy));
}
