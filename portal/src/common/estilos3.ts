/* Estilos de la ficha de equipo y de la página de avisos */
const CSS3 = `
.cbp-volver{display:inline-flex;align-items:center;gap:6px;font-weight:700;font-size:14px;color:#090909;background:none;border:0;padding:6px 0;cursor:pointer;font-family:inherit}
.cbp-volver:hover{color:#B36F00}
.cbp-fhero{background:#090909;color:#fff;border-radius:22px;padding:28px 28px 24px;display:grid;grid-template-columns:minmax(0,1fr) auto;gap:24px;position:relative;overflow:hidden}
.cbp-fhero:after{content:"";position:absolute;right:-60px;top:-60px;width:260px;height:260px;border-radius:50%;border:34px solid rgba(255,159,0,.10);pointer-events:none}
.cbp-fhero .cbp-over{color:#FF9F00}
.cbp-fhero h2{font-family:'Anton','Impact',sans-serif;font-weight:400;text-transform:uppercase;font-size:clamp(34px,5vw,56px);line-height:.95;margin:6px 0 10px;color:#fff}
.cbp-fhero p{margin:0;color:#cfcfcf;font-size:14px;max-width:560px}
.cbp-fdatos{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px}
.cbp-fdato{display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.14);border-radius:999px;padding:6px 12px;font-size:13px;color:#fff}
.cbp-fdato b{color:#FF9F00;font-weight:700}
.cbp-fdato.vacio{color:#9a9a9a;border-style:dashed;cursor:pointer;background:none}
.cbp-fstats{display:grid;grid-template-columns:repeat(3,minmax(84px,1fr));gap:10px;align-self:end;position:relative;z-index:1}
.cbp-fstat{background:#1a1a1a;border-radius:14px;padding:12px 14px;text-align:center}
.cbp-fstat b{display:block;font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:30px;line-height:1;color:#FF9F00}
.cbp-fstat span{font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:#bdbdbd;font-weight:700}
.cbp-facciones{display:flex;flex-wrap:wrap;gap:10px}
.cbp-fgrid{display:grid;grid-template-columns:minmax(0,1.5fr) minmax(0,1fr);gap:18px;align-items:start}
.cbp-fcol{display:flex;flex-direction:column;gap:18px;min-width:0}
.cbp-panel{background:#fff;border:1px solid #E6E3DC;border-radius:18px;overflow:hidden}
.cbp-panel-h{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:16px 20px 12px}
.cbp-panel-h h3{margin:0;font-family:'Anton','Impact',sans-serif;font-weight:400;text-transform:uppercase;font-size:22px;letter-spacing:.01em}
.cbp-panel-h small{font-size:13px;color:#5c5c5c;font-weight:600}
.cbp-panel-b{padding:0 20px 18px;display:flex;flex-direction:column;gap:10px}
.cbp-panel .cbp-partido{margin:0}
.cbp-racha{display:flex;gap:6px;flex-wrap:wrap}
.cbp-racha span{width:30px;height:30px;border-radius:8px;display:inline-flex;align-items:center;justify-content:center;font-weight:800;font-size:13px}
.cbp-racha .v{background:#FF9F00;color:#090909}.cbp-racha .d{background:#EDEAE3;color:#5c5c5c}
.cbp-horario{display:flex;flex-direction:column;gap:8px}
.cbp-hdia{display:grid;grid-template-columns:52px minmax(0,1fr);gap:12px;align-items:center;background:#F6F5F1;border-radius:12px;padding:10px 12px}
.cbp-hdia b{font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:18px;color:#090909;text-align:center;background:#FF9F00;border-radius:8px;padding:4px 0}
.cbp-hdia span{font-size:14px;line-height:1.3}
.cbp-hdia span i{display:block;font-style:normal;color:#5c5c5c;font-size:13px}
.cbp-roster{display:flex;flex-direction:column}
.cbp-jug{display:grid;grid-template-columns:40px minmax(0,1fr) auto auto;gap:12px;align-items:center;padding:9px 0;border-top:1px solid #EFECE6}
.cbp-jug:first-child{border-top:0}
.cbp-dorsal{width:40px;height:40px;border-radius:50%;background:#090909;color:#FF9F00;display:flex;align-items:center;justify-content:center;font-family:'Anton','Impact',sans-serif;font-size:18px}
.cbp-jug .nom{font-weight:600;font-size:15px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.cbp-pct{font-size:12px;font-weight:700;color:#5c5c5c;display:flex;align-items:center;gap:6px}
.cbp-pct i{display:block;width:46px;height:6px;border-radius:3px;background:#EDEAE3;overflow:hidden}
.cbp-pct i u{display:block;height:100%;background:#FF9F00}
.cbp-quitar{border:0;background:none;color:#9a9a9a;font-size:20px;line-height:1;cursor:pointer;width:32px;height:32px;border-radius:8px}
.cbp-quitar:hover{background:#F6F5F1;color:#B00020}
.cbp-alta{display:grid;grid-template-columns:72px minmax(0,1fr) auto;gap:8px;margin-top:6px}
.cbp-alta input{font:inherit;font-size:14px;padding:9px 10px;border:1px solid #D6D2C8;border-radius:10px;min-height:42px;min-width:0}
.cbp-nota{font-size:13px;color:#5c5c5c;margin:0}
@media (max-width:900px){.cbp-fgrid{grid-template-columns:1fr}.cbp-fhero{grid-template-columns:1fr}.cbp-fstats{align-self:start}}
@media (max-width:520px){.cbp-fhero{padding:22px 18px}.cbp-fstat b{font-size:24px}}

/* ---- Avisos ---- */
.cbp-av-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:16px}
.cbp-av{background:#fff;border:1px solid #E6E3DC;border-radius:18px;padding:18px 20px 16px;display:flex;flex-direction:column;gap:10px;position:relative;scroll-margin-top:90px}
.cbp-av.imp{border:2px solid #FF9F00;box-shadow:0 6px 24px rgba(255,159,0,.14)}
.cbp-av.cad{opacity:.62}
.cbp-av.foco{outline:3px solid #090909;outline-offset:2px}
.cbp-av-top{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.cbp-pill{display:inline-flex;align-items:center;gap:5px;border-radius:999px;padding:4px 10px;font-size:11px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;background:#F1EEE8;color:#5c5c5c}
.cbp-pill.imp{background:#FF9F00;color:#090909}
.cbp-pill.auto{background:#090909;color:#FF9F00}
.cbp-pill.urg{background:#B00020;color:#fff}
.cbp-av h3{margin:0;font-size:19px;line-height:1.25;font-weight:800;color:#090909}
.cbp-av p{margin:0;font-size:14.5px;line-height:1.5;color:#333;white-space:pre-line;overflow-wrap:anywhere}
.cbp-av p a{color:#B36F00;font-weight:600}
.cbp-av-pie{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:auto;padding-top:10px;border-top:1px solid #EFECE6;font-size:12.5px;color:#5c5c5c}
.cbp-av-acc{display:flex;gap:4px}
.cbp-av-acc button{border:0;background:none;font:inherit;font-size:13px;font-weight:700;color:#090909;padding:6px 8px;border-radius:8px;cursor:pointer}
.cbp-av-acc button:hover{background:#F6F5F1}
.cbp-av-dest{display:grid;grid-template-columns:auto minmax(0,1fr);gap:18px;align-items:center;background:#090909;color:#fff;border-radius:22px;padding:22px 26px}
.cbp-av-dest .num{font-family:'Anton','Impact',sans-serif;font-size:64px;line-height:.9;color:#FF9F00}
.cbp-av-dest h3{margin:0 0 4px;font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:26px;text-transform:uppercase;color:#fff}
.cbp-av-dest p{margin:0;color:#cfcfcf;font-size:14px}
.cbp-check{display:flex;align-items:center;gap:10px;font-size:14px;font-weight:600;cursor:pointer}
.cbp-check input{width:20px;height:20px;accent-color:#FF9F00}
`;

export function cargarEstilos3(): void {
  if (document.getElementById('cbp-css3')) { return; }
  const s = document.createElement('style'); s.id = 'cbp-css3'; s.textContent = CSS3;
  document.head.appendChild(s);
}

/** Convierte URLs en enlaces dentro de un texto ya escapado */
export function enlazarUrls(htmlEscapado: string): string {
  return htmlEscapado.replace(/(https?:\/\/[^\s<]+[^\s<.,;:)])/g, '<a href="$1" target="_blank" rel="noopener">$1</a>');
}

/** Nombre de archivo .ics publicado para un equipo (igual que en la sincronización) */
export function slugEquipo(n: string): string {
  return n.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
}
export const WEB_CALENDARIOS = 'https://calixtohcc-collab.github.io/gijon-cbp-sync/';
