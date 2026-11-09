/* Oficina de la junta y comisiones */
const CSS4 = `
.cbp-of-hero{background:#090909;color:#fff;border-radius:24px;padding:28px 28px 22px;display:flex;flex-direction:column;gap:20px;position:relative;overflow:hidden}
.cbp-of-hero:before{content:"";position:absolute;right:-90px;bottom:-120px;width:340px;height:340px;border-radius:50%;background:radial-gradient(circle,rgba(255,159,0,.22),rgba(255,159,0,0) 70%);pointer-events:none}
.cbp-of-top{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap;position:relative}
.cbp-of-top .cbp-over{color:#FF9F00}
.cbp-of-top h2{font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:clamp(34px,5vw,54px);line-height:.95;text-transform:uppercase;margin:6px 0 6px;color:#fff}
.cbp-of-top p{margin:0;color:#cfcfcf;font-size:14px}
.cbp-candado{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:700;color:#090909;background:#FF9F00;border-radius:999px;padding:6px 12px;white-space:nowrap}
.cbp-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;position:relative}
.cbp-kpi{background:#1a1a1a;border:1px solid #262626;border-radius:16px;padding:14px 16px;display:flex;flex-direction:column;gap:4px;min-width:0;text-align:left;font:inherit;color:#fff;cursor:pointer}
.cbp-kpi:hover{border-color:#FF9F00}
.cbp-kpi b{font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:32px;line-height:1;color:#FF9F00}
.cbp-kpi span{font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#bdbdbd}
.cbp-kpi small{font-size:12.5px;color:#e5e5e5;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.cbp-kpi.alerta b{color:#FF5A5A}
.cbp-of-tabs{display:flex;gap:6px;flex-wrap:wrap;border-bottom:1px solid #E6E3DC;padding-bottom:10px}
.cbp-of-tab{font:inherit;font-size:14px;font-weight:700;border:0;background:none;padding:10px 14px;border-radius:10px;cursor:pointer;color:#5c5c5c;display:inline-flex;align-items:center;gap:8px}
.cbp-of-tab:hover{background:#F1EEE8;color:#090909}
.cbp-of-tab[aria-selected="true"]{background:#090909;color:#FF9F00}
.cbp-of-tab i{font-style:normal;font-size:11px;background:#E6E3DC;color:#090909;border-radius:999px;padding:1px 7px}
.cbp-of-tab[aria-selected="true"] i{background:#FF9F00}
.cbp-of-grid{display:grid;grid-template-columns:minmax(0,1.4fr) minmax(0,1fr);gap:18px;align-items:start}
.cbp-col{display:flex;flex-direction:column;gap:18px;min-width:0}
.cbp-box{background:#fff;border:1px solid #E6E3DC;border-radius:18px;padding:18px 20px;display:flex;flex-direction:column;gap:12px;min-width:0}
.cbp-box-h{display:flex;justify-content:space-between;align-items:center;gap:10px}
.cbp-box-h h3{margin:0;font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:21px;text-transform:uppercase;letter-spacing:.01em}
.cbp-box-h a,.cbp-box-h button.lnk{font:inherit;font-size:13px;font-weight:700;color:#B36F00;background:none;border:0;cursor:pointer;text-decoration:none}
.cbp-reu{display:grid;grid-template-columns:72px minmax(0,1fr);gap:16px;align-items:start}
.cbp-reu-f{background:#090909;color:#fff;border-radius:14px;text-align:center;padding:10px 4px}
.cbp-reu-f small{display:block;font-size:11px;font-weight:700;letter-spacing:.08em;color:#bdbdbd;text-transform:uppercase}
.cbp-reu-f b{display:block;font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:34px;line-height:1;color:#FF9F00}
.cbp-reu-f span{font-size:12px;font-weight:700;text-transform:uppercase}
.cbp-reu h4{margin:0 0 4px;font-size:18px;font-weight:800}
.cbp-reu .meta{font-size:13px;color:#5c5c5c;display:flex;gap:12px;flex-wrap:wrap}
.cbp-odd{margin:10px 0 0;padding:12px 14px;background:#F6F5F1;border-radius:12px;font-size:14px;line-height:1.55;white-space:pre-line}
.cbp-mini{display:flex;flex-direction:column}
.cbp-mini-it{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:10px;align-items:center;padding:10px 0;border-top:1px solid #EFECE6;font-size:14px}
.cbp-mini-it:first-child{border-top:0}
.cbp-mini-it b{font-weight:700;display:block;overflow:hidden;text-overflow:ellipsis}
.cbp-mini-it small{color:#5c5c5c;font-size:12.5px}
.cbp-mini-it a{color:inherit;text-decoration:none}
.cbp-fpill{font-size:12px;font-weight:700;border-radius:999px;padding:4px 10px;background:#F1EEE8;color:#333;white-space:nowrap}
.cbp-fpill.mal{background:#FDE7E9;color:#B00020}.cbp-fpill.ok{background:#E8F5E9;color:#1B5E20}.cbp-fpill.amb{background:#FFF1D6;color:#8A5300}
.cbp-kanban{display:grid;grid-template-columns:repeat(var(--cols,3),minmax(220px,1fr));gap:14px;overflow-x:auto;padding-bottom:4px}
.cbp-kcol{background:#F1EEE8;border-radius:18px;padding:12px;display:flex;flex-direction:column;gap:10px;min-height:120px}
.cbp-kcol-h{display:flex;justify-content:space-between;align-items:center;padding:2px 4px 4px;font-weight:800;font-size:13px;letter-spacing:.06em;text-transform:uppercase}
.cbp-kcol-h span{background:#fff;border-radius:999px;padding:2px 9px;font-size:12px}
.cbp-kcol-h small{font-weight:700;letter-spacing:0;text-transform:none;color:#5c5c5c}
.cbp-kcard{background:#fff;border-radius:14px;padding:12px 14px;display:flex;flex-direction:column;gap:8px;box-shadow:0 1px 2px rgba(0,0,0,.06);border:1px solid #E6E3DC}
.cbp-kcard h4{margin:0;font-size:15px;font-weight:700;line-height:1.3;cursor:pointer}
.cbp-kcard h4:hover{color:#B36F00}
.cbp-kcard .det{font-size:13px;color:#5c5c5c;display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.cbp-kcard .imp{font-family:'Anton','Impact',sans-serif;font-size:22px;color:#090909}
.cbp-kmov{display:flex;justify-content:space-between;gap:6px}
.cbp-kmov button{font:inherit;font-size:12px;font-weight:700;border:1px solid #E6E3DC;background:#fff;border-radius:8px;padding:4px 10px;cursor:pointer;color:#090909}
.cbp-kmov button:hover{border-color:#090909}
.cbp-avatar{width:30px;height:30px;border-radius:50%;background:#090909;color:#FF9F00;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:800;flex:0 0 auto}
.cbp-docs{display:flex;flex-direction:column}
.cbp-doc{display:grid;grid-template-columns:44px minmax(0,1fr) auto;gap:12px;align-items:center;padding:10px 0;border-top:1px solid #EFECE6;text-decoration:none;color:inherit}
.cbp-doc:first-child{border-top:0}
.cbp-doc:hover b{color:#B36F00}
.cbp-doc-ico{width:44px;height:44px;border-radius:12px;background:#F1EEE8;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:800;color:#090909;text-transform:uppercase}
.cbp-doc-ico.pdf{background:#FDE7E9;color:#B00020}.cbp-doc-ico.docx,.cbp-doc-ico.doc{background:#E3EDFB;color:#1F4E9A}.cbp-doc-ico.xlsx,.cbp-doc-ico.xls{background:#E5F4EA;color:#1B6B3A}
.cbp-doc b{display:block;font-size:14.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.cbp-doc small{color:#5c5c5c;font-size:12.5px}
.cbp-restr{background:#090909;color:#fff;border-radius:24px;padding:40px 28px;text-align:center;display:flex;flex-direction:column;align-items:center;gap:10px}
.cbp-restr h2{font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:36px;margin:0;text-transform:uppercase;color:#fff}
.cbp-restr p{margin:0;color:#cfcfcf;max-width:520px}
/* Comisiones */
.cbp-coms{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:16px}
.cbp-com{background:#fff;border:1px solid #E6E3DC;border-radius:20px;overflow:hidden;display:flex;flex-direction:column;cursor:pointer;text-align:left;font:inherit;color:inherit;padding:0}
.cbp-com:hover{box-shadow:0 10px 30px rgba(0,0,0,.08);transform:translateY(-1px)}
.cbp-com-h{padding:18px 20px 16px;color:#fff;display:flex;flex-direction:column;gap:6px;min-height:96px;justify-content:flex-end}
.cbp-com-h small{font-size:11px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;opacity:.85}
.cbp-com-h h3{margin:0;font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:26px;line-height:1;text-transform:uppercase}
.cbp-com-b{padding:14px 20px 18px;display:flex;flex-direction:column;gap:12px;flex:1}
.cbp-com-b p{margin:0;font-size:14px;color:#333;line-height:1.45;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.cbp-com-pie{display:flex;justify-content:space-between;align-items:center;margin-top:auto;font-size:12.5px;color:#5c5c5c;font-weight:600}
.cbp-pila{display:flex}
.cbp-pila .cbp-avatar{border:2px solid #fff;margin-left:-8px}
.cbp-pila .cbp-avatar:first-child{margin-left:0}
.cbp-com-nueva{border:2px dashed #CFCABF;background:none;min-height:200px;align-items:center;justify-content:center;gap:8px;font-weight:800;color:#5c5c5c;font-size:15px}
.cbp-com-nueva:hover{border-color:#090909;color:#090909;box-shadow:none}
.cbp-com-nueva b{font-size:40px;line-height:1;font-weight:400}
.cbp-chero{border-radius:24px;padding:26px 28px;color:#fff;display:grid;grid-template-columns:minmax(0,1fr) auto;gap:20px;align-items:end}
.cbp-chero h2{font-family:'Anton','Impact',sans-serif;font-weight:400;font-size:clamp(32px,5vw,50px);line-height:.95;text-transform:uppercase;margin:6px 0 8px;color:#fff}
.cbp-chero p{margin:0;font-size:14.5px;max-width:640px;opacity:.92;white-space:pre-line}
.cbp-chero small{font-size:11px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;opacity:.85}
.cbp-chero .cbp-btn{background:#fff;border-color:#fff}
.cbp-miembros{display:flex;flex-direction:column}
.cbp-miembro{display:grid;grid-template-columns:30px minmax(0,1fr) auto;gap:10px;align-items:center;padding:8px 0;border-top:1px solid #EFECE6;font-size:14px}
.cbp-miembro:first-child{border-top:0}
.cbp-miembro span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.cbp-miembro small{display:block;color:#5c5c5c;font-size:12px}
.cbp-feed{display:flex;flex-direction:column;gap:12px}
.cbp-post{background:#fff;border:1px solid #E6E3DC;border-radius:16px;padding:14px 16px;display:grid;grid-template-columns:34px minmax(0,1fr);gap:12px}
.cbp-post .cbp-avatar{width:34px;height:34px}
.cbp-post-h{display:flex;flex-wrap:wrap;gap:8px;align-items:center;font-size:13px;color:#5c5c5c}
.cbp-post-h b{color:#090909}
.cbp-post h4{margin:6px 0 4px;font-size:15.5px}
.cbp-post p{margin:0;font-size:14.5px;line-height:1.5;white-space:pre-line;overflow-wrap:anywhere}
.cbp-post p a{color:#B36F00}
.cbp-post-acc{display:flex;gap:4px;margin-top:6px}
.cbp-post-acc button{font:inherit;font-size:12px;font-weight:700;border:0;background:none;color:#5c5c5c;cursor:pointer;padding:4px 6px;border-radius:6px}
.cbp-post-acc button:hover{background:#F1EEE8;color:#090909}
.cbp-compose{background:#fff;border:1px solid #E6E3DC;border-radius:16px;padding:12px;display:flex;flex-direction:column;gap:10px}
.cbp-compose textarea{font:inherit;font-size:15px;border:0;resize:vertical;min-height:64px;outline:none;padding:4px}
.cbp-compose-b{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
.cbp-tipo{display:flex;gap:4px;flex-wrap:wrap}
.cbp-tipo label{font-size:13px;font-weight:700;border:1px solid #E6E3DC;border-radius:999px;padding:6px 12px;cursor:pointer}
.cbp-tipo input{position:absolute;opacity:0;pointer-events:none}
.cbp-tipo input:checked+span{color:#FF9F00}
.cbp-tipo label:has(input:checked){background:#090909;border-color:#090909;color:#FF9F00}
.cbp-colores{display:flex;gap:8px;flex-wrap:wrap}
.cbp-colores label{width:34px;height:34px;border-radius:50%;cursor:pointer;border:3px solid transparent;position:relative}
.cbp-colores input{position:absolute;opacity:0}
.cbp-colores label:has(input:checked){border-color:#090909;box-shadow:inset 0 0 0 2px #fff}
@media (max-width:900px){.cbp-of-grid{grid-template-columns:1fr}.cbp-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.cbp-chero{grid-template-columns:1fr}}
`;

export function cargarEstilos4(): void {
  if (document.getElementById('cbp-css4')) { return; }
  const s = document.createElement('style'); s.id = 'cbp-css4'; s.textContent = CSS4;
  document.head.appendChild(s);
}

export const COLORES: string[] = ['#090909', '#B36F00', '#1F4E9A', '#1B6B3A', '#8E2A5A', '#5A3E9B', '#B00020', '#2F6F73'];
export function fondoColor(c: string): string { return 'background:' + (c || '#090909') + (c === '#090909' || !c ? ';background-image:linear-gradient(135deg,#090909 60%,#3a2600)' : ''); }
