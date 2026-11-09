/* Generador mínimo de .xlsx (sin dependencias): una hoja, texto, cabecera en negrita y anchos de columna */

let TABLA: number[] | null = null;
function crc32(b: Uint8Array): number {
  if (!TABLA) { TABLA = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) { c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; } TABLA.push(c >>> 0); } }
  let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) { c = TABLA[(c ^ b[i]) & 0xFF] ^ (c >>> 8); }
  return (c ^ 0xFFFFFFFF) >>> 0;
}
function utf8(s: string): Uint8Array { return new TextEncoder().encode(s); }

/** ZIP sin compresión (método «store»), suficiente para Excel */
function zip(archivos: { nombre: string; datos: Uint8Array }[]): Blob {
  const partes: Uint8Array[] = []; const central: Uint8Array[] = []; let off = 0;
  const u16 = (v: number): number[] => [v & 255, (v >>> 8) & 255];
  const u32 = (v: number): number[] => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
  archivos.forEach((f: { nombre: string; datos: Uint8Array }) => {
    const n = utf8(f.nombre); const crc = crc32(f.datos); const len = f.datos.length;
    const loc = new Uint8Array([0x50, 0x4b, 3, 4].concat(u16(20), u16(0x0800), u16(0), u16(0), u16(0x21), u32(crc), u32(len), u32(len), u16(n.length), u16(0)));
    partes.push(loc, n, f.datos);
    central.push(new Uint8Array([0x50, 0x4b, 1, 2].concat(u16(20), u16(20), u16(0x0800), u16(0), u16(0), u16(0x21), u32(crc), u32(len), u32(len), u16(n.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(off))), n);
    off += loc.length + n.length + len;
  });
  const tamC = central.reduce((s: number, x: Uint8Array) => s + x.length, 0);
  const fin = new Uint8Array([0x50, 0x4b, 5, 6].concat(u16(0), u16(0), u16(archivos.length), u16(archivos.length), u32(tamC), u32(off), u16(0)));
  return new Blob(partes.concat(central, [fin]) as BlobPart[], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

const x = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
function col(i: number): string { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }

/** Crea un libro con una hoja. «cabecera» puede ir vacía si la plantilla no la lleva. */
export function libroXlsx(hoja: string, cabecera: string[], filas: string[][], anchos?: number[]): Blob {
  const todas = (cabecera.length ? [cabecera] : []).concat(filas);
  const sheetData = todas.map((f: string[], r: number) => '<row r="' + (r + 1) + '">' + f.map((v: string, c: number) =>
    '<c r="' + col(c) + (r + 1) + '" t="inlineStr"' + (r === 0 && cabecera.length ? ' s="1"' : '') + '><is><t xml:space="preserve">' + x(v || '') + '</t></is></c>').join('') + '</row>').join('');
  const cols = anchos && anchos.length ? '<cols>' + anchos.map((w: number, i: number) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols>' : '';
  const archivos: { nombre: string; datos: Uint8Array }[] = [
    { nombre: '[Content_Types].xml', datos: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>') },
    { nombre: '_rels/.rels', datos: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>') },
    { nombre: 'xl/workbook.xml', datos: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + x(hoja.slice(0, 31)) + '" sheetId="1" r:id="rId1"/></sheets></workbook>') },
    { nombre: 'xl/_rels/workbook.xml.rels', datos: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>') },
    { nombre: 'xl/styles.xml', datos: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf fontId="0"/><xf fontId="1" applyFont="1"/></cellXfs></styleSheet>') },
    { nombre: 'xl/worksheets/sheet1.xml', datos: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + cols + '<sheetData>' + sheetData + '</sheetData></worksheet>') }
  ];
  return zip(archivos);
}
