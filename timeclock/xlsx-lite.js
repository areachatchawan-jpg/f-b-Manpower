/* xlsx-lite: สร้างไฟล์ Excel (.xlsx) พร้อมโลโก้ ไม่ต้องใช้ไลบรารีภายนอก
   XLSX.build({sheets:[{name, title, subtitle, cols:[{h, w, t:'text'|'num'|'int'}], rows:[[...]], total:[...]}], logo:{png:Uint8Array, w, h}}) -> Blob */
(function(){
  const enc = new TextEncoder();
  const xe = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const colName = i => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - m - 1) / 26; } return s; };

  // ---------- zip (store) ----------
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(files) {
    const parts = [], central = []; let off = 0;
    for (const f of files) {
      const name = enc.encode(f.name), data = typeof f.data === 'string' ? enc.encode(f.data) : f.data, crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
      h.setUint16(10, 0, true); h.setUint16(12, 0x21, true); h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
      h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true);
      c.setUint16(12, 0, true); c.setUint16(14, 0x21, true); c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
      c.setUint16(28, name.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), name);
      off += 30 + name.length + data.length;
    }
    const csize = central.reduce((s, p) => s + p.length, 0), e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, csize, true); e.setUint32(16, off, true);
    return new Blob([...parts, ...central, new Uint8Array(e.buffer)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  // ---------- styles ----------
  // 0 ปกติ, 1 หัวเรื่อง, 2 รอง, 3 หัวตาราง, 4 ข้อความ, 5 เงิน, 6 จำนวนเต็ม, 7 รวม-ข้อความ, 8 รวม-เงิน, 9 ทศนิยม, 10 รวม-ทศนิยม
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.00"/><numFmt numFmtId="165" formatCode="0.##"/></numFmts>
<fonts count="5"><font><sz val="11"/><name val="Tahoma"/></font><font><b/><sz val="16"/><color rgb="FF2A292A"/><name val="Tahoma"/></font><font><sz val="10"/><color rgb="FF7B7062"/><name val="Tahoma"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Tahoma"/></font><font><b/><sz val="11"/><name val="Tahoma"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4B4B4D"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEFE9DF"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFD9D2C7"/></left><right style="thin"><color rgb="FFD9D2C7"/></right><top style="thin"><color rgb="FFD9D2C7"/></top><bottom style="thin"><color rgb="FFD9D2C7"/></bottom></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11">
<xf xfId="0" fontId="0"/>
<xf xfId="0" fontId="1" applyFont="1"/>
<xf xfId="0" fontId="2" applyFont="1"/>
<xf xfId="0" fontId="3" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf xfId="0" fontId="0" borderId="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf xfId="0" fontId="0" borderId="1" numFmtId="164" applyBorder="1" applyNumberFormat="1"/>
<xf xfId="0" fontId="0" borderId="1" numFmtId="1" applyBorder="1" applyNumberFormat="1"/>
<xf xfId="0" fontId="4" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf xfId="0" fontId="4" fillId="3" borderId="1" numFmtId="164" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1"/>
<xf xfId="0" fontId="0" borderId="1" numFmtId="165" applyBorder="1" applyNumberFormat="1"/>
<xf xfId="0" fontId="4" fillId="3" borderId="1" numFmtId="165" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  const LOGO_ROWS = 6, TOP = LOGO_ROWS + 3; // แถวหัวตาราง (1-based) = TOP + 1
  function cell(ref, v, s) {
    if (v === null || v === undefined || v === '') return `<c r="${ref}" s="${s}"/>`;
    if (typeof v === 'number' && isFinite(v)) return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
    return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xe(v)}</t></is></c>`;
  }
  function sheetXml(sh, hasLogo) {
    const n = sh.cols.length, last = colName(n - 1), rows = [];
    for (let r = 1; r <= LOGO_ROWS; r++) rows.push(`<row r="${r}" ht="${hasLogo ? 18 : 6}" customHeight="1"/>`);
    rows.push(`<row r="${LOGO_ROWS + 1}" ht="26" customHeight="1">${cell('A' + (LOGO_ROWS + 1), sh.title, 1)}</row>`);
    rows.push(`<row r="${LOGO_ROWS + 2}">${cell('A' + (LOGO_ROWS + 2), sh.subtitle || '', 2)}</row>`);
    const hr = TOP + 1;
    rows.push(`<row r="${hr}" ht="30" customHeight="1">${sh.cols.map((c, i) => cell(colName(i) + hr, c.h, 3)).join('')}</row>`);
    const st = t => t === 'num' ? 5 : t === 'int' ? 6 : t === 'dec' ? 9 : 4;
    sh.rows.forEach((row, ri) => { const r = hr + 1 + ri; rows.push(`<row r="${r}">${sh.cols.map((c, i) => cell(colName(i) + r, row[i], st(c.t))).join('')}</row>`); });
    if (sh.total) { const r = hr + 1 + sh.rows.length; rows.push(`<row r="${r}">${sh.cols.map((c, i) => cell(colName(i) + r, sh.total[i], c.t === 'num' ? 8 : c.t === 'dec' || c.t === 'int' ? 10 : 7)).join('')}</row>`); }
    const endRow = hr + sh.rows.length + (sh.total ? 1 : 0);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<dimension ref="A1:${last}${endRow}"/>
<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="${hr}" topLeftCell="A${hr + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="18"/>
<cols>${sh.cols.map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${c.w || 14}" customWidth="1"/>`).join('')}</cols>
<sheetData>${rows.join('')}</sheetData>
${sh.rows.length ? `<autoFilter ref="A${hr}:${last}${hr + sh.rows.length}"/>` : ''}
<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>
${hasLogo ? '<drawing r:id="rId1"/>' : ''}
</worksheet>`;
  }
  function drawingXml(w, h) {
    const E = 9525;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>${6 * E}</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>${4 * E}</xdr:rowOff></xdr:from><xdr:ext cx="${w * E}" cy="${h * E}"/>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Logo"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * E}" cy="${h * E}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
</xdr:wsDr>`;
  }
  const safeName = (s, used) => { let n = String(s).replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 31) || 'Sheet'; let k = n, i = 2; while (used.has(k)) k = n.slice(0, 28) + ' ' + i++; used.add(k); return k; };

  function build({ sheets, logo }) {
    const used = new Set(), hasLogo = !!(logo && logo.png), f = [];
    const ov = [];
    sheets.forEach((sh, i) => {
      const k = i + 1;
      f.push({ name: `xl/worksheets/sheet${k}.xml`, data: sheetXml(sh, hasLogo) });
      ov.push(`<Override PartName="/xl/worksheets/sheet${k}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
      if (hasLogo) {
        f.push({ name: `xl/worksheets/_rels/sheet${k}.xml.rels`, data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${k}.xml"/></Relationships>` });
        f.push({ name: `xl/drawings/drawing${k}.xml`, data: drawingXml(logo.w, logo.h) });
        f.push({ name: `xl/drawings/_rels/drawing${k}.xml.rels`, data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo.png"/></Relationships>` });
        ov.push(`<Override PartName="/xl/drawings/drawing${k}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
      }
    });
    if (hasLogo) f.push({ name: 'xl/media/logo.png', data: logo.png });
    const names = sheets.map(s => safeName(s.name, used));
    const files = [
      { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${ov.join('')}</Types>` },
      { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
      { name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((n, i) => `<sheet name="${xe(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets><definedNames>${sheets.map((s, i) => s.rows.length ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${xe(names[i]).replace(/'/g, "''")}'!$A$${TOP + 1}:$${colName(s.cols.length - 1)}$${TOP + 1 + s.rows.length}</definedName>` : '').join('')}</definedNames></workbook>` },
      { name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
      { name: 'xl/styles.xml', data: STYLES },
      ...f
    ];
    return zip(files);
  }
  async function loadLogo(url, width) {
    try {
      const r = await fetch(url); if (!r.ok) return null;
      const png = new Uint8Array(await r.arrayBuffer());
      const dv = new DataView(png.buffer), w = dv.getUint32(16), h = dv.getUint32(20);
      return { png, w: width, h: Math.round(width * h / w) };
    } catch { return null; }
  }
  function download(blob, name) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }
  window.XLSX_LITE = { build, loadLogo, download };
})();
