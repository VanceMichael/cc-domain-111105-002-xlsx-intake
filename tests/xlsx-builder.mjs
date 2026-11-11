// 测试辅助：从零合成最小 OOXML 工作簿，用于构造同名工作表、错误关系、横向合并等拒收场景。
import {deflateRawSync} from 'node:zlib';

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function dosTime(date = new Date(1609459200000)) {
  return {
    time: ((date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() / 2)) & 0xffff,
    date: (((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate()) & 0xffff,
  };
}

export function buildZip(entries) {
  const {time, date} = dosTime();
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const {name, data, method = 8} of entries) {
    const nameBuf = Buffer.from(name, 'utf8');
    const payload = Buffer.from(data);
    const stored = method === 0 ? payload : deflateRawSync(payload);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(time, 10);
    local.writeUInt16LE(date, 12);
    local.writeUInt32LE(crc32(payload), 14);
    local.writeUInt32LE(stored.length, 18);
    local.writeUInt32LE(payload.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    locals.push(local, nameBuf, stored);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(method, 10);
    central.writeUInt16LE(time, 12);
    central.writeUInt16LE(date, 14);
    central.writeUInt32LE(crc32(payload), 16);
    central.writeUInt32LE(stored.length, 20);
    central.writeUInt32LE(payload.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBuf);
    offset += local.length + nameBuf.length + stored.length;
  }
  const centralBuf = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);
  return Buffer.concat([...locals, centralBuf, eocd]);
}

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const RNS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

// opts: {sheets:[{name, cells:[{ref,t,s,v,formula,inline}], merges:[ref]}], date1904, fullCalcOnLoad,
//         shared:['字符串'...], zeroPad:'000000', sheetTarget?, includeRels?, sharedPath?, stylesPath?}
export function buildWorkbook(opts = {}) {
  const sheets = opts.sheets ?? [{name: '披露表', cells: [], merges: []}];
  const shared = opts.shared ?? [];
  const sheetEntries = sheets.map((sh, i) => {
    const rows = new Map();
    for (const c of sh.cells) {
      const row = /[A-Z]+(\d+)/.exec(c.ref)[1];
      if (!rows.has(row)) rows.set(row, []);
      rows.get(row).push(c);
    }
    const rowXml = [...rows.entries()]
      .sort((a, b) => Number(a[0]) - Number(b[0]))
      .map(([row, cs]) => {
        const body = cs.map((c) => {
          const attrs = [`r="${c.ref}"`];
          if (c.t) attrs.push(`t="${c.t}"`);
          if (c.s !== undefined) attrs.push(`s="${c.s}"`);
          let inner = '';
          if (c.formula !== undefined) inner += `<f>${c.formula}</f>`;
          if (c.inline !== undefined) inner += `<is><t>${c.inline}</t></is>`;
          else if (c.v !== undefined) inner += `<v>${c.v}</v>`;
          return `<c ${attrs.join(' ')}>${inner}</c>`;
        }).join('');
        return `<row r="${row}">${body}</row>`;
      }).join('');
    const mergeXml = sh.merges?.length
      ? `<mergeCells count="${sh.merges.length}">${sh.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>`
      : '';
    return {
      name: opts.sheetTarget?.(i) ?? `xl/worksheets/sheet${i + 1}.xml`,
      data: `<worksheet ${NS}><sheetData>${rowXml}</sheetData>${mergeXml}</worksheet>`,
    };
  });

  const workbook = `<workbook ${NS} ${RNS}><workbookPr date1904="${opts.date1904 ? '1' : '0'}"/><sheets>${
    sheets.map((sh, i) => `<sheet name="${sh.name}" sheetId="${i + 1}" r:id="s${i + 1}"/>`).join('')
  }</sheets><calcPr fullCalcOnLoad="${opts.fullCalcOnLoad ? '1' : '0'}"/></workbook>`;

  const relTargets = opts.relTargets ?? sheetEntries.map((e) => e.name.replace(/^xl\//, ''));
  const sheetRels = sheets
    .map((sh, i) => `<Relationship Id="s${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${relTargets[i]}"/>`)
    .join('');
  const stylesPath = opts.stylesPath ?? 'styles.xml';
  const sharedPath = opts.sharedPath ?? 'sharedStrings.xml';
  const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheetRels
  }<Relationship Id="st" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="${stylesPath}"/>${
    opts.includeShared === false ? '' : `<Relationship Id="ss" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="${sharedPath}"/>`
  }</Relationships>`;

  const rootRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="root" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const contentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`;
  const sst = `<sst ${NS} count="${shared.length}" uniqueCount="${shared.length}">${
    shared.map((s) => `<si><t>${s}</t></si>`).join('')
  }</sst>`;
  const zeroPad = opts.zeroPad ?? '000000';
  const styles = `<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="${zeroPad}"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>`;

  const entries = [
    {name: '[Content_Types].xml', data: contentTypes},
    {name: '_rels/.rels', data: rootRels},
    {name: 'xl/_rels/workbook.xml.rels', data: opts.relsOverride ?? rels},
    {name: 'xl/workbook.xml', data: workbook},
    ...(opts.includeStyles === false ? [] : [{name: `xl/${stylesPath}`, data: styles}]),
    ...(opts.includeShared === false ? [] : [{name: `xl/${sharedPath}`, data: sst}]),
    ...sheetEntries.map((e) => ({name: e.name, data: e.data})),
  ];
  return buildZip(entries);
}
