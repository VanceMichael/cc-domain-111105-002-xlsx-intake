// OOXML 部件解析：关系、工作簿、共享字符串、样式、工作表。列号一律来自单元格 r 属性，不依赖节点顺序。
import {parseXml} from './xml.mjs';

export class IntakeError extends Error {
  constructor(code, message, detail) {
    super(message);
    this.name = 'IntakeError';
    this.code = code;
    if (detail !== undefined) this.detail = detail;
  }
}

function textOf(parts, path) {
  const buf = parts.get(path);
  if (!buf) throw new IntakeError('PART_MISSING', `包内缺少部件: ${path}`);
  return buf.toString('utf8');
}

// ---------- 关系 ----------

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

function normalizeTarget(baseDir, target) {
  const stack = baseDir.length ? baseDir.split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') {
      if (stack.length === 0) {
        throw new IntakeError('REL_TRAVERSAL', `关系目标越出包根: ${target}`);
      }
      stack.pop();
    } else {
      stack.push(seg);
    }
  }
  return stack.join('/');
}

function decodeTarget(target) {
  try {
    return decodeURIComponent(target);
  } catch {
    return target;
  }
}

function parseRelationships(xml, baseDir) {
  const rels = new Map();
  parseXml(xml, {
    onOpen(path, attrs) {
      if (path[path.length - 1] !== 'Relationship') return;
      const id = attrs.Id;
      const type = attrs.Type;
      const targetMode = attrs.TargetMode;
      if (!id || !type || attrs.Target === undefined) {
        throw new IntakeError('REL_MALFORMED', '关系条目缺少 Id/Type/Target 属性');
      }
      if (targetMode === 'External') {
        throw new IntakeError('REL_EXTERNAL', `关系 ${id} 指向外部目标，拒绝接入`, {id});
      }
      const target = decodeTarget(attrs.Target);
      if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(target) && !target.startsWith('/')) {
        throw new IntakeError('REL_EXTERNAL', `关系 ${id} 使用绝对 URL，拒绝接入`, {id});
      }
      const absolute = target.startsWith('/')
        ? normalizeTarget('', target.slice(1))
        : normalizeTarget(baseDir, target);
      rels.set(id, {id, type, path: absolute});
    },
  });
  return rels;
}

function requireRel(rels, id, expectedSuffix, where) {
  const rel = rels.get(id);
  if (!rel) throw new IntakeError('REL_BROKEN', `${where} 引用的关系 ${id} 不存在`);
  if (!rel.type.endsWith(expectedSuffix)) {
    throw new IntakeError('REL_TYPE', `${where} 的关系 ${id} 类型不符: ${rel.type}`);
  }
  return rel;
}

// ---------- 坐标 ----------

export function columnIndex(ref) {
  const m = /^([A-Z]{1,3})(\d+)$/.exec(ref);
  if (!m) throw new IntakeError('REF_INVALID', `无法解析单元格坐标: ${ref}`);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return {col, row: Number(m[2])};
}

export function columnLetter(index) {
  let n = index;
  let out = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function parseRange(ref) {
  const m = /^([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)$/.exec(ref);
  if (!m) throw new IntakeError('MAPPING_RANGE', `映射范围无法解析: ${ref}`);
  const a = columnIndex(`${m[1]}${m[2]}`);
  const b = columnIndex(`${m[3]}${m[4]}`);
  if (b.col < a.col || b.row < a.row) {
    throw new IntakeError('MAPPING_RANGE', `映射范围起止顺序错误: ${ref}`);
  }
  return {minCol: a.col, maxCol: b.col, minRow: a.row, maxRow: b.row};
}

// ---------- 工作簿 ----------

export function parseWorkbook(parts) {
  const rootRelsXml = textOf(parts, '_rels/.rels');
  const rootRels = parseRelationships(rootRelsXml, '');
  const docRel = [...rootRels.values()].find((r) => r.type === `${REL_NS}officeDocument`);
  if (!docRel) throw new IntakeError('REL_BROKEN', '根关系未指向工作簿部件');
  const workbookPath = docRel.path;
  if (!parts.has(workbookPath)) {
    throw new IntakeError('PART_MISSING', `工作簿部件不存在: ${workbookPath}`);
  }
  const baseDir = workbookPath.includes('/') ? workbookPath.slice(0, workbookPath.lastIndexOf('/')) : '';
  const relsPath = `${baseDir}/_rels/${workbookPath.slice(workbookPath.lastIndexOf('/') + 1)}.rels`;
  const workbookRels = parseRelationships(textOf(parts, relsPath), baseDir);

  let date1904 = false;
  let fullCalcOnLoad = false;
  const sheets = [];
  parseXml(textOf(parts, workbookPath), {
    onOpen(path, attrs) {
      const tag = path[path.length - 1];
      if (tag === 'workbookPr' && (attrs.date1904 === '1' || attrs.date1904 === 'true')) {
        date1904 = true;
      }
      if (tag === 'calcPr' && (attrs.fullCalcOnLoad === '1' || attrs.fullCalcOnLoad === 'true')) {
        fullCalcOnLoad = true;
      }
      if (tag === 'sheet') {
        if (!attrs.name || !attrs['r:id']) {
          throw new IntakeError('WORKBOOK_SHEET', '工作簿 sheet 条目缺少 name 或 r:id');
        }
        sheets.push({name: attrs.name, rid: attrs['r:id']});
      }
    },
  });

  const names = new Set();
  for (const sh of sheets) {
    if (names.has(sh.name)) {
      throw new IntakeError('SHEET_DUPLICATE', `工作簿存在同名工作表: ${sh.name}`, {sheet: sh.name});
    }
    names.add(sh.name);
    const rel = requireRel(workbookRels, sh.rid, '/worksheet', `工作表 ${sh.name}`);
    if (!parts.has(rel.path)) {
      throw new IntakeError('PART_MISSING', `工作表 ${sh.name} 的关系目标不存在: ${rel.path}`, {
        sheet: sh.name,
        path: rel.path,
      });
    }
    sh.path = rel.path;
  }

  let sharedStringsPath = null;
  let stylesPath = null;
  for (const rel of workbookRels.values()) {
    if (rel.type === `${REL_NS}sharedStrings`) sharedStringsPath = rel.path;
    if (rel.type === `${REL_NS}styles`) stylesPath = rel.path;
  }
  return {workbookPath, workbookRels, date1904, fullCalcOnLoad, sheets, sharedStringsPath, stylesPath};
}

// ---------- 共享字符串（富文本拼接多个 r/t） ----------

export function parseSharedStrings(parts, path) {
  if (!path || !parts.has(path)) return [];
  const strings = [];
  let depth = 0;
  let inSi = false;
  let siDepth = 0;
  let pieces = [];
  parseXml(parts.get(path).toString('utf8'), {
    onOpen(openPath) {
      depth++;
      const tag = openPath[openPath.length - 1];
      if (tag === 'si') {
        inSi = true;
        siDepth = depth;
        pieces = [];
      }
    },
    onText(openPath, text) {
      if (!inSi) return;
      const tag = openPath[openPath.length - 1];
      if (tag === 't') pieces.push(text);
    },
    onClose() {
      if (inSi && depth === siDepth) {
        strings.push(pieces.join(''));
        inSi = false;
      }
      depth--;
    },
  });
  return strings;
}

// ---------- 样式：识别纯补零自定义格式 ----------

const BUILTIN_ZERO_PAD = new Set(); // 内置格式中没有纯补零编号格式
const DATE_BUILTIN_IDS = new Set([14, 15, 16, 17, 22, 27, 30, 36, 50, 57]);

function stripFormatLiterals(code) {
  let out = '';
  let quote = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') { quote = !quote; continue; }
    if (quote) continue;
    if (ch === '\\') { i++; continue; }
    out += ch;
  }
  return out;
}

// 纯补零格式：分号分段后第一段仅由 0 与 ? 构成且至少含一个 0
export function isZeroPadFormat(formatCode) {
  if (!formatCode) return false;
  const first = stripFormatLiterals(formatCode).split(';')[0];
  return /^[0?]+$/.test(first) && first.includes('0');
}

export function parseStyles(parts, path) {
  if (!path || !parts.has(path)) return {zeroPadXf: new Set(), zeroPadFormats: new Map(), dateXf: new Set()};
  const customFormats = new Map();
  const xfNumFmtIds = [];
  parseXml(parts.get(path).toString('utf8'), {
    onOpen(openPath, attrs) {
      const tag = openPath[openPath.length - 1];
      if (tag === 'numFmt' && attrs.numFmtId && attrs.formatCode) {
        customFormats.set(Number(attrs.numFmtId), attrs.formatCode);
      }
      if (tag === 'xf' && openPath.includes('cellXfs')) {
        xfNumFmtIds.push(attrs.numFmtId === undefined ? 0 : Number(attrs.numFmtId));
      }
    },
  });
  const zeroPadXf = new Set();
  const zeroPadFormats = new Map();
  const dateXf = new Set();
  xfNumFmtIds.forEach((numFmtId, xf) => {
    const custom = customFormats.get(numFmtId);
    if (custom !== undefined) {
      if (isZeroPadFormat(custom)) {
        zeroPadXf.add(xf);
        zeroPadFormats.set(xf, custom);
      }
      if (/date|time|[ymdh]/.test(stripFormatLiterals(custom).toLowerCase())) dateXf.add(xf);
    } else if (BUILTIN_ZERO_PAD.has(numFmtId)) {
      zeroPadXf.add(xf);
    } else if (DATE_BUILTIN_IDS.has(numFmtId)) {
      dateXf.add(xf);
    }
  });
  return {zeroPadXf, zeroPadFormats, dateXf};
}

// ---------- 工作表 ----------

function inlineText(cell) {
  if (!cell.inline) return null;
  return cell.inline.join('');
}

export function parseSheet(xml, limits = {}) {
  const maxCells = limits.maxCells ?? 1_000_000;
  const cells = new Map();
  const merges = [];
  let cellCount = 0;
  let current = null;
  let cellDepth = 0;
  let depth = 0;
  parseXml(xml, {
    onOpen(openPath, attrs) {
      depth++;
      const tag = openPath[openPath.length - 1];
      if (tag === 'c' && openPath.includes('sheetData')) {
        if (!attrs.r) throw new IntakeError('SHEET_CELL', '单元格缺少 r 坐标属性，无法确定列号');
        if (cells.has(attrs.r)) throw new IntakeError('SHEET_CELL', `单元格坐标重复: ${attrs.r}`);
        if (++cellCount > maxCells) {
          throw new IntakeError('LIMIT_EXCEEDED', `工作表单元格数超过限额 ${maxCells}`);
        }
        current = {
          ref: attrs.r,
          type: attrs.t ?? 'n',
          style: attrs.s === undefined ? null : Number(attrs.s),
          value: null,
          inline: null,
          formula: null,
          depth,
        };
        cellDepth = depth;
        cells.set(attrs.r, current);
      } else if (current && tag === 'is' && depth === cellDepth + 1) {
        current.inline = [];
      } else if (current && tag === 'f' && depth === cellDepth + 1) {
        current.formula = '';
      } else if (current && tag === 'v' && depth === cellDepth + 1) {
        current.value = '';
      } else if (tag === 'mergeCell') {
        if (!attrs.ref) throw new IntakeError('SHEET_MERGE', 'mergeCell 缺少 ref 属性');
        merges.push(attrs.ref);
      }
    },
    onText(openPath, text) {
      if (!current) return;
      const tag = openPath[openPath.length - 1];
      if (tag === 'v') current.value = current.value === null ? text : current.value + text;
      else if (tag === 'f') current.formula = current.formula === null ? text : current.formula + text;
      else if (tag === 't' && current.inline !== null) current.inline.push(text);
    },
    onClose(closePath) {
      if (current && depth === cellDepth) {
        current = null;
      }
      depth--;
    },
  });
  return {cells, merges, cellCount};
}

export {inlineText};
