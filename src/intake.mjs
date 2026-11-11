// 工作簿接入主管线：XLSX 字节 + 映射 -> 结构化记录与逐字段解释。
import {createHash} from 'node:crypto';
import {extractZip} from './zip.mjs';
import {
  IntakeError,
  parseWorkbook,
  parseSharedStrings,
  parseStyles,
  parseSheet,
  parseRange,
  columnIndex,
  columnLetter,
  inlineText,
} from './ooxml.mjs';
import {interpretDecimal, interpretIdentifier, interpretDate, scaleExponent} from './interpret.mjs';

const VALID_KINDS = new Set(['identifier', 'decimal', 'date', 'text']);

export function sha256Hex(buffer) {
  return createHash('sha256').update(buffer).digest('hex');
}

function validateMapping(mapping) {
  if (mapping === null || typeof mapping !== 'object' || Array.isArray(mapping)) {
    throw new IntakeError('MAPPING_SHAPE', '映射必须是 JSON 对象');
  }
  if (typeof mapping.version !== 'string' || mapping.version.trim() === '') {
    throw new IntakeError('MAPPING_SHAPE', '映射缺少非空 version');
  }
  if (typeof mapping.sheet !== 'string' || mapping.sheet.trim() === '') {
    throw new IntakeError('MAPPING_SHAPE', '映射缺少非空 sheet');
  }
  const range = parseRange(mapping.range);
  if (!mapping.columns || typeof mapping.columns !== 'object' || Array.isArray(mapping.columns)) {
    throw new IntakeError('MAPPING_SHAPE', '映射缺少 columns 对象');
  }
  const names = new Set();
  const columns = new Map();
  for (const [letter, spec] of Object.entries(mapping.columns)) {
    if (!/^[A-Z]{1,3}$/.test(letter)) throw new IntakeError('MAPPING_SHAPE', `映射列键不是列字母: ${letter}`);
    const {col} = columnIndex(`${letter}1`);
    if (col < range.minCol || col > range.maxCol) {
      throw new IntakeError('MAPPING_RANGE', `映射列 ${letter} 不在范围 ${mapping.range} 内`);
    }
    if (spec === null || typeof spec !== 'object') {
      throw new IntakeError('MAPPING_SHAPE', `列 ${letter} 的定义必须是对象`);
    }
    if (!VALID_KINDS.has(spec.kind)) {
      throw new IntakeError('MAPPING_SHAPE', `列 ${letter} 的 kind 不支持: ${spec.kind}`);
    }
    if (typeof spec.name !== 'string' || spec.name.trim() === '') {
      throw new IntakeError('MAPPING_SHAPE', `列 ${letter} 缺少字段名 name`);
    }
    if (names.has(spec.name)) throw new IntakeError('MAPPING_SHAPE', `字段名重复: ${spec.name}`);
    names.add(spec.name);
    const entry = {letter, col, name: spec.name, kind: spec.kind, group: spec.group === true};
    if (entry.group && spec.kind !== 'text' && spec.kind !== 'identifier') {
      throw new IntakeError('MAPPING_SHAPE', `分组列 ${letter} 只允许 text/identifier 类型`);
    }
    if (spec.kind === 'decimal') {
      if (spec.scale === undefined) throw new IntakeError('MAPPING_SCALE', `decimal 列 ${letter} 缺少 scale`);
      entry.scale = String(spec.scale);
      entry.scaleExponent = scaleExponent(entry.scale);
    }
    columns.set(col, entry);
  }
  if (columns.size === 0) throw new IntakeError('MAPPING_SHAPE', '映射至少需要一列');
  return {
    version: mapping.version,
    sheet: mapping.sheet,
    rangeText: mapping.range,
    range,
    columns,
    groupCols: new Set([...columns.values()].filter((c) => c.group).map((c) => c.col)),
    digest: sha256Hex(Buffer.from(canonicalJson(mapping))),
  };
}

function canonicalJson(value) {
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(',')}]`;
  const keys = Object.keys(value).sort();
  return `{${keys.map((k) => `${JSON.stringify(k)}:${canonicalJson(value[k])}`).join(',')}}`;
}

// 仅接受与映射范围相交的纵向合并；横向合并、非分组列合并直接拒收
function buildMergeIndex(merges, mapping) {
  const index = new Map();
  for (const ref of merges) {
    const m = /^([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)$/.exec(ref);
    if (!m) throw new IntakeError('SHEET_MERGE', `合并区域坐标无法解析: ${ref}`);
    const top = columnIndex(`${m[1]}${m[2]}`);
    const bottom = columnIndex(`${m[3]}${m[4]}`);
    if (bottom.col !== top.col) {
      throw new IntakeError('MERGE_SHAPE', `合并区域 ${ref} 跨越多列，仅支持单列纵向合并`, {ref});
    }
    const col = top.col;
    if (bottom.row < top.row) throw new IntakeError('SHEET_MERGE', `合并区域方向错误: ${ref}`);
    if (col < mapping.range.minCol || col > mapping.range.maxCol) continue;
    if (bottom.row < mapping.range.minRow || top.row > mapping.range.maxRow) continue;
    if (!mapping.groupCols.has(col)) {
      throw new IntakeError(
        'MERGE_NOT_GROUP',
        `合并区域 ${ref} 位于非分组列 ${m[1]}，指标值不允许向下填充`,
        {ref},
      );
    }
    index.set(col, [...(index.get(col) ?? []), {topRow: top.row, bottomRow: bottom.row, ref}]);
  }
  return index;
}

function findMerge(mergeIndex, col, row) {
  const list = mergeIndex.get(col);
  if (!list) return null;
  return list.find((g) => row > g.topRow && row <= g.bottomRow) ?? null;
}

function stringContent(cell, shared) {
  if (cell.type === 's') {
    const idx = Number(cell.value);
    if (!Number.isInteger(idx) || idx < 0 || idx >= shared.length) {
      return {error: 'SHARED_STRING', message: `共享字符串索引越界: ${cell.value}`};
    }
    return {text: shared[idx]};
  }
  if (cell.type === 'inlineStr') return {text: inlineText(cell) ?? ''};
  if (cell.type === 'str') return {text: cell.value ?? ''};
  return {error: 'TYPE_MISMATCH', message: `文本列遇到类型 ${cell.type}`};
}

function interpretCell(cell, spec, ctx) {
  const {shared, styles, epoch1904, fullCalcOnLoad} = ctx;
  const base = {
    sheet: ctx.sheetName,
    ref: cell.ref,
    rawType: cell.type,
    styleIndex: cell.style,
    formula: cell.formula,
  };

  // 公式单元格：本机不求值，按缓存状态分流
  if (cell.formula !== null) {
    if (fullCalcOnLoad) {
      return {...base, status: 'pending', reason: 'FULL_CALC_ON_LOAD', value: null};
    }
    if (cell.type === 'e') {
      return {...base, status: 'pending', reason: 'CACHED_ERROR', cached: cell.value, value: null};
    }
    if (cell.value === null || cell.value === '') {
      return {...base, status: 'pending', reason: 'NO_CACHED_VALUE', value: null};
    }
    const cached = interpretValue(cell, spec, shared, styles, epoch1904);
    if (cached.error || cached.unavailable) {
      return {
        ...base,
        status: 'pending',
        reason: cached.error ? 'CACHE_TYPE_MISMATCH' : cached.unavailable,
        cached: cell.value,
        value: null,
      };
    }
    return {...base, status: 'ok', cached: true, value: cached.value};
  }

  if (cell.type === 'e') {
    return {...base, status: 'error', reason: 'ERROR_VALUE', raw: cell.value, value: null};
  }
  // inlineStr 的内容在 <is> 中而非 <v>，必须走解释器；其余类型无 <v> 视为空白
  if ((cell.value === null || cell.value === '') && cell.type !== 'inlineStr') {
    return {...base, status: 'empty', reason: 'BLANK', value: null};
  }
  const r = interpretValue(cell, spec, shared, styles, epoch1904);
  if (r.error) return {...base, status: 'error', reason: r.error, message: r.message, value: null};
  if (r.unavailable) return {...base, status: 'unavailable', reason: r.unavailable, message: r.message, value: null};
  return {...base, status: 'ok', value: r.value};
}

function interpretValue(cell, spec, shared, styles, epoch1904) {
  // 文本载体（共享字符串/inlineStr/公式字符串）只允许出现在文本列或标识列
  if (cell.type === 's' || cell.type === 'inlineStr' || cell.type === 'str') {
    const r = stringContent(cell, shared);
    if (r.error) return {error: r.error, message: r.message};
    if (spec.kind === 'text') return {value: {text: r.text}};
    if (spec.kind === 'identifier') return {value: {text: r.text, zeroPadded: false}};
    return {error: 'TYPE_MISMATCH', message: `${spec.kind} 列遇到字符串类型 ${cell.type}`};
  }

  if (cell.type !== 'n' && cell.type !== 'b') {
    return {error: 'TYPE_MISMATCH', message: `不支持的单元格类型 ${cell.type}`};
  }
  const raw = cell.value;
  if (cell.type === 'b') {
    if (spec.kind !== 'identifier') return {error: 'TYPE_MISMATCH', message: '布尔单元格不能用于该列类型'};
    return {value: {text: cell.value === '1' ? '1' : '0', zeroPadded: false}};
  }

  if (spec.kind === 'identifier') {
    return interpretIdentifier(raw, cell.style, styles.zeroPadFormats);
  }
  if (spec.kind === 'decimal') return interpretDecimal(raw, spec.scale, spec.scaleExponent);
  return interpretDate(raw, epoch1904);
}

function blankField(ref, spec, sheetName, reason = 'BLANK') {
  return {
    field: spec.name,
    kind: spec.kind,
    sheet: sheetName,
    ref,
    rawType: null,
    formula: null,
    status: 'empty',
    reason,
    value: null,
  };
}

export function intakeWorkbook(bytes, mappingInput, options = {}) {
  const input = Buffer.from(bytes);
  const mapping = validateMapping(mappingInput);
  const sourceSha256 = sha256Hex(input);
  const intakeId = sha256Hex(
    Buffer.from(`xlsx-intake:${sourceSha256}:${mapping.version}:${mapping.digest}`),
  );
  const identity = {
    intakeId,
    sourceSha256,
    sourceBytes: input.length,
    mappingVersion: mapping.version,
    mappingDigest: mapping.digest,
  };

  let parts;
  try {
    parts = extractZip(input, options.limits);
  } catch (err) {
    return rejection(identity, 'ZIP_INVALID', err.message);
  }

  let workbook;
  let sheetXml;
  let sheetPath;
  try {
    workbook = parseWorkbook(parts);
    const sheet = workbook.sheets.find((sh) => sh.name === mapping.sheet);
    if (!sheet) {
      const names = workbook.sheets.map((sh) => sh.name).join('、') || '(空)';
      throw new IntakeError(
        'SHEET_NOT_FOUND',
        `映射指定工作表「${mapping.sheet}」不存在；工作簿实际包含: ${names}`,
        {sheet: mapping.sheet, available: workbook.sheets.map((sh) => sh.name)},
      );
    }
    if (workbook.sharedStringsPath && !parts.has(workbook.sharedStringsPath)) {
      workbook.sharedStringsPath = null;
    }
    if (workbook.stylesPath && !parts.has(workbook.stylesPath)) {
      workbook.stylesPath = null;
    }
    sheetPath = sheet.path;
    sheetXml = parts.get(sheet.path).toString('utf8');
  } catch (err) {
    return rejection(identity, err instanceof IntakeError ? err.code : 'PACKAGE_INVALID', err.message, err.detail);
  }

  let shared;
  let styles;
  let parsed;
  let mergeIndex;
  try {
    shared = parseSharedStrings(parts, workbook.sharedStringsPath);
    styles = parseStyles(parts, workbook.stylesPath);
    parsed = parseSheet(sheetXml, options.limits);
    mergeIndex = buildMergeIndex(parsed.merges, mapping);
  } catch (err) {
    return rejection(identity, err instanceof IntakeError ? err.code : 'SHEET_INVALID', err.message, err.detail);
  }

  const ctx = {
    sheetName: mapping.sheet,
    shared,
    styles,
    epoch1904: workbook.date1904,
    fullCalcOnLoad: workbook.fullCalcOnLoad,
  };

  const rows = [];
  const stats = {total: 0, ok: 0, pending: 0, error: 0, unavailable: 0, blank: 0};
  for (let rowNum = mapping.range.minRow; rowNum <= mapping.range.maxRow; rowNum++) {
    stats.total++;
    const fields = [];
    let rowStatus = 'ok';
    const escalate = (status) => {
      if (status === 'error') rowStatus = 'error';
      else if (status === 'unavailable' && rowStatus !== 'error') rowStatus = 'unavailable';
      else if (status === 'pending' && rowStatus === 'ok') rowStatus = 'pending';
    };

    for (const spec of [...mapping.columns.values()].sort((a, b) => a.col - b.col)) {
      const ref = `${columnLetter(spec.col)}${rowNum}`;
      let cell = parsed.cells.get(ref);
      let field;
      if (!cell) {
        const merge = findMerge(mergeIndex, spec.col, rowNum);
        if (merge && mapping.groupCols.has(spec.col)) {
          const anchorRef = `${columnLetter(spec.col)}${merge.topRow}`;
          const anchor = parsed.cells.get(anchorRef);
          if (!anchor) {
            field = {
              ...blankField(ref, spec, mapping.sheet, 'MERGE_ANCHOR_MISSING'),
              status: 'error',
              reason: 'MERGE_ANCHOR_MISSING',
              merge: merge.ref,
            };
          } else {
            const anchorField = interpretCell(anchor, spec, ctx);
            field = {
              field: spec.name,
              kind: spec.kind,
              ...anchorField,
              ref,
              anchorRef,
              merge: merge.ref,
              filledFromMerge: true,
            };
          }
        } else {
          field = blankField(ref, spec, mapping.sheet);
        }
      } else {
        const interpreted = interpretCell(cell, spec, ctx);
        field = {field: spec.name, kind: spec.kind, ...interpreted};
      }
      fields.push(field);
      stats[field.status] = (stats[field.status] ?? 0) + 1;
      escalate(field.status);
    }
    stats[rowStatus]++;
    rows.push({row: rowNum, status: rowStatus, fields});
  }

  return {
    accepted: true,
    identity,
    workbook: {
      sheet: mapping.sheet,
      sheetPath,
      range: mapping.rangeText,
      dateSystem: workbook.date1904 ? '1904' : '1900',
      fullCalcOnLoad: workbook.fullCalcOnLoad,
      sharedStringCount: shared.length,
      cellCount: parsed.cellCount,
    },
    stats,
    rows,
  };
}

// BigInt 在 JSON 中以字符串输出，避免序列化为 Number 时丢失精度
export function jsonReplacer(key, value) {
  return typeof value === 'bigint' ? value.toString() : value;
}

function rejection(identity, code, message, detail) {
  return {
    accepted: false,
    identity,
    rejection: {code, message, ...(detail !== undefined ? {detail} : {})},
  };
}
