// 类型解释：标识补零、十进制精确缩放（BigInt，不走浮点）、1900/1904 双基准日期。
import {IntakeError} from './ooxml.mjs';

const DECIMAL_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)$/;
const INT_RE = /^\d+$/;

// scale 必须是正的 10 的整数幂，返回幂次
export function scaleExponent(scaleText) {
  if (typeof scaleText !== 'string' || !/^\d+$/.test(scaleText) || !/^[1-9]/.test(scaleText)) {
    throw new IntakeError('MAPPING_SCALE', `decimal 列的 scale 必须是正整数字符串: ${scaleText}`);
  }
  const m = /^1(0*)$/.exec(scaleText);
  if (!m) throw new IntakeError('MAPPING_SCALE', `decimal 列的 scale 必须是 10 的整数幂: ${scaleText}`);
  return m[1].length;
}

// 十进制原文按 10^k 缩放，结果仍是精确定点小数字符串
export function scaleDecimal(raw, exponent) {
  const negative = raw.startsWith('-');
  const unsigned = negative || raw.startsWith('+') ? raw.slice(1) : raw;
  const dot = unsigned.indexOf('.');
  const intPart = dot < 0 ? unsigned : unsigned.slice(0, dot);
  const fracPart = dot < 0 ? '' : unsigned.slice(dot + 1);
  const digits = `${intPart}${fracPart}`;
  const pointAt = intPart.length + exponent; // 新小数点相对数字串左端的位置
  let out;
  if (pointAt <= 0) {
    out = `0.${'0'.repeat(-pointAt)}${digits}`;
  } else if (pointAt >= digits.length) {
    out = digits + '0'.repeat(pointAt - digits.length);
  } else {
    out = `${digits.slice(0, pointAt)}.${digits.slice(pointAt)}`;
  }
  out = out
    .replace(/^0+(?=\d)/, '')
    .replace(/(\.\d*?)0+$/, '$1')
    .replace(/\.$/, '');
  if (out.startsWith('.')) out = `0${out}`;
  if (out === '' || out === '-') out = '0';
  return negative && out !== '0' ? `-${out}` : out;
}

// 用 BigInt 做一次交叉校验，证明缩放结果没有经过浮点
export function exactIntegerOrNull(scaled) {
  if (/^-?\d+$/.test(scaled)) return BigInt(scaled);
  return null;
}

export function interpretDecimal(raw, scaleText, exponent) {
  if (!DECIMAL_RE.test(raw)) {
    return {error: 'DECIMAL_INVALID', message: `数值原文不是合法十进制: ${raw}`};
  }
  const scaled = scaleDecimal(raw, exponent);
  return {value: {raw, scaled, scale: scaleText, exactInteger: exactIntegerOrNull(scaled)}};
}

function padWidth(formatCode) {
  const first = formatCode.split(';')[0].replace(/[^0?]/g, '');
  return first.length;
}

export function interpretIdentifier(raw, style, zeroPadFormats) {
  if (style !== null && zeroPadFormats.has(style)) {
    if (!INT_RE.test(raw)) {
      return {error: 'IDENTIFIER_NOT_INTEGER', message: `补零标识必须是整数原文: ${raw}`};
    }
    const width = padWidth(zeroPadFormats.get(style));
    if (raw.length >= width) return {value: {text: raw, zeroPadded: true, width}};
    return {value: {text: '0'.repeat(width - raw.length) + raw, zeroPadded: true, width}};
  }
  if (!DECIMAL_RE.test(raw)) {
    return {error: 'IDENTIFIER_INVALID', message: `标识原文无法识别: ${raw}`};
  }
  return {value: {text: raw, zeroPadded: false}};
}

const MS_PER_DAY = 86400000;

function utcIso(ms) {
  return new Date(ms).toISOString();
}

// 1900 基准：序号 60 整日（含当日时间分量）是虚构的 1900-02-29，按要求标为不可用；>=61 补偿一天闰日缺陷
export function interpretDate(raw, epoch1904) {
  const serial = Number(raw);
  if (!Number.isFinite(serial)) return {error: 'DATE_INVALID', message: `日期序号无法识别: ${raw}`};
  if (!epoch1904 && Math.trunc(serial) === 60) {
    return {unavailable: 'DATE_SERIAL_60', message: '1900 基准序号 60（虚构闰日）不可用'};
  }
  const whole = Math.trunc(serial);
  const timeMs = Math.round((Math.abs(serial) - Math.abs(whole)) * MS_PER_DAY);
  let ms;
  if (epoch1904) {
    // 序号 0 = 1904-01-01T00:00:00Z，允许负序号
    ms = Date.UTC(1904, 0, 1) + whole * MS_PER_DAY + (serial < 0 ? -timeMs : timeMs);
  } else if (serial < 0) {
    return {error: 'DATE_INVALID', message: '1900 基准不支持负日期序号'};
  } else if (whole >= 61) {
    ms = Date.UTC(1899, 11, 31) + (whole - 1) * MS_PER_DAY + timeMs;
  } else {
    ms = Date.UTC(1899, 11, 31) + whole * MS_PER_DAY + timeMs;
  }
  return {value: {serial: raw, iso: utcIso(ms), epoch: epoch1904 ? '1904' : '1900'}};
}
