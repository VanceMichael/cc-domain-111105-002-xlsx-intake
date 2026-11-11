// ZIP 读取器：在既有目录检查（source.mjs）之上做完整解压，支持 stored/deflate 与限额控制。
import {inflateRawSync} from 'node:zlib';

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;

function findEocd(buffer) {
  const b = Buffer.from(buffer);
  const floor = Math.max(0, b.length - 65557);
  for (let i = b.length - 22; i >= floor; i--) {
    if (i < 0) break;
    if (b.readUInt32LE(i) === SIG_EOCD && i + 22 + b.readUInt16LE(i + 20) === b.length) return i;
  }
  throw new Error('缺少 ZIP 结束目录');
}

export function listZipEntries(input) {
  const b = Buffer.from(input);
  const end = findEocd(b);
  if (b.readUInt16LE(end + 4) !== 0 || b.readUInt16LE(end + 6) !== 0) {
    throw new Error('不支持分卷 ZIP');
  }
  const count = b.readUInt16LE(end + 10);
  const size = b.readUInt32LE(end + 12);
  const offset = b.readUInt32LE(end + 16);
  if (offset + size > end || count === 65535) throw new Error('ZIP 目录边界无效');
  let p = offset;
  const entries = [];
  for (let n = 0; n < count; n++) {
    if (p + 46 > offset + size || b.readUInt32LE(p) !== SIG_CENTRAL) {
      throw new Error('ZIP 条目不完整');
    }
    const flags = b.readUInt16LE(p + 8);
    const method = b.readUInt16LE(p + 10);
    const nameSize = b.readUInt16LE(p + 28);
    const extra = b.readUInt16LE(p + 30);
    const comment = b.readUInt16LE(p + 32);
    if (p + 46 + nameSize + extra + comment > offset + size) throw new Error('ZIP 名称越界');
    const name = b.subarray(p + 46, p + 46 + nameSize).toString('utf8');
    entries.push({
      name,
      method,
      flags,
      crc32: b.readUInt32LE(p + 16),
      compressed: b.readUInt32LE(p + 20),
      expanded: b.readUInt32LE(p + 24),
      localOffset: b.readUInt32LE(p + 42),
    });
    p += 46 + nameSize + extra + comment;
  }
  if (p !== offset + size) throw new Error('ZIP 目录长度不一致');
  return entries;
}

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

export const DEFAULT_LIMITS = {
  maxEntries: 10000,
  maxExpandedBytes: 100 * 1024 * 1024,
  maxExpandedEntryBytes: 50 * 1024 * 1024,
};

// 返回 Map<部件名, Buffer>；任何部件缺失、越界或超过限额都整体拒收。
export function extractZip(input, limits = {}) {
  const lim = {...DEFAULT_LIMITS, ...limits};
  const b = Buffer.from(input);
  const entries = listZipEntries(b);
  if (entries.length > lim.maxEntries) {
    throw new Error(`ZIP 部件数量 ${entries.length} 超过限额 ${lim.maxEntries}`);
  }
  const seen = new Set();
  let totalExpanded = 0;
  const parts = new Map();
  for (const e of entries) {
    if (seen.has(e.name)) throw new Error(`ZIP 部件重名: ${e.name}`);
    seen.add(e.name);
    if (e.name.endsWith('/')) continue;
    if (e.expanded > lim.maxExpandedEntryBytes) {
      throw new Error(`ZIP 部件 ${e.name} 展开后 ${e.expanded} 字节超过单部件限额`);
    }
    totalExpanded += e.expanded;
    if (totalExpanded > lim.maxExpandedBytes) {
      throw new Error(`ZIP 展开总量超过限额 ${lim.maxExpandedBytes} 字节`);
    }
    const o = e.localOffset;
    if (o + 30 > b.length || b.readUInt32LE(o) !== SIG_LOCAL) {
      throw new Error(`ZIP 部件 ${e.name} 的本地文件头缺失或越界`);
    }
    const localNameSize = b.readUInt16LE(o + 26);
    const localExtraSize = b.readUInt16LE(o + 28);
    const dataStart = o + 30 + localNameSize + localExtraSize;
    if (dataStart + e.compressed > b.length) {
      throw new Error(`ZIP 部件 ${e.name} 的压缩数据被截断`);
    }
    const raw = b.subarray(dataStart, dataStart + e.compressed);
    let data;
    if (e.method === 0) {
      data = Buffer.from(raw);
    } else if (e.method === 8) {
      try {
        data = inflateRawSync(raw, {maxOutputLength: e.expanded || lim.maxExpandedEntryBytes});
      } catch (err) {
        throw new Error(`ZIP 部件 ${e.name} 解压失败: ${err.message}`);
      }
    } else {
      throw new Error(`ZIP 部件 ${e.name} 使用不支持的压缩方式 ${e.method}`);
    }
    if (data.length !== e.expanded) {
      throw new Error(`ZIP 部件 ${e.name} 展开长度与目录不一致`);
    }
    if (crc32(data) !== e.crc32) throw new Error(`ZIP 部件 ${e.name} 校验和不一致`);
    parts.set(e.name, data);
  }
  return parts;
}
