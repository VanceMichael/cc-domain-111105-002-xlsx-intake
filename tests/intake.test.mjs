import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {intakeWorkbook} from '../src/intake.mjs';
import {extractZip} from '../src/zip.mjs';
import {buildWorkbook, buildZip} from './xlsx-builder.mjs';
import {readFixture} from '../src/files.mjs';

const mapping = JSON.parse(await readFile(new URL('../fixtures/mapping.json', import.meta.url)));
const f1900 = await readFixture('disclosure-1900.xlsx');
const f1904 = await readFixture('disclosure-1904.xlsx');
const fRecalc = await readFixture('recalculate.xlsx');
const fTruncated = await readFixture('truncated.xlsx.bin');

function field(row, name) {
  return row.fields.find((f) => f.field === name);
}

test('1900 基准：补零标识、公式缓存、序号 60 不可用与跨行合并', () => {
  const r = intakeWorkbook(f1900, mapping);
  assert.equal(r.accepted, true);
  assert.equal(r.workbook.dateSystem, '1900');
  assert.equal(r.workbook.fullCalcOnLoad, false);

  const [row2, row3, row4] = r.rows;
  // A2 来自共享字符串，A3 由 A2:A3 合并沿分组列向下提供
  assert.equal(field(row2, 'issuer').value.text, '合成甲厂');
  assert.equal(field(row2, 'issuer').rawType, 's');
  const a3 = field(row3, 'issuer');
  assert.equal(a3.status, 'ok');
  assert.equal(a3.value.text, '合成甲厂');
  assert.equal(a3.filledFromMerge, true);
  assert.equal(a3.anchorRef, 'A2');
  assert.equal(a3.merge, 'A2:A3');
  assert.equal(a3.ref, 'A3');
  // 指标列没有被合并值填充
  assert.equal(field(row3, 'water').status, 'pending');

  // 编号 123 + 纯补零样式 000000 -> 000123
  const b2 = field(row2, 'reference');
  assert.equal(b2.status, 'ok');
  assert.equal(b2.value.text, '000123');
  assert.equal(b2.value.zeroPadded, true);
  assert.equal(b2.value.width, 6);
  // 编号 7 -> 000007
  assert.equal(field(row4, 'reference').value.text, '000007');
  // 超长整数以 inlineStr 原样保留，不经浮点
  const b3 = field(row3, 'reference');
  assert.equal(b3.status, 'ok');
  assert.equal(b3.value.text, '9007199254740993');
  assert.equal(b3.rawType, 'inlineStr');

  // D2 公式有数值缓存：记录缓存值与公式来源
  const d2 = field(row2, 'water');
  assert.equal(d2.status, 'ok');
  assert.equal(d2.cached, true);
  assert.equal(d2.formula, 'SUM(1,2)');
  assert.equal(d2.value.raw, '3');
  assert.equal(d2.value.scaled, '3000');
  assert.equal(d2.value.exactInteger, 3000n);
  // D3 公式无缓存 -> 待确认，不混入有效数值
  const d3 = field(row3, 'water');
  assert.equal(d3.status, 'pending');
  assert.equal(d3.reason, 'NO_CACHED_VALUE');
  assert.equal(d3.formula, 'SUM(2,3)');
  assert.equal(d3.value, null);
  // D4 缓存为错误 -> 待确认
  const d4 = field(row4, 'water');
  assert.equal(d4.status, 'pending');
  assert.equal(d4.reason, 'CACHED_ERROR');
  assert.equal(d4.cached, '#DIV/0!');

  // 日期：序号 60 不可用；61 -> 1900-03-01；45000 -> 2023-03-15
  const e2 = field(row2, 'date');
  assert.equal(e2.status, 'unavailable');
  assert.equal(e2.reason, 'DATE_SERIAL_60');
  assert.equal(field(row3, 'date').value.iso, '1900-03-01T00:00:00.000Z');
  assert.equal(field(row4, 'date').value.iso, '2023-03-15T00:00:00.000Z');
  assert.equal(field(row4, 'date').value.serial, '45000');

  // 稀疏列 C 不在映射中；行级隔离：第 3、4 行有 pending，但整批仍然接收
  assert.equal(row2.status, 'unavailable');
  assert.equal(row3.status, 'pending');
  assert.equal(row4.status, 'pending');
  assert.equal(r.stats.total, 3);
});

test('1904 基准：同一序号给出不同日期，序号 60 正常', () => {
  const r = intakeWorkbook(f1904, mapping);
  assert.equal(r.accepted, true);
  assert.equal(r.workbook.dateSystem, '1904');
  const [row2, row3, row4] = r.rows;
  assert.equal(field(row2, 'date').status, 'ok');
  assert.equal(field(row2, 'date').value.iso, '1904-03-01T00:00:00.000Z');
  assert.equal(field(row2, 'date').value.epoch, '1904');
  assert.equal(field(row3, 'date').value.iso, '1904-03-02T00:00:00.000Z');
  assert.equal(field(row4, 'date').value.iso, '2027-03-16T00:00:00.000Z');
});

test('fullCalcOnLoad：所有公式缓存进入待确认', () => {
  const r = intakeWorkbook(fRecalc, mapping);
  assert.equal(r.accepted, true);
  assert.equal(r.workbook.fullCalcOnLoad, true);
  const d2 = field(r.rows[0], 'water');
  assert.equal(d2.status, 'pending');
  assert.equal(d2.reason, 'FULL_CALC_ON_LOAD');
  assert.equal(d2.value, null);
  const d3 = field(r.rows[1], 'water');
  assert.equal(d3.status, 'pending');
  assert.equal(d3.reason, 'FULL_CALC_ON_LOAD');
  // 非公式单元格照常解释
  assert.equal(field(r.rows[2], 'reference').value.text, '000007');
});

test('截断文件整体拒收，不成为正式批次', () => {
  const r = intakeWorkbook(fTruncated, mapping);
  assert.equal(r.accepted, false);
  assert.match(r.rejection.code, /^ZIP_/);
  assert.ok(r.rejection.message.length > 0);
  assert.ok(r.identity.sourceSha256);
  // 再截掉一段同样拒收
  const r2 = intakeWorkbook(f1900.subarray(0, f1900.length - 50), mapping);
  assert.equal(r2.accepted, false);
  assert.match(r2.rejection.code, /^ZIP_/);
  const r3 = intakeWorkbook(Buffer.from('not a zip at all'), mapping);
  assert.equal(r3.accepted, false);
  assert.equal(r3.rejection.code, 'ZIP_INVALID');
});

test('接入身份：同件文件重送可查回，换映射生成独立身份', () => {
  const a = intakeWorkbook(f1900, mapping);
  const b = intakeWorkbook(f1900, JSON.parse(JSON.stringify(mapping)));
  assert.equal(a.identity.intakeId, b.identity.intakeId);
  assert.equal(a.identity.sourceSha256, b.identity.sourceSha256);
  assert.equal(a.identity.mappingDigest, b.identity.mappingDigest);

  const remapped = {...mapping, version: 'water-v2', columns: {...mapping.columns}};
  const c = intakeWorkbook(f1900, remapped);
  assert.notEqual(c.identity.intakeId, a.identity.intakeId);
  assert.equal(c.identity.mappingVersion, 'water-v2');
  assert.equal(c.identity.sourceSha256, a.identity.sourceSha256);

  // 同版本但内容变化也应产生不同摘要
  const changed = {...mapping, columns: {...mapping.columns, D: {...mapping.columns.D, scale: '100'}}};
  const d = intakeWorkbook(f1900, changed);
  assert.notEqual(d.identity.mappingDigest, a.identity.mappingDigest);
  assert.notEqual(d.identity.intakeId, a.identity.intakeId);
});

test('同名工作表与错误关系给出定位明确的拒收原因', () => {
  const dup = buildWorkbook({
    sheets: [
      {name: '披露表', cells: [{ref: 'A1', inline: 'x'}]},
      {name: '披露表', cells: [{ref: 'A1', inline: 'y'}]},
    ],
  });
  const r1 = intakeWorkbook(dup, mapping);
  assert.equal(r1.accepted, false);
  assert.equal(r1.rejection.code, 'SHEET_DUPLICATE');
  assert.match(r1.rejection.message, /同名工作表/);
  assert.deepEqual(r1.rejection.detail, {sheet: '披露表'});

  const broken = buildWorkbook({
    relTargets: ['worksheets/missing.xml'],
  });
  const r2 = intakeWorkbook(broken, mapping);
  assert.equal(r2.accepted, false);
  assert.equal(r2.rejection.code, 'PART_MISSING');
  assert.match(r2.rejection.message, /worksheets\/missing\.xml/);

  const unknownSheet = {...mapping, sheet: '不存在的表'};
  const r3 = intakeWorkbook(f1900, unknownSheet);
  assert.equal(r3.accepted, false);
  assert.equal(r3.rejection.code, 'SHEET_NOT_FOUND');
  assert.deepEqual(r3.rejection.detail.available, ['披露表']);
});

test('横向合并与非分组列纵向合并被拒收，分组列之外不向下填充', () => {
  const horizontal = buildWorkbook({
    sheets: [{name: '披露表', merges: ['A2:B2'], cells: [{ref: 'A2', inline: 'x'}]}],
  });
  const r1 = intakeWorkbook(horizontal, mapping);
  assert.equal(r1.accepted, false);
  assert.equal(r1.rejection.code, 'MERGE_SHAPE');

  const metricMerge = buildWorkbook({
    sheets: [{name: '披露表', merges: ['D2:D3'], cells: [{ref: 'D2', v: '3'}]}],
  });
  const r2 = intakeWorkbook(metricMerge, mapping);
  assert.equal(r2.accepted, false);
  assert.equal(r2.rejection.code, 'MERGE_NOT_GROUP');
  assert.match(r2.rejection.message, /D2:D3/);
});

test('列号来自坐标而非节点顺序：乱序与稀疏单元格', () => {
  const wb = buildWorkbook({
    sheets: [{
      name: '披露表',
      cells: [
        {ref: 'E2', s: 2, v: '45000'},
        {ref: 'B2', s: 1, v: '42'},
        {ref: 'A2', t: 'inlineStr', inline: '乱序厂'},
        {ref: 'D2', v: '9.25'},
      ],
    }],
  });
  const r = intakeWorkbook(wb, mapping);
  assert.equal(r.accepted, true);
  const row = r.rows[0];
  assert.equal(field(row, 'issuer').ref, 'A2');
  assert.equal(field(row, 'reference').ref, 'B2');
  assert.equal(field(row, 'reference').value.text, '000042');
  assert.equal(field(row, 'water').ref, 'D2');
  assert.equal(field(row, 'water').value.raw, '9.25');
  assert.equal(field(row, 'water').value.scaled, '9250');
  assert.equal(field(row, 'date').ref, 'E2');
});

test('十进制精确定点缩放，不经过浮点取整', () => {
  const wb = buildWorkbook({
    sheets: [{
      name: '披露表',
      cells: [
        {ref: 'D2', v: '0.001'},
        {ref: 'D3', v: '12345678901234567890.5'},
        {ref: 'D4', v: '-7.5'},
      ],
    }],
  });
  const r = intakeWorkbook(wb, mapping);
  const [d2, d3, d4] = r.rows.map((row) => field(row, 'water'));
  assert.equal(d2.value.scaled, '1');
  assert.equal(d3.value.raw, '12345678901234567890.5');
  assert.equal(d3.value.scaled, '12345678901234567890500');
  assert.equal(d3.value.exactInteger, 12345678901234567890500n);
  assert.equal(d4.value.scaled, '-7500');
  assert.equal(d4.value.exactInteger, -7500n);
});

test('空值、错误单元格与类型不符按字段隔离', () => {
  const wb = buildWorkbook({
    sheets: [{
      name: '披露表',
      cells: [
        {ref: 'A2', t: 'inlineStr', inline: '厂'},
        {ref: 'D2', t: 'e', v: '#N/A'},
        {ref: 'E2', t: 'inlineStr', inline: '不是日期'},
      ],
    }],
  });
  const r = intakeWorkbook(wb, mapping);
  assert.equal(r.accepted, true);
  const row = r.rows[0];
  assert.equal(row.status, 'error');
  assert.equal(field(row, 'reference').status, 'empty');
  assert.equal(field(row, 'reference').reason, 'BLANK');
  assert.equal(field(row, 'water').status, 'error');
  assert.equal(field(row, 'water').reason, 'ERROR_VALUE');
  assert.equal(field(row, 'date').status, 'error');
  assert.equal(field(row, 'date').reason, 'TYPE_MISMATCH');
  assert.equal(field(row, 'issuer').status, 'ok');
});

test('富文本共享字符串拼接与索引越界', () => {
  const wb = buildWorkbook({
    shared: ['甲厂', '乙厂'],
    sheets: [{name: '披露表', cells: [{ref: 'A2', t: 's', v: '9'}]}],
  });
  const r = intakeWorkbook(wb, mapping);
  assert.equal(r.accepted, true);
  assert.equal(field(r.rows[0], 'issuer').status, 'error');
  assert.equal(field(r.rows[0], 'issuer').reason, 'SHARED_STRING');
});

test('限额：展开总量、部件数、单元格数超限整批拒收', () => {
  const big = buildZip([{name: 'xl/big.xml', data: Buffer.alloc(2048, 0x61)}]);
  const r1 = intakeWorkbook(big, mapping, {limits: {maxExpandedBytes: 1024}});
  assert.equal(r1.accepted, false);
  assert.equal(r1.rejection.code, 'ZIP_INVALID');
  assert.match(r1.rejection.message, /限额/);

  const many = buildZip(Array.from({length: 12}, (_, i) => ({name: `p${i}.xml`, data: 'x'})));
  const r2 = intakeWorkbook(many, mapping, {limits: {maxEntries: 10}});
  assert.equal(r2.accepted, false);
  assert.match(r2.rejection.message, /部件数量/);

  const wb = buildWorkbook({
    sheets: [{name: '披露表', cells: Array.from({length: 30}, (_, i) => ({ref: `A${i + 1}`, inline: 'x'}))}],
  });
  const r3 = intakeWorkbook(wb, mapping, {limits: {maxCells: 20}});
  assert.equal(r3.accepted, false);
  assert.equal(r3.rejection.code, 'LIMIT_EXCEEDED');
});

test('关系越界与外部关系被拒收', () => {
  const traversal = buildWorkbook({relTargets: ['../../etc/passwd']});
  const r1 = intakeWorkbook(traversal, mapping);
  assert.equal(r1.accepted, false);
  assert.equal(r1.rejection.code, 'REL_TRAVERSAL');

  const externalRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="http://example.test/sheet.xml" TargetMode="External"/></Relationships>`;
  const external = buildWorkbook({relsOverride: externalRels});
  const r2 = intakeWorkbook(external, mapping);
  assert.equal(r2.accepted, false);
  assert.equal(r2.rejection.code, 'REL_EXTERNAL');
});

test('映射校验：坏范围、坏 kind、坏 scale 与缺版本', () => {
  assert.throws(() => intakeWorkbook(f1900, {...mapping, range: 'A2'}), (e) => e.code === 'MAPPING_RANGE');
  assert.throws(() => intakeWorkbook(f1900, {
    ...mapping,
    columns: {...mapping.columns, A: {name: 'x', kind: 'float'}},
  }), (e) => e.code === 'MAPPING_SHAPE');
  assert.throws(() => intakeWorkbook(f1900, {
    ...mapping,
    columns: {...mapping.columns, D: {name: 'water', kind: 'decimal', scale: '7'}},
  }), (e) => e.code === 'MAPPING_SCALE');
  assert.throws(() => intakeWorkbook(f1900, {...mapping, version: ''}), (e) => e.code === 'MAPPING_SHAPE');
});

test('ZIP 解压可还原全部部件且校验 CRC', () => {
  const parts = extractZip(f1900);
  assert.ok(parts.has('xl/workbook.xml'));
  assert.ok(parts.has('xl/sharedStrings.xml'));
  assert.ok(parts.has('xl/worksheets/sheet1.xml'));
  assert.match(parts.get('xl/workbook.xml').toString('utf8'), /披露表/);
});

test('压缩字节被篡改时 CRC/解压失败整批拒收', () => {
  const tampered = Buffer.from(f1900);
  tampered[120] ^= 0xff;
  const r = intakeWorkbook(tampered, mapping);
  assert.equal(r.accepted, false);
  assert.equal(r.rejection.code, 'ZIP_INVALID');
});

test('1900 日期序号边界：1/59 可用、60 不可用、61 补偿闰日', () => {
  const wb = buildWorkbook({
    sheets: [{
      name: '披露表',
      cells: [
        {ref: 'E2', v: '1'},
        {ref: 'E3', v: '59'},
        {ref: 'E4', v: '60'},
      ],
    }],
  });
  const r = intakeWorkbook(wb, {
    version: 'v', sheet: '披露表', range: 'E2:E4',
    columns: {E: {name: 'date', kind: 'date'}},
  });
  assert.equal(r.accepted, true);
  assert.equal(r.rows[0].fields[0].value.iso, '1900-01-01T00:00:00.000Z');
  assert.equal(r.rows[1].fields[0].value.iso, '1900-02-28T00:00:00.000Z');
  assert.equal(r.rows[2].fields[0].status, 'unavailable');
  assert.equal(r.rows[2].fields[0].reason, 'DATE_SERIAL_60');
});
