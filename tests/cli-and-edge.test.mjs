import {test} from 'node:test';
import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {writeFile, mkdtemp, rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {parseXml, decodeXml} from '../src/xml.mjs';
import {intakeWorkbook} from '../src/intake.mjs';
import {buildWorkbook} from './xlsx-builder.mjs';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');

test('CLI：XLSX 到结构化记录的可运行入口，拒收时退出码为 1', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'intake-cli-'));
  try {
    const ok = spawnSync(process.execPath, [
      path.join(root, 'src/cli.mjs'),
      path.join(root, 'fixtures/disclosure-1900.xlsx'),
      path.join(root, 'fixtures/mapping.json'),
    ], {encoding: 'utf8'});
    assert.equal(ok.status, 0);
    const parsed = JSON.parse(ok.stdout);
    assert.equal(parsed.accepted, true);
    assert.equal(parsed.identity.sourceSha256, '6e336a888fd9bb785d21b0128019c1731843842d2aea25da024aa61df5badb38');
    assert.equal(parsed.rows.length, 3);

    const bad = spawnSync(process.execPath, [
      path.join(root, 'src/cli.mjs'),
      path.join(root, 'fixtures/truncated.xlsx.bin'),
      path.join(root, 'fixtures/mapping.json'),
    ], {encoding: 'utf8'});
    assert.equal(bad.status, 1);
    const rejected = JSON.parse(bad.stdout);
    assert.equal(rejected.accepted, false);
    assert.equal(rejected.rejection.code, 'ZIP_INVALID');

    // 限额参数生效
    const limited = spawnSync(process.execPath, [
      path.join(root, 'src/cli.mjs'),
      path.join(root, 'fixtures/disclosure-1900.xlsx'),
      path.join(root, 'fixtures/mapping.json'),
      '--max-entries', '2',
    ], {encoding: 'utf8'});
    assert.equal(limited.status, 1);
    assert.match(JSON.parse(limited.stdout).rejection.message, /部件数量/);

    // 坏映射 JSON
    const badMap = path.join(dir, 'bad.json');
    await writeFile(badMap, '{not json');
    const mapFail = spawnSync(process.execPath, [path.join(root, 'src/cli.mjs'), path.join(root, 'fixtures/disclosure-1900.xlsx'), badMap], {encoding: 'utf8'});
    assert.equal(mapFail.status, 1);
    assert.equal(JSON.parse(mapFail.stdout).rejection.code, 'MAPPING_JSON');
  } finally {
    await rm(dir, {recursive: true, force: true});
  }
});

test('XML 实体与数字字符引用解码', () => {
  assert.equal(decodeXml('a&amp;b&lt;c'), 'a&b<c');
  assert.equal(decodeXml('&#65;&#x42;'), 'AB');
  const attrs = {};
  parseXml('<root a="1" b="x &amp; y"><child>文本 &lt;ok&gt;</child></root>', {
    onOpen(p, a) { if (p[0] === 'root') Object.assign(attrs, a); },
    onText(p, t) { if (p[1] === 'child') attrs.childText = t; },
  });
  assert.equal(attrs.a, '1');
  assert.equal(attrs.b, 'x & y');
  assert.equal(attrs.childText, '文本 <ok>');
  assert.throws(() => parseXml('<root><child></root>'), /不匹配/);
});

test('合并锚点单元格缺失时隔离为字段错误', () => {
  const wb = buildWorkbook({
    sheets: [{name: '披露表', merges: ['A2:A3'], cells: [{ref: 'B2', inline: 'x'}]}],
  });
  const r = intakeWorkbook(wb, {
    version: 'v', sheet: '披露表', range: 'A2:A3',
    columns: {A: {name: 'issuer', kind: 'text', group: true}},
  });
  assert.equal(r.accepted, true);
  const [row2, row3] = r.rows;
  assert.equal(row2.fields[0].status, 'empty');
  assert.equal(row3.fields[0].status, 'error');
  assert.equal(row3.fields[0].reason, 'MERGE_ANCHOR_MISSING');
});

test('没有 sharedStrings/styles 部件的工作簿也能接入', () => {
  const wb = buildWorkbook({
    includeShared: false,
    includeStyles: false,
    sheets: [{name: '披露表', cells: [{ref: 'D2', v: '1.5'}]}],
  });
  const r = intakeWorkbook(wb, {
    version: 'v', sheet: '披露表', range: 'A2:D2',
    columns: {
      A: {name: 'issuer', kind: 'text', group: true},
      D: {name: 'water', kind: 'decimal', scale: '1000'},
    },
  });
  assert.equal(r.accepted, true);
  assert.equal(r.workbook.sharedStringCount, 0);
  assert.equal(r.rows[0].fields.find((f) => f.field === 'water').value.scaled, '1500');
});

test('日期时间分量与文本补零边界', () => {
  const wb = buildWorkbook({
    sheets: [{
      name: '披露表',
      cells: [
        {ref: 'E2', s: 2, v: '45000.5'},
        {ref: 'B2', s: 1, v: '1234567'},
        {ref: 'B3', s: 1, v: '42'},
      ],
    }],
  });
  const r = intakeWorkbook(wb, {
    version: 'v', sheet: '披露表', range: 'B2:E3',
    columns: {
      B: {name: 'reference', kind: 'identifier'},
      E: {name: 'date', kind: 'date'},
    },
  });
  const [row2, row3] = r.rows;
  // 超过补零宽度时保留完整原文，不截断
  assert.equal(row2.fields[0].value.text, '1234567');
  assert.equal(row3.fields[0].value.text, '000042');
  // 半日时间分量
  assert.equal(row2.fields[1].value.iso, '2023-03-15T12:00:00.000Z');
});
