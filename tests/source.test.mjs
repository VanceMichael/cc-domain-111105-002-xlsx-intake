import {test} from 'node:test';
import assert from 'node:assert/strict';
import {inspect, listZipEntries} from '../src/source.mjs';
import {readFixture} from '../src/files.mjs';
const sample=await readFixture('disclosure-1900.xlsx');
test('原件结构与错误边界',async()=>{

const entries=listZipEntries(sample);assert.ok(entries.some(e=>e.name==='xl/workbook.xml'));
assert.ok(entries.every(e=>e.expanded>=0));
assert.throws(()=>listZipEntries(sample.subarray(0,sample.length-10)),/ZIP/);
assert.throws(()=>listZipEntries(Buffer.from('bad')),/ZIP/);

});
test('样本读取范围',async()=>{await assert.rejects(readFixture('../package.json'),/越界/);await assert.rejects(readFixture('absent.bin'),/ENOENT/);});
