// 可运行入口：node src/cli.mjs <xlsx 路径> <映射 JSON 路径> [--max-entries N] [--max-bytes N] [--max-cells N]
import {readFile} from 'node:fs/promises';
import {intakeWorkbook, jsonReplacer} from './intake.mjs';
import {IntakeError} from './ooxml.mjs';

function parseArgs(argv) {
  const positional = [];
  const limits = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--max-entries') limits.maxEntries = Number(argv[++i]);
    else if (a === '--max-bytes') limits.maxExpandedBytes = Number(argv[++i]);
    else if (a === '--max-entry-bytes') limits.maxExpandedEntryBytes = Number(argv[++i]);
    else if (a === '--max-cells') limits.maxCells = Number(argv[++i]);
    else if (a.startsWith('--')) throw new IntakeError('CLI_ARG', `未知参数: ${a}`);
    else positional.push(a);
  }
  if (positional.length < 2) {
    throw new IntakeError('CLI_ARG', '用法: cli.mjs <xlsx 路径> <映射 JSON 路径> [限额参数]');
  }
  for (const [k, v] of Object.entries(limits)) {
    if (!Number.isInteger(v) || v <= 0) throw new IntakeError('CLI_ARG', `限额参数 ${k} 必须是正整数`);
  }
  return {xlsxPath: positional[0], mappingPath: positional[1], limits};
}

async function main() {
  const {xlsxPath, mappingPath, limits} = parseArgs(process.argv.slice(2));
  const [bytes, mappingText] = await Promise.all([readFile(xlsxPath), readFile(mappingPath, 'utf8')]);
  let mapping;
  try {
    mapping = JSON.parse(mappingText);
  } catch (err) {
    throw new IntakeError('MAPPING_JSON', `映射文件不是合法 JSON: ${err.message}`);
  }
  const result = intakeWorkbook(bytes, mapping, {limits});
  process.stdout.write(`${JSON.stringify(result, jsonReplacer, 2)}\n`);
  if (!result.accepted) process.exitCode = 1;
}

main().catch((err) => {
  const out = {
    accepted: false,
    rejection: {
      code: err instanceof IntakeError ? err.code : 'UNEXPECTED',
      message: err.message,
    },
  };
  process.stdout.write(`${JSON.stringify(out, null, 2)}\n`);
  process.exit(1);
});
