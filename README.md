# 披露工作簿原件工具

用于接入披露工作簿 XLSX 原件。src/source.mjs 提供 ZIP 目录格式检查（保留原行为），src/files.mjs 提供本地样本读取；src/intake.mjs 是工作簿接入主管线，src/cli.mjs 是从 XLSX 文件到结构化记录的可运行入口。格式约定见 docs/数据约定.md。样本是合成的，文件清单保存字节长度和 SHA-256 摘要。

需要 Node.js 22 或更新版本，当前项目只使用内置模块，不需要安装第三方包。

测试命令：

```sh
npm test
```

编译与语法检查命令：

```sh
npm run build
```

原件 ZIP 目录检查示例：

```sh
npm run inspect
```

可在上述命令末尾使用 `-- 文件名` 选择 fixtures 中其他同格式文件。runtime 用于本地运行数据，不纳入版本管理。

## 工作簿接入

```sh
npm run intake -- fixtures/disclosure-1900.xlsx fixtures/mapping.json
```

输出顶层包含：

- `accepted`：false 表示整批拒收（压缩包不完整、关系越界、同名工作表、合并违规、超限额等），此时只有 `identity` 与 `rejection`（含 code/message，必要时含 detail）。
- `identity`：intakeId、原件 sourceSha256/sourceBytes、映射 mappingVersion/mappingDigest。同件文件重送身份不变，换映射身份独立。
- `workbook`：工作表名与部件路径、范围、1900/1904 基准、fullCalcOnLoad、共享字符串数、单元格数。
- `rows`：映射范围内每一行，每个字段输出工作表名、单元格坐标、原始类型、样式索引、公式原文、解释状态（ok/empty/pending/error/unavailable）与原因码、解释后的值。

可选限额参数：`--max-entries N`、`--max-bytes N`、`--max-entry-bytes N`、`--max-cells N`。拒收时进程退出码为 1。

编程接口：

```js
import {readFile} from 'node:fs/promises';
import {intakeWorkbook} from './src/intake.mjs';
const result = intakeWorkbook(await readFile('披露.xlsx'), mapping, {limits: {maxCells: 100000}});
```
