# 披露工作簿原件工具

用于查看披露工作簿相关数据原件。src/source.mjs 提供当前格式检查，src/files.mjs 提供本地样本读取；格式约定见 docs/数据约定.md。样本是合成的，文件清单保存字节长度和 SHA-256 摘要。

需要 Node.js 22 或更新版本，当前项目只使用内置模块，不需要安装第三方包。

测试命令：

```sh
npm test
```

编译与语法检查命令：

```sh
npm run build
```

原件读取示例：

```sh
npm run inspect
```

可在上述命令末尾使用 `-- 文件名` 选择 fixtures 中其他同格式文件。runtime 用于本地运行数据，不纳入版本管理。
