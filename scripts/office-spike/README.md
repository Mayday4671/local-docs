# Word 编辑的隔离验证

本目录是候选 SDK 的接入样本，不属于应用运行代码或 EXE 内容。`superdoc@2.18.0` 与其 `@superdoc/docx-engine@0.17.0` 依赖仅装入忽略提交的测试目录。其许可与本项目的其他依赖不同，正式分发前需单独确认；本样本不代表已选定正式编辑方案。

在项目根目录运行，使用 Node.js 24 与带 `python-docx`、Pillow、lxml 的 Python：

```powershell
New-Item -ItemType Directory -Force test-results/office-spike
Copy-Item scripts/office-spike/* test-results/office-spike/
npm ci --prefix test-results/office-spike
python test-results/office-spike/fixture.py
node node_modules/vite/bin/vite.js build test-results/office-spike --base=./
node test-results/office-spike/run.mjs
python test-results/office-spike/check.py
```

样本自行生成两页中文 DOCX，包含表格、粗体、图片、页眉页脚及显式分页。关闭 SDK 遥测，测试 Electron 在加载页面前屏蔽并记录所有 HTTP/HTTPS 请求。通过公开的查找替换接口修改一处文字，导出实际 DOCX，重新加载，再用独立的 `python-docx` / XML 读取检查内容与未编辑部分。

结果写入测试目录的 `result.json`、`roundtrip-checks.json`，截图为 `opened.png` / `reopened.png`。保留输入文件，不改写原件。页眉 / 页脚 XML 字节有变化，单独记录；本次检查其文字、样式、对齐和运行格式保持一致，不将语义一致等同于整包字节一致。

此验证没有覆盖中文输入法、全部键盘编辑、复杂修订 / 域 / 嵌入对象、长文档、打印或跨机器字体一致性。正式应用仍没有 Word 保存接口。
