# Office 接入验证

日期：2026-09-27。此记录区分文档调研、已实现代码与实际验证结果。

## 当前决策

0.7.0 将原有逐段 / 逐格表单更换为 Word 页面内改字和 Excel 行列网格，新增空白格输入与合并格显示。保存仍为受限 OOXML 内容补丁，没有接入下文候选引擎，也未实现字体排版、表格结构编辑或公式计算。当前界面与导出验证见 [0.7.0 验收](0.7.0Office页面与网格编辑验收.md)。

保留“单个桌面应用、离线、无需用户配置额外服务”的方向。0.6.0 已提供受限 OOXML 内容编辑与新建：普通段落 / 单元格修改、Word 末尾追加、版本保存及重开，未增加外部服务或新编辑引擎依赖。**完整 Office 排版编辑与公式计算目标仍然未完成。** 当前范围和独立复读结果见 [0.6.0 验收](0.6.0收藏标记与内容编辑验收.md)。下文为此前调研和 0.1–0.4 阶段的历史记录，不能视为候选引擎已接入本版。

## 候选引擎调查

ONLYOFFICE Docs 的网页接入由 DocEditor 创建 iframe，并从已安装的 Document Server 加载编辑器。因此，引入 React 适配组件本身不能解决离线服务打包、进程管理和文件保存回调。其 Desktop Editors 支持离线编辑，但公开的 DMS 接入是在 ONLYOFFICE 桌面应用内打开文档，不能据此认定可直接嵌入本工程的 Electron 页面。[Docs 接入](https://api.onlyoffice.com/docs/docs-api/usage-api/doceditor/) · [Desktop Editors](https://api.onlyoffice.com/docs/desktop-editors/get-started/overview/) · [DMS 文档打开](https://api.onlyoffice.com/docs/desktop-editors/usage-api/adding-a-dms-provider/opening-documents/)

Univer 的当前导入导出文档明确说明 Office 文件转换需要后端，snapshot 路径也需要转换后端，只是不依赖协作文档。因此不能仅凭本地网格显示成功就认定完成离线 XLSX 保真保存。[官方导入导出说明](https://docs.univer.ai/guides/sheets/features/import-export)

SuperDoc 的浏览器编辑器支持导出 DOCX Blob，是 Word 内置编辑的候选方向；其新版 DOCX Engine 使用独立专有许可，和编辑器开源部分不同。初次调研时只查阅公开文档；0.4.0 已补独立样本验证，见文末。商用、分发及费用条件仍未确定，真实复杂文档尚需继续验收。[导出接口](https://docs.superdoc.dev/editor/export-options/) · [DOCX Engine 许可](https://docs.superdoc.dev/resources/docx-engine-license/)

## 本次实现

Word 使用 Apache-2.0 许可的 `docx-preview` 做版式阅读；该库是渲染器，不是编辑器。正文另行提取为有位置标记的段落，用于内容检索、摘要和跳转。版式呈现受 HTML 和库自身能力限制。[项目说明](https://github.com/VolodymyrBaydalka/docxjs)

Excel 从 OOXML 工作簿关系解析工作表，不假设文件名、关系编号或显示序号相同。读取共享字符串、内联字符串、布尔值、数值、公式缓存与合并区域信息；网格保留真实坐标，显示隐藏工作表标识。原始 XLSX 文件不被重写。

所有 Office 文档以原始字节保存在对象库。内容提取只写 SQLite 缓存与搜索文本；失败时保留文件和失败原因，原样导出继续可用。旧版库自动升级到 schema v2。

## 验证证据

`npm test`：14 项测试通过。新增 Office 测试验证段落内不同文字格式的拼接顺序、表格及页眉页脚、稀疏单元格坐标（如 D120、AA120）、非连续工作表关系编号、缓存公式值、无效输入、旧库迁移、重启后的索引和过期结果拒绝。

`npm run test:office`：在实际 Electron 中屏蔽 HTTP/HTTPS，验证 Word 含图片和页眉的版式显示、Word 正文定位、Excel 多工作表与公式缓存、搜索命中摘要、实际工作表和单元格跳转、返回查询保留、损坏文件提示、导出与原件字节一致、原始 Word 文件未改变及重启后的内容搜索。脚本断言没有 HTTP/HTTPS 请求，且没有未捕获的页面异常。

`node scripts/smoke.mjs`：原有 Markdown、保存、版本、回收站及导出流程回归通过。类型检查与生产构建通过。阅读器截图位于忽略提交的 `test-results/office-word.png` 与 `test-results/office-excel.png`。

样本是仓库内程序生成的 OOXML 测试文件，不是用户真实资料，也不是 Word/Excel 全功能兼容性认证。未测试完整 Office 编辑后的格式保真，没有据此承诺复杂文档排版或公式重新计算。

2026-09-27，0.1.0 安装验证记录：已用 `scripts/packaged-smoke.mjs` 直接启动 NSIS 实际安装后的 EXE，通过同一组 Office 测试，验证后台解析线程与阅读器资源在安装包内可以离线运行。Windows 交付详情见《Windows交付说明.md》。

## 下一步验证门槛

正式编辑引擎接入前，需要一个无需单独部署服务的可分发方案，以及包含中文字体、分页、页眉页脚、图片、表格、多个工作表、样式和公式的真实样本集。必须验证保存生成实际文件、重开后内容保留、未编辑部分的保真、失败不会覆盖旧版本，以及应用关闭后无残留服务。阅读器与索引接口可以继续复用，不依赖最终编辑器选型。

## 0.3.0 格式检查补充

修正 Excel 侧栏纯文字摘要、数值 / 日期显示、Office XML 编码换行，以及 Word 宽表在窄侧栏中的可达性。支持 `.md`、`.markdown`、`.docx`、`.xlsx`；大写扩展名也经过导入检查。Office 仍是只读视图，不修改源文件。

新增合成样本覆盖 Markdown 尾部空列、转义竖线、CRLF、UTF-8 BOM、任务列表、删除线、缩进代码与长表；Word 的宽表、合并标题、多行、中文数字字符引用、图片与页眉页脚；Excel 的百分比、日期系统、千分位、补零、错误值、多行文本和格式化公式缓存。详见 [0.3.0 格式与主题验收](0.3.0格式与主题验收.md)。

以上是当前支持格式的代表性回归，不是全部 Office 格式兼容性认证；PDF、图片及旧式 `.doc/.xls` 未纳入支持范围。

## 0.4.0：Word 编辑隔离样本验证

此次实际安装并运行 `superdoc@2.18.0`（依赖 `@superdoc/docx-engine@0.17.0`），使用独立 Electron 测试窗口，没有加入正式依赖和安装包。源码与复现步骤保存在 [scripts/office-spike](../scripts/office-spike/README.md)。依赖获取最初遇到下载等待与 npm 缓存版本解析失败，刷新包元数据后安装、构建及运行成功；这些问题未影响正式应用。

样本由代码生成，包含两页中文、粗体、三行两列表格、图片、页眉、页脚、显式分页。实际完成：加载原 DOCX → 用公开查找替换接口修改正文 → 导出 DOCX Blob → 写入新文件 → 重新加载。SDK 遥测显式关闭，HTTP/HTTPS 在窗口加载前被拦截和记录；本次请求数为零，页面异常为零。

独立读取导出的文件后确认：修改文字已保存，原标记消失，表格文字、图片字节、粗体、页眉页脚文字与样式、未修改的第二页、显式分页和纸张边距保留。`styles.xml` 字节相同；页眉 / 页脚 XML 有重新序列化，字节不同，但本样本的文字、段落样式、对齐和文字格式一致。截图已检查，位于 `docs/screenshots/word-editor-feasibility.png`。

**结论：Word 的无独立服务、离线编辑保存基础流程可行，尚未达到正式接入验收。** 这只是有限合成样本，不是通用 DOCX 格式保真认证。中文输入法、真实用户键盘操作、复杂分页 / 修订 / 域、嵌入对象、长文档、字体差异、保存版本保护和异常中断还需专项验证。

分发条件仍未定：编辑器 npm 包标为 AGPL-3.0，DOCX Engine 是单独的专有许可。官方允许在适用范围内创建由引擎驱动的应用，但不能由一次安装测试推定本项目已经取得所有生产 / 分发权限。因此目前没有将引擎装入交付 EXE，也未作出购买承诺。[官方引擎许可](https://docs.superdoc.dev/resources/docx-engine-license/) · [浏览器编辑示例](https://docs.superdoc.dev/editor/quickstart/)

Excel 本轮仍停留在方案约束核实：Univer 当前官方导入导出转换需要后端，snapshot 路径也不免除转换后端。因此尚未建立符合本项目“EXE 安装后直接离线编辑并保存 XLSX”要求的已验证方案，未将网格编辑或重新生成简化表格当作完成。[官方导入导出说明](https://docs.univer.ai/guides/sheets/features/import-export)
