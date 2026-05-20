# Agent Memory

<!-- usage:agent.invoice.release.entrypoint count=1 since=2026-05-20 last=2026-05-20 -->

最后更新：2026-05-20

## Local Memory System

本 `agent.md` 只维护 `Automation-Invoice-Processinc/` 发布子仓库自己的记忆入口和近期发布注意事项。

| Entry | Marker | Count | Since | Last |
| --- | --- | ---: | --- | --- |
| `agent.md` | `agent.invoice.release.entrypoint` | 1 | 2026-05-20 | 2026-05-20 |
| `.agent/index.md` | `agent.invoice.release.index` | 1 | 2026-05-20 | 2026-05-20 |

下一步入口：`.agent/index.md`。

## Scope

- 本目录是带 `.git` 的发布子仓库，对应 GitHub 仓库 `liuwenhui123/Automation-Invoice-Processing`。
- Windows 发布产物路径：`dist/invoice_renamer.exe`。
- 打包入口：`build_exe.ps1`。

## Recent Changes

- 2026-05-20：修复 `invoice_app/ui_improved.py` 中 `refresh_preview()` 覆盖 `self.records` 的问题。预览刷新只能更新汇总，不能替换 UI 正在编辑的 `InvoiceRecord` 对象列表，否则会出现“只有第一条勾选真正生效、后续行 UI 看似选中但执行不生效”的状态脱钩。
- 2026-05-20：新增 `tests/test_ui_improved.py`，覆盖“预览刷新后仍保持行绑定对象不变”和“第一行勾选后第二行继续勾选仍能生效”两个回归场景。
- 2026-05-20：已重建并发布 `dist/invoice_renamer.exe`，对应 GitHub Release `v2.6.1`。
- 2026-05-20：执行规则调整为“未勾选类别也执行重命名，但不移动/复制到分类文件夹”。只有发票行勾选了具体类别，才进入对应分类目录；未分类但金额可识别的发票状态应为“已就绪”。
- 2026-05-20：右侧分类汇总面板改为 grid 布局，Excel 输出区域固定在底部，状态栏加厚，避免窗口上下缩放时底部输入框被压扁或显示异常。
- 2026-05-20：解析器修复两个 PDF 版式问题：开票日期允许 `2026年 04月 27日` 这类带空格日期并标准化为 `2026年04月27日`；当购买方为“个人”且同一行紧跟公司名、仅出现一个税号时，将“个人”作为购买方，将公司名和税号作为销售方。

## Release Notes

- 现有 tag 序列为：`v1.0`、`v2.0`、`v2.1`、`v2.5`、`v2.6.1`。
- 2026-05-20：GitHub Release 开启 immutable 行为后，某个 tag 一旦被用来创建过已发布 release，就不适合删除重建后复用同名 tag 继续补附件；更稳的流程是直接发布补丁版本号，例如从 `v2.6` 顺延到 `v2.6.1`。
- 发布顺序建议固定为：`git push origin <tag>` -> 创建 draft release -> 上传 exe 资产 -> publish。不要先发布再补传资产。
