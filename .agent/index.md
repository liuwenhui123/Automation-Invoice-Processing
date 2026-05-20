# Release Memory Index

## Scope

本文件是 `Automation-Invoice-Processinc/` 发布子仓库的本地记忆索引。

## Local Entries

| Entry | Purpose |
| --- | --- |
| `agent.md` | 发布子仓库入口、最近修复和发布注意事项 |
| `README.md` | 用户说明和运行方式 |
| `build_exe.ps1` | Windows 单文件 exe 打包入口 |
| `tests/test_ui_improved.py` | UI 复选框状态回归测试 |

## Routing

- GitHub 仓库：`liuwenhui123/Automation-Invoice-Processing`
- 主入口脚本：`invoice_renamer.py`
- UI 核心实现：`invoice_app/ui_improved.py`
- 发布产物：`dist/invoice_renamer.exe`

## Recent Notes

- 2026-05-20：如果 UI 改动涉及逐行类别复选框，必须确认 `refresh_preview()` 不会替换 `self.records`，否则执行链路会读取到与 UI 行控件脱钩的对象。
- 2026-05-20：GitHub Release 采用 immutable 规则时，先上传资产后发布；若误占用了版本号，优先递增补丁版本重新发布，不要依赖删除旧 release 后复用同一 tag。
