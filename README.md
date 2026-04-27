# 发票自动化处理工具

一个面向 Windows 的 PDF 发票批量处理工具，用于扫描发票文件、解析关键字段、人工确认分类、按类别归档，并导出 Excel 汇总表。当前默认启动的是改进版 Tkinter 图形界面，打包产物为 `dist\invoice_renamer.exe`。

## 当前版本状态

- 默认界面入口：`invoice_app/ui_improved.py`
- 备用基础界面：`invoice_app/ui.py`
- 默认扫描目录：源码运行时为项目根目录；EXE 运行时为 EXE 所在目录
- 默认 Excel 输出：`发票信息汇总.xlsx`
- 默认分类：`个人垫付`、`差旅`、`对公转账`
- EXE 图标：使用多尺寸 `.ico` 作为程序资源，运行时使用 `.png` 改善窗口和任务栏高 DPI 显示

## 快速开始

### 直接运行 EXE

双击运行：

```text
dist\invoice_renamer.exe
```

程序会扫描 EXE 所在目录中的 PDF 文件。需要处理其他目录时，可在界面中点击“添加文件”或“添加目录”。

### 从源码运行

安装依赖：

```powershell
python -m pip install pypdf openpyxl
```

启动图形界面：

```powershell
python invoice_renamer.py
```

或显式启动 UI：

```powershell
python invoice_renamer.py --ui
```

也可以使用：

```bat
start.bat
```

说明：`--cli` 参数目前只用于强制不启动图形界面，当前版本没有独立的命令行批处理流程。

## 使用流程

1. 启动程序后，工具会自动解析默认工作目录中的 PDF 文件。
2. 可通过“添加文件”“添加目录”加入更多发票来源。
3. 如需扫描子目录，勾选“递归扫描子目录”。
4. 在发票列表中为每张发票勾选一个或多个类别。
5. 可通过“管理类别”新增、重命名或删除分类。
6. 右侧“分类汇总”会根据当前分类实时统计票数和金额。
7. 在“Excel 输出”中确认或修改汇总表路径。
8. 点击“开始执行”后，程序会移动或复制 PDF，并生成 Excel 汇总文件。

## 处理逻辑

### PDF 解析

解析逻辑位于 `invoice_app/parser.py`。工具使用 `pypdf` 提取 PDF 文本，并识别以下信息：

- 发票号码
- 开票日期
- 购买方名称与税号
- 销售方名称与税号
- 发票金额、税额、价税合计
- 价税合计大写
- 商品或服务名称
- 销货清单、销售清单等附件类型

非发票类 PDF 会被标记为跳过。附件会尝试继承同一发票号码主发票中的日期、购销方和金额信息。

### 预览与确认

预览逻辑位于 `invoice_app/service.py`。发票记录会被标记为：

- `已就绪`：可执行归档
- `需人工确认`：缺少分类或关键金额等信息
- `已跳过`：非发票类 PDF
- `失败`：解析或执行过程中发生错误

只有状态为 `已就绪` 且确认为发票类的记录会进入最终归档流程。

### 文件命名

归档文件名由以下字段组合生成：

```text
发票号后6位_自定义显示名_销售方名称_价税合计金额[_附件].pdf
```

其中自定义显示名为空时会被省略。文件名和目录名会自动清理 Windows 不允许使用的字符。

### 分类归档

归档规划位于 `invoice_app/archive.py`，执行逻辑位于 `invoice_app/service.py`。

每个分类目录名称包含分类名和该分类汇总金额：

```text
类别名称_汇总金额
```

例如：

```text
差旅_1280.50
```

执行时：

- 未分类但已就绪的发票会在原目录中按规范名称重命名。
- 有一个分类的发票会被移动到对应分类目录。
- 有多个分类的发票会先移动到第一个分类目录，再复制到其他分类目录。
- 同名目标已存在时，会自动追加 `_2`、`_3` 等序号避免覆盖。

### Excel 导出

Excel 导出逻辑位于 `invoice_app/export_excel.py`。默认生成 `发票信息汇总.xlsx`，包含两个工作表：

- `发票明细`：逐张记录发票号码、日期、购销方、金额、税额、总额和物品名称。
- `分类汇总`：按分类统计发票数量和汇总金额。

## 项目结构

```text
Automation-Invoice-Processinc/
├── assets/
│   ├── invoice_processing.ico   # EXE 程序图标，多尺寸 ICO
│   ├── invoice_processing.png   # 运行时窗口和任务栏图标
│   └── invoice_processing.svg   # 可编辑源图标
├── dist/
│   └── invoice_renamer.exe      # 已打包的 Windows 可执行文件
├── invoice_app/
│   ├── __init__.py
│   ├── app_icon.py              # 图标路径解析、AppUserModelID、Tk 图标应用
│   ├── archive.py               # 分类目录和归档目标规划
│   ├── classification.py        # 默认分类和分类汇总
│   ├── export_excel.py          # Excel 明细与分类汇总导出
│   ├── models.py                # 发票记录、分类汇总、归档计划数据结构
│   ├── naming.py                # 金额格式化和路径名称清理
│   ├── parser.py                # PDF 文本提取与字段解析
│   ├── service.py               # 预览、分类编辑、执行归档等业务逻辑
│   ├── ui.py                    # 备用基础 Tkinter 界面
│   └── ui_improved.py           # 当前默认 Tkinter 界面
├── build_exe.ps1                # PyInstaller 打包脚本
├── invoice_renamer.py           # 程序入口
├── README.md
└── start.bat                    # 源码快速启动脚本
```


## 开发提示

- 运行源码时请在项目根目录启动，便于默认工作目录和资源路径一致。
- 修改解析字段时，优先更新 `invoice_app/parser.py`。
- 修改分类、预览或执行行为时，优先更新 `invoice_app/service.py`。
- 修改文件命名规则时，检查 `invoice_app/naming.py` 和 `service.py` 中的 `_build_main_stem()`。
- 修改打包资源或图标时，同时检查 `assets/`、`invoice_app/app_icon.py` 和 `build_exe.ps1`。

