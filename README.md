# Excel2Object

[![install from nuget](http://img.shields.io/nuget/v/Chsword.Excel2Object.svg?style=flat-square)](https://www.nuget.org/packages/Chsword.Excel2Object)
[![release](https://img.shields.io/github/release/chsword/Excel2Object.svg?style=flat-square)](https://github.com/chsword/Excel2Object/releases)
[![.NET CI](https://github.com/chsword/Excel2Object/actions/workflows/dotnet-ci.yml/badge.svg)](https://github.com/chsword/Excel2Object/actions/workflows/dotnet-ci.yml)
[![CodeFactor](https://www.codefactor.io/repository/github/chsword/excel2object/badge)](https://www.codefactor.io/repository/github/chsword/excel2object)

Excel 与 .NET 对象互相转换 / Excel convert to .NET Object and vice versa.

[English](README_EN.md) | 中文

- [Top](#excel2object)
    - [安装 NuGet](#安装-nuget)
    - [发布说明和路线图](#发布说明和路线图)
    - [示例代码](#示例代码)
    - [文档](#文档)
    - [开发和发布](#开发和发布)
    - [贡献者](#贡献者)
    - [参考](#参考)

## 平台支持

[![.NET 4.7.2 +](https://img.shields.io/badge/-4.7.2%2B-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET Standard 2.0](https://img.shields.io/badge/-standard2.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET Standard 2.1](https://img.shields.io/badge/-standard2.1-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET 6.0](https://img.shields.io/badge/-6.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET 8.0](https://img.shields.io/badge/-8.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)

## 安装 NuGet

``` powershell
PM> Install-Package Chsword.Excel2Object
```

或使用 .NET CLI:
``` bash
dotnet add package Chsword.Excel2Object
```

### 命令行工具 excel2obj

``` bash
dotnet tool install -g Chsword.Excel2Object.Cli

excel2obj convert orders.xlsx --output orders.json --typed   # Excel -> JSON
excel2obj convert orders.json --output orders.xlsx           # JSON -> Excel
excel2obj generate-model orders.xlsx --class Order           # 由表头生成带 [ExcelTitle] 的 C# 模型类
```

详见 [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md)。

## 发布说明和路线图

### 暂不支持的特性

- [x] CLI 工具 ✅ **v2.2.1 新增** - `dotnet tool install -g Chsword.Excel2Object.Cli`，查看 [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md)
- [x] 支持自动列宽 ✅ **v2.0.4 新增**
- [x] 冻结首行与自动筛选 ✅ **v2.5.0 新增**
- [x] 数据验证（下拉列表）✅ **v2.5.0 新增**
- [x] 条件格式 ✅ **v2.6.0 新增**
- [x] 样式表（CSS 风格）：背景色、边框、隔行底色、表头样式 ✅ **v2.7.0 新增**
- [x] 合并单元格：按列合并连续相同值、指定区域 ✅ **v2.9.1 新增**
- [x] 大文件流式导出：内存不随行数增长 ✅ **v2.10.0 新增**
- [x] 大文件流式导入：逐行读出，不把整个工作簿建进内存 ✅ **v2.11.0 新增** - 查看 [docs/versions/v2.11.0.md](docs/versions/v2.11.0.md) - 查看 [docs/versions/v2.10.0.md](docs/versions/v2.10.0.md)
- [x] 支持 Excel 日期/日期时间/时间格式 ✅ **v2.0.4 新增**，导出为真正的日期单元格 ✅ **v2.4.0 新增** - 查看 [DateTimeFormats.md](DateTimeFormats.md)
- [x] 公式列引用同一工作簿的其他 sheet ✅ **v2.1.0 新增** - 查看 [ExcelFunctions.md](ExcelFunctions.md)
- [x] 公式内置函数库 ✅ **v2.3.0 新增** - 334 个 Excel 函数，10 个类别 - 查看 [ExcelFunctions.md](ExcelFunctions.md)

### 发布说明

* **2026.09.16** - v2.11.0
- [x] ✨ **新增:** 流式导入：`ExcelHelper.ExcelStreamToObject<Order>(stream)` 逐行读出 `.xlsx` 的一张工作表，不把整个工作簿建进内存。实测二十万行五列：整份读入峰值 1265 MB / 8.4 秒，流式导入 194 MB / 2.9 秒，两者读出的数据逐字段一致。取到的序列是惰性的，`Take` 一类的操作真的能少读。三处不同：公式格读的是文件中存着的上一次计算结果而非当场求值（本库导出的文件里公式没有这个结果，读作空白）、只有 `.xlsx` 能逐行读出（`.xls` 照旧整份读入）、工作表在一开始就定位 - 查看 [docs/versions/v2.11.0.md](docs/versions/v2.11.0.md)
- [x] 🔧 导入的类型转换不再依赖 NPOI 的单元格：整份读入与逐行读出共用同一套转换，两条路的行为不会各自漂移。字典形式的导入随之改为惰性给出；标题行中同名的标题以最左一列为准，非文本的标题不再中断导入


* **2026.09.16** - v2.10.0
- [x] ✨ **新增:** 流式导出：`ExcelHelper.ObjectToExcelStream(data, stream, options => ...)` 逐行取数据并写出，内存中只保留 `StreamingRowWindow` 指定的若干行（默认 100），占用不再随行数增长（省下的是内存而非等待时间：写过的行落到临时文件，最终的包在数据取完后一次写入调用方的流）。实测二十万行五列：内存导出峰值 710 MB / 6.4 秒，流式导出 89 MB / 3.7 秒。冻结、筛选、下拉、条件格式、样式表、公式列与合并单元格在流式导出下同样生效；仅 `.xlsx` 能够流式写入，`.xls` 的格式决定了必须先在内存中建好。流式写入基于 NPOI 的 SXSSF，后者刷行时要用 SkiaSharp 测量字符宽度，而 NPOI 把该依赖标为不随包传递，故应用需自行引用 `SkiaSharp`（Linux 上另需 `SkiaSharp.NativeAssets.Linux.NoDependencies`），缺失时会在写出任何内容之前报错说明 - 查看 [docs/versions/v2.10.0.md](docs/versions/v2.10.0.md)
- [x] 🔧 自动列宽与按值合并改在写入过程中算出，不再回看数据：流式导出得以成立，内存导出亦少走一遍数据。开启 `AutoColumnWidth` 时不再把数据源遍历两遍，只能遍历一次的序列因而也可用


* **2026.09.15** - v2.9.1
- [x] ✨ **新增:** 合并单元格：`options.MergeRepeatedColumns.Add("省份")` 将该列中连续相同的值并成一格（仅相邻且相等的行参与，空单元格不参与，各列彼此独立判断），`options.MergedRegions.Add(new MergedRegion("省份", "城市"))` 另行指定区域（列写标题、行写数据行序号，已知布局时也可直接写 `"A1:C1"`）。被并入的单元格仍保留各自的值，导回对象时每一行的数据依然完整；区域重叠会在导出时报错并指出与哪一个重叠 - 查看 [docs/versions/v2.9.1.md](docs/versions/v2.9.1.md)
- [x] 🔧 空引用检查由 `annotations` 改为 `enable`，并将 nullable 警告视为错误。改动过程中修正数处：公式列缺少标题时不再静默导出无名列，而是在加入 `options.FormulaColumns` 时抛出异常；内部列模型的 `Title` 与 `Type` 定为非空；取值为 `null` 的属性不再写入行字典（行为不变）


* **2026.09.13** - v2.8.1
- [x] 🐛 修复 .NET Framework 上样式互相覆盖的问题：`string.Join(string, params object[])` 在该平台首个元素为 `null` 时返回空字符串，而样式与字体的缓存键均以字体色打头，未设字体色的样式因此共用同一个单元格样式。自 v2.7.0 起，隔行底色、边框、加粗、下划线以及 `[ExcelColumn]` 的单元格样式在 .NET Framework 上均不生效；.NET（Core）上的行为一直正确 - 查看 [docs/versions/v2.8.1.md](docs/versions/v2.8.1.md)
- [x] 🐛 修复 .NET Framework 上导入数值丢失精度的问题：该平台默认数值格式为 `G15`，`"R"` 亦对少数值无法往返，日期序列号读入 `double` 属性时与原值不符。现渲染后验证能否往返，必要时退回 17 位有效数字
- [x] 🔧 测试改为多目标框架运行：`net10.0` 与 `net8.0` 在各平台运行，`net472` 在 Windows 上运行，上述两处缺陷即由此查出

* **2026.09.13** - v2.8.0
- [x] 🐛 修复导入过程中读取失败的处理方式：`ExcelImporter` 此前在三处捕获异常后写入标准输出并继续，调用方只能得到空值而无从得知原因。现移除全部标准输出，改由 `ExcelImporterOptions.OnCellError` 回调上报（提供工作表名称、行列序号、单元格地址与原始异常）。读取失败（如日期序列号超出 Excel 日历）时该单元格取默认值并继续导入，与既有版本一致，在回调中抛出异常即可使其中止；值无法转换为目标类型（如文本 `abc` 写入 `int` 属性）时上报之后照旧抛出并中止导入，亦与既有版本一致。文本不是日期、取值不在枚举之中这两类此前既不报错也无提示，现已纳入上报
- [x] ✨ **改进:** 公式求值器由逐单元格创建改为按工作簿创建一次并全程复用，含公式的表格导入不再重复构造求值器 - 查看 [docs/versions/v2.8.0.md](docs/versions/v2.8.0.md)

* **2026.09.13** - v2.7.1
- [x] 🐛 修复样式表中 `Format` 的作用范围：为某一列声明的格式（`Column("标题")`）现优先于特性上的 `Format`（此前 `Cells` 上声明的格式反而会覆盖特性，致使 `[ExcelColumn(Format = "yyyy-MM-dd HH:mm:ss")]` 的时分秒丢失）；在 `Cells` 与奇偶行上声明的格式仅作用于类型相符的列——日期格式不再使 `12.5` 显示为 `1900-01-12`，数字格式亦不再覆盖文本列的 `@`（前导零保护）
- [x] ✨ **改进:** 边框宽度支持 CSS 关键字 `thin` / `medium` / `thick`；写法无法识别时，异常信息将指出未能识别的片段
- [x] ✨ **改进:** 样式改为按「列 × 行奇偶」解析一次并在整表复用，不再逐单元格解析；大表导出可减少数百万次样式合并与缓存键构造

* **2026.09.13** - v2.7.0
- [x] ✨ **新增:** 样式表 `options.Styles`：样式按作用范围声明，无需在每个 `[ExcelColumn]` 上重复书写 —— `Header` / `Cells` / `Column("标题")` / `OddRows` / `EvenRows`，层叠顺序为 `Cells → 奇偶行 → [ExcelColumn] → Column`，仅显式设置的属性参与叠加
- [x] ✨ **新增:** 单元格背景色与边框（此前不支持）：`Background("#4472C4")`、`Border("1px solid #D0D0D0")`，以及 `Wrap`、`VerticalAlign`、`FontSize` 等
- [x] ✨ **新增:** 颜色支持十六进制（`#RRGGBB` / `#RGB`），不再局限于 56 色枚举。`.xlsx` 直接保存所写的颜色；`.xls` 按 CIELAB 感知距离选取调色板中最接近的一色（浅灰匹配为灰而非淡紫）；以 `ExcelStyleColor` 指定的颜色在两种格式下仍按调色板索引写入
- [x] ✨ **改进:** 在 `Column("标题")` 上声明的样式适用于任何列类型，数值列由此可设置 `Format("#,##0.00")`（特性上的 `Format` 仍仅作用于日期列，以保持兼容）
- [x] ✨ **改进:** 外观相同的单元格共用同一个单元格样式，隔行底色等场景不会突破 `.xls` 的 4000 个样式上限

* **2026.09.12** - v2.6.0
- [x] ✨ **新增:** 条件格式 `options.ConditionalFormats`：以列为单位声明规则（`Operator` 与 `Value`，`Between` 另需 `Value2`，亦可直接给出 `Formula`），命中时改变字体颜色、加粗、倾斜与填充背景色；`WholeRow = true` 可高亮整行。规则由 Excel 求值，数据被编辑后颜色随之变化；同一列可声明多条规则，`.xls` 与 `.xlsx` 均受支持

* **2026.09.12** - v2.5.0
- [x] ✨ **新增:** `ExcelExporterOptions.FreezeHeader` 冻结首行，滚动时表头保持可见；`ExcelExporterOptions.AutoFilter` 为表头附加筛选下拉，范围覆盖本次写入的数据行。两项默认关闭，`.xls` 与 `.xlsx` 均受支持
- [x] ✨ **新增:** `AppendObjectToExcelBytes` 新增接受 options 的重载，追加工作表时同样可设置冻结、筛选与下拉
- [x] ✨ **新增:** 数据验证（下拉列表）：固定取值声明于 `[ExcelColumn("状态", Dropdown = new[] {"启用", "停用"})]`，运行时方可确定的取值通过 `options.Dropdowns["列标题"] = values` 传入（优先于特性）。Excel 将拒绝列表之外的输入；列表总长超过 255 字符，或取值含有逗号、引号时，改由隐藏工作表承载（此前该类列表在 `.xls` 上会直接抛出异常）；无数据的导出同样在首行附带下拉，便于作为填写模板

* **2026.09.11** - v2.4.0
- [x] ✨ **新增:** 导出时 `DateTime` 与 `DateTime?` 列写为日期单元格（序列号与日期格式），可在 Excel 中排序、筛选并参与计算，`null` 写为空白单元格；`Dictionary<string, object>` 与 `DataTable` 导出中的 `DateTime` 值同样处理 - 查看 [DateTimeFormats.md](DateTimeFormats.md)
- [x] ✨ **新增:** `[ExcelColumn(Format = ...)]` 的 .NET 日期格式串自动翻译成对应的 Excel 数字格式（`yyyy-MM-dd HH:mm:ss` → `yyyy-mm-dd hh:mm:ss`，`yyyy年MM月dd日` → `yyyy"年"mm"月"dd"日"`），未指定时为 `yyyy-mm-dd hh:mm:ss`；已经是 Excel 拼写的格式（`m/d/yy`、`[$-409]d-mmm-yy`、`yyyy/m/d;@`）原样使用
- [x] ✨ **新增:** `ExcelExporterOptions.DateTimeAsText`，恢复 v2.3 及之前按文本导出日期的方式；`DataTable` 导出新增接受 options 的重载 `ObjectToExcelBytes(DataTable, Action<ExcelExporterOptions>)`
- [x] 🐛 修复公式列接管带自定义 `Format` 的 `DateTime` 列时公式被丢弃、改写为模型自身值的文本的问题：现保留公式并套用日期格式（公式若返回数值而非日期，请以 `FormulaResultType` 声明）
- [x] 🐛 修复日期单元格读入 `string` 属性或 `Dictionary<string, object>` 时得到 `46276.6` 一类序列号的问题：现按单元格格式渲染为 `yyyy-MM-dd`、`HH:mm:ss` 或两者；`[h]:mm` 一类的经过时间仍按数值读取；数值属性（`double`、`decimal` 等）始终取原始数值
- [x] ⚠️ **行为变更:** Excel 的日历不覆盖 1900 年之前（采用 1904 日期系统的工作簿为 1904 年之前），此类日期（尤以未赋值的 `default(DateTime)` 为常见）仍按 `Format` 渲染为文本写入；`hh` 不带 `tt` 时显示 24 小时制，Excel 没有不带 AM/PM 标记的 12 小时制

* **2026.09.11** - v2.3.0
- [x] ✨ **新增:** 公式内置函数库由 34 个扩充至 **334 个** Excel 函数，分为 10 类：数学与三角（73）、统计（80）、逻辑（11）、查找与引用（33）、日期时间（25）、文本（39）、信息（20）、财务（21）、工程（20）、数据库（12）。其中信息、财务、工程、数据库为新增分类，入口分别为 `ExcelFunctions.Information`、`.Financial`、`.Engineering`、`.Database` - 查看 [ExcelFunctions.md](ExcelFunctions.md)
- [x] 🐛 修复 Excel 2007 之后新增的函数写入后打开显示 `#NAME?` 的问题：xlsx 格式要求此类函数存为 `_xlfn.IFS`、`_xlfn.STDEV.P`（`SORT` 与 `FILTER` 为 `_xlfn._xlws.`），现已自动添加前缀；此前已支持的 `DAYS` 即受此影响
- [x] 🐛 修复公式中的数值按当前区域性格式化的问题：`de-DE` 等区域下 `1.5` 会写为 `1,5`，导致公式无法解析，现统一使用不变区域性
- [x] 🐛 修复文本字面量中的双引号未按 Excel 规则转义（`"` 应写为 `""`）的问题
- [x] ✨ **新增:** 公式中可直接书写更多常规 C# 写法：`%`（MOD）、`^`（乘幂）、`&&` / `||` / `!`（AND / OR / NOT）、三元表达式（IF）、`Math.*`、`string` 成员（`ToUpper`、`Substring`、`Length`、`Contains` 等；语义与 Excel 不完全一致的 `Math.Round`、`string.Trim` 等不翻译，需显式调用对应的 `ExcelFunctions` 函数）、`DateTime` 成员与 `AddDays` / `AddMonths` / `AddYears`，以及可空属性的 `.HasValue`
- [x] ✨ **新增:** 公式中可引用 lambda 捕获的局部变量与 `new DateTime(...)` 等常量表达式，求值后写为字面量（此前会生成无效公式）
- [x] ✨ **改进:** 区间（`c.Matrix(...)`、`c.Columns(...)`）可隐式用于接受单值的参数位置，`SUM` 等函数可混合传入单元格与区间引用
- [x] 🐛 修复 `[ExcelColumn]` 的单元格样式被静默忽略的问题：样式缓存键与分支判断不匹配，`CellBold` / `CellFontColor` / `CellAlignment` 等此前未写入单元格；同时修复某列声明 `CellAlignment` 会改动工作簿默认样式、导致其他列一并变化的问题（#47）
- [x] 🐛 修复 `FormulaResultType = typeof(DateTime)` 的公式列缺少日期格式、在 Excel 中显示为 `46282.6` 一类序列号的问题（#47）
- [x] ⚠️ **不兼容变更:** `IStatisticsFunction.Sum` 由 `params ColumnMatrix[]` 改为 `params ColumnValue[]`，`IMathFunction` 多个方法的返回值由 `int` / `double` 统一为 `ColumnValue`。该变更为源码级兼容（现有公式无需修改即可编译），但二进制不兼容，升级后需重新编译

* **2026.09.11** - v2.2.1
- [x] 🔧 修复命令行工具包体超过 NuGet 250 MB 上限、致使 2.2.0 未能发布的问题：命令行工具仅保留 net8.0 并剔除本机 `.pdb` 文件（库代码无变更）

* **2026.09.11** - v2.2.0
- [x] ✨ **新增:** 命令行工具 `excel2obj`（`dotnet tool install -g Chsword.Excel2Object.Cli`）：Excel ↔ JSON 转换、批量转换、由表头生成带 `[ExcelTitle]` 的模型类 - 查看 [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md)（#9）
- [x] ✨ **新增:** 公式列可通过模型属性引用列：`options.FormulaColumns.Add<Order>("Total", (c, m) => m.Price * m.Qty)`，属性名在编译期检查，引用未导出的属性将抛出异常 - 查看 [ExcelFunctions.md](ExcelFunctions.md)（#22）
- [x] ✨ **改进:** 导出时数值列写为数字单元格，`bool` 列写为布尔单元格，`null` 与 `DBNull` 写为空白单元格（此前一律写为文本）；字符串列仍为文本，前导零不会丢失
- [x] 🐛 修复公式中运算符优先级导致括号缺失的问题：`c["A"] * (c["B"] + c["C"])` 此前生成 `A2*B2+C2`
- [x] ✨ **改进:** 公式中引用当前工作表不存在的列标题时抛出 `Excel2ObjectException`，不再静默回退

* **2026.09.08** - v2.1.0
- [x] ✨ **新增:** 公式列支持跨工作表引用：`c.Sheet("Products")["Price", 2]`、`.Matrix(...)`、`.Columns(...)`，可直接用于 `VLOOKUP` 等函数；另新增当前工作表的整列引用 `c.Columns("A列", "D列")` - 查看 [ExcelFunctions.md](ExcelFunctions.md)
- [x] ✨ **改进:** 公式中引用不存在的列或工作表时抛出 `Excel2ObjectException`（附公式列名），不再生成无效公式
- [x] ✨ **更新:** NPOI 更新至 2.8.0（NPOI 自 2.8 起附带 OSMF EULA，本库已在项目文件中声明接受；源代码仍为 Apache-2.0 许可）
- [x] 🔒 修复传递依赖 System.Security.Cryptography.Xml 的安全漏洞（版本钉至 8.0.4）
- [x] 🧹 移除未使用的 SixLabors.ImageSharp 依赖；补全 `FormulaColumnsCollection.CopyTo`；`ColumnValue.Equals` 与 `GetHashCode` 不再抛出异常
- [x] 🔧 持续集成与测试迁移至 .NET 10；修复发布流程中 GitHub Release 创建失败的问题

* **2025.10.11** - v2.0.5
- [x] 🔧 `release.ps1` 支持自动递增补丁版本号，并新增 `-Help` 参数（库代码无变更）

* **2025.10.11** - v2.0.4
- [x] ℹ️ 本版本为 2.0.0.211 之后首个发布至 NuGet 的 2.0.x 版本（v2.0.1 至 v2.0.3 未曾发布），包含此期间的全部改动：
- [x] ✨ **新增:** 基于内容的自动列宽
  - 依据内容计算列宽
  - 支持最小与最大宽度限制
  - 正确处理中文及其他 Unicode 字符的显示宽度
  - 可通过 `ExcelExporterOptions` 配置
- [x] ✨ **新增:** 全面的日期时间格式支持（56 种格式）- 查看 [DateTimeFormats.md](DateTimeFormats.md)
  - ISO 8601 格式（含时区与毫秒）
  - 多种日期分隔符（横线、斜线、点号）
  - 12 小时制与 24 小时制时间格式
  - 仅时间的格式
  - 区域写法支持（美国、欧洲等）
  - 向后兼容既有的中文日期格式（年月日）
- [x] ✨ **更新:** NPOI 更新至 2.7.5
- [x] ✨ **更新:** SixLabors.ImageSharp 更新至 3.1.11（修复安全漏洞）
- [x] 🔧 以 GitHub Actions 取代 AppVeyor；新增由 tag 触发的自动发布流程与 `release.ps1` 一键发布脚本

* **2024.10.21**
- [x] SixLabors.ImageSharp 更新至 2.1.9
- [x] 补充 .NET 8.0 的测试

* **2024.05.10**
- [x] 支持 .NET 8.0、.NET 6.0、.NET Standard 2.1、.NET Standard 2.0 与 .NET Framework 4.7.2
- [x] 清理已弃用的库

* **2023.11.02**
- [x] 支持列标题映射，可在导入时将工作表中的标题动态映射至模型属性 - [Issue39DynamicMappingTitle.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue39DynamicMappingTitle.cs)

* **2023.07.31**
- [x] 支持 `DateTime` 与 `Nullable<DateTime>` 的格式声明，例如 `[ExcelColumn("Title", Format = "yyyy-MM-dd HH:mm:ss")]`

* **2023.03.26**
- [x] 支持列标题中的特殊符号（#37）- [Issue37SpecialCharTest.cs](https://github.com/chsword/Excel2Object/commit/273122275e724367bb6154e03df61702fcec81b3#diff-5f0f5f7558bf7d4207cfa752a4506c4df89d9b491e2501e4862aff0c2276bd61)

* **2023.02.20**
- [x] 支持的平台确定为 .NET Standard 2.0 与 2.1、.NET 6.0、.NET Framework 4.7.2

* **2022.03.19**
- [x] 新增 `ExcelImporterOptions`，可通过 `TitleSkipLine` 跳过标题行之前的若干行 - [Issue32SkipLineImport.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue32SkipLineImport.cs)
- [x] 修复超类属性未被识别的问题 - [Issue31SuperClass.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue31SuperClass.cs)

* **2021.11.4**
- [x] 支持多工作表导出 - [Pr28MultipleSheetTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr28MultipleSheetTest.cs)

* **2021.10.23**
- [x] 修复可空 `DateTime` 的缺陷（由 @SunBrook 贡献）

* **2021.10.22**
- [x] 支持可空类型 - [Pr24NullableTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr24NullableTest.cs)（由 @SunBrook 贡献）

* **2021.5.28**
- [x] 支持表头与单元格样式，新增 `[ExcelColumnAttribute]`
- [x] 支持公式 - [ExcelFunctions.md](./ExcelFunctions.md)

```C#
var list = new List<Pr20Model>
{
        new Pr20Model
        {
            Fullname = "AAA", Mobile = "123456798123"
        },
        new Pr20Model
        {
            Fullname = "BBB", Mobile = "234"
        }
};
var bytes = ExcelHelper.ObjectToExcelBytes(list, ExcelType.Xlsx);
// model
[ExcelTitle("SheetX")]
public class Pr20Model
{
    [ExcelColumn("Full name", CellFontColor = ExcelStyleColor.Red)]
    public string Fullname { get; set; }

    [ExcelColumn("Phone Number",
        HeaderFontFamily = "Normal",
        HeaderBold = true,
        HeaderFontHeight = 30,
        HeaderItalic = true,
        HeaderFontColor = ExcelStyleColor.Blue,
        HeaderUnderline = true,
        HeaderAlignment = HorizontalAlignment.Right,
        //cell
        CellAlignment = HorizontalAlignment.Justify
    )]
    public string Mobile { get; set; }
}
```

* **v2.0.0.113**
- [x] 项目迁移至 .NET Standard 2.0 与 .NET Framework 4.5.2
- [x] 修复 #12 与 #13

* **v1.0.0.80**
- [x] 支持简单公式
- [x] 支持标准的 Excel 模型
  - [x] Excel 与 JSON 互转
  - [x] Excel 与 `Dictionary<string, object>` 互转
- [x] 支持将 `Uri` 写为超链接单元格，并支持将文本单元格读为 `Uri` 类型

* **v1.0.0.43**
- [x] 支持 xlsx 格式（由 Soar360 贡献）
- [x] 支持复合布尔类型

* **v1.0.0.36**
- [x] 新增 `ExcelToObject<T>(bytes)`


## 示例代码

### 定义模型

``` csharp
public class ReportModel
{
    [Excel("My Title", Order=1)]
    public string Title { get; set; }
    
    [Excel("User Name", Order=2)]
    public string Name { get; set; }
}
```

### 创建模型列表

``` csharp
var models = new List<ReportModel>
{
    new ReportModel{Name="a", Title="b"},
    new ReportModel{Name="c", Title="d"},
    new ReportModel{Name="f", Title="e"}
};
```

### 对象转 Excel 文件

``` csharp
var exporter = new ExcelExporter();
var bytes = exporter.ObjectToExcelBytes(models);
File.WriteAllBytes("C:\\demo.xls", bytes);
```

### Excel 文件转对象

``` csharp
var importer = new ExcelImporter();
IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>("c:\\demo.xls");

// 也可以直接使用字节数组
// IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>(bytes);
```

### 自动列宽（新特性）

``` csharp
// 启用自动列宽调整
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    options.ExcelType = ExcelType.Xlsx;
    options.AutoColumnWidth = true;        // 启用自动宽度
    options.MinColumnWidth = 8;            // 最小宽度（字符）
    options.MaxColumnWidth = 50;           // 最大宽度（字符）
    options.DefaultColumnWidth = 16;       // 禁用自动时的默认宽度
});
```

### 冻结首行与自动筛选

``` csharp
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    options.ExcelType = ExcelType.Xlsx;
    options.FreezeHeader = true;           // 滚动时表头始终可见
    options.AutoFilter = true;             // 表头带上筛选下拉
});
```

两个选项默认关闭，`.xls` 与 `.xlsx` 都支持。筛选范围覆盖表头及本次写入的数据行；`AppendObjectToExcelBytes` 追加的 sheet 各自独立，不影响已有 sheet。

### 数据验证（下拉列表）

``` csharp
public class OrderModel
{
    [ExcelTitle("订单号")] public string No { get; set; }

    // 固定取值写在特性上
    [ExcelColumn("状态", Dropdown = new[] {"启用", "停用", "待审"})]
    public string Status { get; set; }

    [ExcelTitle("城市")] public string City { get; set; }
}

var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    // 运行时才知道的取值走 options，优先于特性
    options.Dropdowns["城市"] = cities.Select(c => c.Name).ToArray();
});
```

Excel 会把取值显示为下拉，并拒绝其他输入。追加导出用 `ExcelHelper.AppendObjectToExcelBytes(bytes, models, options => ...)` 同样可以设置这些选项。列表总长超过 255 字符、或取值里含逗号/引号时（Excel 行内列表放不下），自动改写到一张隐藏 sheet 上、由一个定义名称指向该区域，下拉再引用这个名称（`.xls` 无法让数据验证直接跨 sheet 引用），用法不变；`.xls` 与 `.xlsx` 都支持。导出空列表时下拉仍会挂在第一行，方便做填写模板。

### 样式表（CSS 风格）

样式不必再抄在每个 `[ExcelColumn]` 上，可以按"作用于谁"来写：

``` csharp
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    options.Styles
        .Header(s => s.Bold().Background("#4472C4").Color("#FFF").Center())
        .Cells(s => s.Border("1px solid #D0D0D0"))
        .EvenRows(s => s.Background("#F2F2F2"))              // 隔行底色
        .Column("金额", s => s.Format("#,##0.00").Right())
        .Column("备注", s => s.Wrap().VerticalAlign(ExcelVerticalAlignment.Middle));
});
```

可写的属性：`Color` / `Background`（`#RRGGBB`、`#RGB` 或 `ExcelStyleColor`）、`Bold` / `Italic` / `Underline` / `Strikeout`、`FontFamily` / `FontSize`、`Left` / `Center` / `Right` / `Align` / `VerticalAlign`、`Wrap`、`Format`、`Border` 及 `BorderTop` / `BorderRight` / `BorderBottom` / `BorderLeft`（写法同 CSS：`1px solid #D0D0D0`、`medium dashed red`、`none`；宽度可用像素或 `thin` / `medium` / `thick`，颜色可用十六进制或 `red`、`gray`、`navy` 等 CSS 基本色名）。

**层叠顺序**（从宽到窄，只有显式设置的属性参与，所以各层是叠加而不是互相覆盖）：

    Cells → OddRows / EvenRows → [ExcelColumn] 特性 → Column("标题")

`Header(...)` 位于特性的 `Header*` 属性之下，规则相同；写了表头样式后，整行表头字号一致，不再受"带 `[ExcelColumn]` 的表头默认 10pt"这条历史行为影响。

`Format` 按写在哪一层决定作用范围：写在 `Column("标题")` 上是对这一列的明确意图，照单执行（数字、文本、日期列都生效，也会盖过特性上的 `Format`）；写在 `Cells` / 奇偶行这类大范围上则是兜底默认，只落到说得上话的列——数字格式不会把日期变成 `46278.00`，日期格式不会把 `12.5` 变成 `1900-01-12`，也不会顶掉文本列用于保住前导零的 `@`。

**颜色**：`.xlsx` 原样保存十六进制颜色；`.xls` 只有 56 色调色板，会按人眼感知（CIELAB）挑最接近的一个——所以浅灰得到的是灰而不是淡紫，但非常浅的颜色（如 `#F2F2F2`）会落到白色，`.xls` 下想要隔行效果建议用深一点的灰（如 `#C0C0C0`）。用 `ExcelStyleColor` 指定的颜色在两种格式下都仍按调色板索引写入。

外观相同的单元格共用同一个 cell style（`.xls` 上限 4000 个），隔行底色不会因为行数多而撑爆样式表。

### 合并单元格

``` csharp
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    // 该列中连续相同的值并成一格：分组报表里同一省份的若干行只显示一次省名
    options.MergeRepeatedColumns.Add("省份");
    options.MergeRepeatedColumns.Add("城市");

    // 上面这条规则覆盖不到的情形，另行指定区域：列写标题、行写数据行序号
    options.MergedRegions.Add(new MergedRegion("省份", "城市"));                     // 表头行，跨两列
    options.MergedRegions.Add(new MergedRegion("备注") {FirstRow = 1, LastRow = 3}); // 前三行数据
});
```

列以标题指定、行以数据行序号（自 1 计起）指定——列在表中的位置由 `Order`、特性顺序与公式列的插入位置决定，行数取决于数据条数，编写导出代码时都无从得知，故不必去数 A、B、C；确已知道布局时仍可直接写 `options.MergedRegions.Add("A1:C1")`。

仅相邻且相等的行参与合并，空单元格不参与，各列彼此独立判断；数值直接比较其值而非文本。被并入的单元格仍保留各自的值，Excel 只显示左上角那一个，因此本库写出的文件**导回对象时每一行的数据依然完整**（该文件若经 Excel 编辑并另存，这些值是否保留取决于 Excel 自身的处理，不宜依赖）。区域重叠会在导出时报错并指出与哪一个重叠。

公式列不参与按值合并：单元格里写的是公式，结果由 Excel 打开时才算出，导出时无从比较，故明确拒绝。

需要注意：Excel 中含合并单元格的区域无法排序，若同时启用 `AutoFilter`，筛选可用而排序会被 Excel 拒绝。

### 条件格式

``` csharp
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    // 金额大于 1 万的单元格标红加粗
    options.ConditionalFormats.Add(new ConditionalFormat("金额")
    {
        Operator = ConditionalOperator.GreaterThan,
        Value = 10000,
        FontColor = ExcelStyleColor.Red,
        Bold = true
    });

    // 整行高亮：状态为"急件"时整行黄底
    options.ConditionalFormats.Add(new ConditionalFormat("状态")
    {
        Operator = ConditionalOperator.Equal,
        Value = "急件",
        WholeRow = true,
        BackgroundColor = ExcelStyleColor.Yellow
    });

    // 也可以直接写 Excel 条件，按第一行数据（第 2 行）书写，列用 $ 锁定
    options.ConditionalFormats.Add(new ConditionalFormat("金额")
    {
        Formula = "$B2>$C2",
        BackgroundColor = ExcelStyleColor.LightGreen
    });
});
```

规则由 Excel 自己求值，数据被编辑后颜色会跟着变。`Value` 按类型写成 Excel 字面量（数字原样、字符串加引号、`DateTime` 写成 `DATE(y,m,d)`），`Between` / `NotBetween` 用 `Value2` 给出另一端。同一列可以挂多条规则，`.xls` 与 `.xlsx` 都支持。

### 在 ASP.NET MVC 中使用

在 ASP.NET MVC 模型中，`DisplayAttribute` 可以像 `ExcelTitleAttribute` 一样被支持。

## 文档

- [docs/](docs/README.md) - 文档索引：**各版本特性说明**与专题文档
- [ExcelFunctions.md](ExcelFunctions.md) - 公式列：内置函数、按列标题/模型属性引用、跨 sheet 引用
- [DateTimeFormats.md](DateTimeFormats.md) - 支持的日期时间格式与解析规则
- [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md) - 命令行工具 `excel2obj` 用法
- [ROADMAP.md](ROADMAP.md) - 路线图

更多信息请访问：http://www.cnblogs.com/chsword/p/excel2object.html

## 开发和发布

- **[自动化发布脚本 `release.ps1`](release.ps1)** - 一键完成版本更新、提交和 Tag 创建的 PowerShell 脚本
- **[自动化发布系统说明](RELEASE_AUTOMATION.md)** - 包含 Copilot 指令、版本管理和自动发布流程的完整说明
- **[版本管理规范](.github/VERSIONING.md)** - 语义化版本规范和版本号递增规则
- **[发布流程指南](.github/RELEASE_GUIDE.md)** - 详细的发布步骤和故障排查指南
- **[Copilot 使用说明](.github/copilot-instructions.md)** - GitHub Copilot 开发指导和项目规范

### 快速发布新版本

使用自动化脚本一键发布：

```powershell
# Windows：不带参数则自动递增修订号，也可用 -Version 2.3.0 指定
.\release.ps1

# Linux/macOS (需要 PowerShell Core)
pwsh ./release.ps1
```

详细说明请参阅 [RELEASE_AUTOMATION.md](RELEASE_AUTOMATION.md)。

## 贡献者

[![Contributors](https://contrib.rocks/image?repo=chsword/Excel2Object)](https://github.com/chsword/Excel2Object/graphs/contributors)

## 参考

- https://github.com/tonyqus/npoi
- https://github.com/chsword/ctrc

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件。
