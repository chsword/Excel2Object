# Excel2Object 路线图 / Roadmap

本文档列出了 Excel2Object 项目未来可以开发的功能和改进。

This document lists potential features and improvements for the Excel2Object project.

[English](#english-version) | [中文](#中文版本)

---

## 中文版本

### 未来开发任务清单

#### 1. 支持更多 Excel 日期时间格式
**优先级：高**

- 完善日期、日期时间、时间类型在 Excel 中的导入导出支持
- 支持更多自定义日期时间格式
- 处理不同地区的日期格式差异
- 支持时区转换

#### 2. 开发 CLI 命令行工具 ✅ 已完成（#9）
**优先级：中**

- [x] 独立的 dotnet tool `Chsword.Excel2Object.Cli`（命令 `excel2obj`），见 [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md)
- [x] Excel ↔ JSON 转换，支持多文件批量转换到目录
- [x] 由表头和数据推断类型，生成带 `[ExcelTitle]` 的模型类
- [ ] 支持配置文件定义转换规则

示例命令：
```bash
excel2obj convert input.xlsx --output data.json --typed
excel2obj generate-model input.xlsx --output Model.cs --class Model
```

#### 3. 增强公式支持
**优先级：中**

- 支持更复杂的 Excel 公式
- 扩展内置函数库
- 支持自定义函数
- 支持公式的动态计算

#### 4. 支持 Excel 图表导入导出
**优先级：低**

- 支持图表的读取和生成
- 支持多种图表类型（柱状图、折线图、饼图等）
- 支持图表数据的自动绑定
- 支持图表样式自定义

#### 5. 支持 Excel 数据验证
**优先级：中**

- 支持下拉列表数据验证
- 支持数字范围验证
- 支持日期范围验证
- 支持自定义验证规则
- 导出时保留验证规则

#### 6. 性能优化与大文件支持
**优先级：高**

- 优化大数据量的导入导出性能
- 支持流式读写大文件
- 减少内存占用
- 支持异步操作
- 提供进度回调机制

#### 7. 增强样式和格式支持
**优先级：中**

- 支持更多单元格样式（边框、背景、字体等）
- 支持条件格式
- 支持单元格合并
- 支持行高列宽的更精细控制
- 支持主题和模板

#### 8. 支持复杂对象关系
**优先级：中**

- 支持嵌套对象的导入导出
- 支持集合属性的展开
- 支持主从表关系
- 支持多对多关系的处理

#### 9. 提供可视化配置工具
**优先级：低**

- 开发 Web 或桌面配置界面
- 可视化设计 Excel 模板
- 拖拽方式配置列映射
- 实时预览转换结果
- 保存和管理配置方案

#### 10. 增强错误处理和日志
**优先级：高**

- 提供详细的错误信息和堆栈跟踪
- 支持数据验证错误收集
- 提供结构化日志输出
- 支持自定义错误处理策略
- 提供数据转换报告

---

## English Version

### Future Development Task List

#### 1. Support More Excel Date/Time Formats
**Priority: High**

- Enhance import/export support for date, datetime, and time types in Excel
- Support more custom date/time formats
- Handle regional date format differences
- Support timezone conversion

#### 2. Develop CLI Command-Line Tool ✅ Done (#9)
**Priority: Medium**

- [x] Standalone dotnet tool `Chsword.Excel2Object.Cli` (command `excel2obj`), see [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md)
- [x] Excel ↔ JSON conversion, batch conversion of several files into a directory
- [x] Model class generation with `[ExcelTitle]` and types inferred from the header row and data
- [ ] Configuration files to define conversion rules

Example commands:
```bash
excel2obj convert input.xlsx --output data.json --typed
excel2obj generate-model input.xlsx --output Model.cs --class Model
```

#### 3. Enhanced Formula Support
**Priority: Medium**

- Support more complex Excel formulas
- Extend built-in function library
- Support custom functions
- Support dynamic formula calculation

#### 4. Support Excel Chart Import/Export
**Priority: Low**

- Support reading and generating charts
- Support multiple chart types (column, line, pie, etc.)
- Support automatic chart data binding
- Support chart style customization

#### 5. Support Excel Data Validation
**Priority: Medium**

- Support dropdown list validation
- Support number range validation
- Support date range validation
- Support custom validation rules
- Preserve validation rules on export

#### 6. Performance Optimization and Large File Support
**Priority: High**

- Optimize import/export performance for large datasets
- Support streaming read/write for large files
- Reduce memory footprint
- Support asynchronous operations
- Provide progress callback mechanism

#### 7. Enhanced Style and Format Support
**Priority: Medium**

- Support more cell styles (borders, backgrounds, fonts, etc.)
- Support conditional formatting
- Support cell merging
- Support finer control over row heights and column widths
- Support themes and templates

#### 8. Support Complex Object Relationships
**Priority: Medium**

- Support import/export of nested objects
- Support expansion of collection properties
- Support master-detail table relationships
- Handle many-to-many relationships

#### 9. Provide Visual Configuration Tool
**Priority: Low**

- Develop Web or desktop configuration interface
- Visually design Excel templates
- Drag-and-drop configuration for column mapping
- Real-time preview of conversion results
- Save and manage configuration schemes

#### 10. Enhanced Error Handling and Logging
**Priority: High**

- Provide detailed error messages and stack traces
- Support data validation error collection
- Provide structured log output
- Support custom error handling strategies
- Provide data conversion reports

---

## 贡献指南 / Contributing

如果您想参与上述任务的开发，或有其他建议，欢迎：

If you want to participate in the development of the above tasks or have other suggestions, welcome to:

- 提交 Issue 讨论功能需求 / Submit an Issue to discuss feature requests
- 提交 Pull Request 贡献代码 / Submit a Pull Request to contribute code
- 在 Discussions 中分享想法 / Share ideas in Discussions

感谢您对 Excel2Object 的关注和支持！

Thank you for your attention and support for Excel2Object!
