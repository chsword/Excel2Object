# GitHub Copilot 使用说明

本文档为 Excel2Object 项目提供 GitHub Copilot 的使用指导和项目规范说明。

## 项目简介

Excel2Object 是一个用于 Excel 与 .NET 对象互相转换的类库。该项目支持多个 .NET 框架版本，包括 .NET Framework 4.7.2+、.NET Standard 2.0/2.1、.NET 6.0/8.0/9.0/10.0。

仓库包含三个项目：

| 项目 | 说明 |
|-----|------|
| `Chsword.Excel2Object` | 类库本体，发布为 NuGet 包 `Chsword.Excel2Object` |
| `Chsword.Excel2Object.Cli` | 命令行工具，发布为 dotnet tool `Chsword.Excel2Object.Cli`（命令 `excel2obj`） |
| `Chsword.Excel2Object.Tests` | MSTest 测试项目，同时覆盖类库和 CLI |

## 核心功能

1. **Excel 转对象**: 将 Excel 文件内容转换为 .NET 对象集合
2. **对象转 Excel**: 将 .NET 对象集合导出为 Excel 文件
3. **自动列宽**: 基于内容自动调整 Excel 列宽
4. **日期时间格式**: 内置常见日期时间解析格式（见 `ExcelConstants.DateFormats.CommonDateTimeFormats`，另有中文"年月日"格式的特殊处理）
5. **样式支持**: 支持字体、颜色、对齐方式等 Excel 样式设置
6. **公式列**: 用 Lambda 表达式生成 Excel 公式，支持按列标题、模型属性和跨 sheet 引用（见 [ExcelFunctions.md](../ExcelFunctions.md)）

## 技术栈

- **主要依赖**: NPOI 2.8.0 (Excel 操作)
- **目标框架**: 多目标框架 (net472, netstandard2.0, netstandard2.1, net6.0, net8.0, net9.0, net10.0)
- **测试框架**: MSTest
- **CI/CD**: GitHub Actions

## 编码规范

### 命名约定

1. **类名**: 使用 PascalCase，例如 `ExcelExporter`, `ExcelImporter`
2. **方法名**: 使用 PascalCase，例如 `ToExcel()`, `ToObject<T>()`
3. **属性名**: 使用 PascalCase，例如 `FileName`, `SheetName`
4. **私有字段**: 使用下划线前缀 + camelCase，例如 `_cellStyleDict`
5. **常量和静态只读字段**: 使用 PascalCase，例如 `ExcelConstants.DefaultHeaderRowIndex`

### 代码风格

1. **缩进**: 使用 Tab 缩进（项目现有风格）
2. **语言版本**: `<LangVersion>14.0</LangVersion>`，启用隐式 using
3. **可空性**: 启用可空引用类型注解 (`Nullable=annotations`)
4. **注释**: 对公共 API 提供 XML 文档注释
5. **异常处理**: 使用明确的异常类型，提供有意义的错误消息

### 特性使用

项目广泛使用 Attribute 来配置 Excel 映射：

```csharp
[ExcelTitle("工作表名称")]
public class MyModel
{
    [ExcelColumn("列标题", Order = 1)]
    public string Name { get; set; }
    
    [ExcelColumn("日期", Format = "yyyy-MM-dd")]
    public DateTime? Date { get; set; }
    
    [ExcelColumn("金额", CellAlignment = HorizontalAlignment.Right)]
    public decimal Amount { get; set; }
}
```

## 开发指南

### 添加新功能

1. **保持兼容性**: 确保新功能在所有支持的框架版本上正常工作
2. **编写测试**: 为新功能添加 MSTest 测试用例
3. **更新文档**: 在 README.md 中添加功能说明和示例代码
4. **遵循 SemVer**: 根据语义化版本规范更新版本号

### 测试要求

1. **单元测试**: 使用 MSTest 编写测试（CLI 的测试直接调用 `Excel2ObjCli.Run`，用 `StringWriter` 捕获输出）
2. **测试命名**: 使用描述性的测试方法名，例如 `Should_ExportCorrectly_When_ModelHasNullValues`
3. **测试覆盖**: 确保核心功能和边界情况都有测试覆盖
4. **测试数据**: 测试用例应包含中文字符测试，确保 Unicode 支持

### 性能考虑

1. **大文件处理**: 考虑大型 Excel 文件的内存使用
2. **流式处理**: 对于大数据集，优先使用流式处理
3. **缓存机制**: 合理使用缓存避免重复计算

## 提交规范

### Commit Message 格式

建议使用约定式提交（Conventional Commits）格式：

```
<类型>[可选的作用域]: <描述>

[可选的正文]

[可选的脚注]
```

**类型**:
- `feat`: 新功能
- `fix`: 修复 bug
- `docs`: 文档更新
- `style`: 代码格式调整（不影响功能）
- `refactor`: 重构代码
- `test`: 添加或修改测试
- `chore`: 构建过程或辅助工具的变动

**示例**:
```
feat: 支持自定义日期格式导出
fix: 修复空值导致的 NullReferenceException
docs: 更新 README 中的安装说明
```

### Pull Request 规范

1. **标题**: 简明扼要地描述变更内容
2. **描述**: 详细说明变更的原因、实现方式和影响
3. **关联 Issue**: 如果相关，使用 `Closes #123` 关联 Issue
4. **测试**: 说明如何测试变更
5. **Breaking Changes**: 明确标注破坏性变更

## 版本发布

### 版本号规范

项目使用语义化版本（Semantic Versioning）：

- **主版本号 (Major)**: 不兼容的 API 变更
- **次版本号 (Minor)**: 向后兼容的功能新增
- **修订号 (Patch)**: 向后兼容的问题修正

当前版本以 `Chsword.Excel2Object/Chsword.Excel2Object.csproj` 中的 `<Version>` 为准；CLI 项目自动读取该版本号，两个包始终同版本发布。历史版本见 [README 发布说明](../README.md#发布说明和路线图)。

### 发布流程

1. **更新版本号**: 在 `Chsword.Excel2Object.csproj` 中更新 `<Version>` 标签
2. **更新 README**: 在 README.md 和 README_EN.md 的发布说明中添加版本更新内容（工作流会按 `* **YYYY.MM.DD** - vX.Y.Z` 这一行从 README.md 提取 Release 说明）
3. **创建 Git Tag**: 使用 `v{version}` 格式，例如 `v2.2.2`
4. **自动发布**: 推送 tag 后自动触发 NuGet 和 GitHub Release 发布（一次发布两个包：类库和 CLI）

## 常见问题

### 多框架支持

项目支持多个目标框架，编码时注意：

1. 使用条件编译处理框架差异
2. 避免使用特定框架独有的 API
3. 针对不同框架的 API 差异提供兼容层

### NPOI 使用

1. NPOI 是项目的核心依赖，用于 Excel 文件操作
2. 注意 NPOI 的版本兼容性
3. 了解 NPOI 的内存管理，适当释放资源

### 中文支持

1. 确保所有功能正确处理中文字符
2. 字符宽度计算需考虑全角/半角字符差异
3. 日期格式支持中文格式（如"yyyy年MM月dd日"）

## 贡献指南

我们欢迎各种形式的贡献：

1. **报告 Bug**: 提交详细的 Issue，包括重现步骤
2. **功能建议**: 在 Issue 或 Discussion 中讨论新功能
3. **代码贡献**: 提交 Pull Request，遵循上述规范
4. **文档改进**: 改进 README、示例代码和注释

## 资源链接

- **项目主页**: https://github.com/chsword/Excel2Object
- **NuGet 包（类库）**: https://www.nuget.org/packages/Chsword.Excel2Object
- **NuGet 包（命令行工具）**: https://www.nuget.org/packages/Chsword.Excel2Object.Cli
- **NPOI 文档**: https://github.com/tonyqus/npoi
- **问题讨论**: http://www.cnblogs.com/chsword/p/excel2object.html

## 许可证

本项目采用 MIT 许可证。详见 [LICENSE](../LICENSE) 文件。
