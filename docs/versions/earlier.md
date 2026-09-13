# v2.0.4 之前的历史

早期变更以日期而非版本号记录，现整理如下，按时间倒序排列。

## 2024-10-21

SixLabors.ImageSharp 更新至 2.1.9；补充 .NET 8.0 的测试。

## 2024-05-10

支持 .NET 8.0、.NET 6.0、.NET Standard 2.1、.NET Standard 2.0 与 .NET Framework 4.7.2；清理已弃用的库。

## 2023-11-02

支持列标题映射，可在导入时将工作表中的标题动态映射至模型属性。参见 [Issue39DynamicMappingTitle.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue39DynamicMappingTitle.cs)。

## 2023-07-31

支持 `DateTime` 与 `Nullable<DateTime>` 的格式声明，例如 `[ExcelColumn("Title", Format = "yyyy-MM-dd HH:mm:ss")]`。

## 2023-03-26

支持列标题中的特殊符号（#37）。

## 2023-02-20

支持的平台确定为 .NET Standard 2.0 与 2.1、.NET 6.0、.NET Framework 4.7.2。

## 2022-03-19

新增 `ExcelImporterOptions`，可通过 `TitleSkipLine` 跳过标题行之前的若干行；修复超类属性未被识别的问题。

```csharp
var list = ExcelHelper.ExcelToObject<Order>(bytes, options =>
{
    options.TitleSkipLine = 2;   // 标题行之前的行数
});
```

参见 [Issue32SkipLineImport.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue32SkipLineImport.cs) 与 [Issue31SuperClass.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue31SuperClass.cs)。

## 2021-11-04

支持多工作表导出。参见 [Pr28MultipleSheetTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr28MultipleSheetTest.cs)。

## 2021-10-23

修复可空 `DateTime` 的缺陷（由 @SunBrook 贡献）。

## 2021-10-22

支持可空类型。参见 [Pr24NullableTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr24NullableTest.cs)（由 @SunBrook 贡献）。

## 2021-05-28

新增 `[ExcelColumnAttribute]`，支持表头与单元格样式；新增公式支持。参见 [ExcelFunctions.md](../../ExcelFunctions.md)。
