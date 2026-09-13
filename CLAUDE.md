# 项目约定

本文件供参与本仓库开发的人与 AI 助手共同遵循。

## 版本特性文档

每发布一个版本，须在 `docs/versions/` 下新增一篇同名文档（例如 `v2.7.1.md`），并在 `docs/README.md` 的版本表格中追加一行。

**撰写要求**

1. **使用书面语。** 采用完整、规范的中文书面表达，避免口语化说法（例如不写「一滚就没了」「白送」「省得」，改写为「滚动时将移出可视区域」「实现成本较低」「以免」）。避免夸张与营销辞令，只陈述事实。
2. **结构固定。** 依次为：标题与发布日期、「概述」一段、「变更明细」按条目分节，如有不兼容或行为变更则另设「不兼容变更」或「行为变更」小节。
3. **每个特性都要写。** 新增、改进、修复均需覆盖，不得只写新增功能。
4. **给出代码示例。** 凡涉及公开 API 的条目，附可直接运行的 C# 示例；示例中的 API 名称须与源码一致，不得杜撰。
5. **说明取舍与限制。** 若实现中存在权衡（例如 `.xls` 调色板只有 56 色、Excel 将分钟视为月份），在对应条目中写明原因与规避方式。
6. **同步更新。** `README.md` 与 `README_EN.md` 的发布说明仍需按既有格式追加条目，二者与版本文档并行维护。

## 文件编码

本仓库的文件编码并不统一，修改时须逐个保持原样：

- 部分 `.cs` 文件带 UTF-8 BOM（如 `ExcelExporter.cs`、`ExcelColumnAttribute.cs`、`ExcelHelper.cs`），其余不带。
- `Chsword.Excel2Object/Chsword.Excel2Object.csproj` 为 **GBK** 编码，其中含中文 `<Description>`，不得以 UTF-8 覆写。

提交前可执行以下检查，确认 BOM 状态未被改变：

```bash
for f in $(git diff --name-only); do
  a=$(git show HEAD:$f | head -c3 | xxd -p | grep -c efbbbf)
  b=$(head -c3 $f | xxd -p | grep -c efbbbf)
  [ "$a" != "$b" ] && echo "BOM CHANGED $f"
done
```

## 构建与测试

```bash
dotnet build Chsword.Excel2Object          # 全部目标框架，应为 0 警告
dotnet test Chsword.Excel2Object.Tests -f net10.0
```

## 发布

版本号维护在 `Chsword.Excel2Object/Chsword.Excel2Object.csproj` 的 `<Version>`，与 tag 必须一致；推送 `v*` tag 触发发布流程。流程详见 [RELEASE_AUTOMATION.md](RELEASE_AUTOMATION.md)。
