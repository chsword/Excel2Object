# 发布快速参考卡 / Release Quick Reference Card

## 🚀 快速发布步骤 / Quick Release Steps

```bash
# 1️⃣ 更新版本号 / Update Version
vim Chsword.Excel2Object/Chsword.Excel2Object.csproj
# 修改 <Version>X.Y.Z</Version>

# 2️⃣ 更新发布说明 / Update Release Notes
vim README.md README_EN.md
# 添加版本发布说明

# 3️⃣ 提交变更 / Commit Changes
git add .
git commit -m "chore: bump version to X.Y.Z"
git push origin main

# 4️⃣ 创建并推送 Tag / Create & Push Tag
git tag vX.Y.Z
git push origin vX.Y.Z

# 5️⃣ 完成！/ Done!
# 访问 GitHub Actions 监控进度
# https://github.com/chsword/Excel2Object/actions
```

## 📊 版本递增速查表 / Version Increment Quick Reference

| 变更类型 | 版本递增 | 示例 |
|---------|---------|------|
| 🐛 Bug 修复 | `X.Y.Z → X.Y.(Z+1)` | 2.0.2 → 2.0.3 |
| ⚡ 性能优化 | `X.Y.Z → X.Y.(Z+1)` | 2.0.3 → 2.0.4 |
| ✨ 新功能 | `X.Y.Z → X.(Y+1).0` | 2.0.4 → 2.1.0 |
| 🎯 新框架支持 | `X.Y.Z → X.(Y+1).0` | 2.1.0 → 2.2.0 |
| 💥 破坏性变更 | `X.Y.Z → (X+1).0.0` | 2.2.0 → 3.0.0 |

## ⚙️ 自动化流程检查点 / Automation Checkpoints

自动发布流程会执行以下检查：

- ✅ **版本格式验证**: 确保符合 SemVer 规范
- ✅ **版本一致性**: Tag 版本 = csproj 版本
- ✅ **多平台构建**: Ubuntu + Windows
- ✅ **单元测试**: 所有测试必须通过
- ✅ **NuGet 打包**: 生成多框架包
- ✅ **自动发布**: NuGet.org + GitHub Release

## 🔧 必需配置 / Required Configuration

### GitHub Secrets
```
NUGET_API_KEY - NuGet.org API 密钥
```

**获取方式 / How to Get**:
1. 登录 https://www.nuget.org
2. Account Settings → API Keys
3. Create → 复制密钥
4. GitHub 仓库 → Settings → Secrets → New secret

## 📝 发布说明模板 / Release Notes Template

### 中文版本
```markdown
* **YYYY.MM.DD** - vX.Y.Z
- [x] ✨ **新增:** 功能描述
- [x] 🐛 **修复:** 问题描述
- [x] ⚡ **优化:** 改进描述
- [x] 📝 **文档:** 文档更新
```

### 英文版本
```markdown
* **YYYY.MM.DD** - vX.Y.Z
- [x] ✨ **Added:** Feature description
- [x] 🐛 **Fixed:** Issue description
- [x] ⚡ **Improved:** Enhancement description
- [x] 📝 **Docs:** Documentation update
```

## 🚨 常见问题快速解决 / Quick Troubleshooting

### 问题 1: 版本号不匹配
```
错误: Tag version (2.0.3) does not match csproj version (2.0.2)

解决: 确保 .csproj 文件中的 <Version> 与 tag 一致
```

### 问题 2: NuGet 推送失败
```
错误: 401 Unauthorized 或 409 Conflict

解决: 
- 检查 NUGET_API_KEY 是否正确
- 确认版本号未被占用
```

### 问题 3: 测试失败
```
错误: Tests failed

解决:
1. 本地运行测试: dotnet test
2. 修复失败的测试
3. 删除 tag: git push origin :refs/tags/vX.Y.Z
4. 重新发布
```

## 📦 预发布版本 / Pre-release Versions

```bash
# Alpha 版本
git tag v2.1.0-alpha.1
git push origin v2.1.0-alpha.1

# Beta 版本
git tag v2.1.0-beta.1
git push origin v2.1.0-beta.1

# RC 版本
git tag v2.1.0-rc.1
git push origin v2.1.0-rc.1
```

## 🔗 重要链接 / Important Links

- 📚 [完整发布指南](../.github/RELEASE_GUIDE.md)
- 📋 [版本管理规范](../.github/VERSIONING.md)
- 🤖 [Copilot 指令](../.github/copilot-instructions.md)
- 🔄 [发布自动化总览](../RELEASE_AUTOMATION.md)
- 🏃 [CI Workflow](workflows/dotnet-ci.yml)
- 🚀 [Release Workflow](workflows/release.yml)

## ⏱️ 预计时间 / Estimated Time

| 步骤 | 时间 |
|-----|------|
| 准备阶段（更新版本和文档）| 5-15 分钟 |
| 创建并推送 Tag | < 1 分钟 |
| 自动化流程执行 | 5-10 分钟 |
| **总计** | **10-26 分钟** |

## 📞 获取帮助 / Get Help

遇到问题？

- 💬 [GitHub Discussions](https://github.com/chsword/Excel2Object/discussions)
- 🐛 [Issue Tracker](https://github.com/chsword/Excel2Object/issues)
- 📧 Email: chsword@126.com

---

**提示**: 将此页面加入书签以便快速查阅！

**Tip**: Bookmark this page for quick reference!
