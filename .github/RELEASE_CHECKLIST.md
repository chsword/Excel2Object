# 发布检查清单 / Release Checklist

使用此清单确保发布流程的每个步骤都已完成。

Use this checklist to ensure every step of the release process is completed.

---

## 📋 发布前检查 / Pre-Release Checklist

### 代码质量 / Code Quality
- [ ] 所有功能已完成开发
- [ ] 代码已通过 Code Review
- [ ] 所有单元测试通过（本地运行 `dotnet test`）
- [ ] 没有编译警告或错误
- [ ] 代码符合项目编码规范

### 版本管理 / Version Management
- [ ] 确定新版本号（根据 [VERSIONING.md](.github/VERSIONING.md) 规范）
  - [ ] Bug 修复 → Patch +1
  - [ ] 新功能 → Minor +1
  - [ ] 破坏性变更 → Major +1
- [ ] 更新 `Chsword.Excel2Object/Chsword.Excel2Object.csproj` 中的 `<Version>` 标签
- [ ] 版本号格式正确（X.Y.Z 或 X.Y.Z-prerelease）

### 文档更新 / Documentation Update
- [ ] 在 `README.md` 中添加中文发布说明
  - [ ] 日期格式正确（YYYY.MM.DD）
  - [ ] 版本号正确（vX.Y.Z）
  - [ ] 变更说明清晰完整
  - [ ] 使用适当的 emoji 和格式
- [ ] 在 `README_EN.md` 中添加英文发布说明
  - [ ] 与中文版本保持同步
  - [ ] 翻译准确
- [ ] 如有新功能，更新示例代码
- [ ] 如有 API 变更，更新 API 文档

### 功能验证 / Feature Verification
- [ ] 新功能已手动测试
- [ ] 示例代码可以正常运行
- [ ] 在多个框架版本上测试（如可能）
  - [ ] .NET Framework 4.7.2
  - [ ] .NET Standard 2.0
  - [ ] .NET 6.0
  - [ ] .NET 8.0
- [ ] 向后兼容性验证（对于非 Major 版本）

### 破坏性变更检查 / Breaking Changes Check
- [ ] 识别所有破坏性变更
- [ ] 在发布说明中明确标注 💥 Breaking Changes
- [ ] 提供迁移指南（如需要）
- [ ] 更新示例代码以适应新 API

### NuGet 包验证 / NuGet Package Verification
- [ ] 本地构建 NuGet 包成功
  ```bash
  dotnet pack --configuration Release --output ./artifacts
  ```
- [ ] 检查包内容是否正确
  ```bash
  # 解压 .nupkg 文件检查内容
  ```
- [ ] 包大小合理（不包含不必要的文件）
- [ ] 所有目标框架的 DLL 都包含在包中

---

## 🚀 发布执行 / Release Execution

### Git 操作 / Git Operations
- [ ] 所有变更已提交到 main 分支
  ```bash
  git add .
  git commit -m "chore: bump version to X.Y.Z"
  git push origin main
  ```
- [ ] 创建 Git Tag（格式：vX.Y.Z）
  ```bash
  git tag vX.Y.Z
  # 或带注释的 tag
  git tag -a vX.Y.Z -m "Release version X.Y.Z"
  ```
- [ ] 推送 Tag 到远程仓库
  ```bash
  git push origin vX.Y.Z
  ```

### 自动化监控 / Automation Monitoring
- [ ] 访问 GitHub Actions 页面监控流程
  - URL: https://github.com/chsword/Excel2Object/actions
- [ ] 确认 `validate-version` 作业通过
- [ ] 确认 `build-and-test` 作业通过（所有平台）
- [ ] 确认 `pack-nuget` 作业通过
- [ ] 确认 `publish-nuget` 作业通过
- [ ] 确认 `create-release` 作业通过

### 故障处理 / Failure Handling
如果任何作业失败：
- [ ] 查看失败作业的日志
- [ ] 识别失败原因
- [ ] 修复问题
- [ ] 删除失败的 Tag
  ```bash
  git tag -d vX.Y.Z
  git push origin :refs/tags/vX.Y.Z
  ```
- [ ] 递增版本号（如需要）
- [ ] 重新执行发布流程

---

## ✅ 发布后验证 / Post-Release Verification

### NuGet.org 验证 / NuGet.org Verification
- [ ] 访问 NuGet.org 搜索包
  - URL: https://www.nuget.org/packages/Chsword.Excel2Object/
- [ ] 确认新版本已列出
- [ ] 检查包版本号正确
- [ ] 检查包元数据（作者、描述、标签等）
- [ ] 确认包可以被搜索到（可能需要几分钟）
- [ ] 下载统计开始计数

### GitHub Release 验证 / GitHub Release Verification
- [ ] 访问 GitHub Releases 页面
  - URL: https://github.com/chsword/Excel2Object/releases
- [ ] 确认新 Release 已创建
- [ ] 检查 Release 标题和标签正确
- [ ] 验证发布说明内容完整
- [ ] 确认 NuGet 包文件已附加
- [ ] 预发布标记正确（如适用）

### 功能测试 / Functional Testing
- [ ] 在新项目中安装 NuGet 包
  ```bash
  dotnet new console -n TestProject
  cd TestProject
  dotnet add package Chsword.Excel2Object --version X.Y.Z
  ```
- [ ] 运行示例代码验证功能
- [ ] 确认没有运行时错误
- [ ] 验证新功能工作正常

### 文档验证 / Documentation Verification
- [ ] README 徽章显示正确的版本号
- [ ] 文档链接都可访问
- [ ] 示例代码与新版本匹配
- [ ] API 文档（如有）已更新

---

## 📢 发布通知 / Release Announcement

### 内部通知 / Internal Notification
- [ ] 通知团队成员发布完成
- [ ] 在项目 Discussion 中发布公告
- [ ] 更新项目看板或 Milestone

### 社区通知 / Community Notification
- [ ] 在项目主页添加发布公告（如适用）
- [ ] 在博客或社交媒体分享（可选）
- [ ] 回复相关 Issue 通知问题已修复（如适用）

---

## 📊 发布后跟踪 / Post-Release Tracking

### 第一天 / First Day
- [ ] 监控 NuGet 下载量
- [ ] 检查是否有新的 Issue 报告
- [ ] 回应社区反馈

### 第一周 / First Week
- [ ] 收集用户反馈
- [ ] 记录任何问题或改进建议
- [ ] 评估是否需要发布 Patch 版本

---

## 🔄 回滚准备 / Rollback Preparation

如果发现严重问题需要回滚：

### NuGet 回滚 / NuGet Rollback
- [ ] 登录 NuGet.org
- [ ] 找到问题版本
- [ ] 点击 "Unlist" 从搜索中移除
- [ ] 注意：已安装的用户仍可使用

### GitHub Release 回滚 / GitHub Release Rollback
- [ ] 删除或编辑 Release 页面
- [ ] 添加警告说明
- [ ] 可选：删除 Git Tag

### 修复版本 / Fix Version
- [ ] 修复问题
- [ ] 递增版本号（通常是 Patch +1）
- [ ] 重新执行发布流程

---

## 📝 经验总结 / Lessons Learned

发布完成后，记录经验教训：

- [ ] 本次发布遇到的问题
- [ ] 解决方案和改进建议
- [ ] 更新文档和流程（如需要）
- [ ] 分享给团队成员

---

## ✨ 发布完成 / Release Complete

🎉 恭喜！发布流程已完成！

🎉 Congratulations! The release process is complete!

**发布信息 / Release Information:**
- 版本号 / Version: ____________
- 发布日期 / Release Date: ____________
- NuGet 链接 / NuGet Link: https://www.nuget.org/packages/Chsword.Excel2Object/____________
- Release 链接 / Release Link: https://github.com/chsword/Excel2Object/releases/tag/v____________

**执行人 / Released By:** ____________

**特别说明 / Special Notes:** 
____________________________________________
____________________________________________
____________________________________________
