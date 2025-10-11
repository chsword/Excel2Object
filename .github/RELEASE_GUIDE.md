# 发布流程说明 / Release Process Guide

[中文](#中文版本) | [English](#english-version)

---

## 中文版本

### 概述

本文档详细说明 Excel2Object 项目的自动发布流程，包括 NuGet 包发布和 GitHub Release 创建。

### 自动化发布流程

项目使用 GitHub Actions 实现完全自动化的发布流程。当推送符合 `v*` 格式的 Git Tag 时，会自动触发以下流程：

```
推送 Tag (v2.0.3)
    ↓
[1] 验证版本号
    ├─ 检查版本格式（SemVer）
    ├─ 提取 tag 版本号
    ├─ 提取 csproj 版本号
    └─ 验证版本号一致性
    ↓
[2] 构建和测试
    ├─ 多平台构建（Ubuntu/Windows）
    ├─ 运行所有单元测试
    └─ 上传测试结果
    ↓
[3] 打包 NuGet
    ├─ 编译 Release 版本
    ├─ 生成 NuGet 包
    └─ 上传包作为构建产物
    ↓
[4] 发布到 NuGet.org
    └─ 推送包到 NuGet 仓库
    ↓
[5] 创建 GitHub Release
    ├─ 从 README 提取发布说明
    ├─ 附加 NuGet 包文件
    └─ 创建 Release 页面
    ↓
[6] 完成通知
```

### 发布步骤

#### 准备阶段

1. **确定版本号**

   根据变更类型确定新版本号（参考 [版本管理规范](VERSIONING.md)）：
   - Bug 修复 → 修订号 +1（如 2.0.2 → 2.0.3）
   - 新功能 → 次版本号 +1（如 2.0.3 → 2.1.0）
   - 破坏性变更 → 主版本号 +1（如 2.1.0 → 3.0.0）

2. **更新版本号**

   在 `Chsword.Excel2Object/Chsword.Excel2Object.csproj` 文件中更新版本号：
   ```xml
   <Version>2.0.3</Version>
   ```

3. **更新发布说明**

   在 `README.md` 中添加新版本的发布说明：
   ```markdown
   ### 发布说明

   * **2025.10.15** - v2.0.3
   - [x] 修复：修复大文件导出时的内存溢出问题
   - [x] 优化：提升列宽自动计算的性能
   ```

   同步更新 `README_EN.md` 英文版本。

4. **提交变更**

   ```bash
   git add Chsword.Excel2Object/Chsword.Excel2Object.csproj README.md README_EN.md
   git commit -m "chore: bump version to 2.0.3"
   git push origin main
   ```

#### 发布阶段

5. **创建并推送 Tag**

   ```bash
   # 创建 tag（注意使用 v 前缀）
   git tag v2.0.3
   
   # 或者创建带注释的 tag
   git tag -a v2.0.3 -m "Release version 2.0.3"
   
   # 推送 tag 到远程仓库
   git push origin v2.0.3
   ```

6. **监控自动化流程**

   推送 tag 后，访问 GitHub Actions 页面查看发布进度：
   ```
   https://github.com/chsword/Excel2Object/actions
   ```

   整个流程大约需要 5-10 分钟，包括：
   - ✓ 版本验证
   - ✓ 构建和测试
   - ✓ NuGet 打包
   - ✓ 发布到 NuGet.org
   - ✓ 创建 GitHub Release

7. **验证发布结果**

   - **NuGet 包**: 访问 https://www.nuget.org/packages/Chsword.Excel2Object/
   - **GitHub Release**: 访问 https://github.com/chsword/Excel2Object/releases

   注意：NuGet 包可能需要几分钟才能在搜索中显示。

### 预发布版本

如需发布预发布版本（Alpha、Beta、RC），流程相同，但使用带预发布标识的版本号：

```bash
# 更新 csproj 版本号为 2.1.0-beta.1
# 创建并推送 tag
git tag v2.1.0-beta.1
git push origin v2.1.0-beta.1
```

预发布版本会被标记为 "Pre-release"，不会作为最新稳定版本显示。

### 配置要求

#### GitHub Secrets

发布流程需要以下 Secret 配置：

1. **NUGET_API_KEY**: NuGet.org API 密钥
   - 获取方式：登录 NuGet.org → Account Settings → API Keys
   - 权限要求：Push new packages and package versions
   - 配置路径：GitHub 仓库 → Settings → Secrets and variables → Actions

#### GitHub Environments

项目配置了 `nuget-production` 环境用于生产发布：
- 可选：配置审批流程，要求手动批准后才发布到 NuGet
- 路径：GitHub 仓库 → Settings → Environments

### 故障排查

#### 版本号不匹配

**错误信息**：
```
Warning: Tag version (2.0.3) does not match csproj version (2.0.2)
```

**解决方案**：
确保 `Chsword.Excel2Object.csproj` 中的 `<Version>` 与 tag 版本号一致（不包含 v 前缀）。

#### NuGet 推送失败

**可能原因**：
- API Key 无效或过期
- 包版本号已存在（不能重复发布相同版本）
- 网络问题

**解决方案**：
1. 检查 NUGET_API_KEY Secret 配置
2. 确认版本号未被使用
3. 重新运行失败的 workflow

#### 测试失败

如果测试失败，发布流程会自动停止。需要：
1. 查看测试日志定位问题
2. 修复问题后重新提交
3. 删除失败的 tag：
   ```bash
   git tag -d v2.0.3
   git push origin :refs/tags/v2.0.3
   ```
4. 重新创建并推送 tag

### 回滚发布

如需撤回已发布的版本：

1. **从 NuGet.org 撤回**
   - 登录 NuGet.org
   - 找到对应包版本
   - 点击 "Unlist" 将包从搜索中移除（不删除，已安装的用户仍可使用）

2. **删除 GitHub Release**
   - 进入 Releases 页面
   - 删除对应的 Release
   - 可选：删除对应的 tag

3. **发布修复版本**
   - 递增版本号（如 2.0.3 → 2.0.4）
   - 重新执行发布流程

### 最佳实践

1. **版本规划**
   - 在 issue 或 PR 中提前规划版本内容
   - 使用 milestone 追踪版本进度

2. **测试充分**
   - 发布前确保所有测试通过
   - 在多个框架版本上测试

3. **文档同步**
   - 同时更新中英文 README
   - 确保示例代码与新版本兼容

4. **发布节奏**
   - 修订版本（Patch）：按需发布，通常 1-2 周
   - 次版本（Minor）：功能完成后发布，通常 1-3 个月
   - 主版本（Major）：重大变更时发布，较少频率

5. **向后兼容**
   - 尽量保持向后兼容
   - 标记废弃 API 后至少保留一个主版本周期

---

## English Version

### Overview

This document details the automated release process for the Excel2Object project, including NuGet package publishing and GitHub Release creation.

### Automated Release Workflow

The project uses GitHub Actions for a fully automated release process. When a Git Tag matching the `v*` format is pushed, the following workflow is triggered:

```
Push Tag (v2.0.3)
    ↓
[1] Validate Version
    ├─ Check version format (SemVer)
    ├─ Extract tag version
    ├─ Extract csproj version
    └─ Validate version consistency
    ↓
[2] Build and Test
    ├─ Multi-platform build (Ubuntu/Windows)
    ├─ Run all unit tests
    └─ Upload test results
    ↓
[3] Pack NuGet
    ├─ Build Release version
    ├─ Generate NuGet package
    └─ Upload package as artifact
    ↓
[4] Publish to NuGet.org
    └─ Push package to NuGet repository
    ↓
[5] Create GitHub Release
    ├─ Extract release notes from README
    ├─ Attach NuGet package files
    └─ Create Release page
    ↓
[6] Completion Notification
```

### Release Steps

#### Preparation Phase

1. **Determine Version Number**

   Determine the new version number based on the type of changes (see [Versioning Guidelines](VERSIONING.md)):
   - Bug fixes → Patch +1 (e.g., 2.0.2 → 2.0.3)
   - New features → Minor +1 (e.g., 2.0.3 → 2.1.0)
   - Breaking changes → Major +1 (e.g., 2.1.0 → 3.0.0)

2. **Update Version Number**

   Update the version in `Chsword.Excel2Object/Chsword.Excel2Object.csproj`:
   ```xml
   <Version>2.0.3</Version>
   ```

3. **Update Release Notes**

   Add release notes for the new version in `README.md`:
   ```markdown
   ### Release Notes

   * **2025.10.15** - v2.0.3
   - [x] Fixed: Memory overflow issue when exporting large files
   - [x] Improved: Performance of auto column width calculation
   ```

   Sync the English version in `README_EN.md`.

4. **Commit Changes**

   ```bash
   git add Chsword.Excel2Object/Chsword.Excel2Object.csproj README.md README_EN.md
   git commit -m "chore: bump version to 2.0.3"
   git push origin main
   ```

#### Release Phase

5. **Create and Push Tag**

   ```bash
   # Create tag (note the v prefix)
   git tag v2.0.3
   
   # Or create an annotated tag
   git tag -a v2.0.3 -m "Release version 2.0.3"
   
   # Push tag to remote repository
   git push origin v2.0.3
   ```

6. **Monitor Automation**

   After pushing the tag, visit the GitHub Actions page to monitor progress:
   ```
   https://github.com/chsword/Excel2Object/actions
   ```

   The entire process takes approximately 5-10 minutes, including:
   - ✓ Version validation
   - ✓ Build and test
   - ✓ NuGet packaging
   - ✓ Publishing to NuGet.org
   - ✓ Creating GitHub Release

7. **Verify Release**

   - **NuGet Package**: Visit https://www.nuget.org/packages/Chsword.Excel2Object/
   - **GitHub Release**: Visit https://github.com/chsword/Excel2Object/releases

   Note: It may take a few minutes for the NuGet package to appear in search results.

### Pre-release Versions

To publish pre-release versions (Alpha, Beta, RC), follow the same process but use a version number with pre-release identifiers:

```bash
# Update csproj version to 2.1.0-beta.1
# Create and push tag
git tag v2.1.0-beta.1
git push origin v2.1.0-beta.1
```

Pre-release versions will be marked as "Pre-release" and won't be shown as the latest stable version.

### Configuration Requirements

#### GitHub Secrets

The release workflow requires the following Secret configuration:

1. **NUGET_API_KEY**: NuGet.org API key
   - How to obtain: Log in to NuGet.org → Account Settings → API Keys
   - Required permissions: Push new packages and package versions
   - Configuration path: GitHub repository → Settings → Secrets and variables → Actions

#### GitHub Environments

The project configures a `nuget-production` environment for production releases:
- Optional: Configure approval workflow requiring manual approval before publishing to NuGet
- Path: GitHub repository → Settings → Environments

### Troubleshooting

#### Version Mismatch

**Error Message**:
```
Warning: Tag version (2.0.3) does not match csproj version (2.0.2)
```

**Solution**:
Ensure the `<Version>` in `Chsword.Excel2Object.csproj` matches the tag version (without the v prefix).

#### NuGet Push Failure

**Possible Causes**:
- Invalid or expired API Key
- Package version already exists (cannot republish the same version)
- Network issues

**Solutions**:
1. Check NUGET_API_KEY Secret configuration
2. Confirm version number is not already used
3. Re-run the failed workflow

#### Test Failures

If tests fail, the release workflow stops automatically. You need to:
1. Review test logs to identify issues
2. Fix issues and commit again
3. Delete the failed tag:
   ```bash
   git tag -d v2.0.3
   git push origin :refs/tags/v2.0.3
   ```
4. Recreate and push the tag

### Rolling Back a Release

To retract a published version:

1. **Unlist from NuGet.org**
   - Log in to NuGet.org
   - Find the package version
   - Click "Unlist" to remove from search (doesn't delete; existing users can still use it)

2. **Delete GitHub Release**
   - Go to Releases page
   - Delete the corresponding Release
   - Optional: Delete the corresponding tag

3. **Publish Fix Version**
   - Increment version number (e.g., 2.0.3 → 2.0.4)
   - Re-execute release process

### Best Practices

1. **Version Planning**
   - Plan version content in advance in issues or PRs
   - Use milestones to track version progress

2. **Thorough Testing**
   - Ensure all tests pass before releasing
   - Test on multiple framework versions

3. **Documentation Sync**
   - Update both Chinese and English READMEs
   - Ensure example code is compatible with the new version

4. **Release Cadence**
   - Patch versions: As needed, typically 1-2 weeks
   - Minor versions: After features complete, typically 1-3 months
   - Major versions: For significant changes, less frequent

5. **Backward Compatibility**
   - Maintain backward compatibility when possible
   - Keep deprecated APIs for at least one major version cycle
