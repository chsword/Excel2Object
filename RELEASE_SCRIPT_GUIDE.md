# Release Script Usage Guide / 发布脚本使用指南

## English

### Prerequisites

- **Windows**: PowerShell 5.1 or later (comes with Windows 10/11)
- **Linux/macOS**: PowerShell Core 7.0+ ([Install PowerShell](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell))

### Quick Start

```powershell
# Auto-increment version and release
.\release.ps1

# Or release a specific version
.\release.ps1 -Version 2.3.0

# Get help
.\release.ps1 -Help
```

The script will:
1. ✓ Validate version format
2. ✓ Update version in `.csproj` file
3. ✓ Commit changes
4. ✓ Create git tag
5. ✓ Push to remote repository

### Usage Examples

#### Auto-Increment Release
```powershell
.\release.ps1
```
Automatically increments the patch version (e.g., 2.2.1 → 2.2.2) and releases.

#### Show Help
```powershell
.\release.ps1 -Help
```
Displays usage information and examples.

#### Basic Release
```powershell
.\release.ps1 -Version 2.3.0
```
This will prompt for confirmation before each major step.

#### Local-Only Release
```powershell
.\release.ps1 -Version 2.3.0 -SkipPush
```
Creates commit and tag locally, but doesn't push to remote.

#### Force Release (No Confirmations)
```powershell
.\release.ps1 -Version 2.3.0 -Force
```
Skips all confirmation prompts.

#### Combined Flags
```powershell
.\release.ps1 -Version 2.3.0 -SkipPush -Force
```
Creates local commit/tag without pushing, no confirmations.

### Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Version` | No | Version number in format: `Major.Minor.Patch` (e.g., `2.3.0`). If not specified, auto-increments the patch version |
| `-SkipPush` | No | Create local commits and tags only, don't push to remote |
| `-Force` | No | Skip all confirmation prompts |
| `-Help` | No | Display help information |

### Troubleshooting

#### Error: "Tag already exists"
```powershell
# Delete local and remote tag
git tag -d v2.3.0
git push origin :refs/tags/v2.3.0

# Then run the script again
.\release.ps1 -Version 2.3.0
```

#### Error: "Execution policy"
On Windows, you might need to allow script execution:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### Get Help
```powershell
Get-Help .\release.ps1 -Detailed
```

---

## 中文

### 前置要求

- **Windows**: PowerShell 5.1 或更高版本（Windows 10/11 自带）
- **Linux/macOS**: PowerShell Core 7.0+ ([安装 PowerShell](https://docs.microsoft.com/zh-cn/powershell/scripting/install/installing-powershell))

### 快速开始

```powershell
# 自动递增版本并发布
.\release.ps1

# 或发布指定版本
.\release.ps1 -Version 2.3.0

# 获取帮助
.\release.ps1 -Help
```

脚本将执行：
1. ✓ 验证版本号格式
2. ✓ 更新 `.csproj` 文件中的版本号
3. ✓ 提交更改
4. ✓ 创建 git 标签
5. ✓ 推送到远程仓库

### 使用示例

#### 自动递增发布
```powershell
.\release.ps1
```
自动递增修订号（例如：2.2.1 → 2.2.2）并发布。

#### 显示帮助
```powershell
.\release.ps1 -Help
```
显示使用说明和示例。

#### 基本发布
```powershell
.\release.ps1 -Version 2.3.0
```
在每个主要步骤前会提示确认。

#### 仅本地发布
```powershell
.\release.ps1 -Version 2.3.0 -SkipPush
```
在本地创建提交和标签，但不推送到远程。

#### 强制发布（无确认）
```powershell
.\release.ps1 -Version 2.3.0 -Force
```
跳过所有确认提示。

#### 组合参数
```powershell
.\release.ps1 -Version 2.3.0 -SkipPush -Force
```
创建本地提交/标签但不推送，无确认提示。

### 参数说明

| 参数 | 必需 | 说明 |
|------|------|------|
| `-Version` | 否 | 版本号格式：`主版本.次版本.修订版`（例如：`2.3.0`）。如果不指定，则自动递增修订号 |
| `-SkipPush` | 否 | 仅创建本地提交和标签，不推送到远程 |
| `-Force` | 否 | 跳过所有确认提示 |
| `-Help` | 否 | 显示帮助信息 |

### 故障排除

#### 错误："Tag 已存在"
```powershell
# 删除本地和远程标签
git tag -d v2.3.0
git push origin :refs/tags/v2.3.0

# 然后重新运行脚本
.\release.ps1 -Version 2.3.0
```

#### 错误："执行策略"
在 Windows 上，您可能需要允许脚本执行：
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### 获取帮助
```powershell
Get-Help .\release.ps1 -Detailed
```

---

## See Also / 另见

- [RELEASE_AUTOMATION.md](RELEASE_AUTOMATION.md) - Complete release automation documentation
- [.github/VERSIONING.md](.github/VERSIONING.md) - Versioning guidelines
- [.github/RELEASE_GUIDE.md](.github/RELEASE_GUIDE.md) - Detailed release process guide
