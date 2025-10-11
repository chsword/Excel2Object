#!/usr/bin/env pwsh
<#
.SYNOPSIS
    自动化发布脚本 - 更新版本号、提交代码、创建并推送Tag
    Release automation script - Update version, commit changes, create and push tag

.DESCRIPTION
    该脚本自动执行以下操作：
    1. 更新 Chsword.Excel2Object.csproj 中的版本号
    2. 提交代码到 Git 仓库
    3. 创建 Git Tag (格式: v{version})
    4. 推送代码和 Tag 到远程仓库

    This script automates the following operations:
    1. Update version in Chsword.Excel2Object.csproj
    2. Commit changes to Git repository
    3. Create Git Tag (format: v{version})
    4. Push code and tag to remote repository

.PARAMETER Version
    新版本号，格式：主版本.次版本.修订版 (例如: 2.0.3)
    New version number, format: Major.Minor.Patch (e.g., 2.0.3)

.PARAMETER SkipPush
    仅创建本地提交和标签，不推送到远程仓库
    Only create local commits and tags, do not push to remote

.PARAMETER Force
    强制执行，跳过确认提示
    Force execution, skip confirmation prompts

.EXAMPLE
    ./release.ps1 -Version 2.0.3
    更新版本到 2.0.3，提交并推送

.EXAMPLE
    ./release.ps1 -Version 2.1.0 -SkipPush
    更新版本到 2.1.0，仅本地提交，不推送

.EXAMPLE
    ./release.ps1 -Version 3.0.0 -Force
    强制更新版本到 3.0.0，跳过确认提示
#>

param(
    [Parameter(Mandatory = $true, HelpMessage = "版本号 (例如: 2.0.3)")]
    [string]$Version,
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipPush,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# 设置错误处理
$ErrorActionPreference = "Stop"

# 颜色输出函数
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Write-Success {
    param([string]$Message)
    Write-ColorOutput "✓ $Message" "Green"
}

function Write-Error {
    param([string]$Message)
    Write-ColorOutput "✗ $Message" "Red"
}

function Write-Info {
    param([string]$Message)
    Write-ColorOutput "ℹ $Message" "Cyan"
}

function Write-Warning {
    param([string]$Message)
    Write-ColorOutput "⚠ $Message" "Yellow"
}

# 验证版本号格式 (语义化版本)
function Test-VersionFormat {
    param([string]$Ver)
    
    if ($Ver -match '^\d+\.\d+\.\d+$') {
        return $true
    }
    return $false
}

# 提取当前版本号
function Get-CurrentVersion {
    $csprojPath = "Chsword.Excel2Object/Chsword.Excel2Object.csproj"
    
    if (-not (Test-Path $csprojPath)) {
        throw "找不到项目文件: $csprojPath"
    }
    
    $content = Get-Content $csprojPath -Raw
    if ($content -match '<Version>([^<]+)</Version>') {
        return $Matches[1]
    }
    
    throw "无法从项目文件中提取版本号"
}

# 更新版本号
function Update-Version {
    param(
        [string]$NewVersion
    )
    
    $csprojPath = "Chsword.Excel2Object/Chsword.Excel2Object.csproj"
    
    $content = Get-Content $csprojPath -Raw
    $newContent = $content -replace '<Version>[^<]+</Version>', "<Version>$NewVersion</Version>"
    
    Set-Content -Path $csprojPath -Value $newContent -NoNewline
    
    Write-Success "版本号已更新为: $NewVersion"
}

# 检查 Git 状态
function Test-GitClean {
    $status = git status --porcelain
    if ($status) {
        return $false
    }
    return $true
}

# 检查是否在 Git 仓库中
function Test-GitRepository {
    try {
        git rev-parse --git-dir 2>&1 | Out-Null
        return $LASTEXITCODE -eq 0
    }
    catch {
        return $false
    }
}

# 检查 Tag 是否已存在
function Test-TagExists {
    param([string]$TagName)
    
    $tags = git tag -l $TagName
    return ($tags -eq $TagName)
}

# 主流程
try {
    Write-Info "===== Excel2Object 自动化发布脚本 ====="
    Write-Info ""
    
    # 验证版本号格式
    if (-not (Test-VersionFormat -Ver $Version)) {
        Write-Error "版本号格式错误！请使用语义化版本格式 (例如: 2.0.3)"
        exit 1
    }
    
    # 检查是否在 Git 仓库中
    if (-not (Test-GitRepository)) {
        Write-Error "当前目录不是 Git 仓库！"
        exit 1
    }
    
    # 获取当前版本
    $currentVersion = Get-CurrentVersion
    Write-Info "当前版本: $currentVersion"
    Write-Info "目标版本: $Version"
    Write-Info ""
    
    # 检查版本是否相同
    if ($currentVersion -eq $Version) {
        Write-Warning "目标版本与当前版本相同，无需更新。"
        if (-not $Force) {
            $continue = Read-Host "是否继续？(y/N)"
            if ($continue -ne 'y' -and $continue -ne 'Y') {
                Write-Info "操作已取消。"
                exit 0
            }
        }
    }
    
    # 检查工作目录是否干净
    if (-not (Test-GitClean)) {
        Write-Warning "工作目录有未提交的更改！"
        git status --short
        Write-Info ""
        
        if (-not $Force) {
            $continue = Read-Host "是否继续？这将提交所有更改 (y/N)"
            if ($continue -ne 'y' -and $continue -ne 'Y') {
                Write-Info "操作已取消。"
                exit 0
            }
        }
    }
    
    # 检查 Tag 是否已存在
    $tagName = "v$Version"
    if (Test-TagExists -TagName $tagName) {
        Write-Error "Tag '$tagName' 已存在！"
        Write-Info "请使用不同的版本号或删除现有 Tag："
        Write-Info "  git tag -d $tagName"
        Write-Info "  git push origin :refs/tags/$tagName"
        exit 1
    }
    
    # 确认操作
    if (-not $Force) {
        Write-Info ""
        Write-Warning "即将执行以下操作："
        Write-Info "  1. 更新版本号: $currentVersion → $Version"
        Write-Info "  2. 提交更改: 'chore: bump version to $Version'"
        Write-Info "  3. 创建 Tag: $tagName"
        if (-not $SkipPush) {
            Write-Info "  4. 推送到远程仓库"
        }
        Write-Info ""
        
        $confirm = Read-Host "确认执行？(y/N)"
        if ($confirm -ne 'y' -and $confirm -ne 'Y') {
            Write-Info "操作已取消。"
            exit 0
        }
        Write-Info ""
    }
    
    # 步骤 1: 更新版本号
    Write-Info "步骤 1/4: 更新版本号..."
    Update-Version -NewVersion $Version
    
    # 步骤 2: 提交更改
    Write-Info "步骤 2/4: 提交更改..."
    git add .
    git commit -m "chore: bump version to $Version"
    Write-Success "代码已提交"
    
    # 步骤 3: 创建 Tag
    Write-Info "步骤 3/4: 创建 Tag..."
    git tag $tagName
    Write-Success "Tag '$tagName' 已创建"
    
    # 步骤 4: 推送到远程
    if (-not $SkipPush) {
        Write-Info "步骤 4/4: 推送到远程仓库..."
        
        # 获取当前分支
        $currentBranch = git rev-parse --abbrev-ref HEAD
        
        # 推送代码
        Write-Info "推送分支 '$currentBranch'..."
        git push origin $currentBranch
        Write-Success "代码已推送到 origin/$currentBranch"
        
        # 推送 Tag
        Write-Info "推送 Tag '$tagName'..."
        git push origin $tagName
        Write-Success "Tag 已推送"
    }
    else {
        Write-Info "步骤 4/4: 跳过推送 (使用了 -SkipPush 参数)"
        Write-Warning "请手动推送代码和 Tag："
        Write-Info "  git push origin $(git rev-parse --abbrev-ref HEAD)"
        Write-Info "  git push origin $tagName"
    }
    
    # 完成
    Write-Info ""
    Write-Success "===== 发布流程完成 ====="
    Write-Info ""
    Write-Info "版本 $Version 已准备就绪！"
    
    if (-not $SkipPush) {
        Write-Info ""
        Write-Info "GitHub Actions 将自动开始构建和发布流程："
        Write-Info "  → https://github.com/chsword/Excel2Object/actions"
        Write-Info ""
        Write-Info "预计 5-10 分钟后可在以下位置查看发布结果："
        Write-Info "  → NuGet: https://www.nuget.org/packages/Chsword.Excel2Object"
        Write-Info "  → GitHub: https://github.com/chsword/Excel2Object/releases"
    }
}
catch {
    Write-Error "发生错误: $_"
    Write-Info ""
    Write-Info "如需帮助，请查看文档："
    Write-Info "  → RELEASE_AUTOMATION.md"
    exit 1
}
