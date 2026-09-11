# 版本管理规范 / Versioning Guidelines

[中文](#中文版本) | [English](#english-version)

---

## 中文版本

### 版本号格式

Excel2Object 项目遵循 [语义化版本 2.0.0](https://semver.org/lang/zh-CN/) 规范。

版本号格式：`主版本号.次版本号.修订号` (例如: `2.0.2`)

### 版本号递增规则

#### 1. 主版本号 (Major Version) - X.0.0

**何时递增**：当你做了不兼容的 API 修改

**示例场景**：
- 删除或重命名公共 API
- 更改方法签名导致破坏性变更
- 移除对旧框架版本的支持
- 重大架构调整

**递增方式**：
- 主版本号 +1
- 次版本号和修订号重置为 0
- 例如：`2.0.2` → `3.0.0`

#### 2. 次版本号 (Minor Version) - 0.X.0

**何时递增**：当你做了向下兼容的功能性新增

**示例场景**：
- 添加新的公共方法或属性
- 新增对新 Excel 格式的支持
- 添加新的特性或功能增强
- 添加对新框架版本的支持

**递增方式**：
- 次版本号 +1
- 修订号重置为 0
- 例如：`2.0.2` → `2.1.0`

#### 3. 修订号 (Patch Version) - 0.0.X

**何时递增**：当你做了向下兼容的问题修正

**示例场景**：
- 修复 Bug
- 性能优化（不改变 API）
- 依赖项版本更新（安全更新）
- 文档修正

**递增方式**：
- 修订号 +1
- 例如：`2.0.2` → `2.0.3`

### 版本号管理流程

#### 开发阶段

1. 在功能分支上开发新功能或修复 bug
2. 确保所有测试通过
3. 在 PR 中说明变更类型（Major/Minor/Patch）

#### 发布前准备

1. **确定版本号**：根据变更内容确定新版本号
2. **更新 csproj 文件**：
   ```xml
   <Version>2.0.3</Version>
   ```
3. **更新 README.md**：在发布说明部分添加新版本信息
   ```markdown
   * **YYYY.MM.DD** - vX.Y.Z
   - [x] 变更说明1
   - [x] 变更说明2
   ```
4. **更新 README_EN.md**：同步更新英文文档

#### 发布流程

1. **合并到主分支**：将包含版本更新的 PR 合并到 `main` 分支
2. **创建 Git Tag**：
   ```bash
   git tag v2.0.3
   git push origin v2.0.3
   ```
3. **自动发布**：
   - 推送 tag 后，GitHub Actions 自动触发
   - 自动构建并发布 NuGet 包
   - 自动创建 GitHub Release

### 版本历史参考

完整的版本历史见 [README 的发布说明](../README.md#发布说明和路线图)，那里同时也是发布工作流提取 GitHub Release 说明的数据来源。

### 预发布版本

如需发布预发布版本（Alpha、Beta、RC），使用以下格式：

- Alpha: `2.1.0-alpha.1`
- Beta: `2.1.0-beta.1`
- RC: `2.1.0-rc.1`

预发布版本与正式版本走同一条发布流水线：推送 `v*` tag 后同样会被推送到 NuGet.org，只是在 GitHub Release 上标记为 Pre-release。

### 自动化版本检查

项目使用 GitHub Actions 进行自动化版本检查：

1. **版本格式验证**：确保版本号符合 SemVer 格式
2. **版本号一致性**：检查 csproj 中的版本号与 tag 版本号一致
3. **版本号递增**：确保新版本号大于当前最新版本

---

## English Version

### Version Number Format

The Excel2Object project follows [Semantic Versioning 2.0.0](https://semver.org/).

Version format: `MAJOR.MINOR.PATCH` (e.g., `2.0.2`)

### Version Increment Rules

#### 1. Major Version - X.0.0

**When to increment**: When you make incompatible API changes

**Example scenarios**:
- Removing or renaming public APIs
- Changing method signatures that break compatibility
- Dropping support for old framework versions
- Major architecture changes

**How to increment**:
- Major +1
- Minor and Patch reset to 0
- Example: `2.0.2` → `3.0.0`

#### 2. Minor Version - 0.X.0

**When to increment**: When you add functionality in a backward compatible manner

**Example scenarios**:
- Adding new public methods or properties
- Adding support for new Excel formats
- Adding new features or enhancements
- Adding support for new framework versions

**How to increment**:
- Minor +1
- Patch reset to 0
- Example: `2.0.2` → `2.1.0`

#### 3. Patch Version - 0.0.X

**When to increment**: When you make backward compatible bug fixes

**Example scenarios**:
- Bug fixes
- Performance improvements (without API changes)
- Dependency updates (security patches)
- Documentation corrections

**How to increment**:
- Patch +1
- Example: `2.0.2` → `2.0.3`

### Version Management Workflow

#### Development Phase

1. Develop features or fixes in feature branches
2. Ensure all tests pass
3. Specify change type (Major/Minor/Patch) in PR

#### Pre-release Preparation

1. **Determine version number**: Based on change type
2. **Update csproj file**:
   ```xml
   <Version>2.0.3</Version>
   ```
3. **Update README.md**: Add new version to release notes
   ```markdown
   * **YYYY.MM.DD** - vX.Y.Z
   - [x] Change description 1
   - [x] Change description 2
   ```
4. **Update README_EN.md**: Sync English documentation

#### Release Process

1. **Merge to main**: Merge version update PR to `main` branch
2. **Create Git Tag**:
   ```bash
   git tag v2.0.3
   git push origin v2.0.3
   ```
3. **Automatic Release**:
   - GitHub Actions triggers automatically on tag push
   - Automatically builds and publishes NuGet package
   - Automatically creates GitHub Release

### Version History Reference

See the [release notes in the README](../README_EN.md#release-notes-and-roadmap) for the full version history; that section is also where the release workflow reads the GitHub Release body from.

### Pre-release Versions

For pre-release versions (Alpha, Beta, RC), use the following format:

- Alpha: `2.1.0-alpha.1`
- Beta: `2.1.0-beta.1`
- RC: `2.1.0-rc.1`

Pre-release versions go through the same pipeline as stable ones: pushing a `v*` tag publishes them to NuGet.org too, they are merely marked as a Pre-release on GitHub.

### Automated Version Checks

The project uses GitHub Actions for automated version checks:

1. **Version format validation**: Ensures version follows SemVer format (`X.Y.Z` or `X.Y.Z-prerelease`)
2. **Version consistency**: Checks csproj version matches tag version, and aborts the release when they differ

The workflow does not check that the new version is greater than the published one — that is up to the releaser; duplicate versions are skipped by `dotnet nuget push --skip-duplicate`.
