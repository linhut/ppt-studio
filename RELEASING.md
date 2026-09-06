<!--
  (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->

# ppt-studio 发布流程（Release Process）

本文档定义 ppt-studio 的**版本规范、发布流程与验证清单**。
参考 gongwen-skill 的发布实践制定，但 ppt-studio **不做 npm / pip 发布**——只做
**Git 仓库发布（GitHub 主仓库 + 可选国内镜像）+ GitHub Release**。

> 本仓库只同步**代码 / 文档 / 模板文本**，非必要文件（图片、PPTX 产物、预览截图、
> 本地脚本、缓存）一律不提交（见 `.gitignore`）。

---

## 1. 版本规范

采用标准 [SemVer](https://semver.org/lang/zh-CN/)：`MAJOR.MINOR.PATCH`

| 段 | 含义 | 示例 |
|----|------|------|
| `MAJOR` | 重大重构 / 不向后兼容 | `1.0.0` → `2.0.0` |
| `MINOR` | 功能迭代（向后兼容） | `1.0.0` → `1.1.0` |
| `PATCH` | 缺陷修复 | `1.0.0` → `1.0.1` |

规则：
- 禁止重复使用同一版本号；已打 tag 的版本不可覆盖
- 预发布后缀（`rc`/`dev`）仅内部验证，正式发布不用
- 版本号唯一来源 = `CHANGELOG.md` 条目 + git tag（本仓库无 npm/pip 多端同步需求）

---

## 2. 发布流程（标准操作）

```text
1. 确认代码就绪（功能完成、本地检查通过）
2. 清理非必要文件（图片/PPTX/预览/本地脚本不提交，见 §4）
3. 更新 CHANGELOG.md（把 Unreleased 收敛为 vX.Y.Z 条目）
4. 提交（chore: release vX.Y.Z）
5. 打注解 tag（git tag -a vX.Y.Z -m "..."）
6. 推送 main + tag 到远程（触发/手动创建 GitHub Release）
7. 创建 GitHub Release（编号 = tag）
8. 验证发布（gh release view / 页面核对）
```

### 2.1 详细步骤

```bash
# ① 确认工作区干净
git status

# ② 冒烟自检：编译示例仍可用
py -3 scripts/quality_check.py examples/json-demo/deck.json --preset tech --auto-only

# ③ 更新 CHANGELOG.md（Unreleased → vX.Y.Z 条目），并核对 README 描述

# ④ 提交
git add CHANGELOG.md README.md SKILL.md scripts references templates examples
git commit -m "chore: release vX.Y.Z"

# ⑤ 打注解 tag
git tag -a vX.Y.Z -m "vX.Y.Z - 发布说明摘要"

# ⑥ 推送
git push origin main --tags

# ⑦ 创建 GitHub Release（编号 = tag）
gh release create vX.Y.Z --title "ppt-studio vX.Y.Z" --generate-notes

# ⑧ 验证
gh release view vX.Y.Z
```

---

## 3. 版本号更新点

| # | 文件 | 位置 | 类型 |
|---|------|------|------|
| 1 | `CHANGELOG.md` | 顶部 `## vX.Y.Z (YYYY-MM-DD)` 条目 | 文档（唯一来源） |
| 2 | `README.md` | 项目描述/示例中的版本引用（如有） | 文档 |
| 3 | `SKILL.md` | frontmatter `metadata.version`（如启用） | 文档 |
| 4 | git tag | `vX.Y.Z`（必须与 CHANGELOG 一致） | 代码 |

> 核对命令：`grep -E "^## v[0-9]+\.[0-9]+\.[0-9]+" CHANGELOG.md | head -1` 与
> `git tag --sort=-v:refname | head -1` 一致。

---

## 4. 发布前检查清单

- [ ] `git status` 干净（无未提交改动）
- [ ] 无未跟踪的非必要文件（`git status` 中不应出现 `*.png` / `*.pptx` / `_export.ps1` / `__pycache__`）
- [ ] 敏感信息扫描：`grep -rnE "C:\\\\Users|oauth2:|password|token|api[_-]?key" --include='*.py' --include='*.md' --include='*.json' --include='*.yaml' .`
- [ ] 本地冒烟：`py -3 scripts/quality_check.py examples/json-demo/deck.json --preset tech --auto-only`
- [ ] CHANGELOG 顶部已有目标版本条目（对应 tag）
- [ ] 所有代码/文档文件已带作者标注（`Jose AI`），新增文件须补
- [ ] 仓库描述已更新（`gh repo edit --description "..."`）

---

## 5. 发布后验证

```bash
# ① GitHub Release 存在且编号 = tag
gh release view vX.Y.Z

# ② 远程与本地一致
git fetch origin && git status
git ls-remote origin refs/tags/vX.Y.Z
```

---

## 6. 仓库镜像（可选）

GitHub 为主仓库（唯一正式发布渠道）。如需国内镜像（GitCode / AtomGit），
参考 gongwen-skill 的做法配置 `gc` / `atomgit` remote 并推送 `main --tags`：

```bash
git remote add gc  https://oauth2:<token>@gitcode.com/linhut/ppt-studio.git
git remote add atomgit https://oauth2:<token>@atomgit.com/linhut/ppt-studio.git
git push gc main --tags
git push atomgit main --tags
```

> ⚠️ token 只存在本机 `.git/config`，**绝不写入任何仓库文件**；镜像平台 Release
> 需与 GitHub tag 保持相同编号。

---

## 7. 作者信息与许可（固定要求）

- 每个代码/文档文件必须带作者标注（只保留本项目作者，不加第三方署名）：
  - `.py` / `.ps1` 等脚本：文件头 `#` 注释
  - `.md` 文档：文件头 `<!-- -->` HTML 注释
  - `SKILL.md`：frontmatter `metadata.author`（不能破坏 frontmatter）
  - 数据/模板文件（`.json` / `.page` / `.yaml` / `.pptd`）：不加注释（解析安全），版权由 `LICENSE` + `README` 统一声明
- 仓库 `LICENSE`：MIT，Copyright (c) 2026 Jose AI
- 第三方组件许可证随附保留：`scripts/LICENSE.np-ppt`（编译内核，源自 dsh-np-ppt，MIT）
- 作者信息类微调（增删作者标注、版权行）**不做版本发布**：直接修改、提交并推送即可，
  不 bump 版本、不打 tag、不建 Release