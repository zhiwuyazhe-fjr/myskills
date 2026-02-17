# myskills

`myskills` 是知无涯者的个人 Skills 仓库，用于集中存储、维护与复用日常高频使用的 Claude Skills。

## 仓库目标

- 沉淀可复用的技能能力，减少重复工作。
- 统一技能组织方式，便于检索、扩展和迭代。
- 保持技能描述与实现一致，支持长期维护。

## Claude Skills 官方格式规范（本仓库约定）

每个 Skill 使用独立目录管理，目录内至少包含一个 `SKILL.md`，并遵循以下规范：

1. 使用 `skill` 代码块包裹元信息与说明内容。
2. 在文档开头提供 YAML 元信息：
   - `name`: Skill 唯一名称
   - `description`: Skill 功能与适用场景说明
3. 在正文中按清晰结构描述：
   - 能力范围与适用边界
   - 工作流/执行步骤
   - 强制规则（如格式、约束、优先级）
   - 调用入口与核心实现位置

推荐目录结构：

```text
myskills/
  <skill-name>/
    SKILL.md
    assets/
    references/
    scripts/
```

## 1 ：docx

位置：`docx/SKILL.md`

`docx` skill 用于将 Markdown 或 md 风格纯文本转换为符合中文排版规范的 `.docx` 文件，适用于对标题层级、字体字号、行距、列表缩进与表格渲染有严格要求的文档生成场景。

核心特点：

- 固定并可复用的中文文档排版规则（标题、正文、列表、表格）。
- 支持 Markdown 常用结构解析与渲染（标题、列表、表格、加粗）。
- 提供命令行与模块两种调用方式，便于脚本化集成。

命令行示例：

```bash
python scripts/md2docx.py -i input.md -o output.docx
```

---

后续可按同一规范继续新增更多 skills，并在本 README 中持续更新索引与简介。
