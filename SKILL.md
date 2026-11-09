---
name: english-resume
description: 用固定模板生成一页英文简历（DOCX + PDF），面向外企/跨国公司投递；内含 Times New Roman 制表位排版、自动压页、外企简历内容红线与中译英改写规则。English resume builder for MNC internship/job applications — fixed one-page A4 layout, YAML content, DOCX + PDF output.
whenToUse: 需要写、改、或把中文简历翻成一页英文简历投外企（尤其是数据分析、AI 应用、咨询、市场研究类实习或校招）时
---

# 一页英文简历（外企版）

固定版式 + YAML 内容。改内容不用碰排版，排版不会因为改内容而散。

**这个 skill 解决三件容易做错的事**：外企简历该删什么、一页压不下时砍哪里、中文经历怎么改成有说服力的英文。

## 30 秒上手

```bash
# 1. 复制示例，改内容
cp examples/haiyang-zhao.yaml resume.yaml

# 2. 生成（自动缩小字号直到 1 页）
python scripts/build.py resume.yaml --png

# 3. 看 out/*.png 检查版面，投递时发 out/*.pdf
```

依赖：`pip install -r requirements.txt`（`python-docx`、`PyMuPDF`、`PyYAML`）。

## 必须遵守的六条

1. **只投 PDF**，DOCX 只用来改。外企 ATS 系统解析 PDF 更稳。
2. **必须有 Availability 行**：`Available: <开始> – <结束> | <每周天数> | <城市>`。
   简历筛选第一眼就是这一行，缺了直接出局。
3. **删掉**：照片、出生日期、政治面貌、婚姻状况、身份证号、技能自评分数。
   真外企（尤其美资）看到这些会判定"不懂国际规范"。详见 `references/content-rules.md`。
4. **章节顺序固定**：`PROFILE → EDUCATION → EXPERIENCE → PROJECTS → AWARDS → SKILLS`。
   PROFILE 必须在最前面（放最后的没人读）；获奖必须独立成节（塞进项目 bullet 里等于扔掉）。
5. **每条 bullet 以过去式动词开头 + 一个数字或一个决策**。
   禁止 `Responsible for` / `Familiar with` / `significantly`。
6. **不许编数字**。没有数字就量化范围（多少家公司、多少条记录、对比了几种方案）。

## 内容怎么写

- 结构、该删什么、一页压不下时按什么顺序砍 → `references/content-rules.md`
- 动词表、量化模板、中译英改写对照（含机构名/竞赛名标准译法）→ `references/phrasing.md`
- 版式参数（不要随意改）→ `references/layout-spec.md`

## 版式要点

A4，页边距上 1.70 / 下 1.60 / 左右 1.90 cm，正文 Times New Roman 10pt 起（自动往下缩到压进一页），
小节标题 11.5pt 全大写带下划线，日期 | 机构 | 岗位 | 地点 用制表位对齐（3.35 cm / 8.55 cm / 右对齐 17.20 cm）。
完整参数见 `references/layout-spec.md`。

## 命令参数

```
python scripts/build.py <content.yaml> [-o 输出目录] [-n 文件名] [--png]
                                       [--max-pages 1] [--body-pt 9.5]
```

- `--max-pages 1`（默认）：不断缩小正文字号，直到排进一页。压到 9.0pt 以下会提示，那时该砍内容而不是继续缩。
- `--body-pt 9.5`：锁死字号，跳过自动压页（想自己控制时用）。
- `--png`：额外输出首页预览图，改完先看这个再决定要不要动。

输出：`out/<姓氏-名字>-Resume-EN.docx` 和 `.pdf`。

## YAML 结构

```yaml
name: HAIYANG ZHAO
contact: "城市  |  电话  |  邮箱"
availability: "Available: Dec 1, 2026 – Feb 28, 2027  |  4–5 days per week  |  Beijing or Shanghai"

sections:
  - type: paragraph          # 自由段落，用于 PROFILE
    title: PROFILE
    text: >
      ...

  - type: entries            # 日期 | 机构 | 岗位 | 地点 + bullets
    title: EDUCATION
    entries:
      - date: "2023.09 – 2027.06"
        org: Harbin Engineering University
        role: "Big Data Management and Application (Bachelor)"
        location: Harbin, China
        bullets:
          - "..."

  - type: skills             # 分组技能
    title: SPECIAL SKILLS
    groups:
      - label: "Professional Skills"
        items:
          - "..."
```

`type` 只有 `paragraph` / `entries` / `skills` 三种。`entries` 的 `role`、`location`、`bullets` 都可省略；
`date` 太长的行会自动换成两行（岗位移到第二行斜体），不会溢出页面。

## 常见错误

| 症状 | 原因 | 处理 |
|---|---|---|
| 只有一页但字小得难受 | 内容太多 | 按 `content-rules.md` 第 5 节的顺序砍，不要靠缩字号 |
| 第二页只有一两行 | 差一点点 | 砍一句冗余括注通常就够；或 `--max-pages 1` 让它自动缩 0.1pt |
| 日期和机构挤在一起 | 机构名太长 | 正常，会自动换行；想压回一行就缩短 `org` |
| 中文引号/全角标点混进 PDF | 从中文稿复制 | 全文搜 `，。；：""` 替换成半角 |
| Word 里打开是 2 页 | Word 分页与 PDF 渲染微差 | 以 PDF 为准；投递只用 PDF |

## 目录

```
SKILL.md                 本文件
manifest.yaml            按需加载配置
references/              content-rules / phrasing / layout-spec
examples/                可构建的真实样例
scripts/build.py         生成器（YAML → DOCX + PDF）
templates/               原始模板 DOCX（视觉基准）
```
