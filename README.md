# deck-craft

一个面向 Claude Code / Codex 的 skill：**从零做一份学术或工作汇报 PPT**，也能接着改已有的 pptx。适用于组会、学术报告、开题/中期/结题答辩、基金会评、项目汇报、课程讲解。

它把一套经过多轮真实答辩、会评改稿验证的做法固化下来：

1. **定需求**：场景、听众、时长、材料、模板
2. **叙事骨架**：按场景选骨架，一条主线贯穿全篇
3. **逐页故事板**：每页一句结论 + 视觉主体 + 版式块 + 素材来源
4. **（可选）版式样图**：Codex 或 gpt-image-2 先出构图参考
5. **一键搭建**：JSON 页面描述 → `deck.py` → 原生可编辑 pptx，结构上不重叠，同组同字号
6. **核验**：越界、重叠、溢出检查 + 逐页渲染 PNG 目视检查
7. **迭代改稿**：用户手改优先，主文件只读，文本锚点定位

核心口径：每个数字都能在材料里找到出处；一页一个中心；同层元素同维度、同句式、同字数；白底黑字加一种强调色。

## 文件

| 文件 | 作用 |
|---|---|
| `SKILL.md` | 完整流程与内容、视觉、讲稿、核验约束 |
| `references/storylines.md` | 学术报告、答辩、会评、工作汇报、科普的叙事骨架 |
| `references/spec-format.md` | spec.json 页面类型与版式块说明 |
| `references/codex-mockup.md` | 可选的样图生成与提示词模板 |
| `references/pptx-pitfalls.md` | python-pptx / OOXML 踩坑记录 |
| `scripts/deck.py` | 版式引擎：spec.json → pptx |
| `scripts/example_spec.json` | 完整示例 |
| `scripts/check_deck.py` | 越界、重叠、文字溢出检查 |
| `scripts/render.py` | 逐页导出 PNG（PowerPoint COM，或 LibreOffice + PyMuPDF） |
| `scripts/dump_deck.py` | 逐页导出现有 pptx 的文本与几何 |
| `scripts/slidecopy.py` | 跨文件整页复制、文本置换、安全删页 |
| `scripts/docstate.py` | 检测用户在两轮之间的手改 |

## 安装

```bash
# Claude Code
git clone https://github.com/wenzyan/deck-craft ~/.claude/skills/deck-craft
# Codex
git clone https://github.com/wenzyan/deck-craft ~/.codex/skills/deck-craft
```

依赖：`pip install python-pptx`。渲染在 Windows 上需要 `pywin32` 和本机 PowerPoint，其他平台需要 LibreOffice 和 `pymupdf`。

## 快速试用

```bash
python scripts/deck.py scripts/example_spec.json demo.pptx
python scripts/check_deck.py demo.pptx
python scripts/render.py demo.pptx png/
```

在对话里直接说“根据这篇论文做一份 15 分钟组会 PPT”“按申报书做答辩 PPT”“帮我检查这份 PPT 的层级对称”，并给出材料路径即可。
