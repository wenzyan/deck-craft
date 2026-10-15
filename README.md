# review-deck

面向评审专家的学术汇报 PPT（基金会评、项目答辩、开题、中期、结题）的 Claude Code / Codex skill。

它不负责生成花哨的模板，而是把一套经过多轮会评改稿验证的**流程和口径**固化下来：

- 申报书、论文等依据文件是唯一事实来源，不编造数字
- 一页一个中心，同层元素同维度、同句式、同字数
- 白底黑字加一种强调色，少用色块，结果图只用正式论文图
- 用户手改优先，主文件只读，改动落副本
- 改完必做几何核验和逐页渲染目视检查
- 可选：用 Codex 或 gpt-image-2 生成版式样图，先定构图再做原生页面

## 文件

| 文件 | 作用 |
|---|---|
| `SKILL.md` | 流程、页面内容约束、视觉约束、讲稿、收尾核验 |
| `references/pptx-pitfalls.md` | python-pptx / OOXML 的真实踩坑记录 |
| `references/codex-mockup.md` | 可选的样图生成（Codex 内置 image_gen 或 gpt-image-2 API）与提示词模板 |
| `scripts/dump_deck.py` | 逐页导出文本与几何 |
| `scripts/check_deck.py` | 越界、重叠、文字溢出检查 |
| `scripts/render.py` | 逐页导出 PNG（PowerPoint COM，或 LibreOffice + PyMuPDF） |
| `scripts/slidecopy.py` | 跨文件整页复制、文本置换、安全删页 |
| `scripts/docstate.py` | 记录与比对文件哈希，发现用户在两轮之间的手改 |

## 安装

```bash
# Claude Code
git clone https://github.com/wenzyan/review-deck-skill ~/.claude/skills/review-deck
# Codex
git clone https://github.com/wenzyan/review-deck-skill ~/.codex/skills/review-deck
```

依赖：`pip install python-pptx`；Windows 渲染需要 `pywin32` 和本机 PowerPoint，其他平台需要 LibreOffice 和 `pymupdf`。

## 用法

在对话里直接说“帮我改会评 PPT”“按申报书做答辩 PPT”“检查这份汇报 PPT 的层级对称”，并给出 pptx 和依据文件路径。
