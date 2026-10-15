# 可选：用 Codex / gpt-image 生成版式样图

**用途**：某页版式拿不准时，先生成 1 到 3 张样图，和用户确认构图后再用 python-pptx 做原生页面。

**红线**：
- 样图只定构图，**不进成稿**。整页位图无法局部修改，评审追问时也改不动。
- 样图里的文字一律用英文占位。中文和数字以依据文件为准，由原生文本框写入。
- 不用于结果图表。结果图只取已发表论文或正式海报里的图。

## 通道 A：Codex CLI 内置 `image_gen`

前提：`codex features list` 里 `image_generation` 为 `true`，并且用 ChatGPT 账号登录官方 provider。
**使用自定义 `model_provider` 时，内置工具可能不暴露**（实测 codex-cli 0.154 + 自定义 provider 时返回“内置 image_gen 工具当前不可用”），这种情况改走通道 B。

```powershell
codex exec --skip-git-repo-check -s workspace-write -C <输出目录> "<下方提示词>"
```

```bash
codex exec --skip-git-repo-check -s workspace-write -C <outdir> "<prompt below>"
```

生成的图默认存在 `$CODEX_HOME/generated_images/`，所以提示词末尾要求它复制到当前目录。

## 通道 B：gpt-image-2 API 脚本

用 [gpt-image-gen](https://github.com/wenzyan/gpt-image-gen) 的 `gen_figure.py`，需要 `OPENAI_API_KEY`（可选 `OPENAI_BASE_URL`）：

```bash
python gen_figure.py --prompt mockup_p05.txt --size 1536x1024 --out mockups/p05.png
```

## 提示词模板

英文提示词的效果明显好于中文。每页一张，横版 1536x1024。

```
Use the imagegen skill with the built-in image_gen tool.   <- 仅通道 A 需要这一行
Generate ONE 16:9 slide layout mockup (1536x1024) for an academic grant-review presentation.
It is a layout reference only, not final content.

Slide purpose (one sentence): <这一页要让评审记住的结论，英文>
Layout: <例：title top-left; one conclusion banner under the title;
         three equal cards side by side, each with header, a simple line icon, three short lines;
         a thin footnote band at the bottom>
Style: white background, black text, ONE accent color <#9D2933> used only for emphasis
       (banner outline, header underline, numbers); flat vector; thin lines; generous margins;
       consistent line-icon style; no photos, no gradients, no 3D, no drop shadows.
Text: short English placeholders only; keep sibling elements the same length.

Copy the final PNG to ./mockup_p<NN>.png in the current directory and reply with the saved path only.
```

## 用样图做什么

1. 和用户确认构图：几栏、视觉主体放哪、留白多少。
2. 量出样图中的栏宽比例、边距、卡片间距，换算成 pt 后用原生形状实现。
3. 检查实现后的页面是否仍满足 SKILL.md 第 2 节的层级对称约束；样图常常为了好看牺牲对称，以约束为准。
