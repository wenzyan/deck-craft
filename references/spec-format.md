# spec.json 格式

```json
{
  "template": "可选，16:9 母版 pptx 路径",
  "theme": {"accent": "9D2933", "text": "1A1A1A", "grey": "595959", "line": "BFBFBF",
            "title": "1A1A1A", "font_ea": "微软雅黑", "font_latin": "Arial", "size_body": 14, "size_title": 28},
  "slides": [ ...页面... ]
}
```

`theme` 只写要改的键。画布 960×540pt，正文区 y = 92 到 500pt，左右边距 48pt。

## 页面类型

| type | 字段 | 用途 |
|---|---|---|
| `cover` | `title` · `subtitle` · `meta`（行列表） | 封面 |
| `section` | `no`（如 "01"）· `title` | 章节过渡页，强调色满底 |
| `content`（缺省） | `title` · `blocks` · `notes` | 内容页，自动加页码 |
| `end` | `title` | 致谢页 |

所有页面都可带 `notes`（讲稿）。

## 版式块

`blocks` 是从上到下的列表，每项是只有一个键的对象。

| 块 | 写法 | 适合 |
|---|---|---|
| `banner` | `"一句话结论"` | 页面核心结论、过渡句，每页最多一条 |
| `cards` | `[{"head": "表头", "lines": ["行1", "行2"]}, ...]` | 2 到 4 个并列要点；各卡行数要相等 |
| `grid` | `{"rows": ["槽位1", ...], "cols": [{"head": "列名", "cells": ["...", ...]}, ...]}` | 对比矩阵；每列 `cells` 与 `rows` 等长 |
| `bignums` | `[{"num": "3.13 亿", "label": "说明"}, ...]` | 2 到 4 个关键数字，须同量纲 |
| `chain` | `["步骤1", "步骤2", ...]` | 流程、技术路线，3 到 6 步 |
| `figure` | `{"path": "图片路径", "caption": "图注", "h": 200}` | 论文图、照片；没有图时省略 `path`、写 `todo`，生成虚线占位框 |
| `bullets` | `["要点1", "要点2"]` | 少量补充要点；能用 cards/grid 的不要用它 |
| `note` | `"脚注或数据来源"` | 页面底部小字 |
| `split` | `{"ratio": 0.55, "left": [块...], "right": [块...]}` | 左右分栏，常见于图在左、要点在右 |

## 选块经验

- 先问这一页的中心是**一个结论**、**几个并列点**、**一组对比**、**一个流程**还是**一张图**，再分别选 `banner`、`cards`、`grid`、`chain`、`figure`。
- 典型组合：`banner` + `cards`；`chain` + `split(figure, bullets)`；`grid` + `banner`；`bignums` + `cards`。
- 同一组 `cards`、`grid` 的文字由引擎统一取字号（最长一行能单行放下）。字号被压得太小，就是字太多，删字而不是换块。
