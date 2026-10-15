# python-pptx / OOXML 陷阱

全部是真实改稿中踩过的坑。每条写清现象和正确做法。

## 形状与对象

- **每次遍历 `slide.shapes` 都会新建代理对象。** 用 `is` / `is not` 比较对象身份必然失败，曾因此把标题占位符一并删掉。比较时用 `sh._element`。
- **判断占位符用 `sh.is_placeholder`**，`shape_type` 不可靠。
- **复制形状会沿用源 `cNvPr/@id`**，导致同页重复 shape id。PowerPoint 打开时会静默修复，但属于格式错误。复制后要重编 id（`slidecopy.renumber_ids`）。

## 组合（GROUP）坐标

- 组内子形状的 `left/top` 是**组内坐标**，不是页面绝对坐标。要换算到页面：

  ```
  page_x = grp.off.x + (child.x - grp.chOff.x) * grp.ext.cx / grp.chExt.cx
  ```

  嵌套组合要递归套用。直接拿组坐标当绝对坐标，下划线、标注会落到错误位置。
- 量页面密度、查重叠时，组合只取**自身包围盒**，不要递归进子形状按绝对坐标累计，否则整组会被算飞，得出“这页很空”的错误结论。
- 缩放组合：改 `a:xfrm/a:ext`，保留 `a:chExt`，子形状会自动跟着缩放。

## 删页

- 只从 `sldIdLst` 摘掉条目**不够**，presentation 到 slide 的关系还在，孤儿部件留在包里，保存时报 `Duplicate name: ppt/slides/slideN.xml`。正确做法：先 `prs.part.drop_rel(sldId.rId)`，再从 `sldIdLst` 移除（`slidecopy.delete_slide`）。
- 删页、并页之后核对：幻灯片部件数 = `sldIdLst` 条目数，`zipfile.testzip()` 通过。

## 文字溢出估算的误报来源

`check_deck.py` 是按“字宽 × 字号 / 可用宽度”估算行数，以下情况会误报，**报出来也不要改**：

1. 竖排文本框（`bodyPr/@vert`）；
2. `word_wrap=False`，或 `autosize = SHAPE_TO_FIT_TEXT`，PowerPoint 会自己扩展；
3. 段内混排大小字号（48pt 数字 + 20pt 文字）时，用最大字号乘全段；脚本已改为逐 run 累计；
4. 字号继承自母版、未显式设置时，按默认 18pt 估。

最终以渲染图为准。

## 页码

- 用 `slidenum` 域的文本框，插页删页会自动重编；不要写死数字。
- 页码颜色按实际背景取：渲染后取页码位置一小块区域的中位亮度，小于 150 判为深色底用白字，其余用灰字。

## 图片

- 换图时**重指 `r:embed`** 原地替换，位置、尺寸、层级都不变；不要删了重插。
- 整页位图（例如 AI 生成的整页图）无法局部修改。需要改时按原图 1:1 重建成原生形状：EMU 与像素的换算比例、`srcRect` 裁切偏移都要反解，重建后与原渲染逐像素套合核对。

## 手动断行

- 自然换行会把 `Q1).` 这种短尾甩成孤字行时，在合适位置插入 `<a:br/>`，让整块一起换到下一行。

## 文件占用

- 渲染时 PowerPoint COM 会打开文件。用户的 PowerPoint 里如果同时开着这份 pptx，之后从 PowerPoint 另存会覆盖脚本写入的改动。
