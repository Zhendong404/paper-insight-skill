# Paper Insight Skill · 单页论文洞察 PPT

把一篇论文转成**一页中文技术洞察 PPT**的 skill：红黑灰、三栏（问题 / 方案 / 收益）、
结论式标题 + 洞察/启示 + 关键技术 + 来源，证据写进备注。适合组会速览、会议跟踪、技术规划评审。

模板是作者按参考样例手工复刻的版式，不是任何厂商的官方规范；公开版已移除第三方品牌标识与
密级字样，四个示例配图也换成了中性占位图。

## 内容

| 路径 | 说明 |
|---|---|
| `SKILL.md` | skill 主文件：内容组织方式、文案约束、视觉规则、编辑锚点、DSH/Codex 用法 |
| `assets/reference.pptx` | 单页参考模板（只读；复制后原位编辑），已中性化品牌标识 |
| `assets/preview.png`、`assets/reference-render.pdf` | 模板渲染预览 |
| `references/style-spec.json` | 实测几何、颜色、字号、对象 ID、段落属性与母版映射 |
| `references/paper-to-insight.md` | 原文事实与洞察/启示的边界说明 |
| `references/dsh-validation.md` | DSH 适配与像素级验证记录 |
| `scripts/build-report.py` | DSH 环境生成器：复制模板、替换 8 处文本槽与 4 张配图、写备注 |
| `scripts/crop-figure.py` | 从论文 PDF 渲染页面并裁图（自动去白边） |
| `scripts/fit-figures.py` | 把四张配图放大到栏内剩余空间（保持纵横比） |
| `scripts/merge-decks.py` | 把多篇单页 deck 合并为一个多页 PPT |
| `scripts/build-report.mjs` | Codex 版生成器（依赖 `@oai/artifact-tool`，可选） |
| `examples/aiperf-report.json` | 一份 report 输入示例（仅内容稿） |

## 依赖

- Python 3.10+，`python-pptx`、`Pillow`、`lxml`（DSH 会话里用 `load_workspace_dependencies` 提供的解释器）。
- 渲染/PDF 导出可选：DSH 自带 LibreOffice Kit（`@deepseek-ai/libreoffice-kit`）或本机 LibreOffice。

## 快速开始

```bash
PY=python3   # 换成带 python-pptx/Pillow 的解释器
SKILL=$(pwd)

# 1) 写 report.json（字段见 examples/aiperf-report.json 与 SKILL.md）
# 2) 生成单页
"$PY" scripts/build-report.py --input report.json --output out/draft.pptx
# 3) 若配图偏小，放大到栏内空白
"$PY" scripts/fit-figures.py --input out/draft.pptx --output out/draft-fit.pptx
# 4) 裁图（需要 PDF 与渲染器）
"$PY" scripts/crop-figure.py --paper paper.pdf --page 5 --box 0.49,0.09,1.0,0.372 --out figs/fig4.png --trim
# 5) 多篇合并
"$PY" scripts/merge-decks.py --output combined.pptx 01.pptx 02.pptx
```

`report.json` 的 `figures` 四个键（context/method/evaluation/result）必须显式给对象或 `null`；
`null` 表示删除该图位。每个对象包含 `path`、`source`、`alt`，可选 `box_in`（页面绝对英寸）。

## 版权与声明

- 本仓库**不是**任何厂商的官方项目；版式为作者手工复刻的红黑灰技术洞察样式。
- 模板中的第三方品牌标识、密级与版权字样已移除；**请勿**在未授权的情况下把厂商商标加回公开产物。
- `examples/` 只保留作者撰写的内容稿；原示例论文的 PDF 与由其裁出的截图因版权不在仓库内。
- `assets/reference.pptx` 的四个示例配图已替换为中性占位图，使用时替换为你自己的论文配图。
- 代码与文档以 MIT 许可发布（见 `LICENSE`）；模板版式与示例内容归作者所有，引用的第三方事实与
  商标归各自权利人。

## English summary

A skill that turns one paper into a single-page Chinese insight slide (red/black/grey, three columns:
problem, method, results) with evidence in the speaker notes. It includes a read-only template, a
python-pptx builder, a figure cropper, a figure fitter and a multi-deck merger. The public template has
vendor brand marks removed, and paper-specific example figures are replaced with neutral placeholders.
