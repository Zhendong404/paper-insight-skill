> 公开仓库版说明：`examples/aiperf-paper.pdf`、由论文裁出的示例图与 `examples/aiperf-validated.pptx` 因版权不在仓库内；
> `assets/reference.pptx` 中的品牌标识与密级字样已移除，四个示例配图替换为中性占位图；下文为本地验证记录。

# DSH 适配与验证记录

本 skill 原包使用 Codex 的 `@oai/artifact-tool`（`scripts/build-report.mjs`）。DSH 会话里改用本目录的
`scripts/build-report.py`（python-pptx）与 `scripts/crop-figure.py`（PDF 渲染裁图），参考文件
`assets/reference.pptx` 仍只读，产物结构与视觉基准保持一致。

## 环境

- Python：`load_workspace_dependencies` 返回的 Python，含 python-pptx 1.0.2、Pillow 12.3.0、lxml。
- 渲染：DSH 自带 LibreOffice Kit（`@deepseek-ai/libreoffice-kit`：PDF 用 pdfium，PPTX 用 LibreOffice）。
- 视觉基准：`assets/preview.png` 是 PowerPoint 原生导出的参考页；LibreOffice 与 PowerPoint 在换行和字宽上
  存在差异，溢出判断以 PowerPoint 为准。

## 等价性验证（AIPerf 样例）

```text
<python> scripts/build-report.py --input examples/aiperf-report.json \
    --output work/draft.pptx --skill-dir <skill 根目录>
```

- 对照物：`examples/aiperf-validated.pptx`（Codex artifact-tool 生成的已验证产物）。
- 方法：两份 PPTX 都用 LibreOffice Kit 渲染为 1921×1081 PNG，逐像素比较。
- 结果：任一通道差 >32 的像素 **0 个（0.0000%）**；>16 的像素 5 个；最大通道差 17（抗锯齿）。
- 结构：单页；画布 12196763×6858000 EMU；母版/版式/主题与参考一致；图片对象 ID、位置、尺寸不变。

## 真实论文端到端验证

- 论文：MICRO 2026, *Perf-IR: A Portable Intermediate Representation for Cross-Platform GPU
  Performance Prediction*（DOI 10.1109/MICRO67795.2026.00083）。
- 流程：PDF → `crop-figure.py` 渲染页面并裁出 Fig.1/4/7/8 → 写 report.json → `build-report.py`
  → LibreOffice 渲染预览 → 与 `assets/preview.png` 对照检查层级、比例、密度、图文位置与页脚。
- 校验结果：单页；8 处文本槽全部替换；备注 841 字符（含页码/图表定位、量化条件与局限）；4 张图按原始
  纵横比 contain 放入，未拉伸、未裁掉内容；内容预算全部落在参考范围内（标题 40、洞察 56、启示 56、
  问题 109、方案 141、主收益 41、次收益 40、关键技术 133）。
- 产物见工作区 `insight-report/dsh-validation/`。

## 与 Codex 产物的已知差异

- 组合形状（`grpSp`，洞察/启示带状区域）保持原样，artifact-tool 导出会把组合展平；渲染一致。
- 文本框沿用模板高度（`spAutoFit`），artifact-tool 会把高度归一到内容；视觉无差异。
- Python 版用 python-pptx 保存，没有执行 PowerPoint 重新导出；交付前建议在 PowerPoint 打开确认一次。

## 边界

- `build-report.py` 只编辑参考幻灯片里的既有对象，不新建幻灯片或版式，不替代读论文、事实核对与视觉验收。
- `crop-figure.py` 默认调用 DSH 自带渲染器；也可用 `--page-image` 对外部已渲染的页面 PNG 裁图。
- `scripts/build-report.mjs` 保留给 Codex 环境，两条路径的 report.json 结构相同。
