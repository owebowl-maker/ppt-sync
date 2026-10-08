# PPT 同类页匹配与数据自动更新：实施方案（供执行 Agent 使用）

> 本文档是给执行 Agent 的任务说明书。请按阶段顺序实现，每个阶段完成后对照"验收标准"自检，再进入下一阶段。
> 背景调研见同目录 `PPT数据自动更新方案调研.md`。

---

## 0. 背景与目标

**现状**
- 公司有多份 PPT，每季度需要更新销售额、合作厂家、业绩、合作案例数、时间等数据。
- 同一指标散落在多个 PPT 中，比如"合作案例数"同时出现在 PPT-1 第 6 页和 PPT-3 第 13 页，这些页面外观基本一样。
- 页面特点：文字少，多为数字、图片（Logo、案例照片）、图表；排版自由，没有固定的阅读顺序。

**总目标（分三阶段）**
1. **阶段 A：同类页匹配**。扫描所有 PPT，自动找出跨文件的同类页，输出对照表，并标出数值不一致的页面。
2. **阶段 B：打标固化**。人工确认对照表后，把指标标签写入幻灯片本身，以后关联页面直接精确查找，不再依赖相似度。
3. **阶段 C：模板化自动更新**。用一张 Excel 作为单一数据源，脚本把最新数据写入所有打过标的页面，生成新版 PPT。

**核心技术决策（已确定，执行时不要更改）**
- 以 **PPTX 结构特征** 为主（图片哈希、图表元数据、文本关键词集合、版式），**渲染图像相似度** 为兜底。不用 OCR 读页面文字。
- 所有特征都按 **无序集合** 比较，不依赖阅读顺序。
- 不修改原始 PPT；所有写操作都输出到新文件。

---

## 1. 技术栈与环境

| 用途 | 选型 |
|---|---|
| 语言 | Python 3.10+ |
| 读写 PPTX | `python-pptx`；必要时用 `lxml` 直接处理 XML |
| 渲染页面为图片 | LibreOffice headless：`soffice --headless --convert-to pdf`，再用 `pdf2image` 或 `pymupdf` 把 PDF 转成每页 PNG |
| 感知哈希 | `imagehash`（pHash + dHash） |
| 图像向量（可选） | `open_clip_torch` 或 DINOv2（`transformers`）；没有 GPU 时可以只用 pHash |
| 文本相似 | `rapidfuzz`（关键词集合 Jaccard / token_set_ratio）；可选 `sentence-transformers` 中文模型（如 `BAAI/bge-small-zh-v1.5`） |
| 数据与输出 | `pandas`、`openpyxl` |
| CLI | `typer` 或 `argparse` |

开始前先验证环境：`soffice` 能否真正转换一份样例 PPT（`command -v` 能找到不代表能用），以及所有库能否 import。LibreOffice 不可用时，阶段 A 降级为纯结构特征，并在报告里注明。

---

## 2. 目录结构（建议）

```
ppt-sync/
  ppt_sync/
    extract.py      # 解析 pptx → 每页特征
    render.py       # 渲染每页为 PNG
    features.py     # 文本归一化、数字抽取、哈希
    match.py        # 两两打分、聚类
    report.py       # 输出对照表 Excel
    tag.py          # 阶段 B：写入标签
    render_data.py  # 阶段 C：按数据源更新 PPT
    cli.py
  tests/
  config.yaml       # 权重、阈值、停用词、指标关键词
  README.md
```

CLI 命令：

```
ppt-sync index   <pptx_dir> --out work/index.parquet     # 阶段 A-1
ppt-sync match   work/index.parquet --out work/matches.xlsx   # 阶段 A-2
ppt-sync tag     <pptx_dir> --mapping work/matches_confirmed.xlsx --out <out_dir>   # 阶段 B
ppt-sync find    <pptx_file> <slide_no>                  # 查询某页的所有同类页
ppt-sync render  <pptx_dir> --data data.xlsx --out <out_dir>   # 阶段 C
ppt-sync check   <pptx_dir>                              # 跨文件数值一致性体检
```

---

## 3. 阶段 A：同类页匹配

### A-1 每页特征抽取（`extract.py`）

为每个 `(文件, 页码)` 生成一条记录。递归遍历所有形状，包括组合形状（GroupShape）内部、表格单元格、图表、占位符；还要读取页面所用版式（layout）的名称。

| 字段 | 说明 |
|---|---|
| `deck`, `slide_no`, `slide_id` | 文件名、页码（从 1 开始）、`slide.slide_id`（稳定 ID） |
| `layout_name` | 所用版式名称 |
| `image_hashes` | 页面上每张图片 blob 的 SHA1 集合（`shape.image.sha1`），**排除母版/版式上的图片** |
| `image_phashes` | 每张图片的 pHash，用于识别"同一 Logo 但被重新压缩"的情况 |
| `chart_sigs` | 每个图表的签名：`(chart_type, 系列名集合, 类别标签集合)`；类别标签中的年份/季度要归一化，如 `2026Q2` → `<PERIOD>` |
| `text_tokens` | 页面所有文字（含表格、图表标题），数字替换为 `<NUM>`，分词后去掉停用词，得到关键词**集合** |
| `numbers` | 抽取出的数字及其上下文，格式为 `[{value, unit, context_tokens, shape_id}]`；context 取同一文本框内或几何上最近的标签文字 |
| `shape_layout` | 每个形状的 `(类型, 归一化中心 x/y, 归一化宽高)` 列表，坐标除以页面尺寸 |
| `render_png` | 渲染后 PNG 的路径（见 A-1b） |
| `img_phash`, `img_embed` | 整页 pHash；可选的整页图像向量 |
| `is_boilerplate` | 是否是封面、目录、致谢等通用页（规则：形状少于 3 个，或标题命中配置里的通用页关键词） |

**A-1b 渲染**：把每份 PPT 用 LibreOffice 转成 PDF，再按页转 PNG（宽 640px 即可）。渲染出来的页数必须与 python-pptx 读到的页数一致，否则把该文件标记为渲染失败，该文件只使用结构特征。

**全局去噪（必须做）**：统计每个图片哈希、每个文本 token 在所有页面中出现的比例。出现在超过 30% 页面上的（例如公司 Logo、页脚、"机密"字样）视为模板噪声，从特征中剔除。阈值写在 `config.yaml`。

### A-2 打分与聚类（`match.py`）

对**不同文件**之间的页面两两计算相似度（同一文件内的页面也计算，但在报告中单独标注）：

| 子分 | 计算方法 | 默认权重 |
|---|---|---|
| `s_img` | `image_hashes` 的 Jaccard；若有 pHash 汉明距离 ≤ 6 的图片，也算作相同 | 0.30 |
| `s_chart` | 图表签名匹配：类型相同且系列名/类别标签的 Jaccard 取平均 | 0.20 |
| `s_text` | `text_tokens` 的 Jaccard（或 rapidfuzz token_set_ratio / 100） | 0.25 |
| `s_layout` | 形状布局相似度：按类型配对，用中心点距离的匈牙利匹配，分数为 1 − 平均距离 | 0.10 |
| `s_visual` | 整页 pHash 相似度（1 − 汉明距离 / 64）；有向量时取余弦相似度 | 0.15 |

- **缺失处理**：某个子分两页都没有对应元素时（比如都没有图表），该项不参与计算，其权重按比例分给其余项。
- **总分** `score = Σ w_i · s_i`。
- **候选阈值**：`score ≥ 0.55` 进入候选，`≥ 0.75` 标记为"高置信"。权重和阈值都放在 `config.yaml`，后续用样例校准。
- **性能**：页数 N 较大（> 2000）时，先用 `text_tokens ∪ image_hashes` 做 MinHash-LSH（`datasketch`）粗筛候选对，再精算。
- **聚类**：以候选对为边建图，取连通分量作为"同类页组"；组内任意两页的分数都要 ≥ 0.45，否则拆开，避免链式误连。
- 跳过 `is_boilerplate=True` 的页面，除非两页的分数 ≥ 0.9。

### A-3 数值一致性检查

在每个同类页组内，按 `numbers.context_tokens` 把数字配对（上下文关键词的 Jaccard ≥ 0.5 即视为同一指标），**同一指标数值不同就标记为不一致**。这正是"某页更新了、另一页忘记更新"的信号。

### A-4 输出对照表（`report.py`）

输出 `matches.xlsx`，包含以下 Sheet：

1. **groups**：`group_id, 建议指标名, 成员(deck#页码 列表), 组内最低分, 置信度, 是否有数值不一致`
2. **pairs**：`deck_a, slide_a, deck_b, slide_b, score, s_img, s_chart, s_text, s_layout, s_visual, 主要依据`（例如"共享 3 张图片 + 图表系列名相同"）
3. **numbers**：`group_id, deck, slide, 指标上下文, 数值`；不一致的行标红
4. **confirm**：留给人工确认的空列 `确认(Y/N), 指标key(英文), 指标名称`

另外每组生成一张缩略图拼图（组内各页 PNG 横排），存到 `work/groups/<group_id>.png`，并在 groups 表里放上路径，方便人工快速核对。

### A-5 验收标准

- 用 Domi 提供的 2～3 份样例 PPT 运行，已知的同类页（如 PPT-1 第 6 页 ↔ PPT-3 第 13 页）必须出现在 groups 中，且置信度为"高"。
- 人工抽查 20 个候选对，准确率 ≥ 80%；不达标就调整权重/阈值并记录在 README。
- `ppt-sync find PPT-1.pptx 6` 能在 3 秒内列出同类页及各页当前数值。
- 每个模块有单元测试：至少覆盖组合形状遍历、数字抽取（`1.2亿`、`3,200`、`35%`、`128家` 等格式）、全局去噪。

---

## 4. 阶段 B：打标固化

**输入**：人工在 `confirm` 表中填好的 `matches_confirmed.xlsx`。

**做法**
1. 对每个已确认的组，给组内每一页写入标签。
   - 页面级标签：写在 slide 的 `name` 属性（`p:cSld/@name`）中，格式 `ppt-sync:group=<group_id>;metric=<metric_key>`。也可以同时写入 slide 的 custom XML / tags 部件（`p:custDataLst`），二选一即可，在 README 里说明选择。
   - 形状级标签：把承载指标数值的文本框/表格单元格/图表，其 `name` 改为 `metric:<metric_key>`；保留原名称时用 `原名|metric:<key>`。这一步是阶段 C 精确替换的基础。
2. 输出到新目录，绝不覆盖原文件。
3. 生成 `registry.xlsx`，列为 `metric_key, 指标名称, deck, slide_id, 当前页码, shape_name, 当前数值`。
4. `ppt-sync find` 优先读取标签做精确查找；页面没有标签时才回退到阶段 A 的相似度结果。

**注意**
- 标签必须基于 `slide_id`，不能基于页码（页码在插删页后会变）。
- 打标后用 PowerPoint 能正常打开、外观完全不变；用 python-pptx 重新读取时，标签仍然存在。

**验收**：对打标后的文件运行 `find`，结果与人工确认表完全一致；打标前后各页渲染 PNG 的像素差为 0（或 pHash 差 ≤ 2）。

---

## 5. 阶段 C：单一数据源 + 自动更新

**数据源 `data.xlsx`**（Sheet `metrics`）

| metric_key | 指标名称 | 值 | 单位 | 显示格式 | 更新周期 | 负责人 | 数据来源 |
|---|---|---|---|---|---|---|---|
| case_count | 合作案例数 | 128 | 个 | `{:,}` | 季度 | … | … |
| sales_total | 销售额 | 3.2 | 亿元 | `{:.1f}` | 季度 | … | … |
| report_period | 更新时间 | 2026Q3 | | | 季度 | … | … |

图表数据另建 Sheet，Sheet 名为 `chart:<metric_key>`，第一列为类别，后面各列为系列。

**渲染逻辑（`render_data.py`）**
1. 遍历目录中的 PPT，找出所有带 `metric:` 标签的形状。
2. 文本框/单元格：只替换其中的**数字部分**，保留原文字和格式。实现时要先处理 run 被拆分的问题，即在段落内把各 run 拼接后定位数字，再写回第一个 run 并清空被合并的 run，保持该 run 的字体格式。
3. 图表：用 `chart.replace_data()` 写入 `chart:<key>` 的数据。
4. 新文件命名为 `<原名>_<report_period>.pptx`，输出到新目录。
5. 生成 `render_log.xlsx`：每处替换的旧值 → 新值；数据源里有、但 PPT 中找不到的 key；PPT 里有标签、但数据源缺失的 key（这种情况要报错，不能静默跳过）。

**验收**
- 修改 `data.xlsx` 中 `case_count` 的值后运行 `render`，所有相关页面的数值同步变化，其余内容不变（渲染对比时只有数字区域发生变化）。
- 输出文件能用 PowerPoint 正常打开，不触发"修复"提示。
- 运行 `check` 时，各 PPT 中同一 metric 的数值全部一致。

**定时运行（最后一步，可选）**：提供 Windows 计划任务和 cron 的示例配置；默认只生成文件，不自动覆盖共享盘上的正式版本。

---

## 6. 边界情况清单

- 组合形状嵌套多层；SmartArt（python-pptx 读不到其中文字，需直接解析 `dgm` 部件的 XML，或依赖渲染图兜底）。
- 图片作为形状填充或背景（`blipFill`），而不是 picture 形状；也要纳入图片哈希。
- 数字格式：`1.2亿`、`12,000`、`35%`、`50+`、`近百家`（非数字表述，记为无法解析并在报告中列出）。
- 同一页面有多个指标：一页可以属于多个 metric，形状级标签解决这一问题。
- 隐藏的幻灯片要标注，但仍然参与匹配。
- `.ppt` 旧格式：先用 LibreOffice 转换成 `.pptx`。
- 文件名含中文和空格：所有路径处理使用 `pathlib`。

---

## 7. 执行顺序与交付物

| 步骤 | 交付物 | 需要人工的地方 |
|---|---|---|
| 1. 环境验证 + 读取样例 | 环境检查日志、样例 PPT 概况（页数、图表数、图片数） | 提供样例 PPT |
| 2. 阶段 A | `index.parquet`、`matches.xlsx`、组缩略图、调参记录 | 确认对照表 |
| 3. 阶段 B | 打标后的 PPT、`registry.xlsx`、`find` 命令 | 抽查打标结果 |
| 4. 阶段 C | `data.xlsx` 模板、`render` / `check` 命令、`render_log.xlsx` | 填写真实数据 |
| 5. 文档 | README（安装、命令、配置项说明、已知限制） | — |

**执行约束**
- 不修改、不覆盖任何原始 PPT；所有输出写到新目录。
- 涉及业务数据的 PPT 不上传到任何外部服务。若使用在线模型（向量、LLM），必须先得到 Domi 同意；默认只使用本地模型。
- 每个阶段结束时汇报：做了什么、指标结果（准确率、耗时）、遗留问题。
