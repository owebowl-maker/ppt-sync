# ppt-sync — PPT 同类页匹配与数据自动更新

解决"同一指标散落在多份 PPT、每季度要逐页手改"的问题：

1. **阶段 A**：扫描目录下所有 PPT，自动找出跨文件的同类页（如 PPT-1 第 6 页 ↔ PPT-3 第 13 页），输出对照表 Excel，并标出数值不一致的页面。
2. **阶段 B**：人工确认对照表后，把指标标签写进幻灯片本身，之后关联页直接精确查找。
3. **阶段 C**：以一张 Excel 为单一数据源，一条命令把最新数据写入所有打标页面，生成新版 PPT。

**不修改任何原始 PPT；所有输出写新文件；全部本地运行，不上传数据。**

## 安装

```bash
cd 项目06-
uv venv --python 3.12 .venv          # 系统仅 3.9，用 uv 建 3.12 环境
uv pip install --python .venv/bin/python -e "ppt-sync[dev]"
source .venv/bin/activate
ppt-sync --help
```

**方向2/3 的两个前置条件**（缺了会自动降级，不影响其余功能）：

1. **LibreOffice**（渲染兜底）：`soffice` 在 PATH、macOS `/Applications/LibreOffice.app`、Windows `C:\Program Files\LibreOffice` 下均可自动探测；其他位置在 `config.yaml` 的 `render_visual.soffice_path` 指定。
2. **CLIP 视觉编码器模型**：`ppt-sync/models/vision_model_quantized.onnx`（约 85MB，来自 Xenova/clip-vit-base-patch32 量化版）。**拷贝工具到别的电脑时必须连 `models/` 一起拷。** 缺模型时整页比对降级为 pHash。

Windows 电脑的完整安装与使用步骤见项目根目录《另一台电脑使用说明.md》。

## 快速上手

```bash
# 1. 生成索引 + 匹配（阶段 A）
ppt-sync index  <pptx目录> --out work/index.json
ppt-sync match  --index work/index.json --out work
#    → work/matches.xlsx：groups / pairs / numbers / confirm 四个 Sheet
#    → 人工在 confirm Sheet 填：确认(Y/N)=Y、指标key(英文)、指标名称

# 2. 打标（阶段 B）
ppt-sync tag <pptx目录> --index work/index.json \
    --mapping matches_confirmed.xlsx --out tagged_dir/
#    → tagged_dir/ 下的 PPT 带标签；registry.xlsx 是指标登记表

# 3. 查询某页的同类页（打标后精确匹配，未打标回退相似度）
ppt-sync find tagged_dir/PPT-1.pptx 6

# 4. 数据更新（阶段 C）
python scripts/make_data_template.py data.xlsx    # 生成数据源模板
#    → 每季度只需改 data.xlsx 里的数值
ppt-sync render tagged_dir/ --data data.xlsx --out new_dir/ --period 2026Q3
#    → new_dir/PPT-1_2026Q3.pptx 等；render_log.xlsx 记录每处旧值→新值

# 5. 体检：各文件同一指标的数值是否一致（cron 里用，不一致退出码 1）
ppt-sync check tagged_dir/
```

## 数据源 data.xlsx 格式

Sheet `metrics`：

| metric_key | 指标名称 | 值 | 单位 | 显示格式 | 更新周期 | 负责人 | 数据来源 |
|---|---|---|---|---|---|---|---|
| case_count | 合作案例数 | 128 | 个 | `{:,.0f}` | 季度 | … | … |

图表数据：另建 Sheet，名为 `chart_<metric_key>`（Excel 不允许冒号，故用下划线），第一列是类别，后面各列是系列。

## 标签机制（阶段 B 决策说明）

- **页面级**：写入 slide 的 `p:cSld/@name`（python-pptx 的 `slide.name`），格式 `ppt-sync:group=<gid>;metric=<key1,key2>`。一页多指标用逗号合并。
- **形状级**：承载指标的文本框/表格/图表，`shape.name` 改为 `原名|metric:<key>`。PowerPoint 渲染用 shape id 引用形状，改名不影响外观；唯一注意点是**如果公司有宏按形状名引用，会受影响**（罕见）。
- 不使用 `p:custDataLst` custom tags：需要手工拼 OPC 部件，有触发 PowerPoint 修复提示的风险，收益为零。
- 标签基于 `slide_id` 定位，插页删页后不失效。
- 打标前后 slide XML 逐属性对比验证：只有上述两处 name 变化，外观零改变。

## 配置（config.yaml）

权重（s_img 0.30 / s_chart 0.20 / s_text 0.25 / s_layout 0.10 / s_visual 0.15）、候选阈值 0.55、高置信 0.75、组内凝聚度 0.45、全局去噪比例 0.30、通用页关键词等都在 `ppt-sync/config.yaml`。**用真实 PPT 校准后请把调参记录补在这里。** 某个子分缺失（如两页都无图表）时其权重自动按比例摊给其余项。

## 已知限制

- **渲染兜底与 CV 比对（方向2/3）**：`index` 时自动把每页渲染成 PNG（`work/renders/`），有 CLIP 模型时用视觉编码器算 512 维嵌入（方向3），否则退到整页 pHash；两者都没有（无 LibreOffice）时只用结构特征。CLIP 相似度 ≥0.80 的页面对会绕过文本/图片预筛进入精算（`match.visual_prescreen` 可调）。**真实 PPT 上的 CLIP 阈值需要校准**：同模板不同页的余弦可能偏高，校准时重点看 `pairs` 表的 `s_visual` 列。
- SmartArt 只读文字（直接解析 diagram 数据部件），未在样例中验证，遇到问题以渲染兜底。
- GroupShape 作为整体参与版式打分，不展开子形状坐标（子形状文字仍全部收集）。
- `.ppt` 旧格式请先用 PowerPoint/LibreOffice 另存为 `.pptx`。
- 数字识别支持 `1.2亿 / 12,000 / 35% / 50+ / 128家` 等格式；"近百家"这类非数字表述会列进报告但无法自动更新。

## 定时运行（可选）

cron（Mac/Linux，每周一 07:03）：

```cron
3 7 * * 1 cd /path/to/ppt-sync && /path/to/.venv/bin/ppt-sync render tagged_dir/ --data data.xlsx --out out_$(date +\%Y\%m\%d)/ --period $(date +\%Yq\%-m) && /path/to/.venv/bin/ppt-sync check out_$(date +\%Y\%m\%d)/
```

Windows 计划任务：任务计划程序 → 创建基本任务 → 每季度 → 操作填
`<venv>\Scripts\ppt-sync.exe render tagged_dir\ --data data.xlsx --out out\ --period 2026Q3`。

默认只生成新文件到新目录，**不自动覆盖共享盘上的正式版本**——确认无误后人工替换。

## 测试

```bash
python -m pytest tests/ -q      # 36 个用例，全部本地合成样例，秒级完成
```

覆盖：数字抽取（1.2亿/12,000/35%/50+/128家）、组合形状递归、图表签名、全局去噪、防链式误连聚类、已知对应关系端到端、打标 XML 零外观变化、跨 run 数字替换、图表 replace_data、缺 key 报错退出、find 3 秒内响应。

## 开发/校准流程

1. `python scripts/make_samples.py` 生成 3 份带已知对应关系的样例（对应关系写在 `tests/conftest.py` 的 `KNOWN_PAIRS`）。
2. 换上真实 PPT 后：跑 `index` + `match` → 抽查 20 个候选对（准确率目标 ≥80%）→ 调 `config.yaml` 权重/阈值 → 重跑，直到满意。
