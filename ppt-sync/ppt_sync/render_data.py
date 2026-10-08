"""阶段 C：以 data.xlsx 为单一数据源，更新所有打标页面，输出新版 PPT。

文本替换：只替换数字部分（NUM_RE 的 num 组），保留单位后缀与字体格式。
  - 段落文本按子元素（a:r / a:br / a:fld）拼接构建字符流，跨 run 的数字也能精确替换。
  - 从后往前替换，前面的 span 不失效。
  - 被合并的中间 run 只清空 text，不删元素（保留 rPr 结构）。
图表替换：chart.replace_data() 写入 `chart:<key>` sheet 的数据。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn

from .features import NUM_RE
from .tag import shape_metric_key

A_R = qn("a:r")
A_BR = qn("a:br")
A_FLD = qn("a:fld")
A_T = qn("a:t")


@dataclass
class Replacement:
    deck: str
    slide_id: int
    slide_no: int
    shape_name: str
    metric_key: str
    old_raw: str
    new_raw: str


@dataclass
class RenderLog:
    replacements: list[Replacement] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


# ---------------------------------------------------------------- 段落文本构建

def build_paragraph_text(para) -> tuple[str, list]:
    """返回 (全文本, [(run, 该 run 文本在段落字符流中的起始偏移)])。

    a:br 在字符流中记为 '\\v'（不入 run 表）；a:fld 的文本参与拼接但其内部
    span 上的数字不会被替换（调用方检查）。
    """
    full: list[str] = []
    run_offsets: list[tuple[object, int]] = []
    fld_spans: list[tuple[int, int]] = []
    for child in para._p:
        if child.tag == A_R:
            t_el = child.find(A_T)
            text = t_el.text or "" if t_el is not None else ""
            run_offsets.append((child, sum(len(s) for s in full)))
            full.append(text)
        elif child.tag == A_BR:
            full.append("\v")
        elif child.tag == A_FLD:
            t_el = child.find(A_T)
            text = t_el.text or "" if t_el is not None else ""
            fld_spans.append((sum(len(s) for s in full), sum(len(s) for s in full) + len(text)))
            full.append(text)
    return "".join(full), run_offsets, fld_spans


def _fmt_value(num_text: str, fmt: str | None, value: float | None) -> str:
    """按显示格式生成新数字文本。fmt 为空时尽量沿用原写法（千分位/小数位）。"""
    if fmt:
        try:
            return fmt.format(value)
        except Exception:
            pass
    # 沿用原格式：有千分位则千分位，小数位取原位数
    if "," in num_text:
        decimals = len(num_text.split(".")[1]) if "." in num_text else 0
        return f"{value:,.{decimals}f}"
    if "." in num_text:
        decimals = len(num_text.split(".")[1])
        return f"{value:.{decimals}f}"
    return str(int(value)) if value == int(value) else str(value)


def replace_numbers_in_paragraph(para, new_value: float | None,
                                 fmt: str | None, log_hint: str,
                                 log: RenderLog) -> int:
    """把段落中所有数字替换为 new_value（按 fmt 格式化）。返回替换次数。"""
    full, run_offsets, fld_spans = build_paragraph_text(para)
    if not full:
        return 0

    def run_at(pos: int):
        """覆盖字符位置 pos 的 run 与其在 run 内的偏移；不在任何 run 内返回 (None, 0)。"""
        cur = None
        for run, off in run_offsets:
            rtext = None
            t_el = run.find(A_T)
            rtext = (t_el.text or "") if t_el is not None else ""
            if off <= pos < off + len(rtext):
                return run, pos - off
            cur = (run, off)
        return None, 0

    def in_fld(s: int, e: int) -> bool:
        return any(not (e <= fs or s >= fe) for fs, fe in fld_spans)

    replaced = 0
    for m in reversed(list(NUM_RE.finditer(full))):
        s, e = m.start("num"), m.end("num")
        old = m.group("num")
        if in_fld(s, e):
            log.warnings.append(f"{log_hint}: 数字 {old} 位于字段(fld)中，跳过")
            continue
        new = _fmt_value(old, fmt, new_value)
        if new == old:
            continue
        run_s, off_s = run_at(s)
        run_e, off_e = run_at(e - 1)
        if run_s is None or run_e is None:
            log.warnings.append(f"{log_hint}: 数字 {old} 跨越换行/字段，无法安全替换")
            continue
        if run_s is run_e:
            t_el = run_s.find(A_T)
            rtext = t_el.text or ""
            t_el.text = rtext[:off_s] + new + rtext[off_s + (e - s):]
        else:
            t_s = run_s.find(A_T)
            t_e = run_e.find(A_T)
            head = (t_s.text or "")[:off_s]
            tail = (t_e.text or "")[off_e + (e - s):]
            t_s.text = head + new          # 新数字继承数字起始处的格式
            t_e.text = tail
            # 清空被合并的中间 run
            start_off = dict((id(r), o) for r, o in run_offsets)
            for run, off in run_offsets:
                if start_off[id(run)] > s and start_off[id(run)] < e - 1:
                    t = run.find(A_T)
                    if t is not None:
                        t.text = ""
        replaced += 1
    return replaced


# ---------------------------------------------------------------- 图表

def update_chart(chart, df: pd.DataFrame, hint: str, log: RenderLog) -> bool:
    try:
        cd = CategoryChartData()
        cd.categories = [str(c) for c in df.iloc[:, 0]]
        for col in df.columns[1:]:
            cd.add_series(str(col), [float(v) for v in df[col]])
        chart.replace_data(cd)
        return True
    except Exception as e:
        log.warnings.append(f"{hint}: 图表替换失败（{e}），跳过")
        return False


# ---------------------------------------------------------------- 主流程

def load_data(data_xlsx: Path) -> tuple[dict, dict]:
    """返回 (metrics: key -> {value, unit, fmt, ...}, charts: key -> DataFrame)"""
    sheets = pd.read_excel(data_xlsx, sheet_name=None)
    metrics: dict = {}
    charts: dict = {}
    for name, df in sheets.items():
        if name == "metrics":
            for _, r in df.iterrows():
                key = str(r["metric_key"]).strip()
                metrics[key] = {
                    "value": r["值"],
                    "unit": "" if pd.isna(r.get("单位")) else str(r.get("单位")),
                    "fmt": None if pd.isna(r.get("显示格式")) else str(r.get("显示格式")),
                    "period": "" if pd.isna(r.get("更新周期")) else str(r.get("更新周期")),
                }
        elif name.startswith("chart:") or name.startswith("chart_"):
            # Excel sheet 名不允许冒号，推荐 chart_<key>；chart: 为兼容写法
            key = name.split(":", 1)[1] if name.startswith("chart:") else name.split("_", 1)[1]
            charts[key] = df
    return metrics, charts


def render_all(pptx_dir: Path, data_xlsx: Path, out_dir: Path,
               report_period: str = "") -> tuple[list[Path], RenderLog, list[str]]:
    log = RenderLog()
    metrics, charts = load_data(data_xlsx)
    out_dir.mkdir(parents=True, exist_ok=True)
    used_keys: set[str] = set()
    written: list[Path] = []

    decks = sorted(pptx_dir.glob("*.pptx"))
    for src in decks:
        prs = Presentation(str(src))
        changed = False
        for slide_no, slide in enumerate(prs.slides, start=1):
            for shape in slide.shapes:
                key = shape_metric_key(shape)
                if key is None:
                    continue
                if key not in metrics:
                    log.errors.append(
                        f"{src.name}#P{slide_no}: 形状 {shape.name!r} 标签 metric:{key} "
                        f"在 data.xlsx 中不存在")
                    continue
                used_keys.add(key)
                m = metrics[key]
                hint = f"{src.name}#P{slide_no}:{key}"
                if shape.has_chart:
                    df = charts.get(key)
                    if df is None:
                        log.errors.append(f"{hint}: 缺少 chart:{key} sheet")
                        continue
                    if update_chart(shape.chart, df, hint, log):
                        changed = True
                        log.replacements.append(Replacement(
                            deck=src.name, slide_id=slide.slide_id, slide_no=slide_no,
                            shape_name=shape.name, metric_key=key,
                            old_raw="(chart data)", new_raw=f"chart:{key} sheet"))
                    continue
                # 文本框 / 表格
                n_rep = 0
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            for para in cell.text_frame.paragraphs:
                                n_rep += replace_numbers_in_paragraph(
                                    para, float(m["value"]), m["fmt"], hint, log)
                elif shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        n_rep += replace_numbers_in_paragraph(
                            para, float(m["value"]), m["fmt"], hint, log)
                if n_rep:
                    changed = True
                    log.replacements.append(Replacement(
                        deck=src.name, slide_id=slide.slide_id, slide_no=slide_no,
                        shape_name=shape.name, metric_key=key,
                        old_raw=f"{n_rep} 处", new_raw=str(m["value"])))

        if changed:
            suffix = f"_{report_period}" if report_period else ""
            out_path = out_dir / f"{src.stem}{suffix}.pptx"
            prs.save(str(out_path))
            written.append(out_path)

    # 数据源有、但 PPT 无标签的 key
    for key in metrics:
        if key not in used_keys:
            log.warnings.append(f"data.xlsx 中的 {key} 未在任何打标形状中找到")
    return written, log, sorted(used_keys)


def save_render_log(log: RenderLog, out_path: Path):
    rep_df = pd.DataFrame([{
        "deck": r.deck, "slide_no": r.slide_no, "slide_id": r.slide_id,
        "shape_name": r.shape_name, "metric_key": r.metric_key,
        "旧值": r.old_raw, "新值": r.new_raw,
    } for r in log.replacements])
    warn_df = pd.DataFrame({"警告": log.warnings}) if log.warnings else pd.DataFrame(columns=["警告"])
    err_df = pd.DataFrame({"错误": log.errors}) if log.errors else pd.DataFrame(columns=["错误"])
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as w:
        (rep_df if not rep_df.empty else pd.DataFrame(columns=["deck"])).to_excel(
            w, sheet_name="replacements", index=False)
        warn_df.to_excel(w, sheet_name="warnings", index=False)
        err_df.to_excel(w, sheet_name="errors", index=False)
