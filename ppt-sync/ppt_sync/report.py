"""阶段 A-3/A-4：数值一致性检查 + 对照表 Excel 输出。"""
from __future__ import annotations

import pandas as pd

from .extract import SlideFeatures
from .match import PairScore
from .features import jaccard


def check_number_consistency(groups: list[list[int]], slides: list[SlideFeatures],
                             cfg: dict) -> pd.DataFrame:
    """组内按 context_tokens 配对数字，同一指标数值不同 → 不一致。

    返回 numbers 表：group_id, deck, slide, 指标上下文, 数值, 是否不一致。
    """
    th = cfg.get("report", {}).get("inconsistency_context_jaccard", 0.5)
    rows = []
    for gid, members in enumerate(groups):
        # 汇总组内所有数字，按上下文聚类
        hits = [(slides[k], n) for k in members for n in slides[k].numbers]
        used: list[list[tuple[SlideFeatures, object]]] = []  # 每簇 = [(slide, hit)]
        for s, h in hits:
            placed = False
            for cluster_ in used:
                rep = cluster_[0][1]
                if jaccard(set(h.context_tokens), set(rep.context_tokens)) >= th:
                    cluster_.append((s, h))
                    placed = True
                    break
            if not placed:
                used.append([(s, h)])
        for ci, cluster_ in enumerate(used):
            values = {(round(h.value, 4), h.unit) for _, h in cluster_}
            inconsistent = len(values) > 1
            # 指标上下文：取簇内 token 最多的上下文前几个词
            ctx = cluster_[0][1].context_tokens
            ctx_label = " ".join(sorted(ctx)[:6]) if ctx else "(无上下文)"
            for s, h in cluster_:
                rows.append({
                    "group_id": gid,
                    "deck": s.deck,
                    "slide": s.slide_no,
                    "簇": ci,
                    "指标上下文": ctx_label,
                    "数值": h.raw,
                    "是否不一致": "❌" if inconsistent else "",
                })
    return pd.DataFrame(rows)


def _suggest_group_name(slides: list[SlideFeatures]) -> str:
    """建议指标名：取组内共有 token 中较有信息量的前几个。"""
    common: set[str] | None = None
    for s in slides:
        common = set(s.text_tokens) if common is None else common & s.text_tokens
    if not common:
        return ""
    return " ".join(sorted(common)[:5])


def write_report(index: dict, pairs: list[PairScore], groups: list[list[int]],
                 cfg: dict, out_path) -> pd.DataFrame:
    """输出 matches.xlsx（groups / pairs / numbers / confirm 四个 Sheet）。"""
    slides: list[SlideFeatures] = index["slides"]
    pairs_df = pd.DataFrame([p.to_jsonable() for p in pairs])
    if not pairs_df.empty:
        pairs_df = pairs_df.sort_values("score", ascending=False)
        pairs_df = pairs_df.rename(columns={
            "deck_a": "文件A", "slide_a": "页码A",
            "deck_b": "文件B", "slide_b": "页码B",
            "score": "总分", "s_img": "图片分", "s_chart": "图表分",
            "s_text": "文本分", "s_layout": "版式分", "s_visual": "视觉分",
            "evidence": "主要依据", "high_confidence": "高置信",
            "same_deck": "同文件内",
        })

    num_df = check_number_consistency(groups, slides, cfg)

    group_rows = []
    for gid, members in enumerate(groups):
        member_slides = [slides[k] for k in members]
        member_label = ", ".join(f"{s.deck}#{s.slide_no}" for s in member_slides)
        min_pair = None
        for i in range(len(members)):
            for j in range(i + 1, len(members)):
                for p in pairs:
                    if ({(p.deck_a, p.slide_a), (p.deck_b, p.slide_b)} ==
                        {(slides[members[i]].deck, slides[members[i]].slide_no),
                         (slides[members[j]].deck, slides[members[j]].slide_no)}):
                        min_pair = p.score if min_pair is None else min(min_pair, p.score)
        inconsistent = (not num_df.empty and
                        (num_df[num_df["group_id"] == gid]["是否不一致"] == "❌").any())
        conf = "高" if (min_pair is not None and
                        min_pair >= cfg.get("match", {}).get("high_confidence_threshold", 0.75)) \
            else "中"
        group_rows.append({
            "group_id": gid,
            "建议指标名": _suggest_group_name(member_slides),
            "成员": member_label,
            "组内最低分": min_pair if min_pair is not None else 1.0,
            "置信度": conf,
            "是否有数值不一致": "❌ 是" if inconsistent else "",
        })
    groups_df = pd.DataFrame(group_rows)

    # confirm sheet：留给人工填写
    confirm_rows = [{
        "group_id": r["group_id"],
        "建议指标名": r["建议指标名"],
        "成员": r["成员"],
        "确认(Y/N)": "",
        "指标key(英文)": "",
        "指标名称": "",
    } for _, r in groups_df.iterrows()]
    confirm_df = pd.DataFrame(confirm_rows)

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as w:
        groups_df.to_excel(w, sheet_name="groups", index=False)
        pairs_df.to_excel(w, sheet_name="pairs", index=False)
        num_df.to_excel(w, sheet_name="numbers", index=False)
        confirm_df.to_excel(w, sheet_name="confirm", index=False)
        # numbers 不一致行标红
        if not num_df.empty:
            wb = w.book
            ws = w.sheets["numbers"]
            red = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
            from xlsxwriter.utility import xl_col_to_name
            flag_col = xl_col_to_name(num_df.columns.get_loc("是否不一致"))
            nrow = len(num_df)
            ws.conditional_format(1, 0, nrow, len(num_df.columns) - 1, {
                "type": "formula",
                "criteria": f'=${flag_col}2="❌"',
                "format": red,
            })
    return confirm_df
