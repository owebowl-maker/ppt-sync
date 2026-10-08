# -*- coding: utf-8 -*-
"""打分与聚类：子分公式、权重重分配、防链式误连、端到端对应关系。"""
import pytest

from ppt_sync.extract import SlideFeatures
from ppt_sync.match import score_pair, cluster
from conftest import CONFIG, KNOWN_PAIRS


def _feat(deck, no, **kw):
    kw.setdefault("slide_id", no)
    return SlideFeatures(deck=deck, slide_no=no, **kw)


def test_score_weight_redistribution():
    """两页都无图表/图片时，权重摊给文本与版式。"""
    a = _feat("A.pptx", 1, text_tokens={"合作", "案例", "数据"})
    b = _feat("B.pptx", 1, text_tokens={"合作", "案例", "数据"})
    ps = score_pair(a, b, CONFIG)
    assert ps.s_chart is None and ps.s_img is None
    assert ps.score == 1.0   # 文本、版式全同（都是空布局 → 版式缺失）

    b2 = _feat("B.pptx", 1, text_tokens={"完全", "不同", "内容"})
    assert score_pair(a, b2, CONFIG).score < 0.55


def test_image_hash_score():
    a = _feat("A.pptx", 1, image_hashes={"x", "y"})
    b = _feat("B.pptx", 1, image_hashes={"x", "y", "z"})
    assert score_pair(a, b, CONFIG).s_img == pytest.approx(2 / 3)
    c = _feat("B.pptx", 2, image_hashes={"z"})
    assert score_pair(a, c, CONFIG).s_img == 0.0


def test_cluster_no_chain():
    """A-B 高分、B-C 高分、A-C 低分 → 不得合并成一组（链式误连）。"""
    cfg = dict(CONFIG)
    cfg["match"] = {"candidate_threshold": 0.55, "high_confidence_threshold": 0.75,
                    "group_cohesion_min": 0.8}

    def key(i):
        return (f"D{i}.pptx", i)

    class P:
        def __init__(self, da, sa, db, sb, score):
            self.deck_a, self.slide_a, self.deck_b, self.slide_b = da, sa, db, sb
            self.score = score
            self.same_deck = da == db

    slides = [_feat(f"D{i}.pptx", i) for i in (1, 2, 3)]
    pairs = [P("D1.pptx", 1, "D2.pptx", 2, 0.95),
             P("D2.pptx", 2, "D3.pptx", 3, 0.9),
             P("D1.pptx", 1, "D3.pptx", 3, 0.3)]  # 弱边
    groups = cluster(pairs, slides, cfg)
    assert len(groups) == 1 and set(groups[0]) == {0, 1}  # D3 被挡在组外


def test_known_pairs_grouped(index, match_result):
    """已知同类页（PPT-1#5 ↔ PPT-3#13 等）必须入组且高置信。"""
    pairs, groups = match_result
    idx = {(s.deck, s.slide_no): k for k, s in enumerate(index["slides"])}
    for pa, pb in KNOWN_PAIRS.items():
        found_group = None
        for g in groups:
            members = {(index["slides"][k].deck, index["slides"][k].slide_no) for k in g}
            if pa in members and pb in members:
                found_group = members
        assert found_group, f"{pa} ↔ {pb} 应在同一组"
        # 组内不应混入其他页
        assert found_group == {pa, pb}, f"组内混入其他页：{found_group}"


def test_sales_inconsistency(index, match_result):
    """PPT-1#3（3.2亿元）与 PPT-2#2（2.9亿元）应报数值不一致。"""
    from ppt_sync.report import check_number_consistency
    pairs, groups = match_result
    num_df = check_number_consistency(groups, index["slides"], CONFIG)
    bad = num_df[num_df["是否不一致"] == "❌"]
    assert not bad.empty
    assert set(bad["deck"]) == {"PPT-1.pptx", "PPT-2.pptx"}
