"""阶段 A-2：两两打分、聚类成同类页组。

子分：s_img / s_chart / s_text / s_layout / s_visual（渲染未启用则缺失摊权）。
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict

from .features import jaccard, phash_hamming
from .extract import SlideFeatures


@dataclass
class PairScore:
    deck_a: str
    slide_a: int
    deck_b: str
    slide_b: int
    same_deck: bool
    score: float = 0.0
    s_img: float | None = None
    s_chart: float | None = None
    s_text: float | None = None
    s_layout: float | None = None
    s_visual: float | None = None   # CLIP 余弦（方向3），无渲染时为整页 pHash，无渲染兜底则 None
    evidence: str = ""              # 主要依据
    high_confidence: bool = False

    def to_jsonable(self) -> dict:
        return asdict(self)


def _score_img(a: SlideFeatures, b: SlideFeatures) -> float | None:
    if not a.image_hashes and not b.image_hashes:
        return None
    inter = a.image_hashes & b.image_hashes
    if inter:
        return len(inter) / len(a.image_hashes | b.image_hashes)
    # SHA 不同但 pHash 汉明距离 <= 6：同一图被重新压缩
    for pa in a.image_phashes:
        for pb in b.image_phashes:
            d = phash_hamming(pa, pb)
            if d is not None and d <= 6:
                return 1.0
    return 0.0


def _score_chart(a: SlideFeatures, b: SlideFeatures) -> float | None:
    if not a.chart_sigs and not b.chart_sigs:
        return None
    best = 0.0
    for ta, sa, ca in a.chart_sigs:
        for tb, sb, cb in b.chart_sigs:
            name_sim = jaccard(set(sa), set(sb))
            cat_sim = jaccard(set(ca), set(cb))
            best = max(best, (name_sim + cat_sim) / 2 if ta == tb else 0.0)
    return best


def _score_text(a: SlideFeatures, b: SlideFeatures) -> float | None:
    if not a.text_tokens and not b.text_tokens:
        return None
    j = jaccard(a.text_tokens, b.text_tokens)
    try:
        from rapidfuzz import fuzz
        # token_set 比较容忍长短文案差异；拼接已排序 token 保证无序比较
        sa = " ".join(sorted(a.text_tokens))
        sb = " ".join(sorted(b.text_tokens))
        j = max(j, fuzz.token_set_ratio(sa, sb) / 100.0)
    except ImportError:
        pass
    return j


def _score_layout(a: SlideFeatures, b: SlideFeatures) -> float | None:
    if not a.shape_layout and not b.shape_layout:
        return None
    # 按类型分组做贪心配对（组内元素少，贪心足够；匈牙利收益边际）
    from collections import defaultdict
    by_type_a: dict = defaultdict(list)
    by_type_b: dict = defaultdict(list)
    for s in a.shape_layout:
        by_type_a[s.stype].append(s)
    for s in b.shape_layout:
        by_type_b[s.stype].append(s)
    dists: list[float] = []
    for t in set(by_type_a) | set(by_type_b):
        la, lb = by_type_a.get(t, []), by_type_b.get(t, [])
        if not la or not lb:
            dists.append(1.0)   # 一方缺失该类型形状：按完全不匹配计
            continue
        # 贪心：每次取中心距离最近的一对
        pairs = [(pa, pb) for pa in la for pb in lb]
        used_a: set = set()
        used_b: set = set()
        matched: list[float] = []
        for pa, pb in sorted(pairs, key=lambda x: (x[0].cx - x[1].cx) ** 2 + (x[0].cy - x[1].cy) ** 2):
            if pa.shape_id in used_a or pb.shape_id in used_b:
                continue
            used_a.add(pa.shape_id)
            used_b.add(pb.shape_id)
            d = ((pa.cx - pb.cx) ** 2 + (pa.cy - pb.cy) ** 2) ** 0.5
            matched.append(d)
        unmatched = max(len(la), len(lb)) - len(matched)
        avg_d = sum(matched) / len(matched) if matched else 1.0
        # 未配对的形状按距离 1 计
        dists.append((avg_d * len(matched) + unmatched * 1.0) / max(len(la), len(lb)))
    return 1.0 - (sum(dists) / len(dists) if dists else 1.0)


def score_pair(a: SlideFeatures, b: SlideFeatures, cfg: dict,
               visual_sim: float | None = None) -> PairScore:
    mcfg = cfg.get("match", {})
    weights = dict(mcfg.get("weights", {}))
    if visual_sim is None:
        # CLIP 嵌入缺失时降级为整页 pHash 相似度（无渲染则仍为 None）
        from .visual import phash_similarity
        visual_sim = phash_similarity(a.img_phash, b.img_phash)
    subs = {
        "s_img": _score_img(a, b),
        "s_chart": _score_chart(a, b),
        "s_text": _score_text(a, b),
        "s_layout": _score_layout(a, b),
        "s_visual": visual_sim,
    }
    # 缺失项权重摊给其余项
    active = {k: v for k, v in subs.items() if v is not None}
    wsum = sum(weights.get(k, 0.0) for k in active)
    score = 0.0
    if wsum > 0:
        for k, v in active.items():
            score += weights.get(k, 0.0) / wsum * v
    same_deck = a.deck == b.deck
    ps = PairScore(
        deck_a=a.deck, slide_a=a.slide_no,
        deck_b=b.deck, slide_b=b.slide_no,
        same_deck=same_deck, score=round(score, 4), **active,
    )
    # 主要依据（可读性）
    ev = []
    if (subs["s_img"] or 0) >= 0.5:
        n_common = len(a.image_hashes & b.image_hashes)
        ev.append(f"共享 {n_common} 张图片")
    if (subs["s_chart"] or 0) >= 0.5:
        ev.append("图表系列名/类别相同")
    if (subs["s_text"] or 0) >= 0.4:
        ev.append("文本关键词相似")
    if (subs["s_layout"] or 0) >= 0.7:
        ev.append("版式布局一致")
    if (subs["s_visual"] or 0) >= 0.9:
        ev.append("整页视觉高度相似")
    ps.evidence = "；".join(ev) if ev else "综合打分"
    ps.high_confidence = score >= mcfg.get("high_confidence_threshold", 0.75)
    return ps


def _visual_sim_map(slides: list[SlideFeatures]) -> dict[frozenset, float]:
    """有 CLIP 嵌入的页两两算余弦相似度（归一化向量点积，矩阵乘法一次算完）。"""
    idx = [k for k, s in enumerate(slides) if s.img_embed]
    if len(idx) < 2:
        return {}
    import numpy as np
    m = np.array([slides[k].img_embed for k in idx], dtype=np.float32)
    sim = m @ m.T
    out: dict[frozenset, float] = {}
    for a in range(len(idx)):
        for b in range(a + 1, len(idx)):
            out[frozenset((idx[a], idx[b]))] = float(sim[a, b])
    return out


def match_all(index: dict, cfg: dict) -> list[PairScore]:
    """候选对打分。先做便宜预筛（无共同图片且文本 Jaccard 低则跳过）；
    CLIP 整页相似度高的对绕过该预筛（方向3）。"""
    slides: list[SlideFeatures] = index["slides"]
    mcfg = cfg.get("match", {})
    cand_th = mcfg.get("candidate_threshold", 0.55)
    visual_prescreen = mcfg.get("visual_prescreen", 0.80)
    vis_sims = _visual_sim_map(slides)
    pairs: list[PairScore] = []
    n = len(slides)
    for i in range(n):
        for j in range(i + 1, n):
            a, b = slides[i], slides[j]
            if a.is_boilerplate != b.is_boilerplate:
                continue
            # 封面/目录等通用页互配没有业务价值，直接跳过（与实施方案的差异：原稿允许 ≥0.9 保留）
            if a.is_boilerplate and b.is_boilerplate:
                continue
            vs = vis_sims.get(frozenset((i, j)))
            # 便宜预筛：无共同图片、无共同 token 的组合直接跳过；
            # 但图表签名相同或 CLIP 整页相似度高的组合要进来
            if not (a.image_hashes & b.image_hashes) and not (a.text_tokens & b.text_tokens):
                has_chart = a.chart_sigs and b.chart_sigs
                visual_hit = vs is not None and vs >= visual_prescreen
                if not has_chart and not visual_hit:
                    continue
            ps = score_pair(a, b, cfg, visual_sim=vs)
            if ps.score >= cand_th:
                pairs.append(ps)
    return pairs


# ---------------------------------------------------------------- 聚类

class _UF:
    def __init__(self, n: int):
        self.p = list(range(n))

    def find(self, x: int) -> int:
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x

    def union(self, a: int, b: int):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.p[ra] = rb


def cluster(pairs: list[PairScore], slides: list[SlideFeatures], cfg: dict) -> list[list[int]]:
    """贪心合并：按分数降序遍历候选对，仅当两组间任意两页分数 >= cohesion_min 才合并，
    避免链式误连（A 分数高、C 分数低却因 B 桥接进同一组）。"""
    mcfg = cfg.get("match", {})
    cohesion = mcfg.get("group_cohesion_min", 0.45)
    key2idx = {(s.deck, s.slide_no): k for k, s in enumerate(slides)}
    score_map: dict[frozenset[int], float] = {}
    edges = []
    for p in pairs:
        ia, ib = key2idx[(p.deck_a, p.slide_a)], key2idx[(p.deck_b, p.slide_b)]
        score_map[frozenset((ia, ib))] = p.score
        edges.append((p.score, ia, ib))
    edges.sort(reverse=True)

    groups: dict[int, list[int]] = {k: [k] for k in range(len(slides))}
    root_of = {k: k for k in range(len(slides))}

    def pair_score(i: int, j: int) -> float:
        return score_map.get(frozenset((i, j)), 0.0)

    for score, ia, ib in edges:
        if score < cohesion:
            break
        ra, rb = root_of[ia], root_of[ib]
        if ra == rb:
            continue
        ga, gb = groups[ra], groups[rb]
        # 两组间任意两页都要 >= cohesion（组内已有的对直接查表，缺失的按 0 计则拒绝合并）
        if all(pair_score(x, y) >= cohesion for x in ga for y in gb):
            merged = ga + gb
            groups[ra] = merged
            del groups[rb]
            for k in gb:
                root_of[k] = ra

    result = [m for m in groups.values() if len(m) >= 2]
    result.sort(key=lambda g: min(g))
    return result
