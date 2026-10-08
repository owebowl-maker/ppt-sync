# -*- coding: utf-8 -*-
"""特征抽取：组合形状递归、图表签名、boilerplate、全局去噪。"""
from ppt_sync.extract import SlideFeatures, build_index, extract_deck


def _get(index, deck, slide_no):
    for s in index["slides"]:
        if s.deck == deck and s.slide_no == slide_no:
            return s
    raise KeyError((deck, slide_no))


def test_group_recursion(index):
    # PPT-1 第 6 页：嵌套组合形状内的文本应被收集
    s = _get(index, "PPT-1.pptx", 6)
    assert "服务" in s.text_tokens or "客户" in s.text_tokens
    values = {n.value for n in s.numbers}
    assert 200.0 in values and 30.0 in values


def test_chart_signature(index):
    s = _get(index, "PPT-1.pptx", 3)
    assert s.chart_sigs, "销售额页应抽取到图表签名"
    ctype, series, cats = next(iter(s.chart_sigs))
    assert "销售额" in "".join(series)


def test_numbers_extracted(index):
    s = _get(index, "PPT-1.pptx", 5)
    assert any(n.value == 128 and n.unit == "个" for n in s.numbers)


def test_boilerplate(index):
    for deck in ("PPT-1.pptx", "PPT-2.pptx", "PPT-3.pptx"):
        assert _get(index, deck, 1).is_boilerplate   # 封面
        assert not _get(index, "PPT-1.pptx", 3).is_boilerplate


def test_slide_ids_stable(index):
    a = _get(index, "PPT-1.pptx", 5)
    b = _get(index, "PPT-3.pptx", 13)
    assert a.slide_id > 0 and b.slide_id > 0 and a.slide_id != b.slide_id


def test_denoise():
    """token 出现在超过 denoise_ratio 页面时被剔除（合成数据控制比例）。"""
    from ppt_sync.extract import denoise
    feats = []
    for i in range(10):
        toks = {"页", "logo"} if i < 5 else {f"独有{i}", "页"}
        feats.append(SlideFeatures(deck=f"D{i}.pptx", slide_no=1, slide_id=i,
                                   text_tokens=set(toks)))
    cfg = {"extract": {"denoise_ratio": 0.3}}
    out, noise_imgs, noise_toks = denoise(feats, cfg)
    assert "logo" in noise_toks       # 5/10 > 0.3 → 剔除
    assert "页" in noise_toks         # 10/10 > 0.3 → 同样剔除
    assert "独有0" not in noise_toks  # 低频保留
