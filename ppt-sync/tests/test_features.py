# -*- coding: utf-8 -*-
"""数字抽取 / 文本归一化 / 哈希 的单元测试（对应方案 A-5）。"""
import pytest

from ppt_sync.features import (
    NUM_RE, extract_numbers, jaccard, normalize_text, phash_of_blob,
    sha1_of_blob, text_tokens, vague_number_texts,
)


def num_and_unit(text):
    hits = extract_numbers(text, frozenset(), 0, 0)
    return [(h.value, h.unit) for h in hits]


@pytest.mark.parametrize("text,expected", [
    ("销售额1.2亿", [(1.2, "亿")]),
    ("12,000", [(12000.0, "")]),
    ("3,200.5", [(3200.5, "")]),
    ("增长率35%", [(35.0, "%")]),
    ("128家", [(128.0, "家")]),
    ("50+", [(50.0, "+")]),
    ("合计1,200万元", [(1200.0, "万元")]),
])
def test_number_formats(text, expected):
    assert num_and_unit(text) == expected


def test_no_number():
    assert num_and_unit("近百家") == []
    assert vague_number_texts("服务了近百家企业")  # 非数字表述被记录


def test_multiple_numbers():
    assert num_and_unit("客户3,200家，营收1.2亿元") == [(3200.0, "家"), (1.2, "亿元")]


def test_normalize_and_tokens():
    assert "<NUM>" in normalize_text("客户3,200家")
    toks = text_tokens("合作案例128个")
    assert "合作" in toks and "案例" in toks     # 中文 bigram
    assert "128" not in toks and "num" not in toks   # 数字与占位符都不算关键词


def test_jaccard():
    assert jaccard({"a", "b"}, {"a", "b"}) == 1.0
    assert jaccard(set(), set()) == 0.0
    assert jaccard({"a"}, {"a", "b"}) == 1 / 2


def test_hashes():
    blob = b"\x89PNG fake image data"
    assert sha1_of_blob(blob) == sha1_of_blob(blob)
    assert sha1_of_blob(b"x") != sha1_of_blob(b"y")
    assert phash_of_blob(blob) is None          # 坏图不崩溃
