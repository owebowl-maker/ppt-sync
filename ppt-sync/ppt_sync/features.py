"""文本归一化、数字抽取、图片哈希。

所有比较都基于无序集合，不依赖阅读顺序。
"""
from __future__ import annotations

import hashlib
import io
import re
from dataclasses import dataclass, field

import yaml
from PIL import Image
import imagehash

# ---------------------------------------------------------------- config

DEFAULT_CONFIG_PATH = None  # 由 cli 注入；单元测试可直接传 dict


def load_config(path: str | None) -> dict:
    if path is None:
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


# ---------------------------------------------------------------- 文本归一化

# 数字正则：12,000 / 3,200 / 1.2亿 / 35% / 50+ / 128家 / 3.2万元
# 支持全角逗号。单位后缀单独捕获。
NUM_RE = re.compile(
    r"(?P<num>\d+(?:[,，]\d{3})*(?:\.\d+)?)(?P<unit>\s*(?:%|亿元|亿|万元|万|千|家\+|\+家|\+|家|个|人|元|单|项|份|年|月|季)?)"
)

# 不可解析的数字式表述：近百家 / 数十个 / 几十万 等
VAGUE_NUM_RE = re.compile(r"(近|超过|约|大概|数十|数百万|数十万|几|多)百|近\d|超\d")

PLACEHOLDER = "<NUM>"

_CJK_RE = re.compile(r"[一-鿿]+")


def _cjk_tokens(s: str) -> set[str]:
    """中文按二元组（bigram）切分，单字保留。"""
    out: set[str] = set()
    for m in _CJK_RE.finditer(s):
        w = m.group(0)
        if len(w) == 1:
            out.add(w)
        else:
            out.update(w[i : i + 2] for i in range(len(w) - 1))
    return out


def normalize_text(text: str) -> str:
    """数字替换为 <NUM>，用于 token 化与指标上下文。"""
    return NUM_RE.sub(PLACEHOLDER, text)


def text_tokens(text: str) -> set[str]:
    """关键词集合：英数词 + 中文 bigram。数字已占位。"""
    norm = normalize_text(text)
    tokens: set[str] = set()
    for w in re.findall(r"[A-Za-z0-9_]+", norm):
        if w.lower() != "num":   # <NUM> 占位符不算关键词
            tokens.add(w.lower())
    tokens.update(_cjk_tokens(norm))
    return tokens


def jaccard(a: set, b: set) -> float:
    if not a and not b:
        return 0.0
    inter = len(a & b)
    if inter == 0:
        return 0.0
    return inter / (len(a) + len(b) - inter)


# ---------------------------------------------------------------- 数字抽取

STOPWORDS = {
    "the", "a", "an", "of", "and", "or", "to", "in", "on", "for", "with", "by",
    "is", "are", "at", "as", "num",
}


@dataclass
class NumberHit:
    value: float            # 数值（去掉千分位与 % 后的浮点；"1.2亿" -> 1.2，unit 单独存）
    raw: str                # 原文，如 "12,000" / "35%"
    unit: str               # 单位后缀，如 "%" / "亿" / "家"
    context_tokens: frozenset[str]   # 同一文本框内的上下文关键词（去数字）
    shape_id: int           # 承载该数字的形状 id
    paragraph_idx: int      # 段落序号（同形状内可能有多个数字）
    char_start: int         # 在段落拼接文本中的起点（供阶段 C 精确替换）
    char_end: int


def extract_numbers(paragraph_text: str, context_tokens: frozenset[str],
                    shape_id: int, paragraph_idx: int) -> list[NumberHit]:
    """从一段文字里抽取数字及其上下文。"""
    hits: list[NumberHit] = []
    for m in NUM_RE.finditer(paragraph_text):
        raw_num = m.group("num").replace(",", "").replace("，", "")
        try:
            value = float(raw_num)
        except ValueError:
            continue
        unit = (m.group("unit") or "").strip()
        hits.append(NumberHit(
            value=value,
            raw=m.group(0),
            unit=unit,
            context_tokens=context_tokens,
            shape_id=shape_id,
            paragraph_idx=paragraph_idx,
            char_start=m.start(),
            char_end=m.end(),
        ))
    return hits


def vague_number_texts(text: str) -> list[str]:
    """非数字表述（近百家、数十个），记下来进报告。"""
    return [m.group(0) for m in VAGUE_NUM_RE.finditer(text)]


# ---------------------------------------------------------------- 哈希

def sha1_of_blob(blob: bytes) -> str:
    return hashlib.sha1(blob).hexdigest()


def phash_of_blob(blob: bytes) -> str | None:
    """图片感知哈希，用于识别同一 Logo 被重新压缩的情况。解码失败返回 None。"""
    try:
        img = Image.open(io.BytesIO(blob)).convert("L")
        return str(imagehash.phash(img))
    except Exception:
        return None


def phash_of_image(path) -> str | None:
    """整页 PNG 的感知哈希（渲染兜底用）。"""
    try:
        img = Image.open(path).convert("L")
        return str(imagehash.phash(img))
    except Exception:
        return None


def phash_hamming(h1: str | None, h2: str | None) -> int | None:
    if h1 is None or h2 is None:
        return None
    return imagehash.hex_to_hash(h1) - imagehash.hex_to_hash(h2)
