"""阶段 B：把人工确认的同类页组写入标签（打标固化）。

标签载体（决策，见 README）：
- 页面级：slide.name（即 p:cSld/@name），python-pptx 原生读写。
- 形状级：shape.name 追加 `|metric:<key>`。渲染不受影响（PowerPoint 用 shape id 引用）。
不做 p:custDataLst custom tags：需手工拼 OPC 部件，有触发 PowerPoint 修复提示的风险。

标签基于 slide_id 定位，不基于页码。
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from pptx import Presentation

from .extract import SlideFeatures

PAGE_TAG_RE = re.compile(r"^ppt-sync:group=([^;]+);metric=([^;]+)$")
SHAPE_TAG_RE = re.compile(r"^(?P<orig>.*?)\|metric:(?P<key>[A-Za-z0-9_\-]+)$")


@dataclass
class TagReport:
    tagged_decks: list[str] = field(default_factory=list)
    slides_tagged: int = 0
    shapes_tagged: int = 0
    skipped: list[str] = field(default_factory=list)


def read_confirmed_mapping(xlsx_path: Path) -> list[dict]:
    """读 matches_confirmed.xlsx 的 confirm sheet，只保留 确认=Y 的行。

    若没有名为 confirm 的 sheet，则取第一个 sheet（人工填表时可能另存为单表文件）。
    """
    sheets = pd.read_excel(xlsx_path, sheet_name=None)
    df = sheets.get("confirm") or next(iter(sheets.values()))
    need = {"group_id", "确认(Y/N)", "指标key(英文)", "指标名称"}
    missing = need - set(df.columns)
    if missing:
        raise ValueError(f"confirm sheet 缺少列：{missing}")
    out = []
    for _, r in df.iterrows():
        if str(r["确认(Y/N)"]).strip().upper() == "Y":
            key = str(r["指标key(英文)"]).strip()
            if not key or not re.fullmatch(r"[A-Za-z0-9_\-]+", key):
                raise ValueError(f"group {r['group_id']} 的指标key 非法：{key!r}（只允许英文/数字/下划线）")
            out.append({
                "group_id": int(r["group_id"]),
                "metric_key": key,
                "metric_name": str(r["指标名称"]).strip(),
                "members": [m.strip() for m in str(r["成员"]).split(",")],
            })
    return out


def parse_member(m: str) -> tuple[str, int]:
    """'PPT-1.pptx#5' -> ('PPT-1.pptx', 5)"""
    deck, _, no = m.rpartition("#")
    return deck.strip(), int(no)


def tag_all(pptx_dir: Path, confirmed_xlsx: Path, out_dir: Path,
            index: dict) -> tuple[list[Path], list[dict], TagReport]:
    """打标所有确认组。返回 (打标后文件列表, registry 行, 报告)。"""
    mapping = read_confirmed_mapping(confirmed_xlsx)
    out_dir.mkdir(parents=True, exist_ok=True)
    report = TagReport()

    # (deck, slide_no) -> (group_id, metric_key)
    slide_tags: dict[tuple[str, int], tuple[int, str]] = {}
    for item in mapping:
        for m in item["members"]:
            deck, no = parse_member(m)
            slide_tags[(deck, no)] = (item["group_id"], item["metric_key"])

    # slide_no -> slide_id 映射（从 index 拿，打标时按 slide_id 匹配，页码变了也不怕）
    slideid_by_no: dict[tuple[str, int], int] = {
        (s.deck, s.slide_no): s.slide_id for s in index["slides"]
    }
    # metric_key -> 指标名称
    metric_names = {item["metric_key"]: item["metric_name"] for item in mapping}
    # registry 需要每个 metric 的承载形状：取特征里该页 numbers 的 shape_id
    numbers_by_slide: dict[tuple[str, int], list] = {
        (s.deck, s.slide_no): s.numbers for s in index["slides"]
    }

    decks = sorted({d for d, _ in slide_tags})
    written: list[Path] = []
    registry: list[dict] = []
    for deck in decks:
        src = pptx_dir / deck
        prs = Presentation(str(src))
        id2slide = {s.slide_id: s for s in prs.slides}
        changed = False
        for (d, no), (gid, key) in slide_tags.items():
            if d != deck:
                continue
            sid = slideid_by_no.get((d, no))
            slide = id2slide.get(sid) if sid is not None else None
            if slide is None:
                report.skipped.append(f"{deck}#{no}: 找不到对应 slide_id {sid}")
                continue
            # 页面级标签（保留已有其他 metric，用逗号合并）
            m = PAGE_TAG_RE.match(slide.name or "")
            if m:
                keys = [k for k in m.group(2).split(",") if k] + [key]
                slide.name = f"ppt-sync:group={m.group(1)};metric={','.join(sorted(set(keys)))}"
            else:
                slide.name = f"ppt-sync:group=g{gid};metric={key}"
            changed = True
            report.slides_tagged += 1
            # 图表形状：打标页上的图表默认承载页面级指标（如销售额页的图表）
            for sp in slide.shapes:
                if sp.has_chart:
                    orig, existing = _parse_shape_name(sp.name)
                    if key not in existing:
                        sp.name = "|".join(filter(None, [orig] +
                                                  [f"metric:{k}" for k in existing + [key]]))
                        report.shapes_tagged += 1
                    registry.append({
                        "metric_key": key,
                        "指标名称": metric_names.get(key, ""),
                        "deck": deck,
                        "slide_id": sid,
                        "当前页码": no,
                        "shape_name": sp.name,
                        "当前数值": "(图表数据，见 chart sheet)",
                    })
            # 形状级标签：承载该数字的形状（按 shape_id）
            for n in numbers_by_slide.get((d, no), []):
                shape = None
                for sp in slide.shapes:
                    if sp.shape_id == n.shape_id:
                        shape = sp
                        break
                if shape is None:
                    continue
                orig, existing = _parse_shape_name(shape.name)
                if key in existing:
                    pass  # 已打标
                else:
                    new_tags = existing + [key]
                    shape.name = "|".join(filter(None, [orig] +
                                                 [f"metric:{k}" for k in new_tags]))
                report.shapes_tagged += 1
                registry.append({
                    "metric_key": key,
                    "指标名称": metric_names.get(key, ""),
                    "deck": deck,
                    "slide_id": sid,
                    "当前页码": no,
                    "shape_name": shape.name,
                    "当前数值": n.raw,
                })
        if not changed:
            continue
        out_path = out_dir / deck
        prs.save(str(out_path))
        written.append(out_path)
        report.tagged_decks.append(deck)
    return written, registry, report


def _parse_shape_name(name: str) -> tuple[str, list[str]]:
    """'原名|metric:a|metric:b' -> ('原名', ['a','b'])"""
    if not name:
        return "", []
    parts = name.split("|")
    orig_parts, keys = [], []
    for p in parts:
        mm = re.fullmatch(r"metric:([A-Za-z0-9_\-]+)", p)
        if mm:
            keys.append(mm.group(1))
        else:
            orig_parts.append(p)
    return "|".join(orig_parts), keys


def shape_metric_key(shape) -> str | None:
    """取出形状上的第一个 metric key；无标签返回 None。"""
    _, keys = _parse_shape_name(shape.name or "")
    return keys[0] if keys else None


def save_registry(registry: list[dict], out_path: Path, names_json: Path | None = None):
    df = pd.DataFrame(registry, columns=[
        "metric_key", "指标名称", "deck", "slide_id", "当前页码", "shape_name", "当前数值",
    ])
    df.to_excel(out_path, index=False)
    if names_json is not None:
        names_json.parent.mkdir(parents=True, exist_ok=True)
        names_json.write_text(json.dumps(
            {r["metric_key"]: r["指标名称"] for r in registry}, ensure_ascii=False),
            encoding="utf-8")
