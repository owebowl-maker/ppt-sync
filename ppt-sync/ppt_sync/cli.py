"""ppt-sync 命令行入口。

命令：
  index  <pptx_dir> --out work/index.json        阶段 A-1 特征索引
  match  --index work/index.json --out work/     阶段 A-2/3/4 打分聚类 + matches.xlsx
  tag    <pptx_dir> --mapping xlsx --out dir     阶段 B 打标
  find   <pptx_file> <slide_no>                  查某页所有同类页
  render <pptx_dir> --data data.xlsx --out dir   阶段 C 按数据源更新
  check  <pptx_dir>                              跨文件数值一致性体检
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from .features import load_config
from .extract import build_index, save_index, load_index
from .match import match_all, cluster
from .report import write_report, check_number_consistency
from .tag import PAGE_TAG_RE
from .render_data import render_all, save_render_log

DEFAULT_CONFIG = Path(__file__).resolve().parent.parent / "config.yaml"


def _cfg(args) -> dict:
    path = getattr(args, "config", None) or DEFAULT_CONFIG
    return load_config(str(path)) if Path(path).exists() else {}


# ---------------------------------------------------------------- index

def cmd_index(args):
    cfg = _cfg(args)
    pptx_dir = Path(args.pptx_dir)
    if not pptx_dir.exists():
        sys.exit(f"目录不存在：{pptx_dir}")
    out = Path(args.out)
    # 渲染兜底：PNG 输出到索引同级的 renders/ 目录
    index = build_index(pptx_dir, cfg, render_dir=out.parent / "renders")
    save_index(index, out)
    n = len(index["slides"])
    print(f"已索引 {n} 页 → {out}")
    if index.get("render_enabled"):
        n_clip = sum(1 for s in index["slides"] if s.img_embed is not None)
        n_hash = sum(1 for s in index["slides"] if s.img_phash and s.img_embed is None)
        print(f"渲染兜底已启用：{n_clip} 页有 CLIP 嵌入（方向3），"
              f"{n_hash} 页仅整页 pHash")
    else:
        print("注：渲染兜底未启用（无 LibreOffice），仅使用 PPTX 结构特征。")
    if index["noise_images"] or index["noise_tokens"]:
        print(f"去噪：剔除 {len(index['noise_images'])} 个高频图片、"
              f"{len(index['noise_tokens'])} 个高频文本 token（视为模板噪声）")
    for f in index["failed_decks"]:
        print(f"⚠ {f}")


# ---------------------------------------------------------------- match

def cmd_match(args):
    cfg = _cfg(args)
    index = load_index(Path(args.index))
    pairs = match_all(index, cfg)
    groups = cluster(pairs, index["slides"], cfg)
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx = out_dir / "matches.xlsx"
    write_report(index, pairs, groups, cfg, xlsx)
    # find 命令用的快速检索文件
    (out_dir / "matches.json").write_text(json.dumps({
        "pairs": [p.to_jsonable() for p in pairs],
        "groups": [[index["slides"][k].deck + "#" + str(index["slides"][k].slide_no)
                    for k in g] for g in groups],
    }, ensure_ascii=False), encoding="utf-8")
    print(f"候选对 {len(pairs)} 个，同类页组 {len(groups)} 个 → {xlsx}")
    n_vis = sum(1 for p in pairs if p.s_visual is not None)
    if index.get("render_enabled"):
        print(f"其中 {n_vis} 对带整页视觉分（CLIP/pHash）。")
    else:
        print("注：渲染兜底未启用（无 LibreOffice），仅使用 PPTX 结构特征。")


# ---------------------------------------------------------------- tag

def cmd_tag(args):
    cfg = _cfg(args)
    pptx_dir = Path(args.pptx_dir)
    index = load_index(Path(args.index))
    from .tag import tag_all, save_registry
    written, registry, report = tag_all(
        pptx_dir, Path(args.mapping), Path(args.out), index)
    save_registry(registry, Path(args.out) / "registry.xlsx")
    print(f"打标 {len(written)} 份文件：{[p.name for p in written]}")
    print(f"页面标签 {report.slides_tagged} 个，形状标签 {report.shapes_tagged} 个")
    print(f"registry → {Path(args.out) / 'registry.xlsx'}")
    for s in report.skipped:
        print(f"⚠ 跳过：{s}")


# ---------------------------------------------------------------- find

def _slide_record(index: dict, deck: str, slide_no: int):
    for s in index["slides"]:
        if s.deck == deck and s.slide_no == slide_no:
            return s
    return None


def cmd_find(args):
    cfg = _cfg(args)
    pptx = Path(args.pptx_file)
    work = Path(args.work)
    index = load_index(work / "index.json")
    matches = {}
    mp = work / "matches.json"
    if mp.exists():
        matches = json.loads(mp.read_text(encoding="utf-8"))
    # 读该文件的页面标签（直接打开 pptx 拿 slide.name）
    from pptx import Presentation
    prs = Presentation(str(pptx))
    slide_no = int(args.slide_no)
    if not 1 <= slide_no <= len(prs.slides._sldIdLst):
        sys.exit(f"页码越界：1~{len(prs.slides._sldIdLst)}")
    slide = list(prs.slides)[slide_no - 1]
    tagged = slide.name or ""
    results: list = []
    m = PAGE_TAG_RE.match(tagged)
    if m:
        gid, keys = m.group(1), m.group(2)
        print(f"页面标签：group={gid} metric={keys}（精确匹配）")
        # 遍历目录内其他文件的标签
        for other in sorted(pptx.parent.glob("*.pptx")):
            if other == pptx:
                continue
            oprs = Presentation(str(other))
            for no, s in enumerate(oprs.slides, start=1):
                om = PAGE_TAG_RE.match(s.name or "")
                if om and om.group(1) == gid:
                    nums = _numbers_of(index, other.name, no)
                    results.append((other.name, no, 1.0, True, nums))
    else:
        print("页面无标签，回退到相似度匹配结果：")
        for g in matches.get("groups", []):
            me = f"{pptx.name}#{slide_no}"
            if me in g:
                for other in g:
                    if other == me:
                        continue
                    d, _, no = other.rpartition("#")
                    results.append((d, int(no), None, False, _numbers_of(index, d, int(no))))
    if not results:
        print("未找到同类页。")
        return
    print(f"共 {len(results)} 个同类页：")
    for deck, no, score, exact, nums in results:
        score_s = "精确" if exact else (f"{score:.2f}" if score else "-")
        nums_s = "；".join(f"{n.raw}" for n in nums) if nums else ""
        print(f"  {deck} 第 {no} 页  相似度 {score_s}  数值: {nums_s}")


def _numbers_of(index: dict, deck: str, slide_no: int):
    s = _slide_record(index, deck, slide_no)
    return s.numbers if s else []


# ---------------------------------------------------------------- render

def cmd_render(args):
    cfg = _cfg(args)
    written, log, used = render_all(
        Path(args.pptx_dir), Path(args.data), Path(args.out),
        report_period=args.period or "")
    out_dir = Path(args.out)
    save_render_log(log, out_dir / "render_log.xlsx")
    print(f"更新 {len(written)} 份文件，替换 {len(log.replacements)} 处 → {out_dir}")
    for w in log.warnings:
        print(f"⚠ {w}")
    if log.errors:
        for e in log.errors:
            print(f"✗ {e}")
        sys.exit(1)


# ---------------------------------------------------------------- check

def cmd_check(args):
    cfg = _cfg(args)
    pptx_dir = Path(args.pptx_dir)
    index = build_index(pptx_dir, cfg)
    pairs = match_all(index, cfg)
    groups = cluster(pairs, index["slides"], cfg)
    # 有标签的页面按标签分组补进 groups（标签比相似度准）
    tagged: dict[str, list] = {}
    from pptx import Presentation
    for deck in sorted(pptx_dir.glob("*.pptx")):
        prs = Presentation(str(deck))
        for no, s in enumerate(prs.slides, start=1):
            m = PAGE_TAG_RE.match(s.name or "")
            if m:
                tagged.setdefault(m.group(1), []).append((deck.name, no))
    num_df = check_number_consistency(groups, index["slides"], cfg)
    bad = num_df[num_df["是否不一致"] == "❌"] if not num_df.empty else num_df
    if bad.empty:
        print("✓ 数值一致性检查通过：无不一致项。")
        return
    print(f"✗ 发现 {len(bad)} 处数值不一致：")
    for _, r in bad.iterrows():
        print(f"  [{r['指标上下文']}] {r['deck']} 第 {r['slide']} 页 = {r['数值']}")
    sys.exit(1)


# ---------------------------------------------------------------- main

def main(argv=None):
    ap = argparse.ArgumentParser(prog="ppt-sync", description="PPT 同类页匹配与数据自动更新")
    ap.add_argument("--config", default=None, help="config.yaml 路径")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("index", help="阶段 A-1：特征索引")
    p.add_argument("pptx_dir")
    p.add_argument("--out", default="work/index.json")
    p.set_defaults(func=cmd_index)

    p = sub.add_parser("match", help="阶段 A-2/3/4：打分聚类 + 对照表")
    p.add_argument("--index", default="work/index.json")
    p.add_argument("--out", default="work")
    p.set_defaults(func=cmd_match)

    p = sub.add_parser("tag", help="阶段 B：打标固化")
    p.add_argument("pptx_dir")
    p.add_argument("--index", default="work/index.json")
    p.add_argument("--mapping", required=True, help="人工确认后的 matches_confirmed.xlsx")
    p.add_argument("--out", required=True, help="打标后输出目录")
    p.set_defaults(func=cmd_tag)

    p = sub.add_parser("find", help="查某页的所有同类页")
    p.add_argument("pptx_file")
    p.add_argument("slide_no")
    p.add_argument("--work", default="work")
    p.set_defaults(func=cmd_find)

    p = sub.add_parser("render", help="阶段 C：按数据源更新")
    p.add_argument("pptx_dir")
    p.add_argument("--data", required=True, help="data.xlsx")
    p.add_argument("--out", required=True)
    p.add_argument("--period", default="", help="报告期，如 2026Q3，用作文件名后缀")
    p.set_defaults(func=cmd_render)

    p = sub.add_parser("check", help="跨文件数值一致性体检")
    p.add_argument("pptx_dir")
    p.set_defaults(func=cmd_check)

    args = ap.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
