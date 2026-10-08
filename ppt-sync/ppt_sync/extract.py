"""阶段 A-1：解析 pptx，为每个 (文件, 页码) 生成特征记录。

递归遍历所有形状：组合形状、表格、图表、图片、背景填充（blipFill）。
不读母版/版式上的内容（只走 slide.shapes）。
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

from .features import (
    NumberHit, extract_numbers, phash_of_blob, sha1_of_blob, text_tokens,
    vague_number_texts,
)

DIAGRAM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
BLIP_QN = qn("a:blip")
R_EMBED_QN = qn("r:embed")


@dataclass
class ShapeInfo:
    shape_id: int
    name: str
    stype: str            # picture / textbox / table / chart / group / other
    cx: float             # 归一化中心 x（0~1）
    cy: float
    w: float
    h: float


@dataclass
class SlideFeatures:
    deck: str
    slide_no: int         # 从 1 开始
    slide_id: int
    hidden: bool = False
    layout_name: str = ""
    image_hashes: set[str] = field(default_factory=set)      # blob SHA1
    image_phashes: set[str] = field(default_factory=set)
    chart_sigs: set[tuple] = field(default_factory=set)      # (type, 系列 frozenset, 类别 frozenset)
    text_tokens: set[str] = field(default_factory=set)       # 关键词集合（数字已占位）
    numbers: list[NumberHit] = field(default_factory=list)
    vague_texts: list[str] = field(default_factory=list)
    shape_layout: list[ShapeInfo] = field(default_factory=list)
    is_boilerplate: bool = False
    render_png: str | None = None
    img_phash: str | None = None
    img_embed: list[float] | None = None   # CLIP 512 维嵌入（方向3），L2 归一化

    def to_jsonable(self) -> dict:
        d = asdict(self)
        d["text_tokens"] = sorted(self.text_tokens)
        d["image_hashes"] = sorted(self.image_hashes)
        d["image_phashes"] = sorted(self.image_phashes)
        if self.img_embed is not None:
            d["img_embed"] = [round(x, 5) for x in self.img_embed]  # 省体积
        d["chart_sigs"] = [
            [sig[0], sorted(sig[1]), sorted(sig[2])] for sig in self.chart_sigs
        ]
        d["numbers"] = [asdict(n) for n in self.numbers]
        d["numbers"] = [
            {**n, "context_tokens": sorted(n["context_tokens"])}
            for n in d["numbers"]
        ]
        return d

    @classmethod
    def from_jsonable(cls, d: dict) -> "SlideFeatures":
        d = dict(d)
        d["text_tokens"] = set(d["text_tokens"])
        d["image_hashes"] = set(d["image_hashes"])
        d["image_phashes"] = set(d["image_phashes"])
        d["chart_sigs"] = {
            (s[0], frozenset(s[1]), frozenset(s[2])) for s in d["chart_sigs"]
        }
        d["numbers"] = [
            NumberHit(
                value=n["value"], raw=n["raw"], unit=n["unit"],
                context_tokens=frozenset(n["context_tokens"]),
                shape_id=n["shape_id"], paragraph_idx=n["paragraph_idx"],
                char_start=n["char_start"], char_end=n["char_end"],
            )
            for n in d["numbers"]
        ]
        d["shape_layout"] = [ShapeInfo(**s) for s in d["shape_layout"]]
        return cls(**d)


# ---------------------------------------------------------------- 遍历

def _norm_pos(shape, slide_w: int, slide_h: int) -> tuple[float, float, float, float]:
    w = shape.width or 0
    h = shape.height or 0
    left = shape.left or 0
    top = shape.top or 0
    cx = (left + w / 2) / slide_w if slide_w else 0.0
    cy = (top + h / 2) / slide_h if slide_h else 0.0
    return cx, cy, w / slide_w if slide_w else 0.0, h / slide_h if slide_h else 0.0


def _walk_shape(shape, slide_w, slide_h, feats: SlideFeatures,
                image_parts: dict[str, str]):
    """递归处理一个形状。image_parts: rId -> sha1（本 slide 内已处理的图片部件）。"""
    st = shape.shape_type

    if st == MSO_SHAPE_TYPE.GROUP:
        _record_shape(shape, "group", slide_w, slide_h, feats)
        for sub in shape.shapes:
            _walk_shape(sub, slide_w, slide_h, feats, image_parts)
        return

    if st == MSO_SHAPE_TYPE.PICTURE:
        _record_shape(shape, "picture", slide_w, slide_h, feats)
        try:
            blob = shape.image.blob
            _add_image_blob(blob, feats)
        except Exception:
            pass  # 占位符图片等读不出 blob
        return

    if shape.has_table:
        _record_shape(shape, "table", slide_w, slide_h, feats)
        for row in shape.table.rows:
            for cell in row.cells:
                _collect_text(cell.text, shape.shape_id, feats)
        return

    if shape.has_chart:
        _record_shape(shape, "chart", slide_w, slide_h, feats)
        _collect_chart_sig(shape.chart, feats)
        return

    if shape.has_text_frame:
        _record_shape(shape, "textbox" if st == MSO_SHAPE_TYPE.TEXT_BOX else "other",
                      slide_w, slide_h, feats)
        _collect_text_frame(shape, feats)
        return

    _record_shape(shape, "other", slide_w, slide_h, feats)
    # 其他形状若带文字（如自选图形上的标注）
    if shape.has_text_frame:
        _collect_text_frame(shape, feats)


def _record_shape(shape, stype: str, slide_w, slide_h, feats: SlideFeatures):
    cx, cy, w, h = _norm_pos(shape, slide_w, slide_h)
    feats.shape_layout.append(ShapeInfo(
        shape_id=shape.shape_id, name=shape.name or "", stype=stype,
        cx=round(cx, 4), cy=round(cy, 4), w=round(w, 4), h=round(h, 4),
    ))


def _collect_text_frame(shape, feats: SlideFeatures):
    sid = shape.shape_id
    shape_tokens: set[str] = set()
    paragraphs_text: list[str] = []
    for p in shape.text_frame.paragraphs:
        runs_text = "".join(r.text or "" for r in p.runs)
        if not runs_text and p.text:
            runs_text = p.text
        paragraphs_text.append(runs_text)
        shape_tokens |= text_tokens(runs_text)
    full_text = "\n".join(paragraphs_text)
    feats.text_tokens |= shape_tokens
    feats.vague_texts.extend(vague_number_texts(full_text))
    # 数字：上下文 = 同一文本框内全部关键词
    ctx = frozenset(shape_tokens)
    for i, ptext in enumerate(paragraphs_text):
        feats.numbers.extend(extract_numbers(ptext, ctx, sid, i))


def _collect_text(text: str, shape_id: int, feats: SlideFeatures):
    toks = text_tokens(text)
    feats.text_tokens |= toks
    ctx = frozenset(toks)
    for i, line in enumerate(text.split("\n")):
        feats.numbers.extend(extract_numbers(line, ctx, shape_id, i))


def _collect_chart_sig(chart, feats: SlideFeatures):
    from .features import normalize_text
    try:
        ctype = str(chart.chart_type)
        series_names: set[str] = set()
        for s in chart.plots[0].series:
            series_names.add(normalize_text(s.name or ""))
        cats: set[str] = set()
        for plot in chart.plots:
            for c in plot.categories:
                cats.add(normalize_text(str(c)))
        feats.chart_sigs.add((ctype, frozenset(series_names), frozenset(cats)))
        # 图表标题进文本 token
        if chart.has_title:
            feats.text_tokens |= text_tokens(chart.chart_title.text_frame.text)
    except Exception:
        pass


def _add_image_blob(blob: bytes, feats: SlideFeatures):
    h = sha1_of_blob(blob)
    feats.image_hashes.add(h)
    ph = phash_of_blob(blob)
    if ph:
        feats.image_phashes.add(ph)


def _collect_fill_images(slide, image_parts: dict[str, str], feats: SlideFeatures):
    """背景/形状填充里的图片（blipFill），排除 p:pic 内部（那些已按 Picture 处理）。"""
    slide_el = slide._element
    for blip in slide_el.iter(BLIP_QN):
        # 是否位于 p:pic 内
        parent = blip.getparent()
        in_pic = False
        while parent is not None:
            if parent.tag == qn("p:pic"):
                in_pic = True
                break
            parent = parent.getparent()
        if in_pic:
            continue
        rid = blip.get(R_EMBED_QN)
        if not rid:
            continue
        try:
            part = slide.part.related_part(rid)
            blob = part.blob
        except Exception:
            continue
        h = sha1_of_blob(blob)
        image_parts[rid] = h
        feats.image_hashes.add(h)
        ph = phash_of_blob(blob)
        if ph:
            feats.image_phashes.add(ph)


def _collect_smartart_text(slide, feats: SlideFeatures):
    """SmartArt：直接读 diagram 数据部件里的 a:t 文本。"""
    slide_el = slide._element
    found = False
    for gf in slide_el.iter(qn("p:graphicFrame")):
        uri_el = gf.find(".//" + qn("a:graphicData"))
        if uri_el is None or DIAGRAM_URI not in (uri_el.get("uri") or ""):
            continue
        found = True
    if not found:
        return
    try:
        for rel in slide.part.rels.values():
            if "diagramData" in rel.reltype:
                data_part = rel.target_part
                from lxml import etree
                root = etree.fromstring(data_part.blob)
                text = "".join(t.text or "" for t in root.iter(qn("a:t")))
                _collect_text(text, -1, feats)
    except Exception:
        pass


def _is_hidden(slide) -> bool:
    show = slide._element.get("show")
    return show is not None and show == "0"


def extract_deck(path: Path, cfg: dict) -> list[SlideFeatures]:
    prs = Presentation(str(path))
    slide_w = prs.slide_width
    slide_h = prs.slide_height
    boiler_kw = set(cfg.get("extract", {}).get("boilerplate_keywords", []))
    out: list[SlideFeatures] = []
    for idx, slide in enumerate(prs.slides, start=1):
        feats = SlideFeatures(
            deck=path.name, slide_no=idx, slide_id=slide.slide_id,
            hidden=_is_hidden(slide),
            layout_name=slide.slide_layout.name or "",
        )
        image_parts: dict[str, str] = {}
        for shape in slide.shapes:
            _walk_shape(shape, slide_w, slide_h, feats, image_parts)
        _collect_fill_images(slide, image_parts, feats)
        _collect_smartart_text(slide, feats)
        # 通用页判定：形状 < 3 或标题命中通用关键词
        title = ""
        if slide.shapes.title is not None:
            title = slide.shapes.title.text or ""
        n_shapes = len(slide.shapes)
        feats.is_boilerplate = (
            n_shapes < 3
            or any(k in title for k in boiler_kw)
            or any(k in (feats.layout_name or "") for k in {"标题幻灯片", "Title Slide"})
        )
        out.append(feats)
    return out


# ---------------------------------------------------------------- 全局去噪 + 索引

def denoise(all_feats: list[SlideFeatures], cfg: dict) -> tuple[list[SlideFeatures], list[str], list[str]]:
    """出现在超过 denoise_ratio 页面上的图片哈希/token 视为模板噪声。

    返回 (去噪后的特征列表, 被剔除的图片哈希, 被剔除的 token)。
    """
    ratio = cfg.get("extract", {}).get("denoise_ratio", 0.30)
    n = len(all_feats)
    if n == 0:
        return all_feats, [], []
    from collections import Counter
    img_cnt: Counter = Counter()
    tok_cnt: Counter = Counter()
    for f in all_feats:
        img_cnt.update(f.image_hashes)
        tok_cnt.update(f.text_tokens)
    noise_imgs = {h for h, c in img_cnt.items() if c / n > ratio and c > 1}
    noise_toks = {t for t, c in tok_cnt.items() if c / n > ratio and c > 1}
    for f in all_feats:
        f.image_hashes -= noise_imgs
        f.text_tokens -= noise_toks
    return all_feats, sorted(noise_imgs), sorted(noise_toks)


def build_index(pptx_dir: Path, cfg: dict, render_dir: Path | None = None) -> dict:
    """扫描目录下所有 .pptx，生成索引。返回可 json 序列化的 dict。

    render_dir 不为 None 且 LibreOffice 可用时，额外渲染每页 PNG 并计算
    整页 pHash + CLIP 嵌入（方向2/3）；否则只用结构特征。
    """
    decks = sorted(pptx_dir.glob("*.pptx"))
    all_feats: list[SlideFeatures] = []
    render_failed: list[str] = []
    for d in decks:
        try:
            feats = extract_deck(d, cfg)
        except Exception as e:
            render_failed.append(f"{d.name}: {e}")
            continue
        if render_dir is not None:
            _attach_visual(d, feats, cfg, render_dir, render_failed)
        all_feats.extend(feats)
    all_feats, noise_imgs, noise_toks = denoise(all_feats, cfg)
    n_visual = sum(1 for s in all_feats if s.img_embed is not None or s.img_phash)
    return {
        "slides": all_feats,   # SlideFeatures 对象；save_index 时才序列化
        "noise_images": noise_imgs,
        "noise_tokens": noise_toks,
        "failed_decks": render_failed,
        "render_enabled": n_visual > 0,
    }


def _attach_visual(pptx_path: Path, feats: list[SlideFeatures], cfg: dict,
                   render_dir: Path, failed_out: list[str]):
    """渲染一份 PPT 的每页 PNG，填整页 pHash 与 CLIP 嵌入。"""
    from .render import render_deck, soffice_available
    from .features import phash_of_image
    if not soffice_available(cfg):
        return
    deck_dir = render_dir / pptx_path.stem
    res = render_deck(pptx_path, deck_dir, cfg)
    if not res.ok or len(res.png_paths) != len(feats):
        failed_out.append(f"{pptx_path.name}: 渲染失败或页数不符"
                          f"（{len(res.png_paths) if res.ok else res.reason}）")
        return
    for f, png in zip(feats, res.png_paths):
        f.render_png = str(png)
        f.img_phash = phash_of_image(png)
    # 方向3：CLIP 批量嵌入
    try:
        from .visual import clip_available, embed_pngs
        if clip_available(cfg):
            B = 32
            for start in range(0, len(res.png_paths), B):
                chunk = res.png_paths[start:start + B]
                emb = embed_pngs(chunk, cfg)
                for f, row in zip(feats[start:start + B], emb):
                    f.img_embed = [float(x) for x in row]
    except Exception as e:
        failed_out.append(f"{pptx_path.name}: CLIP 嵌入失败，降级 pHash（{e}）")


def save_index(index: dict, out_path: Path):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    data = dict(index)
    data["slides"] = [s.to_jsonable() for s in index["slides"]] \
        if index["slides"] and isinstance(index["slides"][0], SlideFeatures) \
        else index["slides"]
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)


def load_index(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        idx = json.load(f)
    idx["slides"] = [SlideFeatures.from_jsonable(d) for d in idx["slides"]]
    return idx
