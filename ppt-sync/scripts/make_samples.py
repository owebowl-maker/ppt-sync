# -*- coding: utf-8 -*-
"""生成合成样例 PPT，模拟"同类页散落在多份 PPT"的场景。

对应关系（用于自动断言）：
  PPT-1#3  ==  PPT-2#2   （销售额页，数值不同 → 应报数值不一致）
  PPT-1#5  ==  PPT-3#13  （合作案例页，完全相同 → 已知同类页）
  PPT-1#4  ==  PPT-2#3   （合作厂家页，Logo 相同）
"""
from __future__ import annotations

import io
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

OUT = Path(__file__).resolve().parent.parent / "work" / "samples"
OUT.mkdir(parents=True, exist_ok=True)

# ---------- 素材图片（确定性生成） ----------

def make_png(color, label: str, size=(240, 160), pattern="border") -> bytes:
    img = Image.new("RGB", size, color)
    d = ImageDraw.Draw(img)
    if pattern == "border":
        d.rectangle([10, 10, size[0] - 10, size[1] - 10], outline=(255, 255, 255), width=3)
        d.text((20, size[1] // 2 - 5), label, fill=(255, 255, 255))
    elif pattern == "diagonal":
        for i in range(0, size[0] + size[1], 24):
            d.line([(i, 0), (i - size[1], size[1])], fill=(255, 255, 255), width=6)
        d.text((size[0] // 4, size[1] // 2 - 5), label, fill=(0, 0, 0))
    elif pattern == "circle":
        d.ellipse([size[0] // 4, size[1] // 6, size[0] * 3 // 4, size[1] * 5 // 6],
                  outline=(255, 255, 255), width=8)
        d.text((size[0] // 3, size[1] // 2 - 5), label, fill=(255, 255, 255))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


LOGO = make_png((30, 90, 200), "COMPANY LOGO", pattern="border")
CASE_PHOTO = make_png((200, 120, 40), "CASE PHOTO", size=(420, 300), pattern="diagonal")
CASE_PHOTO2 = make_png((40, 160, 90), "CASE B", size=(420, 300), pattern="circle")

# ---------- 页面构造辅助 ----------

def add_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s.shapes.add_textbox(Inches(2), Inches(2), Inches(6), Inches(1))
    tb.text_frame.text = "公司介绍 封面"
    return s


def add_toc(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(4))
    tf = tb.text_frame
    tf.text = "目录"
    for line in ("1. 业绩概览", "2. 合作厂家", "3. 合作案例"):
        p = tf.add_paragraph()
        p.text = line
    return s


def add_sales_page(prs, value: str):
    """销售额页：图表 + 数字。不同 deck 传不同数值制造不一致。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    title = s.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(6), Inches(0.6))
    title.text_frame.text = "业绩概览"
    big = s.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(4), Inches(1))
    big.text_frame.text = f"销售额 {value}"
    chart_data = CategoryChartData()
    chart_data.categories = ["2025Q3", "2025Q4", "2026Q1", "2026Q2"]
    chart_data.add_series("销售额(亿元)", (2.1, 2.4, 2.8, 3.2))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                            Inches(0.5), Inches(2.2), Inches(8), Inches(4.5), chart_data)
    gf.chart.has_legend = False
    return s


def add_partner_page(prs, count: str):
    """合作厂家页：共享 LOGO 图。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    title = s.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(6), Inches(0.6))
    title.text_frame.text = "合作厂家"
    s.shapes.add_picture(io.BytesIO(LOGO), Inches(0.5), Inches(1.2), Inches(2), Inches(1.33))
    tb = s.shapes.add_textbox(Inches(3), Inches(2.5), Inches(5), Inches(1))
    tb.text_frame.text = f"累计合作厂家 {count}"
    return s


def add_case_page(prs, count: str, photo):
    """合作案例页：案例照片 + 数字。跨文件同类页的核心。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    title = s.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(6), Inches(0.6))
    title.text_frame.text = "合作案例"
    s.shapes.add_picture(io.BytesIO(photo), Inches(0.5), Inches(1.2), Inches(3), Inches(2))
    tb = s.shapes.add_textbox(Inches(4), Inches(1.5), Inches(5), Inches(2))
    tf = tb.text_frame
    tf.text = "最新合作案例数"
    p = tf.add_paragraph()
    p.text = f"{count} 个成功落地"
    # 组合形状：案例清单
    from pptx.enum.shapes import MSO_SHAPE
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                             Inches(0.5), Inches(4), Inches(8.5), Inches(1.8))
    box.text_frame.text = "重点案例：某制造企业数字化改造项目、某零售集团供应链项目"
    return s


def add_filler_page(prs, n: int):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(4))
    tb.text_frame.text = f"填充页 {n}：业务板块介绍"
    return s


def add_group_page(prs):
    """含嵌套组合形状的页，测试递归遍历。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    from pptx.enum.shapes import MSO_SHAPE
    outer = s.shapes.add_group_shape()
    inner1 = outer.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
    inner1.text_frame.text = "服务客户 200 家"
    inner2 = outer.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4), Inches(1), Inches(2), Inches(1))
    inner2.text_frame.text = "覆盖城市 30 个"
    return s


# ---------- 三份 PPT ----------

def build_ppt1():
    prs = Presentation()
    add_cover(prs)
    add_toc(prs)
    add_sales_page(prs, "3.2亿元")
    add_partner_page(prs, "45家")
    add_case_page(prs, "128", CASE_PHOTO)
    add_group_page(prs)
    return prs


def build_ppt2():
    prs = Presentation()
    add_cover(prs)
    add_sales_page(prs, "2.9亿元")      # 与 PPT-1#3 同页不同数 → 不一致
    add_partner_page(prs, "45家")       # 与 PPT-1#4 同类
    add_filler_page(prs, 1)
    return prs


def build_ppt3():
    prs = Presentation()
    add_cover(prs)
    for i in range(1, 12):              # 11 页填充，让"案例页"落在第 13 页
        add_filler_page(prs, i)
    add_case_page(prs, "128", CASE_PHOTO)   # == PPT-1#5（第 13 页）
    return prs


def main(out_dir: Path | None = None):
    out = Path(out_dir) if out_dir else OUT
    out.mkdir(parents=True, exist_ok=True)
    paths = {}
    for name, builder in [("PPT-1.pptx", build_ppt1),
                          ("PPT-2.pptx", build_ppt2),
                          ("PPT-3.pptx", build_ppt3)]:
        p = out / name
        builder().save(str(p))
        paths[name] = p
        print(f"已生成 {p}")
    return paths


if __name__ == "__main__":
    main()


if __name__ == "__main__":
    main()
