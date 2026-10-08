"""方向2：LibreOffice 渲染兜底。

soffice --headless --convert-to pdf 把整份 PPT 转 PDF，
再用 pymupdf 按页转 PNG（宽 640px），并计算整页 pHash。
页数与 python-pptx 读到的不一致时标记该文件渲染失败，只用结构特征。

Windows 下 soffice 不在 PATH 时的常见安装路径会自动探测，
也可在 config.yaml 的 render.soffice_path 显式指定。
"""
from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass, field
from pathlib import Path

# Windows 常见安装位置（soffice 不在 PATH 时依次探测）
_WIN_SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


@dataclass
class RenderResult:
    ok: bool
    reason: str = ""
    png_paths: list[Path] = field(default_factory=list)


def soffice_available(cfg: dict | None = None) -> bool:
    return find_soffice(cfg) is not None


def find_soffice(cfg: dict | None = None) -> str | None:
    """定位 soffice 可执行文件：config 指定 > PATH > 常见安装路径。"""
    explicit = (cfg or {}).get("render_visual", {}).get("soffice_path")
    if explicit and Path(explicit).exists():
        return explicit
    found = shutil.which("soffice")
    if found:
        return found
    # macOS 上 soffice 不在 PATH 时
    mac = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if Path(mac).exists():
        return mac
    for cand in _WIN_SOFFICE_CANDIDATES:
        if Path(cand).exists():
            return cand
    return None


def render_deck(pptx_path: Path, out_dir: Path, cfg: dict | None = None) -> RenderResult:
    """渲染失败/不可用时返回 ok=False，调用方降级为纯结构特征。"""
    soffice = find_soffice(cfg)
    if not soffice:
        return RenderResult(ok=False, reason="libreoffice_unavailable")
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        proc = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", str(out_dir), str(pptx_path)],
            capture_output=True, timeout=300,
        )
        pdf = out_dir / (pptx_path.stem + ".pdf")
        if proc.returncode != 0 or not pdf.exists():
            err = (proc.stderr or b"").decode(errors="replace")[:200]
            return RenderResult(ok=False, reason=f"convert_failed: {err}")
        width = (cfg or {}).get("render_visual", {}).get("png_width", 640)
        pngs = _pdf_to_pngs(pdf, out_dir, width)
        return RenderResult(ok=True, png_paths=pngs)
    except Exception as e:
        return RenderResult(ok=False, reason=f"render_error: {e}")


def _pdf_to_pngs(pdf: Path, out_dir: Path, width: int = 640) -> list[Path]:
    import fitz  # pymupdf
    doc = fitz.open(str(pdf))
    pngs: list[Path] = []
    try:
        for i, page in enumerate(doc, start=1):
            zoom = width / page.rect.width
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            p = out_dir / f"{pdf.stem}_page{i:03d}.png"
            pix.save(str(p))
            pngs.append(p)
    finally:
        doc.close()
    return pngs
