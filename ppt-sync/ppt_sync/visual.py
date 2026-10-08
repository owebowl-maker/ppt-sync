"""方向3：CV 模型（轻量 CLIP，ONNX）整页比对两张幻灯片是否是同一页。

只使用视觉编码器（vision_model_quantized.onnx，~89MB），纯 CPU 推理：
  渲染 PNG → CLIP 预处理（224×224 中心裁剪 + CLIP 均值方差）→ 512 维嵌入 → L2 归一化
两页余弦相似度即为 s_visual。全部本地运行，不上传数据。

模型文件放在 ppt-sync/models/ 下，随工具一起拷贝；
缺失或未安装 onnxruntime 时自动降级为整页 pHash（s_visual 仍可用，区分度略低）。
"""
from __future__ import annotations

import io
from pathlib import Path

import numpy as np
from PIL import Image

# CLIP 官方预处理参数
IMG_SIZE = 224
CLIP_MEAN = np.array([0.48145466, 0.4578275, 0.40821073], dtype=np.float32)
CLIP_STD = np.array([0.26862954, 0.26130258, 0.27577711], dtype=np.float32)

DEFAULT_MODEL_DIR = Path(__file__).resolve().parent.parent / "models"
# 依次尝试：fp32 → 量化版
_MODEL_FILES = ["vision_model.onnx", "vision_model_quantized.onnx"]

_session = None
_input_name: str | None = None


def model_path(cfg: dict | None = None) -> Path | None:
    d = Path((cfg or {}).get("visual", {}).get("model_dir") or DEFAULT_MODEL_DIR)
    for name in _MODEL_FILES:
        p = d / name
        if p.exists():
            return p
    return None


def clip_available(cfg: dict | None = None) -> bool:
    try:
        import onnxruntime  # noqa: F401
    except ImportError:
        return False
    return model_path(cfg) is not None


def _get_session(cfg: dict | None):
    global _session, _input_name
    if _session is not None:
        return _session
    import onnxruntime as ort
    p = model_path(cfg)
    if p is None:
        raise FileNotFoundError("未找到 CLIP 视觉编码器 ONNX 模型")
    _session = ort.InferenceSession(str(p), providers=["CPUExecutionProvider"])
    _input_name = _session.get_inputs()[0].name
    return _session


def preprocess(png_path: Path) -> np.ndarray:
    """PNG → (1,3,224,224) float32，CLIP 标准预处理。"""
    img = Image.open(png_path).convert("RGB")
    # 短边缩放到 224（双三次插值，与 CLIP 原实现一致）
    w, h = img.size
    scale = IMG_SIZE / min(w, h)
    img = img.resize((round(w * scale), round(h * scale)), Image.BICUBIC)
    # 中心裁剪 224×224
    w, h = img.size
    left, top = (w - IMG_SIZE) // 2, (h - IMG_SIZE) // 2
    img = img.crop((left, top, left + IMG_SIZE, top + IMG_SIZE))
    x = np.asarray(img, dtype=np.float32) / 255.0
    x = (x - CLIP_MEAN) / CLIP_STD
    return x.transpose(2, 0, 1)[np.newaxis]  # HWC → NCHW


def embed_pngs(png_paths: list[Path], cfg: dict | None = None) -> np.ndarray:
    """批量嵌入，返回 (N, 512) L2 归一化矩阵。调用前先用 clip_available() 判断。"""
    sess = _get_session(cfg)
    batch = np.concatenate([preprocess(p) for p in png_paths], axis=0)
    out = sess.run(None, {_input_name: batch})[0]
    emb = np.asarray(out, dtype=np.float32).reshape(len(png_paths), -1)
    # L2 归一化（余弦相似度 = 点积）
    norms = np.linalg.norm(emb, axis=1, keepdims=True)
    norms[norms == 0] = 1.0
    return emb / norms


def phash_similarity(h1: str | None, h2: str | None) -> float | None:
    """整页 pHash 相似度（无 CV 模型时的降级）。"""
    if not h1 or not h2:
        return None
    from .features import phash_hamming
    d = phash_hamming(h1, h2)
    return max(0.0, 1.0 - d / 64.0)
