# -*- coding: utf-8 -*-
"""测试共享 fixture：生成一次合成样例，供各测试模块使用。"""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

from make_samples import main as make_samples_main  # noqa: E402
from ppt_sync.features import load_config  # noqa: E402

CONFIG = load_config(str(ROOT / "config.yaml"))


@pytest.fixture(scope="session")
def samples(tmp_path_factory):
    out = tmp_path_factory.mktemp("samples")
    make_samples_main(out_dir=out)
    return out


@pytest.fixture(scope="session")
def index(samples, tmp_path_factory):
    from ppt_sync.extract import build_index
    return build_index(samples, CONFIG)


@pytest.fixture(scope="session")
def match_result(index):
    from ppt_sync.match import match_all, cluster
    pairs = match_all(index, CONFIG)
    groups = cluster(pairs, index["slides"], CONFIG)
    return pairs, groups


# 已知对应关系（与 make_samples.py 中的设计一致）
KNOWN_PAIRS = {
    ("PPT-1.pptx", 3): ("PPT-2.pptx", 2),    # 销售额页（数值不同 → 不一致）
    ("PPT-1.pptx", 4): ("PPT-2.pptx", 3),    # 合作厂家页
    ("PPT-1.pptx", 5): ("PPT-3.pptx", 13),   # 合作案例页（核心已知同类页）
}


@pytest.fixture(scope="session")
def tagged(samples, index, tmp_path_factory):
    """阶段 B 打标输出（确认表按 KNOWN_PAIRS 人工填写）。"""
    import pandas as pd
    from ppt_sync.tag import tag_all
    out = tmp_path_factory.mktemp("tagged")
    rows = [
        {"group_id": 0, "成员": "PPT-1.pptx#3, PPT-2.pptx#2",
         "确认(Y/N)": "Y", "指标key(英文)": "sales_total", "指标名称": "销售额"},
        {"group_id": 1, "成员": "PPT-1.pptx#4, PPT-2.pptx#3",
         "确认(Y/N)": "Y", "指标key(英文)": "partner_count", "指标名称": "合作厂家数"},
        {"group_id": 2, "成员": "PPT-1.pptx#5, PPT-3.pptx#13",
         "确认(Y/N)": "Y", "指标key(英文)": "case_count", "指标名称": "合作案例数"},
    ]
    mapping = tmp_path_factory.mktemp("conf") / "confirmed.xlsx"
    pd.DataFrame(rows).to_excel(mapping, index=False)
    written, registry, report = tag_all(samples, mapping, out, index)
    return out, registry, report
