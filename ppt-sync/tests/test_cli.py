# -*- coding: utf-8 -*-
"""CLI 六个命令 smoke 测试 + find 响应时间。"""
import json
import time

import pytest

from ppt_sync import cli


@pytest.fixture(scope="module")
def cli_run(samples, tmp_path_factory):
    work = tmp_path_factory.mktemp("work")
    def run(*argv):
        cli.main(list(argv))
    return run, work, samples


def test_index_and_match(cli_run, capsys):
    run, work, samples = cli_run
    cli.main(["index", str(samples), "--out", str(work / "index.json")])
    cli.main(["match", "--index", str(work / "index.json"), "--out", str(work)])
    assert (work / "matches.xlsx").exists()
    m = json.loads((work / "matches.json").read_text(encoding="utf-8"))
    assert any("PPT-3.pptx#13" in g for g in m["groups"])


def test_find_performance(cli_run, capsys):
    """find 须在 3 秒内返回（方案 A-5）。"""
    run, work, samples = cli_run
    t0 = time.time()
    cli.main(["find", str(samples / "PPT-1.pptx"), "5", "--work", str(work)])
    assert time.time() - t0 < 3.0
    out = capsys.readouterr().out
    assert "PPT-3.pptx" in out and "13" in out


def test_check_reports_inconsistency(cli_run):
    run, work, samples = cli_run
    with pytest.raises(SystemExit) as e:
        cli.main(["check", str(samples)])
    assert e.value.code == 1


def test_render_and_check_pass(samples, tagged, tmp_path_factory, monkeypatch):
    tagged = tagged[0]
    """render 输出目录的 check 应通过（数值一致）。"""
    import pandas as pd
    from ppt_sync.render_data import render_all
    out = tmp_path_factory.mktemp("rend")
    data = tmp_path_factory.mktemp("d") / "data.xlsx"
    metrics = pd.DataFrame([
        {"metric_key": "case_count", "值": 156, "单位": "个", "显示格式": "{:,.0f}"},
        {"metric_key": "sales_total", "值": 3.8, "单位": "亿元", "显示格式": "{:.1f}"},
        {"metric_key": "partner_count", "值": 52, "单位": "家", "显示格式": "{:,.0f}"},
    ])
    metrics.to_excel(data, index=False, sheet_name="metrics")
    render_all(tagged, data, out)
    cli.main(["check", str(out)])   # 不抛 SystemExit 即通过
