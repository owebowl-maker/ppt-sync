# -*- coding: utf-8 -*-
"""生成 data.xlsx 数据源模板（阶段 C 用）。"""
import sys
from pathlib import Path

import pandas as pd

TEMPLATE = pd.DataFrame([
    {"metric_key": "case_count", "指标名称": "合作案例数", "值": 128, "单位": "个",
     "显示格式": "{:,.0f}", "更新周期": "季度", "负责人": "", "数据来源": ""},
    {"metric_key": "sales_total", "指标名称": "销售额", "值": 3.2, "单位": "亿元",
     "显示格式": "{:.1f}", "更新周期": "季度", "负责人": "", "数据来源": ""},
    {"metric_key": "partner_count", "指标名称": "合作厂家数", "值": 45, "单位": "家",
     "显示格式": "{:,.0f}", "更新周期": "季度", "负责人": "", "数据来源": ""},
    {"metric_key": "report_period", "指标名称": "更新时间", "值": "2026Q3", "单位": "",
     "显示格式": "", "更新周期": "季度", "负责人": "", "数据来源": ""},
])


def main(out: str = "data.xlsx"):
    out_path = Path(out)
    if out_path.exists():
        sys.exit(f"{out_path} 已存在，不覆盖。")
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as w:
        TEMPLATE.to_excel(w, sheet_name="metrics", index=False)
        # 图表数据模板：第一列为类别，之后各列为系列；sheet 名 = chart_<metric_key>
        pd.DataFrame({"季度": ["2025Q3", "2025Q4", "2026Q1", "2026Q2"],
                      "销售额(亿元)": [2.1, 2.4, 2.8, 3.2]}).to_excel(
            w, sheet_name="chart_sales_total", index=False)
    print(f"模板已生成：{out_path}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "data.xlsx")
