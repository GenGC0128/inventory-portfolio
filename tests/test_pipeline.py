import shutil
import subprocess
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]


def test_end_to_end(tmp_path):
    data, out = tmp_path / "reports", tmp_path / "out"
    shutil.copytree(ROOT / "data" / "reports", data)

    r = subprocess.run(
        [sys.executable, str(ROOT / "inventory_pipeline.py"),
         "--input", str(data), "--config", str(ROOT / "config" / "normalization.json"),
         "--out", str(out)],
        capture_output=True, text=True,
    )
    assert r.returncode == 0, r.stderr

    summary = pd.read_excel(out / "summary.xlsx")
    assert len(summary) > 0
    assert {"date", "base", "identify", "quantity", "spec_std", "supplier_std"} <= set(summary.columns)

    # 三种日期模式（datetime / serial / 行内日期）都应解析成功
    assert summary["date"].notna().all()

    # 序列日期应落在 2024-03 附近而非 1900 年
    assert summary["date"].dt.year.min() >= 2024

    # 口径归一化生效
    assert {"dense", "loose"} <= set(summary["density"].unique())

    # 损坏文件被隔离
    failed = pd.read_csv(out / "failed.csv")
    assert "corrupted.xlsx" in set(failed["file"])
