from __future__ import annotations

import csv
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parent.parent
RESULTS_DIR = ROOT_DIR / "results"


def _write_rows(path: Path, rows: list[dict], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def _write_single_row(path: Path, row: dict) -> None:
    _write_rows(path, [row], list(row.keys()))


def export_case_results(case_name, results_raw, indicators) -> None:
    out_dir = RESULTS_DIR / case_name
    out_dir.mkdir(parents=True, exist_ok=True)

    _write_rows(out_dir / "node_voltages.csv", results_raw["node_voltages"], ["node", "L1", "L2", "L3"])

    tr = results_raw["transformer"]
    tr_rows = [{"phase": p, "p_kw": tr["p_kw"][p], "i_a": tr["i_a"][p], "u_pu": tr["u_pu"][p]} for p in ("L1", "L2", "L3")]
    _write_rows(out_dir / "transformer_phase_results.csv", tr_rows, ["phase", "p_kw", "i_a", "u_pu"])

    _write_rows(out_dir / "pv_setpoints.csv", results_raw["pv_setpoints"], ["inverter", "phase", "q_kvar"])
    _write_rows(out_dir / "storage_setpoints.csv", results_raw["storage_setpoints"], ["phase", "p_kw", "q_kvar"])
    _write_rows(out_dir / "branch_losses.csv", results_raw["branch_losses"], ["branch", "p_loss_kw"])
    _write_single_row(out_dir / "indicators.csv", indicators)
