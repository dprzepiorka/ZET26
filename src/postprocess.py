from __future__ import annotations

import csv
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parent.parent
SUMMARY_DIR = ROOT_DIR / "results" / "summary"

MINIMIZED_METRICS = [
    "Udev_mean_pu",
    "Udev_max_pu",
    "dU_phase_max_pu",
    "kU2_max_percent",
    "kU2_mean_percent",
    "I_unbalance_tr_percent",
    "P_export_total_kW",
    "P_loss_total_kW",
    "constraint_violations",
]


def _safe_pct(base: float, case: float) -> float:
    if abs(base) <= 1e-9:
        return 0.0
    return (base - case) / base * 100.0


def _write_csv(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        return
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def generate_summary(all_indicators: dict[str, dict]) -> None:
    summary_rows = [{"case_name": case, **metrics} for case, metrics in all_indicators.items()]
    _write_csv(SUMMARY_DIR / "indicators_all_cases.csv", summary_rows)

    base = all_indicators["base_no_control"]
    pso = all_indicators.get("pso_global")

    comparison_base = []
    for case_name, metrics in all_indicators.items():
        if case_name == "base_no_control":
            continue
        row = {"case_name": case_name}
        for m in MINIMIZED_METRICS:
            row[f"{m}_improvement_vs_base_percent"] = _safe_pct(float(base[m]), float(metrics[m]))
        comparison_base.append(row)
    _write_csv(SUMMARY_DIR / "comparison_to_base.csv", comparison_base)

    comparison_pso = []
    if pso:
        for case_name, metrics in all_indicators.items():
            if case_name in {"base_no_control", "pso_global"}:
                continue
            row = {"case_name": case_name}
            for m in MINIMIZED_METRICS:
                denom = float(base[m]) - float(pso[m])
                if abs(denom) <= 1e-9:
                    row[f"{m}_effectiveness_vs_pso_percent"] = 0.0
                else:
                    row[f"{m}_effectiveness_vs_pso_percent"] = (float(base[m]) - float(metrics[m])) / denom * 100.0
            comparison_pso.append(row)
    _write_csv(SUMMARY_DIR / "comparison_to_pso.csv", comparison_pso)
