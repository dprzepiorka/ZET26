from __future__ import annotations

from src.powerfactory_interface import PHASES


def _safe_mean(values: list[float]) -> float:
    return sum(values) / len(values) if values else 0.0


def calculate_indicators(results_raw) -> dict:
    node_voltages = results_raw["node_voltages"]
    all_u = [float(row[phase]) for row in node_voltages for phase in PHASES]
    udev = [abs(u - 1.0) for u in all_u]
    d_u_node = [max(float(row[p]) for p in PHASES) - min(float(row[p]) for p in PHASES) for row in node_voltages]

    ku2_values = []
    for row in node_voltages:
        vals = [float(row[p]) for p in PHASES]
        u_avg = _safe_mean(vals)
        max_dev = max(abs(v - u_avg) for v in vals) if vals else 0.0
        ku2_values.append(0.0 if u_avg <= 1e-9 else 100.0 * max_dev / u_avg)

    tr_i = results_raw["transformer"]["i_a"]
    tr_p = results_raw["transformer"]["p_kw"]
    i_vals = [float(tr_i[p]) for p in PHASES]
    i_avg = _safe_mean(i_vals)

    pv_q_abs = [abs(float(x["q_kvar"])) for x in results_raw["pv_setpoints"]]
    storage = {entry["phase"]: float(entry["p_kw"]) for entry in results_raw["storage_setpoints"]}
    losses = sum(float(item["p_loss_kw"]) for item in results_raw["branch_losses"])
    violations = results_raw.get("constraint_violations", [])

    indicators = {
        "Umax_pu": max(all_u),
        "Umin_pu": min(all_u),
        "Udev_mean_pu": _safe_mean(udev),
        "Udev_max_pu": max(udev) if udev else 0.0,
        "dU_phase_max_pu": max(d_u_node) if d_u_node else 0.0,
        "kU2_max_percent": max(ku2_values) if ku2_values else 0.0,
        "kU2_mean_percent": _safe_mean(ku2_values),
        "I_tr_L1_A": float(tr_i["L1"]),
        "I_tr_L2_A": float(tr_i["L2"]),
        "I_tr_L3_A": float(tr_i["L3"]),
        "I_unbalance_tr_percent": 0.0 if i_avg <= 1e-9 else 100.0 * (max(i_vals) - min(i_vals)) / i_avg,
        "P_tr_L1_kW": float(tr_p["L1"]),
        "P_tr_L2_kW": float(tr_p["L2"]),
        "P_tr_L3_kW": float(tr_p["L3"]),
        "P_export_total_kW": sum(max(0.0, -float(tr_p[phase])) for phase in PHASES),
        "P_loss_total_kW": losses,
        "Q_pv_abs_sum_kvar": sum(pv_q_abs),
        "Q_pv_abs_max_kvar": max(pv_q_abs) if pv_q_abs else 0.0,
        "P_storage_L1_kW": storage["L1"],
        "P_storage_L2_kW": storage["L2"],
        "P_storage_L3_kW": storage["L3"],
        "P_storage_total_kW": sum(storage.values()),
        "constraint_violations": len(violations),
    }
    return indicators
