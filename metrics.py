from __future__ import annotations

from typing import Dict, List


MINIMIZE_METRICS = {
    "Udev_mean_pu",
    "Udev_max_pu",
    "dU_phase_max_pu",
    "kU2_max_percent",
    "kU2_mean_percent",
    "I_unbalance_tr_percent",
    "P_export_total_kW",
    "P_loss_total_kW",
    "constraint_violations",
}


def _all_phase_values(rows: List[Dict], keys=("L1", "L2", "L3")) -> List[float]:
    vals: List[float] = []
    for row in rows:
        for k in keys:
            v = row.get(k)
            if isinstance(v, (int, float)):
                vals.append(float(v))
    return vals


def calculate_indicators(results_raw: Dict) -> Dict[str, float]:
    node_voltages = results_raw.get("node_voltages", [])
    tr = results_raw.get("transformer_phase_results", [{}])[0] if results_raw.get("transformer_phase_results") else {}
    pv_setpoints = results_raw.get("pv_setpoints", [])
    storage_setpoints = results_raw.get("storage_setpoints", [])
    branch_losses = results_raw.get("branch_losses", [])

    u_vals = _all_phase_values(node_voltages)
    if not u_vals:
        raise ValueError("Brak fazowych napięć w wynikach (node_voltages).")

    u_dev = [abs(u - 1.0) for u in u_vals]
    d_u_phase = []
    ku2_list = []
    for row in node_voltages:
        phases = [row.get("L1"), row.get("L2"), row.get("L3")]
        if all(isinstance(x, (int, float)) for x in phases):
            p = [float(x) for x in phases]
            avg = sum(p) / 3.0
            d_u_phase.append(max(p) - min(p))
            ku2_list.append(0.0 if avg == 0 else (max(abs(x - avg) for x in p) / avg) * 100.0)

    i_l1 = float(tr.get("I_tr_L1_A", 0.0))
    i_l2 = float(tr.get("I_tr_L2_A", 0.0))
    i_l3 = float(tr.get("I_tr_L3_A", 0.0))
    i_avg = (i_l1 + i_l2 + i_l3) / 3.0
    i_unbalance = 0.0 if i_avg == 0 else (max(i_l1, i_l2, i_l3) - i_avg) / i_avg * 100.0

    p_l1 = float(tr.get("P_tr_L1_kW", 0.0))
    p_l2 = float(tr.get("P_tr_L2_kW", 0.0))
    p_l3 = float(tr.get("P_tr_L3_kW", 0.0))
    p_export = max(0.0, -(p_l1 + p_l2 + p_l3))

    q_abs = [abs(float(r.get("Q_kvar", 0.0))) for r in pv_setpoints]

    p_storage_l1 = sum(float(r.get("P_L1_kW", 0.0)) for r in storage_setpoints)
    p_storage_l2 = sum(float(r.get("P_L2_kW", 0.0)) for r in storage_setpoints)
    p_storage_l3 = sum(float(r.get("P_L3_kW", 0.0)) for r in storage_setpoints)

    out = {
        "Umax_pu": max(u_vals),
        "Umin_pu": min(u_vals),
        "Udev_mean_pu": sum(u_dev) / len(u_dev),
        "Udev_max_pu": max(u_dev),
        "dU_phase_max_pu": max(d_u_phase) if d_u_phase else 0.0,
        "kU2_max_percent": max(ku2_list) if ku2_list else 0.0,
        "kU2_mean_percent": (sum(ku2_list) / len(ku2_list)) if ku2_list else 0.0,
        "I_tr_L1_A": i_l1,
        "I_tr_L2_A": i_l2,
        "I_tr_L3_A": i_l3,
        "I_unbalance_tr_percent": i_unbalance,
        "P_tr_L1_kW": p_l1,
        "P_tr_L2_kW": p_l2,
        "P_tr_L3_kW": p_l3,
        "P_export_total_kW": p_export,
        "P_loss_total_kW": sum(float(r.get("P_loss_kW", 0.0)) for r in branch_losses),
        "Q_pv_abs_sum_kvar": sum(q_abs),
        "Q_pv_abs_max_kvar": max(q_abs) if q_abs else 0.0,
        "P_storage_L1_kW": p_storage_l1,
        "P_storage_L2_kW": p_storage_l2,
        "P_storage_L3_kW": p_storage_l3,
        "P_storage_total_kW": p_storage_l1 + p_storage_l2 + p_storage_l3,
        "constraint_violations": float(len(results_raw.get("violations", []))),
    }
    return out


def _ratio(base: float, case: float) -> float:
    if base == 0:
        return 0.0
    return (base - case) / base * 100.0


def comparison_to_base(all_indicators: Dict[str, Dict[str, float]], base_case: str = "base_no_control") -> List[Dict[str, float]]:
    if base_case not in all_indicators:
        return []
    base = all_indicators[base_case]
    rows: List[Dict[str, float]] = []
    for case_name, ind in all_indicators.items():
        if case_name == base_case:
            continue
        row = {"case": case_name}
        for m in MINIMIZE_METRICS:
            row[f"{m}_improvement_vs_base_percent"] = _ratio(float(base.get(m, 0.0)), float(ind.get(m, 0.0)))
        rows.append(row)
    return rows


def comparison_to_pso(all_indicators: Dict[str, Dict[str, float]], base_case: str = "base_no_control", pso_case: str = "pso_global") -> List[Dict[str, float]]:
    if base_case not in all_indicators or pso_case not in all_indicators:
        return []
    base = all_indicators[base_case]
    pso = all_indicators[pso_case]
    rows: List[Dict[str, float]] = []
    for case_name, ind in all_indicators.items():
        if case_name in (base_case, pso_case):
            continue
        row = {"case": case_name}
        for m in MINIMIZE_METRICS:
            base_m = float(base.get(m, 0.0))
            pso_m = float(pso.get(m, 0.0))
            local_m = float(ind.get(m, 0.0))
            denom = base_m - pso_m
            row[f"{m}_effectiveness_vs_pso_percent"] = 0.0 if denom == 0 else (base_m - local_m) / denom * 100.0
        rows.append(row)
    return rows
