from __future__ import annotations

import math


def _clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def _node_voltage_map(node_voltages: list[dict]) -> dict:
    return {row["node"]: row for row in node_voltages}


def _qu_from_curve(u_pu: float, qmax: float, curve: dict) -> float:
    u_min_full = float(curve["u_min_full"])
    u_min_deadband = float(curve["u_min_deadband"])
    u_max_deadband = float(curve["u_max_deadband"])
    u_max_full = float(curve["u_max_full"])

    if u_pu <= u_min_full:
        return -qmax
    if u_min_full < u_pu < u_min_deadband:
        ratio = (u_pu - u_min_full) / (u_min_deadband - u_min_full)
        return -qmax * (1.0 - ratio)
    if u_min_deadband <= u_pu <= u_max_deadband:
        return 0.0
    if u_max_deadband < u_pu < u_max_full:
        ratio = (u_pu - u_max_deadband) / (u_max_full - u_max_deadband)
        return qmax * ratio
    return qmax


def apply_pv_qu_control(case_data, pf_objects, params) -> None:
    node_voltages = _node_voltage_map(case_data["node_voltages"])
    warnings = case_data.setdefault("warnings", [])
    for pv in pf_objects["pv_data"]:
        s_inv = float(pv["s_kva"])
        p_pv = float(pv["p_kw"])
        if p_pv > s_inv:
            warnings.append(f"PV {pv['name']} {pv['phase']}: P({p_pv}) > S({s_inv}); P clipped to S.")
            p_pv = s_inv
        qmax_available = math.sqrt(max(s_inv * s_inv - p_pv * p_pv, 0.0))
        u_pu = float(node_voltages[pv["node"]][pv["phase"]])
        q_curve = _qu_from_curve(u_pu, qmax_available, params["qu_curve"])
        q_command = _clamp(q_curve, -qmax_available, qmax_available)
        pf_objects["interface"].set_pv_q_setpoint(pv["name"], pv["phase"], q_command)
