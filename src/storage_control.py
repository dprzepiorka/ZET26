from __future__ import annotations

from src.powerfactory_interface import PHASES


def _clamp(value: int | float, low: int | float, high: int | float):
    return max(low, min(high, value))


def _step_from_export(p_export_kw: float, p_start_kw: float, p_phase_max_kw: float) -> int:
    if p_export_kw <= p_start_kw:
        return 0
    if p_export_kw <= 0.25 * p_phase_max_kw:
        return 1
    if p_export_kw <= 0.50 * p_phase_max_kw:
        return 2
    if p_export_kw <= 0.75 * p_phase_max_kw:
        return 3
    return 4


def _unbalance_step(di: float, threshold_1: float, threshold_2: float) -> int:
    if di <= threshold_1:
        return 0
    if di <= threshold_2:
        return 1
    return 2


def _step_to_p(step: int, p_phase_max_kw: float) -> float:
    return float(step) / 4.0 * p_phase_max_kw


def apply_storage_tr_control(case_data, pf_objects, params) -> None:
    storage_cfg = params["storage"]
    tr_rule = params["transformer_rule"]
    p_phase_max_kw = float(storage_cfg["p_phase_max_kw"])
    s_phase_kva = float(storage_cfg["s_phase_kva"])
    p_start_kw = float(storage_cfg["p_start_kw"])
    th1 = float(tr_rule["current_unbalance_threshold_1"])
    th2 = float(tr_rule["current_unbalance_threshold_2"])

    transformer_data = case_data["transformer"] if "transformer" in case_data else case_data
    tr_p = transformer_data.get("p_kw")
    if tr_p is None:
        tr_p = pf_objects["interface"].get_transformer_phase_powers(prefer_direct=False)
    tr_i = transformer_data["i_a"]
    i_avg = sum(float(tr_i[p]) for p in PHASES) / len(PHASES)

    for phase in PHASES:
        p_export_phase = max(0.0, -float(tr_p[phase]))
        step_export = _step_from_export(p_export_phase, p_start_kw, p_phase_max_kw)
        di = 0.0 if i_avg <= 1e-9 else (float(tr_i[phase]) - i_avg) / i_avg
        step_unbalance = _unbalance_step(di, th1, th2)
        step_final = int(_clamp(step_export + step_unbalance, -4, 4))
        p_phase = _step_to_p(step_final, p_phase_max_kw)
        p_phase = float(_clamp(p_phase, -p_phase_max_kw, p_phase_max_kw))
        p_phase = float(_clamp(p_phase, -s_phase_kva, s_phase_kva))
        pf_objects["interface"].set_storage_p_setpoint(phase, p_phase)


def apply_storage_end_control(case_data, pf_objects, params) -> None:
    storage_cfg = params["storage"]
    end_rule = params["end_node_rule"]
    p_phase_max_kw = float(storage_cfg["p_phase_max_kw"])
    s_phase_kva = float(storage_cfg["s_phase_kva"])

    u_l1 = float(case_data["critical_node_voltages"]["L1"])
    u_l2 = float(case_data["critical_node_voltages"]["L2"])
    u_l3 = float(case_data["critical_node_voltages"]["L3"])
    u_values = {"L1": u_l1, "L2": u_l2, "L3": u_l3}
    u_avg = (u_l1 + u_l2 + u_l3) / 3.0

    for phase in PHASES:
        u = u_values[phase]
        if u <= float(end_rule["u_start_pu"]):
            step = 0
        elif u <= float(end_rule["u_step_25_pu"]):
            step = 1
        elif u <= float(end_rule["u_step_50_pu"]):
            step = 2
        elif u <= float(end_rule["u_step_75_pu"]):
            step = 3
        else:
            step = 4
        if u - u_avg > float(end_rule["u_unbalance_boost_pu"]):
            step = min(4, step + 1)

        p_phase = _step_to_p(step, p_phase_max_kw)
        p_phase = float(_clamp(p_phase, -p_phase_max_kw, p_phase_max_kw))
        p_phase = float(_clamp(p_phase, -s_phase_kva, s_phase_kva))
        pf_objects["interface"].set_storage_p_setpoint(phase, p_phase)
