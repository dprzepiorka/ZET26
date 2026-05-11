from __future__ import annotations

import math
from typing import Dict, List, Tuple

from PSO import PSO
from metrics import calculate_indicators


def _f(v, default=0.0) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return float(default)


def qu_curve_q_from_u(u_pu: float, qmax: float, params: Dict[str, float]) -> float:
    u_min_full = _f(params.get("u_min_full", 0.95), 0.95)
    u_min_deadband = _f(params.get("u_min_deadband", 0.97), 0.97)
    u_max_deadband = _f(params.get("u_max_deadband", 1.03), 1.03)
    u_max_full = _f(params.get("u_max_full", 1.08), 1.08)

    if u_pu <= u_min_full:
        return -qmax
    if u_min_full < u_pu < u_min_deadband:
        t = (u_pu - u_min_full) / (u_min_deadband - u_min_full)
        return -qmax * (1.0 - t)
    if u_min_deadband <= u_pu <= u_max_deadband:
        return 0.0
    if u_max_deadband < u_pu < u_max_full:
        t = (u_pu - u_max_deadband) / (u_max_full - u_max_deadband)
        return qmax * t
    return qmax


def _pv_available_qmax(p_kw: float, s_inv_kva: float) -> float:
    if p_kw > s_inv_kva:
        p_kw = s_inv_kva
    return math.sqrt(max(0.0, s_inv_kva * s_inv_kva - p_kw * p_kw))


def apply_pv_qu_control(case_data, pf_objects, params) -> Dict[str, Dict[str, float]]:
    del pf_objects
    node_voltages = case_data["node_voltages"]
    pv_units = case_data["pv_units"]
    node_u = {str(n.get("node_id")): n for n in node_voltages}
    curve = params.get("qu_curve", {})

    q_setpoints: Dict[str, Dict[str, float]] = {}
    for pv in pv_units:
        pid = str(pv.get("object_id"))
        node_id = str(pv.get("node_id") or pid)
        v = node_u.get(node_id, {"L1": 1.0, "L2": 1.0, "L3": 1.0})
        s_inv = _f(pv.get("S_inv_kVA", 0.0))
        per_phase = {}
        for ph in ("L1", "L2", "L3"):
            p = _f(pv.get(f"{ph}_kW", 0.0))
            qmax = _pv_available_qmax(p, s_inv)
            q_curve = qu_curve_q_from_u(_f(v.get(ph, 1.0), 1.0), qmax, curve)
            per_phase[ph] = max(-qmax, min(qmax, q_curve))
        q_setpoints[pid] = per_phase
    return q_setpoints


def _step_from_export(export_kw: float, p_phase_max_kw: float, p_start_kw: float) -> int:
    if export_kw <= p_start_kw:
        return 0
    if export_kw <= 0.25 * p_phase_max_kw:
        return 1
    if export_kw <= 0.50 * p_phase_max_kw:
        return 2
    if export_kw <= 0.75 * p_phase_max_kw:
        return 3
    return 4


def _step_from_voltage(u_pu: float, params: Dict[str, float]) -> int:
    if u_pu <= _f(params.get("u_start_pu", 1.03), 1.03):
        return 0
    if u_pu <= _f(params.get("u_step_25_pu", 1.04), 1.04):
        return 1
    if u_pu <= _f(params.get("u_step_50_pu", 1.05), 1.05):
        return 2
    if u_pu <= _f(params.get("u_step_75_pu", 1.06), 1.06):
        return 3
    return 4


def apply_storage_tr_control(case_data, pf_objects, params) -> Dict[str, Dict[str, float]]:
    del pf_objects
    tr = case_data["transformer_phase_results"][0]
    storage = case_data["storage"]

    storage_params = params.get("storage", {})
    tr_rule = params.get("transformer_rule", {})
    p_phase_max = _f(storage_params.get("p_phase_max_kw", 20.0), 20.0)
    p_start = _f(storage_params.get("p_start_kw", 1.0), 1.0)
    th1 = _f(tr_rule.get("current_unbalance_threshold_1", 0.05), 0.05)
    th2 = _f(tr_rule.get("current_unbalance_threshold_2", 0.10), 0.10)

    i = [_f(tr.get("I_tr_L1_A", 0.0)), _f(tr.get("I_tr_L2_A", 0.0)), _f(tr.get("I_tr_L3_A", 0.0))]
    i_avg = sum(i) / 3.0 if i else 0.0

    per_phase = {}
    for idx, ph in enumerate(("L1", "L2", "L3")):
        p_export = max(0.0, -_f(tr.get(f"P_tr_{ph}_kW", 0.0), 0.0))
        step_export = _step_from_export(p_export, p_phase_max, p_start)
        d_i = 0.0 if i_avg == 0 else (i[idx] - i_avg) / i_avg
        step_unb = 0
        if d_i > th2:
            step_unb = 2
        elif d_i > th1:
            step_unb = 1
        step_final = max(-4, min(4, step_export + step_unb))
        per_phase[ph] = (step_final / 4.0) * p_phase_max

    sid = str(storage[0].get("object_id", "storage_1")) if storage else "storage_1"
    return {sid: per_phase}


def apply_storage_end_control(case_data, pf_objects, params) -> Dict[str, Dict[str, float]]:
    del pf_objects
    storage = case_data["storage"]
    node_voltages = case_data["node_voltages"]
    critical_node = case_data["critical_node"]

    end_rule = params.get("end_node_rule", {})
    p_phase_max = _f(params.get("storage", {}).get("p_phase_max_kw", 20.0), 20.0)
    boost_thr = _f(end_rule.get("u_unbalance_boost_pu", 0.005), 0.005)

    row = next((r for r in node_voltages if str(r.get("node_id")) == str(critical_node)), None)
    if row is None:
        row = node_voltages[-1]
    phases = [_f(row.get("L1", 1.0), 1.0), _f(row.get("L2", 1.0), 1.0), _f(row.get("L3", 1.0), 1.0)]
    u_avg = sum(phases) / 3.0

    per_phase = {}
    for idx, ph in enumerate(("L1", "L2", "L3")):
        step = _step_from_voltage(phases[idx], end_rule)
        if phases[idx] > u_avg + boost_thr:
            step = min(4, step + 1)
        per_phase[ph] = (step / 4.0) * p_phase_max

    sid = str(storage[0].get("object_id", "storage_1")) if storage else "storage_1"
    return {sid: per_phase}


def run_pso_optimization(case_data, pf_objects, params) -> Dict[str, Dict[str, float]]:
    pf = pf_objects
    base_ind = case_data["base_indicators"]
    pv_units = case_data["pv_units"]
    storage = case_data["storage"]

    pso_cfg = params.get("pso", {})
    w = params.get("objective_weights", {"wV": 0.20, "wVU": 0.30, "wIU": 0.25, "wEXP": 0.20, "wPEN": 0.05})
    p_phase_max = _f(params.get("storage", {}).get("p_phase_max_kw", 20.0), 20.0)

    pv_phase_keys: List[Tuple[str, str]] = []
    lb: List[float] = []
    ub: List[float] = []
    for pv in pv_units:
        pid = str(pv.get("object_id"))
        s_inv = _f(pv.get("S_inv_kVA", 0.0), 0.0)
        for ph in ("L1", "L2", "L3"):
            p = _f(pv.get(f"{ph}_kW", 0.0), 0.0)
            qmax = _pv_available_qmax(p, s_inv)
            pv_phase_keys.append((pid, ph))
            lb.append(-qmax)
            ub.append(qmax)

    storage_phases = ["L1", "L2", "L3"]
    for _ in storage_phases:
        lb.append(-p_phase_max)
        ub.append(p_phase_max)

    if not storage:
        raise ValueError("Brak danych magazynu dla pso_global.")
    storage_id = str(storage[0].get("object_id", "storage_1"))

    def norm(metric_name: str, val: float) -> float:
        base_val = float(base_ind.get(metric_name, 0.0))
        if base_val == 0:
            return val
        return val / base_val

    def objective(x):
        q_set = {}
        for idx, (pid, ph) in enumerate(pv_phase_keys):
            q_set.setdefault(pid, {})[ph] = float(x[idx])

        st = {
            storage_id: {
                "L1": float(x[len(pv_phase_keys) + 0]),
                "L2": float(x[len(pv_phase_keys) + 1]),
                "L3": float(x[len(pv_phase_keys) + 2]),
            }
        }

        pf.set_pv_q_setpoints(q_set)
        pf.set_storage_p_setpoints(st)
        raw = pf.run_and_collect()
        raw.setdefault("violations", [])
        ind = calculate_indicators(raw)

        jv = norm("Udev_mean_pu", ind["Udev_mean_pu"])
        jvu = norm("kU2_mean_percent", ind["kU2_mean_percent"])
        jiu = norm("I_unbalance_tr_percent", ind["I_unbalance_tr_percent"])
        jexp = norm("P_export_total_kW", ind["P_export_total_kW"])
        jpen = norm("constraint_violations", ind["constraint_violations"])

        return (
            _f(w.get("wV", 0.20), 0.20) * jv
            + _f(w.get("wVU", 0.30), 0.30) * jvu
            + _f(w.get("wIU", 0.25), 0.25) * jiu
            + _f(w.get("wEXP", 0.20), 0.20) * jexp
            + _f(w.get("wPEN", 0.05), 0.05) * jpen
        )

    pso = PSO(
        func=objective,
        n_particles=int(_f(pso_cfg.get("particles", 16), 16)),
        dim=len(lb),
        lb=lb,
        ub=ub,
        max_iter=int(_f(pso_cfg.get("iterations", 20), 20)),
        w=_f(pso_cfg.get("w", 0.7), 0.7),
        c1=_f(pso_cfg.get("c1", 1.4), 1.4),
        c2=_f(pso_cfg.get("c2", 1.4), 1.4),
    )
    result = pso.optimize()
    best = result["gbest"]

    q_set = {}
    for idx, (pid, ph) in enumerate(pv_phase_keys):
        q_set.setdefault(pid, {})[ph] = float(best[idx])

    storage_set = {
        storage_id: {
            "L1": float(best[len(pv_phase_keys) + 0]),
            "L2": float(best[len(pv_phase_keys) + 1]),
            "L3": float(best[len(pv_phase_keys) + 2]),
        }
    }

    return {"pv_q_setpoints": q_set, "storage_p_setpoints": storage_set, "objective_best": float(result["gbest_val"])}
