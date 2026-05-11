from __future__ import annotations

import math
from typing import List

import numpy as np

from PSO import PSO
from src.indicators import calculate_indicators
from src.powerfactory_interface import PHASES
from src.validation import validate_allowed_pso_variables


def _safe_div(n: float, d: float) -> float:
    return n / d if abs(d) > 1e-9 else n


def run_pso_optimization(case_data, pf_objects, params) -> dict:
    interface = pf_objects["interface"]
    pv_data = pf_objects["pv_data"]
    base = case_data["base_indicators"]

    decision_names: List[str] = []
    lb = []
    ub = []
    for pv in pv_data:
        p = min(float(pv["p_kw"]), float(pv["s_kva"]))
        s = float(pv["s_kva"])
        qmax = math.sqrt(max(s * s - p * p, 0.0))
        decision_names.append(f"QPV_{pv['name']}_{pv['phase']}")
        lb.append(-qmax)
        ub.append(qmax)

    p_max = float(params["control_params"]["storage"]["p_phase_max_kw"])
    for phase in PHASES:
        decision_names.append(f"PSTOR_{phase}")
        lb.append(-p_max)
        ub.append(p_max)

    violations = validate_allowed_pso_variables(decision_names)
    if violations:
        raise ValueError("; ".join(violations))

    weights = params["pso_config"]["objective_weights"]
    pso_cfg = params["pso_config"]["pso"]
    np.random.seed(int(pso_cfg["seed"]))

    def objective(vector):
        interface.reset_setpoints()
        idx = 0
        for pv in pv_data:
            interface.set_pv_q_setpoint(pv["name"], pv["phase"], float(vector[idx]))
            idx += 1
        for phase in PHASES:
            interface.set_storage_p_setpoint(phase, float(vector[idx]))
            idx += 1

        raw = interface.run_powerflow()
        ind = calculate_indicators(raw)
        jv = _safe_div(ind["Udev_mean_pu"], base["Udev_mean_pu"])
        jvu = _safe_div(ind["kU2_mean_percent"], base["kU2_mean_percent"])
        jiu = _safe_div(ind["I_unbalance_tr_percent"], base["I_unbalance_tr_percent"])
        jexp = _safe_div(ind["P_export_total_kW"], base["P_export_total_kW"])
        jpen = float(ind["constraint_violations"])
        return (
            float(weights["wV"]) * jv
            + float(weights["wVU"]) * jvu
            + float(weights["wIU"]) * jiu
            + float(weights["wEXP"]) * jexp
            + float(weights["wPEN"]) * jpen
        )

    pso = PSO(
        func=objective,
        n_particles=int(pso_cfg["particles"]),
        dim=len(decision_names),
        lb=lb,
        ub=ub,
        max_iter=int(pso_cfg["iterations"]),
        w=float(pso_cfg["inertia"]),
        c1=float(pso_cfg["c1"]),
        c2=float(pso_cfg["c2"]),
        autosave_every_iters=int(pso_cfg.get("autosave_every_iters", 0)),
        autosave_path="pso_checkpoint.npz",
    )
    res = pso.optimize()

    interface.reset_setpoints()
    best = res["gbest"]
    idx = 0
    for pv in pv_data:
        interface.set_pv_q_setpoint(pv["name"], pv["phase"], float(best[idx]))
        idx += 1
    for phase in PHASES:
        interface.set_storage_p_setpoint(phase, float(best[idx]))
        idx += 1

    return {
        "pso_result": res,
        "decision_names": decision_names,
    }
