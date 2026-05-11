from __future__ import annotations

import math

from src.powerfactory_interface import PHASES


def validate_phase_exported(results_raw: dict) -> list[str]:
    violations = []
    for row in results_raw.get("node_voltages", []):
        for phase in PHASES:
            if phase not in row:
                violations.append(f"Missing voltage phase {phase} at node {row.get('node', '<unknown>')}")
    tr = results_raw.get("transformer", {})
    for group in ("p_kw", "i_a"):
        for phase in PHASES:
            if phase not in tr.get(group, {}):
                violations.append(f"Missing transformer {group} phase {phase}")
    return violations


def validate_pv_limits(pv_data: list[dict], pv_setpoints: list[dict]) -> list[str]:
    violations = []
    q_map = {(x["inverter"], x["phase"]): float(x["q_kvar"]) for x in pv_setpoints}
    for pv in pv_data:
        p = float(pv["p_kw"])
        s = float(pv["s_kva"])
        if p > s + 1e-9:
            violations.append(f"PV {pv['name']} {pv['phase']} exceeds Smax: P={p} > S={s}")
        qmax = math.sqrt(max(s * s - min(p, s) * min(p, s), 0.0))
        q = abs(q_map[(pv["name"], pv["phase"])])
        if q > qmax + 1e-6:
            violations.append(f"PV {pv['name']} {pv['phase']} exceeds Qmax_available")
    return violations


def validate_storage_limits(storage_setpoints: list[dict], params: dict) -> list[str]:
    violations = []
    p_phase_max = float(params["storage"]["p_phase_max_kw"])
    s_phase = float(params["storage"]["s_phase_kva"])
    p_total_max = float(params["storage"]["p_total_max_kw"])
    p_total = 0.0
    for item in storage_setpoints:
        p = float(item["p_kw"])
        q = float(item.get("q_kvar", 0.0))
        p_total += p
        if abs(p) > p_phase_max + 1e-9:
            violations.append(f"Storage phase {item['phase']} exceeds P phase max")
        if p * p + q * q > s_phase * s_phase + 1e-9:
            violations.append(f"Storage phase {item['phase']} exceeds S phase max")
    if abs(p_total) > p_total_max + 1e-9:
        violations.append("Storage total active power exceeds configured max")
    return violations


def validate_transformer_signs(transformer: dict) -> list[str]:
    violations = []
    p_map = transformer.get("p_kw", {})
    for phase in PHASES:
        if phase not in p_map:
            continue
        p = float(p_map[phase])
        export = max(0.0, -p)
        if p < 0 and export <= 0:
            violations.append(f"Inconsistent export sign at {phase}")
    return violations


def validate_allowed_pso_variables(variable_names: list[str]) -> list[str]:
    violations = []
    for name in variable_names:
        if not (name.startswith("QPV_") or name.startswith("PSTOR_")):
            violations.append(f"PSO uses disallowed variable: {name}")
    return violations


def validate_case_configuration(cases: dict) -> list[str]:
    violations = []
    required = {"base_no_control", "pso_global", "local_qu_storage_tr", "local_qu_storage_end"}
    missing = required - set(cases.keys())
    if missing:
        violations.append(f"Missing required cases: {sorted(missing)}")
    if "pso_global" in cases and not bool(cases["pso_global"].get("pso", False)):
        violations.append("Case pso_global must have pso: true")
    return violations
