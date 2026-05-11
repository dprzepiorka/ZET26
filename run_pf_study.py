from __future__ import annotations

import argparse
import math
from pathlib import Path
from typing import Dict, List

from controls import (
    apply_pv_qu_control,
    apply_storage_end_control,
    apply_storage_tr_control,
    run_pso_optimization,
)
from excel_io import create_sample_input_excel, read_input_excel, write_case_results_excel, write_summary_excel
from metrics import calculate_indicators, comparison_to_base, comparison_to_pso
from powerfactory_interface import PowerFactoryInterface


CASES = ["base_no_control", "pso_global", "local_qu_storage_tr", "local_qu_storage_end"]
_CURRENT_CONTEXT: Dict = {}
VALIDATION_EPS = 1e-6


def _f(v, default=0.0) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return float(default)


def _to_bool(v) -> bool:
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    return str(v).strip().lower() in {"1", "true", "yes", "on"}


def _validate_phase_columns(rows: List[Dict], row_name: str, warnings: List[str]) -> None:
    for idx, r in enumerate(rows):
        for ph in ("L1", "L2", "L3"):
            if ph not in r and f"{ph}_kW" not in r and f"{ph}_A" not in r:
                warnings.append(f"{row_name}[{idx}] brak kolumny fazowej {ph}.")


def _validate_constraints(raw: Dict, input_data: Dict, params: Dict, case_name: str) -> List[str]:
    warnings: List[str] = []

    pv_rows = raw.get("pv_setpoints", [])
    for r in pv_rows:
        p = _f(r.get("P_kW", 0.0), 0.0)
        q = _f(r.get("Q_kvar", 0.0), 0.0)
        s = _f(r.get("S_inv_kVA", 0.0), 0.0)
        if p * p + q * q > s * s + VALIDATION_EPS:
            warnings.append(f"PV {r.get('object_id')} {r.get('phase')}: przekroczenie S_inv.")

    storage_rows = raw.get("storage_setpoints", [])
    p_phase_max = _f(params.get("storage", {}).get("p_phase_max_kw", 20.0), 20.0)
    s_phase = _f(params.get("storage", {}).get("s_phase_kva", 20.0), 20.0)
    p_total_max = _f(params.get("storage", {}).get("p_total_max_kw", 60.0), 60.0)
    for r in storage_rows:
        phases = ["L1", "L2", "L3"]
        total_abs = 0.0
        for ph in phases:
            p = _f(r.get(f"P_{ph}_kW", 0.0), 0.0)
            total_abs += abs(p)
            if abs(p) > p_phase_max + VALIDATION_EPS:
                warnings.append(f"Storage {r.get('object_id')} {ph}: |P|>{p_phase_max} kW.")
            if p * p > s_phase * s_phase + VALIDATION_EPS:
                warnings.append(f"Storage {r.get('object_id')} {ph}: przekroczenie S_phase.")
        if total_abs > p_total_max + VALIDATION_EPS:
            warnings.append(f"Storage {r.get('object_id')}: przekroczenie P_total_max.")

    tr = raw.get("transformer_phase_results", [])
    if tr:
        t = tr[0]
        p_sum = _f(t.get("P_tr_L1_kW", 0.0), 0.0) + _f(t.get("P_tr_L2_kW", 0.0), 0.0) + _f(t.get("P_tr_L3_kW", 0.0), 0.0)
        p_export = max(0.0, -p_sum)
        p_export_reported = t.get("P_export_total_kW", None)
        if p_export_reported is not None and abs(_f(p_export_reported, 0.0) - p_export) > VALIDATION_EPS:
            warnings.append("Niespójna konwencja znaków mocy transformatora (P_export_total_kW vs fazy).")

    _validate_phase_columns(raw.get("node_voltages", []), "node_voltages", warnings)
    _validate_phase_columns(raw.get("transformer_phase_results", []), "transformer_phase_results", warnings)

    if case_name == "pso_global":
        pso_allowed = {"Q_PV", "P_storage_L1", "P_storage_L2", "P_storage_L3"}
        pso_used = set(raw.get("pso_used_variables", []))
        if pso_used and pso_used != pso_allowed:
            warnings.append("PSO używa niedozwolonych zmiennych.")

    local_or_pso = {"local_qu_storage_tr", "local_qu_storage_end", "pso_global"}
    if case_name in local_or_pso:
        if not input_data.get("storage"):
            warnings.append("Brak magazynu dla wybranego wariantu.")

    return warnings


def _merge_params(input_data: Dict) -> Dict:
    cp = input_data.get("control_params", {})
    merged = {
        "storage": {
            "s_total_kva": 60.0,
            "p_total_max_kw": 60.0,
            "s_phase_kva": 20.0,
            "p_phase_max_kw": 20.0,
            "steps_percent": [0, 25, 50, 75, 100],
            "p_start_kw": 1.0,
        },
        "transformer_rule": {
            "current_unbalance_threshold_1": 0.05,
            "current_unbalance_threshold_2": 0.10,
        },
        "end_node_rule": {
            "u_start_pu": 1.03,
            "u_step_25_pu": 1.04,
            "u_step_50_pu": 1.05,
            "u_step_75_pu": 1.06,
            "u_unbalance_boost_pu": 0.005,
        },
        "qu_curve": {
            "u_min_full": 0.95,
            "u_min_deadband": 0.97,
            "u_max_deadband": 1.03,
            "u_max_full": 1.08,
        },
        "objective_weights": {
            "wV": 0.20,
            "wVU": 0.30,
            "wIU": 0.25,
            "wEXP": 0.20,
            "wPEN": 0.05,
        },
        "pso": {"particles": 16, "iterations": 20, "w": 0.7, "c1": 1.4, "c2": 1.4},
    }
    for group, values in cp.items():
        merged.setdefault(group, {}).update(values)
    return merged


def _critical_node_from_base(base_raw: Dict) -> str:
    node_voltages = base_raw.get("node_voltages", [])
    if not node_voltages:
        return "TR_NODE"
    best = max(node_voltages, key=lambda r: max(_f(r.get("L1", 0.0)), _f(r.get("L2", 0.0)), _f(r.get("L3", 0.0))))
    return str(best.get("node_id", "TR_NODE"))


def run_case(case_name: str, context: Dict | None = None) -> Dict:
    ctx = context or _CURRENT_CONTEXT
    pf = ctx["pf"]
    input_data = ctx["input_data"]
    params = ctx["params"]

    if case_name not in CASES:
        raise ValueError(f"Nieznany case_name: {case_name}")

    pf.set_pv_q_setpoints({})
    pf.set_storage_p_setpoints({})

    if case_name == "base_no_control":
        raw = pf.run_and_collect()

    elif case_name == "local_qu_storage_tr":
        pre = pf.run_and_collect()
        q_set = apply_pv_qu_control({"node_voltages": pre["node_voltages"], "pv_units": input_data["pv_units"]}, pf, params)
        pf.set_pv_q_setpoints(q_set)
        mid = pf.run_and_collect()
        p_storage = apply_storage_tr_control(
            {"transformer_phase_results": mid["transformer_phase_results"], "storage": input_data["storage"]}, pf, params
        )
        pf.set_storage_p_setpoints(p_storage)
        raw = pf.run_and_collect()

    elif case_name == "local_qu_storage_end":
        pre = pf.run_and_collect()
        q_set = apply_pv_qu_control({"node_voltages": pre["node_voltages"], "pv_units": input_data["pv_units"]}, pf, params)
        pf.set_pv_q_setpoints(q_set)
        mid = pf.run_and_collect()
        p_storage = apply_storage_end_control(
            {
                "node_voltages": mid["node_voltages"],
                "storage": input_data["storage"],
                "critical_node": ctx.get("critical_node", "TR_NODE"),
            },
            pf,
            params,
        )
        pf.set_storage_p_setpoints(p_storage)
        raw = pf.run_and_collect()

    elif case_name == "pso_global":
        if "base_indicators" not in ctx:
            raise ValueError("Dla pso_global wymagany jest wynik base_no_control.")
        pso_result = run_pso_optimization(
            {
                "base_indicators": ctx["base_indicators"],
                "pv_units": input_data["pv_units"],
                "storage": input_data["storage"],
            },
            pf,
            params,
        )
        pf.set_pv_q_setpoints(pso_result["pv_q_setpoints"])
        pf.set_storage_p_setpoints(pso_result["storage_p_setpoints"])
        raw = pf.run_and_collect()
        raw["pso_best_objective"] = pso_result["objective_best"]
        raw["pso_used_variables"] = ["Q_PV", "P_storage_L1", "P_storage_L2", "P_storage_L3"]

    else:
        raise ValueError(case_name)

    raw["violations"] = _validate_constraints(raw, input_data, params, case_name)
    indicators = calculate_indicators(raw)
    return {"case": case_name, "raw": raw, "indicators": indicators}


def _resolve_cases(input_data: Dict, cli_case: str) -> List[str]:
    if cli_case != "auto":
        if cli_case == "all":
            return CASES.copy()
        return [cli_case]
    cfg_case = str(input_data.get("study_config", {}).get("case_name", "all")).strip()
    if cfg_case.lower() == "all":
        return CASES.copy()
    return [cfg_case]


def _case_output_sheets(result: Dict) -> Dict[str, List[Dict]]:
    raw = result["raw"]
    return {
        "node_voltages": raw.get("node_voltages", []),
        "transformer_phase_results": raw.get("transformer_phase_results", []),
        "pv_setpoints": raw.get("pv_setpoints", []),
        "storage_setpoints": raw.get("storage_setpoints", []),
        "branch_losses": raw.get("branch_losses", []),
        "indicators": [{"case": result["case"], **result["indicators"]}],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Prosty runner badań PowerFactory (Excel -> case -> Excel).")
    parser.add_argument("--input", default="input_study.xlsx", help="Ścieżka do pliku wejściowego Excel.")
    parser.add_argument("--output-dir", default="results", help="Katalog wynikowy.")
    parser.add_argument("--case", default="auto", choices=["auto", "all", *CASES], help="Wybór case'u.")
    parser.add_argument("--demo", action="store_true", help="Wymuś tryb demo bez DIgSILENT PowerFactory.")
    parser.add_argument("--create-sample", action="store_true", help="Utwórz przykładowy plik wejściowy Excel i zakończ.")
    args = parser.parse_args()

    if args.create_sample:
        create_sample_input_excel(args.input)
        print(f"Utworzono przykładowy plik: {args.input}")
        return

    input_data = read_input_excel(args.input)
    params = _merge_params(input_data)
    cases = _resolve_cases(input_data, args.case)

    force_demo_cfg = _to_bool(input_data.get("study_config", {}).get("demo_mode", False))
    pf = PowerFactoryInterface(use_demo=args.demo or force_demo_cfg)
    pf.set_data(input_data.get("object_map", []), input_data.get("loads", []), input_data.get("pv_units", []), input_data.get("storage", []))

    context = {"pf": pf, "input_data": input_data, "params": params}
    _CURRENT_CONTEXT.update(context)

    all_results: Dict[str, Dict] = {}

    if any(c in cases for c in ["pso_global", "local_qu_storage_end"]) and "base_no_control" not in cases:
        base = run_case("base_no_control", context)
        all_results["base_no_control"] = base
        context["base_indicators"] = base["indicators"]
        context["critical_node"] = _critical_node_from_base(base["raw"])

    for case_name in cases:
        if case_name == "base_no_control" and case_name in all_results:
            continue
        if case_name in {"pso_global", "local_qu_storage_end"} and "base_no_control" in all_results:
            context["base_indicators"] = all_results["base_no_control"]["indicators"]
            context["critical_node"] = _critical_node_from_base(all_results["base_no_control"]["raw"])
        result = run_case(case_name, context)
        all_results[case_name] = result

    out_root = Path(args.output_dir)
    for case_name, result in all_results.items():
        case_dir = out_root / case_name
        case_xlsx = case_dir / "results.xlsx"
        write_case_results_excel(str(case_xlsx), _case_output_sheets(result))

    all_ind = {k: v["indicators"] for k, v in all_results.items()}
    rows_all = [{"case": k, **v} for k, v in all_ind.items()]
    cmp_base = comparison_to_base(all_ind, "base_no_control")
    cmp_pso = comparison_to_pso(all_ind, "base_no_control", "pso_global")

    summary = {
        "indicators_all_cases": rows_all,
        "comparison_to_base": cmp_base,
        "comparison_to_pso": cmp_pso,
    }
    write_summary_excel(str(out_root / "summary" / "summary.xlsx"), summary)

    print("Zakończono. Wyniki zapisano do:")
    print(f"- przypadki: {out_root}/<case>/results.xlsx")
    print(f"- podsumowanie: {out_root}/summary/summary.xlsx")


if __name__ == "__main__":
    main()
