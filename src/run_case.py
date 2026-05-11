from __future__ import annotations

from src.config_loader import load_all_configs
from src.export_results import export_case_results
from src.indicators import calculate_indicators
from src.local_qu_control import apply_pv_qu_control
from src.postprocess import generate_summary
from src.powerfactory_interface import PHASES, PowerFactoryInterface
from src.pso_optimizer import run_pso_optimization
from src.storage_control import apply_storage_end_control, apply_storage_tr_control
from src.validation import (
    validate_case_configuration,
    validate_phase_exported,
    validate_pv_limits,
    validate_storage_limits,
    validate_transformer_signs,
)


def _find_critical_node(base_raw: dict) -> tuple[str, dict]:
    best_name = ""
    best_u = -1.0
    best_row = {}
    for row in base_raw["node_voltages"]:
        row_max = max(float(row[p]) for p in PHASES)
        if row_max > best_u:
            best_u = row_max
            best_name = row["node"]
            best_row = row
    return best_name, best_row


def _prepare_pf_objects(interface: PowerFactoryInterface) -> dict:
    return {"interface": interface, "pv_data": interface.get_pv_data()}


def _collect_violations(raw: dict, pv_data: list[dict], control_params: dict) -> list[str]:
    violations = []
    violations.extend(validate_phase_exported(raw))
    violations.extend(validate_pv_limits(pv_data, raw["pv_setpoints"]))
    violations.extend(validate_storage_limits(raw["storage_setpoints"], control_params))
    violations.extend(validate_transformer_signs(raw["transformer"]))
    return violations


def run_case(case_name: str) -> dict:
    cfg = load_all_configs()
    cases = cfg["cases"]["cases"]
    config_violations = validate_case_configuration(cases)
    if config_violations:
        raise ValueError("; ".join(config_violations))
    case_cfg = cases[case_name]
    interface = PowerFactoryInterface()
    pf_objects = _prepare_pf_objects(interface)

    interface.reset_setpoints()
    base_raw = interface.run_powerflow()
    base_indicators = calculate_indicators(base_raw)
    critical_node_name, critical_node_row = _find_critical_node(base_raw)

    interface.reset_setpoints()
    interface.set_storage_location(case_cfg["storage_location"], critical_node_name)
    pso_meta = {}

    if case_name == "base_no_control":
        pass
    elif case_cfg["pso"]:
        pso_meta = run_pso_optimization(
            case_data={"base_indicators": base_indicators},
            pf_objects=pf_objects,
            params={"control_params": cfg["control_params"], "pso_config": cfg["pso_config"]},
        )
    else:
        if case_cfg["pv_control"] == "local_qu":
            apply_pv_qu_control(
                case_data={"node_voltages": base_raw["node_voltages"], "warnings": []},
                pf_objects=pf_objects,
                params=cfg["control_params"],
            )
        if case_cfg["storage_control"] == "step_transformer_current_power":
            after_pv = interface.run_powerflow()
            apply_storage_tr_control(after_pv, pf_objects, cfg["control_params"])
        elif case_cfg["storage_control"] == "step_local_voltage":
            apply_storage_end_control(
                {"critical_node_voltages": critical_node_row},
                pf_objects,
                cfg["control_params"],
            )

    raw = interface.run_powerflow()
    raw["constraint_violations"] = _collect_violations(raw, pf_objects["pv_data"], cfg["control_params"])
    indicators = calculate_indicators(raw)
    export_case_results(case_name, raw, indicators)
    return {"case_name": case_name, "results_raw": raw, "indicators": indicators, "pso_meta": pso_meta}


def run_all_cases() -> dict[str, dict]:
    cfg = load_all_configs()
    output = {}
    for case_name in cfg["cases"]["cases"].keys():
        result = run_case(case_name)
        output[case_name] = result["indicators"]
    generate_summary(output)
    return output
