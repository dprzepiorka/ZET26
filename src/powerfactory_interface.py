from __future__ import annotations

from copy import deepcopy
from typing import Dict, List


PHASES = ("L1", "L2", "L3")


class PowerFactoryInterface:
    """
    Sign convention for transformer active power:
    - P_tr_phase > 0: import from MV to LV
    - P_tr_phase < 0: export from LV to MV
    """

    def __init__(self, force_stub: bool = False):
        self.use_stub = True
        if not force_stub:
            try:
                import powerfactory  # type: ignore  # noqa: F401

                self.use_stub = False
            except Exception:
                self.use_stub = True
        self.storage_location = "none"
        self.critical_node = None
        self._build_stub_model()
        self.reset_setpoints()
        self._last_results = {}

    def _build_stub_model(self) -> None:
        self._base_node_voltages = [
            {"node": "TR_NN", "L1": 1.026, "L2": 1.031, "L3": 1.034},
            {"node": "N_A", "L1": 1.037, "L2": 1.042, "L3": 1.047},
            {"node": "N_B", "L1": 1.041, "L2": 1.048, "L3": 1.054},
            {"node": "N_C", "L1": 1.046, "L2": 1.052, "L3": 1.058},
        ]
        self._base_transformer_p_kw = {"L1": -18.0, "L2": -20.0, "L3": -22.0}
        self._base_transformer_i_a = {"L1": 72.0, "L2": 80.0, "L3": 88.0}
        self._base_branch_losses_kw = [
            {"branch": "BR_TR_A", "p_loss_kw": 0.65},
            {"branch": "BR_A_B", "p_loss_kw": 0.48},
            {"branch": "BR_B_C", "p_loss_kw": 0.42},
        ]
        self._pv_units = [
            {"name": "PV_A", "node": "N_B", "s_kva": {"L1": 14.0, "L2": 14.0, "L3": 14.0}, "p_kw": {"L1": 10.5, "L2": 11.5, "L3": 10.0}},
            {"name": "PV_B", "node": "N_C", "s_kva": {"L1": 14.0, "L2": 14.0, "L3": 14.0}, "p_kw": {"L1": 11.0, "L2": 11.8, "L3": 11.2}},
            {"name": "PV_C", "node": "N_A", "s_kva": {"L1": 14.0, "L2": 14.0, "L3": 14.0}, "p_kw": {"L1": 8.0, "L2": 8.2, "L3": 8.0}},
        ]

    def reset_setpoints(self) -> None:
        self._pv_q_setpoints = {unit["name"]: {phase: 0.0 for phase in PHASES} for unit in self._pv_units}
        self._storage_p_setpoints = {phase: 0.0 for phase in PHASES}

    def set_storage_location(self, location: str, critical_node: str | None = None) -> None:
        self.storage_location = location
        self.critical_node = critical_node

    def get_pv_data(self) -> List[dict]:
        data = []
        for unit in self._pv_units:
            for phase in PHASES:
                data.append(
                    {
                        "name": unit["name"],
                        "node": unit["node"],
                        "phase": phase,
                        "p_kw": float(unit["p_kw"][phase]),
                        "s_kva": float(unit["s_kva"][phase]),
                    }
                )
        return data

    def set_pv_q_setpoint(self, inverter_name: str, phase: str, q_kvar: float) -> None:
        self._pv_q_setpoints[inverter_name][phase] = float(q_kvar)

    def set_storage_p_setpoint(self, phase: str, p_kw: float) -> None:
        self._storage_p_setpoints[phase] = float(p_kw)

    def get_storage_setpoints(self) -> Dict[str, float]:
        return dict(self._storage_p_setpoints)

    def _sum_pv_q_per_phase(self) -> Dict[str, float]:
        q_sum = {phase: 0.0 for phase in PHASES}
        for q_by_phase in self._pv_q_setpoints.values():
            for phase in PHASES:
                q_sum[phase] += q_by_phase[phase]
        return q_sum

    def _find_node(self, node_name: str) -> dict:
        for node in self._base_node_voltages:
            if node["node"] == node_name:
                return node
        raise KeyError(f"Unknown node: {node_name}")

    def get_transformer_phase_powers(self, prefer_direct: bool = True) -> Dict[str, float]:
        if prefer_direct and self._last_results:
            return dict(self._last_results["transformer"]["p_kw"])
        if not self._last_results:
            self.run_powerflow()
        phasors = self._last_results["transformer"]["phasors"]
        powers = {}
        for phase in PHASES:
            s = phasors["u_complex"][phase] * phasors["i_complex"][phase].conjugate() / 1000.0
            powers[phase] = float(s.real)
        return powers

    def run_powerflow(self) -> dict:
        q_phase = self._sum_pv_q_per_phase()
        storage = dict(self._storage_p_setpoints)

        node_voltages = []
        for node in self._base_node_voltages:
            factor = 1.0 if node["node"] == "TR_NN" else (1.15 if node["node"] == "N_A" else 1.25)
            row = {"node": node["node"]}
            for phase in PHASES:
                u = node[phase] + 0.0012 * q_phase[phase] * factor - 0.0025 * storage[phase] * factor
                row[phase] = round(u, 6)
            node_voltages.append(row)

        tr_p = {}
        tr_i = {}
        for phase in PHASES:
            tr_p[phase] = round(self._base_transformer_p_kw[phase] + storage[phase] + 0.08 * q_phase[phase], 6)
            tr_i[phase] = round(abs(self._base_transformer_i_a[phase] + 1.15 * storage[phase] + 0.35 * q_phase[phase]), 6)

        branch_losses = []
        current_factor = sum(tr_i.values()) / max(sum(self._base_transformer_i_a.values()), 1e-9)
        for item in self._base_branch_losses_kw:
            branch_losses.append({"branch": item["branch"], "p_loss_kw": round(item["p_loss_kw"] * current_factor * current_factor, 6)})

        u_tr = self._find_node("TR_NN")
        phasor_u = {phase: complex(u_tr[phase] * 230.0, 0.0) for phase in PHASES}
        phasor_i = {}
        for phase in PHASES:
            sign = -1.0 if tr_p[phase] < 0 else 1.0
            phasor_i[phase] = complex(sign * tr_i[phase], 0.0)

        self._last_results = {
            "node_voltages": node_voltages,
            "transformer": {
                "p_kw": tr_p,
                "i_a": tr_i,
                "u_pu": {phase: float(u_tr[phase]) for phase in PHASES},
                "phasors": {"u_complex": phasor_u, "i_complex": phasor_i},
            },
            "pv_setpoints": [
                {"inverter": name, "phase": phase, "q_kvar": q}
                for name, values in self._pv_q_setpoints.items()
                for phase, q in values.items()
            ],
            "storage_setpoints": [{"phase": phase, "p_kw": storage[phase], "q_kvar": 0.0} for phase in PHASES],
            "branch_losses": branch_losses,
            "metadata": {
                "pf_mode": "stub" if self.use_stub else "real",
                "storage_location": self.storage_location,
                "critical_node": self.critical_node,
            },
        }
        return deepcopy(self._last_results)
