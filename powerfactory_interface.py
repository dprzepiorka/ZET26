from __future__ import annotations

import math
from typing import Dict, List


class PowerFactoryInterface:
    def __init__(self, use_demo: bool = False):
        self.use_demo = use_demo
        self.app = None
        if not use_demo:
            try:
                import powerfactory  # type: ignore

                self.app = powerfactory.GetApplication()
                if self.app is None:
                    self.use_demo = True
            except Exception:
                self.use_demo = True

        self.object_map: List[Dict] = []
        self.loads: List[Dict] = []
        self.pv_units: List[Dict] = []
        self.storage_units: List[Dict] = []
        self.pv_q_setpoints: Dict[str, Dict[str, float]] = {}
        self.storage_p_setpoints: Dict[str, Dict[str, float]] = {}

    def set_data(self, object_map: List[Dict], loads: List[Dict], pv_units: List[Dict], storage_units: List[Dict]) -> None:
        self.object_map = object_map
        self.loads = loads
        self.pv_units = pv_units
        self.storage_units = storage_units

    def set_pv_q_setpoints(self, q_setpoints: Dict[str, Dict[str, float]]) -> None:
        self.pv_q_setpoints = q_setpoints

    def set_storage_p_setpoints(self, p_setpoints: Dict[str, Dict[str, float]]) -> None:
        self.storage_p_setpoints = p_setpoints

    def run_loadflow(self) -> None:
        if self.use_demo:
            return
        # Miejsce na realny load-flow w PowerFactory (np. ComLdf.Execute())

    def _value(self, row: Dict, key: str) -> float:
        v = row.get(key, 0.0)
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    def _node_id(self, row: Dict) -> str:
        return str(row.get("node_id") or row.get("object_id") or "NODE")

    def run_and_collect(self) -> Dict:
        self.run_loadflow()
        if self.use_demo:
            return self._collect_demo()
        return self._collect_real_pf_placeholder()

    def _collect_real_pf_placeholder(self) -> Dict:
        raise NotImplementedError(
            "Tryb PowerFactory: uzupełnij metody pobierania obiektów, ustawiania mocy i eksportu wyników fazowych. "
            "Skrypt działa od razu w trybie demo (--demo)."
        )

    def _collect_demo(self) -> Dict:
        nodes = sorted({self._node_id(r) for r in (self.loads + self.pv_units + self.storage_units)} | {"TR_NODE"})
        phase_names = ["L1", "L2", "L3"]

        load_by_phase = {ph: 0.0 for ph in phase_names}
        for l in self.loads:
            load_by_phase["L1"] += self._value(l, "L1_kW")
            load_by_phase["L2"] += self._value(l, "L2_kW")
            load_by_phase["L3"] += self._value(l, "L3_kW")

        pv_by_phase = {ph: 0.0 for ph in phase_names}
        for p in self.pv_units:
            pv_by_phase["L1"] += self._value(p, "L1_kW")
            pv_by_phase["L2"] += self._value(p, "L2_kW")
            pv_by_phase["L3"] += self._value(p, "L3_kW")

        storage_phase = {"L1": 0.0, "L2": 0.0, "L3": 0.0}
        for _, vals in self.storage_p_setpoints.items():
            storage_phase["L1"] += float(vals.get("L1", 0.0))
            storage_phase["L2"] += float(vals.get("L2", 0.0))
            storage_phase["L3"] += float(vals.get("L3", 0.0))

        p_tr = {
            "L1": load_by_phase["L1"] - pv_by_phase["L1"] + storage_phase["L1"],
            "L2": load_by_phase["L2"] - pv_by_phase["L2"] + storage_phase["L2"],
            "L3": load_by_phase["L3"] - pv_by_phase["L3"] + storage_phase["L3"],
        }
        i_tr = {ph: abs(p_tr[ph]) * 1.7 + 15.0 for ph in phase_names}

        total_load = sum(load_by_phase.values())
        total_pv = sum(pv_by_phase.values())
        stress = (total_pv - total_load) / total_load if total_load > 0 else 0.0

        node_voltages: List[Dict] = []
        for idx, node in enumerate(nodes):
            bias = 0.004 * idx
            l1 = 1.0 + 0.01 * stress + bias + (pv_by_phase["L1"] - load_by_phase["L1"] - storage_phase["L1"]) / 3000.0
            l2 = 1.0 + 0.01 * stress + bias + (pv_by_phase["L2"] - load_by_phase["L2"] - storage_phase["L2"]) / 3000.0
            l3 = 1.0 + 0.01 * stress + bias + (pv_by_phase["L3"] - load_by_phase["L3"] - storage_phase["L3"]) / 3000.0
            node_voltages.append({"node_id": node, "L1": l1, "L2": l2, "L3": l3})

        pv_setpoints: List[Dict] = []
        for p in self.pv_units:
            pid = str(p.get("object_id"))
            q = self.pv_q_setpoints.get(pid, {})
            for ph in phase_names:
                pv_setpoints.append(
                    {
                        "object_id": pid,
                        "phase": ph,
                        "P_kW": self._value(p, f"{ph}_kW"),
                        "Q_kvar": float(q.get(ph, 0.0)),
                        "S_inv_kVA": self._value(p, "S_inv_kVA"),
                    }
                )

        storage_setpoints: List[Dict] = []
        if self.storage_units:
            for s in self.storage_units:
                sid = str(s.get("object_id"))
                vals = self.storage_p_setpoints.get(sid, {"L1": 0.0, "L2": 0.0, "L3": 0.0})
                storage_setpoints.append(
                    {
                        "object_id": sid,
                        "P_L1_kW": float(vals.get("L1", 0.0)),
                        "P_L2_kW": float(vals.get("L2", 0.0)),
                        "P_L3_kW": float(vals.get("L3", 0.0)),
                    }
                )

        branch_losses = [
            {"branch_id": "branch_1", "P_loss_kW": 0.01 * sum(abs(v) for v in p_tr.values())},
            {"branch_id": "branch_2", "P_loss_kW": 0.008 * sum(abs(v) for v in p_tr.values())},
        ]

        transformer = [
            {
                "P_tr_L1_kW": p_tr["L1"],
                "P_tr_L2_kW": p_tr["L2"],
                "P_tr_L3_kW": p_tr["L3"],
                "I_tr_L1_A": i_tr["L1"],
                "I_tr_L2_A": i_tr["L2"],
                "I_tr_L3_A": i_tr["L3"],
            }
        ]

        return {
            "node_voltages": node_voltages,
            "transformer_phase_results": transformer,
            "pv_setpoints": pv_setpoints,
            "storage_setpoints": storage_setpoints,
            "branch_losses": branch_losses,
        }

    @staticmethod
    def transformer_export_phase_kw(p_tr_phase_kw: float) -> float:
        return max(0.0, -float(p_tr_phase_kw))

    @staticmethod
    def calc_phase_power_from_phasor(u_v: complex, i_a: complex) -> float:
        return float((u_v * complex(i_a).conjugate()).real / 1000.0)

    @staticmethod
    def current_unbalance_percent(i_l1: float, i_l2: float, i_l3: float) -> float:
        i_avg = (i_l1 + i_l2 + i_l3) / 3.0
        if abs(i_avg) < 1e-9:
            return 0.0
        return max(abs(i_l1 - i_avg), abs(i_l2 - i_avg), abs(i_l3 - i_avg)) / i_avg * 100.0
