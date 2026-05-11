from pathlib import Path

import yaml


ROOT_DIR = Path(__file__).resolve().parent.parent
CONFIG_DIR = ROOT_DIR / "config"


def load_yaml(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def load_all_configs() -> dict:
    return {
        "cases": load_yaml(CONFIG_DIR / "cases.yaml"),
        "control_params": load_yaml(CONFIG_DIR / "control_params.yaml"),
        "pso_config": load_yaml(CONFIG_DIR / "pso_config.yaml"),
    }
