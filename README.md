# ZET26 – badanie sterowania w asymetrycznej sieci nn

Zestaw skryptów porównuje 4 przypadki:

1. `base_no_control`
2. `pso_global`
3. `local_qu_storage_tr`
4. `local_qu_storage_end`

Implementacja działa:
- z DIgSILENT PowerFactory (gdy środowisko jest dostępne),
- albo w trybie demonstracyjnym (`stub`) bez PowerFactory.

## Struktura

- `config/cases.yaml` – konfiguracja przypadków.
- `config/control_params.yaml` – parametry Q(U), magazynu i reguł lokalnych.
- `config/pso_config.yaml` – konfiguracja PSO i wagi funkcji celu.
- `src/powerfactory_interface.py` – warstwa integracyjna PF/stub.
- `src/run_case.py` – uruchamianie przypadków.
- `src/local_qu_control.py` – lokalna charakterystyka Q(U) PV.
- `src/storage_control.py` – reguły magazynu (transformator / węzeł krytyczny).
- `src/pso_optimizer.py` – centralna optymalizacja PSO (Q PV + P magazynu fazowo).
- `src/indicators.py` – obliczanie wskaźników.
- `src/export_results.py` – eksport CSV per przypadek.
- `src/postprocess.py` – porównania `vs base` i `vs pso`.

## Uruchomienie

```bash
python -m pip install -r requirements.txt
python main.py
```

## Wyniki

Per przypadek:
- `results/<case_name>/node_voltages.csv`
- `results/<case_name>/transformer_phase_results.csv`
- `results/<case_name>/pv_setpoints.csv`
- `results/<case_name>/storage_setpoints.csv`
- `results/<case_name>/branch_losses.csv`
- `results/<case_name>/indicators.csv`

Podsumowanie:
- `results/summary/indicators_all_cases.csv`
- `results/summary/comparison_to_base.csv`
- `results/summary/comparison_to_pso.csv`

## Uwagi dot. znaków mocy transformatora

Przyjęto:
- `P_tr_phase > 0`: import z SN do nn,
- `P_tr_phase < 0`: eksport z nn do SN (przepływ zwrotny).

Stąd:
- `P_export_total_kW = sum(max(0, -P_tr_phase))`.
