# ZET26 – prosty skrypt badawczy PowerFactory (Excel -> case -> Excel)

Repozytorium zawiera uproszczone rozwiązanie z jednym głównym skryptem:

- `run_pf_study.py` – uruchamia przypadki badawcze
- `powerfactory_interface.py` – interfejs PF + tryb demo (mock)
- `controls.py` – Q(U), lokalne reguły magazynu i PSO
- `metrics.py` – wskaźniki i porównania
- `excel_io.py` – odczyt/zapis Excela
- `PSO.py` – lekki optimizer PSO

## Szybki start

1. Zainstaluj zależności:

```bash
python -m pip install numpy openpyxl
```

2. Utwórz przykładowy plik wejściowy:

```bash
python run_pf_study.py --create-sample --input input_study.xlsx
```

3. Uruchom obliczenia (demo):

```bash
python run_pf_study.py --input input_study.xlsx --output-dir results --demo
```

## Uruchomienie w DIgSILENT PowerFactory

- Jeśli środowisko PF ma moduł `powerfactory`, skrypt spróbuje użyć realnego API.
- W `powerfactory_interface.py` są przygotowane miejsca do:
  - wyszukiwania obiektów po mapowaniu z Excela,
  - ustawiania mocy odbiorów/PV,
  - ustawiania `Q_PV` i `P_storage`,
  - uruchamiania load flow,
  - pobierania fazowych wyników.
- Gdy `powerfactory` nie jest dostępny, działa tryb demo (`--demo` lub `study_config.demo_mode=true`).

## Obsługiwane przypadki

- `base_no_control`
- `pso_global`
- `local_qu_storage_tr`
- `local_qu_storage_end`

Wybór przypadku:

- z CLI: `--case base_no_control|pso_global|local_qu_storage_tr|local_qu_storage_end|all`
- lub z Excela: `study_config.case_name` (`all` albo nazwa przypadku)

## Format Excela wejściowego

Wymagane arkusze:

- `study_config` (`key`, `value`), min.:
  - `case_name`
  - `output_dir` (opcjonalnie)
  - `output_file` (opcjonalnie)
  - `demo_mode`
- `loads`: `object_id`, `node_id`, `L1_kW`, `L2_kW`, `L3_kW`
- `pv_units`: `object_id`, `node_id`, `L1_kW`, `L2_kW`, `L3_kW`, `S_inv_kVA`
- `storage`: `object_id`, `node_id`, `location`, `S_phase_kVA`, `P_phase_max_kW`
- `object_map`: `object_id`, `pf_path`, `node_id`, `object_type`
- `control_params`: `group`, `key`, `value`

Przykładowe grupy w `control_params`:

- `storage`: `s_total_kva`, `p_total_max_kw`, `s_phase_kva`, `p_phase_max_kw`, `p_start_kw`
- `transformer_rule`: `current_unbalance_threshold_1`, `current_unbalance_threshold_2`
- `end_node_rule`: `u_start_pu`, `u_step_25_pu`, `u_step_50_pu`, `u_step_75_pu`, `u_unbalance_boost_pu`
- `qu_curve`: `u_min_full`, `u_min_deadband`, `u_max_deadband`, `u_max_full`

## Co zapisuje skrypt

Dla każdego przypadku (`results/<case>/results.xlsx`):

- `node_voltages`
- `transformer_phase_results`
- `pv_setpoints`
- `storage_setpoints`
- `branch_losses`
- `indicators`

Podsumowanie (`results/summary/summary.xlsx`):

- `indicators_all_cases`
- `comparison_to_base`
- `comparison_to_pso`

## Założenia techniczne zaimplementowane w skrypcie

- Model trójfazowy/asymetryczny (`L1`, `L2`, `L3`)
- PV: sterowanie tylko `Q`, brak curtailmentu jako strategii
- Ograniczenie falownika: `P^2 + Q^2 <= S_inv^2`
- Magazyn: sterowanie tylko `P`, znak `P>0` = ładowanie
- Ograniczenia magazynu fazowego i całkowitego
- PSO używa wyłącznie: `Q_PV`, `P_storage_L1/L2/L3`
