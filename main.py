from src.run_case import run_all_cases


if __name__ == "__main__":
    summary = run_all_cases()
    for case_name, metrics in summary.items():
        print(f"{case_name}: P_export_total_kW={metrics['P_export_total_kW']:.3f}, Udev_mean_pu={metrics['Udev_mean_pu']:.5f}")
