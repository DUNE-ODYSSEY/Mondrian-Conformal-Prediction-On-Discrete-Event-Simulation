"""
Publication-rigor upgrade, part 4: generalizability check on a second
department.

Reuses extract_distributions.py's functions unchanged (department filter,
arrival/ESI distribution logic) but points at Department B (166,497 visits,
the second-largest site: 1 academic + 2 community EDs in the dataset,
Department A already used is the academic site) and writes to
results/tables/dept_b/ instead of results/tables/ - Department A's files
are never touched.

Service time params are literature-derived, not department-specific, so
they're identical to Department A's by construction - copied here rather
than re-derived.
"""

import pandas as pd

from src.utils.extract_distributions import (
    SERVICE_TIME_PARAMS_BY_ESI,
    STUDY_PERIOD_DAYS,
    arrival_distribution,
    esi_distribution,
    load_data,
)

DEPARTMENT = "B"
OUT_TABLES = "results/tables/dept_b"


def main():
    df = load_data(department=DEPARTMENT)
    visits_per_day = len(df) / STUDY_PERIOD_DAYS

    by_hour_bin, by_day, by_month = arrival_distribution(df)
    esi_mix = esi_distribution(df)

    by_hour_bin.to_csv(f"{OUT_TABLES}/arrivals_by_hour_bin.csv", header=["count"])
    by_day.to_csv(f"{OUT_TABLES}/arrivals_by_day.csv", header=["count"])
    by_month.to_csv(f"{OUT_TABLES}/arrivals_by_month.csv", header=["count"])
    esi_mix.to_csv(f"{OUT_TABLES}/esi_mix.csv", header=["proportion"])

    service_time_df = pd.DataFrame(
        SERVICE_TIME_PARAMS_BY_ESI, index=["mean_minutes", "sd_minutes"]
    ).T
    service_time_df.index.name = "esi"
    service_time_df.to_csv(f"{OUT_TABLES}/service_time_params.csv")

    pd.Series(
        {"department": DEPARTMENT, "study_period_days": STUDY_PERIOD_DAYS,
         "total_visits": len(df), "visits_per_day": visits_per_day},
    ).to_csv(f"{OUT_TABLES}/calibration_meta.csv", header=["value"])

    print(f"Department {DEPARTMENT}: {len(df)} visits over {STUDY_PERIOD_DAYS} days"
          f" = {visits_per_day:.1f} visits/day\n")
    print("Arrivals by 4h bin:\n", by_hour_bin, "\n")
    print("ESI mix:\n", esi_mix, "\n")


if __name__ == "__main__":
    main()
