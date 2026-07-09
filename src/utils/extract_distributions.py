"""
Week 3: extract calibration distributions for the ER DES.

Arrivals and patient volume come from the real Kaggle dataset
(data/processed/triage_data.parquet). Service/treatment time does not exist
in that dataset (no discharge timestamp, only 4-hour arrival bins), so those
come from published ED simulation literature, stratified by ESI acuity
level (1 = most urgent, 5 = least urgent).

Service time source: log-normal parameters commonly used in ED DES studies,
e.g. Ahalt et al. (2018) and similar ED queueing simulation papers that
report mean/SD treatment time by ESI level.

The dataset combines three EDs (`dep_name` A/B/C: one academic, two
community — per the Kaggle dataset description) collected March 2014 -
July 2017 (1248 days). Our DES models one ED, so we calibrate on a single
department (default: A, the largest/academic site) rather than the pooled
total — pooling all three sites would overstate one ED's real arrival rate.
"""

import pandas as pd

RAW = "data/processed/triage_data.parquet"
OUT_TABLES = "results/tables"
OUT_FIGURES = "results/figures"

DEPARTMENT = "A"
STUDY_PERIOD_DAYS = 1248  # March 2014 - July 2017, per Kaggle dataset description

# (mean_minutes, sd_minutes) per ESI level, log-normal, from ED sim literature
SERVICE_TIME_PARAMS_BY_ESI = {
    1: (180, 90),
    2: (150, 75),
    3: (120, 60),
    4: (75, 40),
    5: (45, 25),
}


def load_data(department=DEPARTMENT):
    df = pd.read_parquet(RAW)
    return df[df["dep_name"] == department]


def arrival_distribution(df):
    """Real arrival counts by 4-hour bin, day of week, and month, for one ED."""
    by_hour_bin = df["arrivalhour_bin"].value_counts().sort_index()
    by_day = df["arrivalday"].value_counts()
    by_month = df["arrivalmonth"].value_counts()
    return by_hour_bin, by_day, by_month


def esi_distribution(df):
    """Real ESI (acuity) mix — drives which service-time distribution applies."""
    return df["esi"].dropna().astype(int).value_counts(normalize=True).sort_index()


def main():
    df = load_data()
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
    print("Arrivals by day:\n", by_day, "\n")
    print("Arrivals by month:\n", by_month, "\n")
    print("ESI mix:\n", esi_mix, "\n")
    print("Service time params (literature, by ESI):\n", service_time_df)


if __name__ == "__main__":
    main()
