import pandas as pd

# ==============================
# BASIC COUNTS
# ==============================

total_monthly = len(monthly_filtered)
total_quarterly = len(quarterly_filtered)

monthly_unique = monthly_filtered["scorecard_id"].nunique()
quarterly_unique = quarterly_filtered["scorecard_id"].nunique()

monthly_ids = set(monthly_filtered["scorecard_id"])
quarterly_ids = set(quarterly_filtered["scorecard_id"])

common_ids = monthly_ids & quarterly_ids
only_in_monthly = monthly_ids - quarterly_ids
only_in_quarterly = quarterly_ids - monthly_ids

monthly_duplicates = monthly_filtered.duplicated(subset=["scorecard_id"]).sum()
quarterly_duplicates = quarterly_filtered.duplicated(subset=["scorecard_id"]).sum()

# ==============================
# MONTH-WISE COUNT
# ==============================

monthly_filtered["month"] = monthly_filtered["date_of_approval_void_decline"].dt.strftime("%Y-%m")
quarterly_filtered["month"] = quarterly_filtered["date_of_approval_void_decline"].dt.strftime("%Y-%m")

monthly_by_month = monthly_filtered.groupby("month").size().reset_index(name="monthly_count")
quarterly_by_month = quarterly_filtered.groupby("month").size().reset_index(name="quarterly_count")

month_comparison = monthly_by_month.merge(
    quarterly_by_month, on="month", how="outer"
)

month_comparison["difference"] = (
    month_comparison["monthly_count"] - month_comparison["quarterly_count"]
)

# ==============================
# SUMMARY TABLE
# ==============================

summary = pd.DataFrame({
    "Metric": [
        "Total Rows (Monthly)",
        "Total Rows (Quarterly)",
        "Distinct Scorecard IDs (Monthly)",
        "Distinct Scorecard IDs (Quarterly)",
        "Common Scorecard IDs",
        "Only in Monthly",
        "Only in Quarterly",
        "Duplicate Rows (Monthly)",
        "Duplicate Rows (Quarterly)"
    ],
    "Value": [
        total_monthly,
        total_quarterly,
        monthly_unique,
        quarterly_unique,
        len(common_ids),
        len(only_in_monthly),
        len(only_in_quarterly),
        monthly_duplicates,
        quarterly_duplicates
    ]
})

# ==============================
# EXPORT CSV
# ==============================

summary.to_csv("01_summary_counts.csv", index=False)
month_comparison.to_csv("02_monthly_counts.csv", index=False)
