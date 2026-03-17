import pandas as pd

# ==============================
# Step 1 — Setup
# ==============================
date_col = "date_of_approval_void_decline"

# Ensure datetime
monthly_filtered[date_col] = pd.to_datetime(monthly_filtered[date_col], errors='coerce')
quarterly_filtered[date_col] = pd.to_datetime(quarterly_filtered[date_col], errors='coerce')

# Create month column
monthly_filtered["month"] = monthly_filtered[date_col].dt.strftime("%Y-%m")
quarterly_filtered["month"] = quarterly_filtered[date_col].dt.strftime("%Y-%m")

# ==============================
# OPTION A — Record Count Trend
# ==============================

# Step 2 — Count per month
monthly_count = monthly_filtered.groupby("month").size().reset_index(name="monthly_count")
quarterly_count = quarterly_filtered.groupby("month").size().reset_index(name="quarterly_count")

trend_df = monthly_count.merge(
    quarterly_count, on="month", how="outer"
).sort_values("month")

# Step 3 — Month-over-month difference
trend_df["monthly_mom_diff"] = trend_df["monthly_count"].diff()
trend_df["quarterly_mom_diff"] = trend_df["quarterly_count"].diff()

# Step 4 — 3-month rolling average
trend_df["monthly_3m_avg"] = trend_df["monthly_count"].rolling(3).mean()
trend_df["quarterly_3m_avg"] = trend_df["quarterly_count"].rolling(3).mean()

# Step 5 — Compare trends
trend_df["mom_diff_gap"] = (
    trend_df["monthly_mom_diff"] - trend_df["quarterly_mom_diff"]
)

trend_df["avg_3m_gap"] = (
    trend_df["monthly_3m_avg"] - trend_df["quarterly_3m_avg"]
)

# ==============================
# OPTION B — Fill Rate Trend
# ==============================

def fill_rate_by_month(df):
    return df.groupby("month").apply(lambda x: x.notna().mean() * 100)

monthly_fill_trend = fill_rate_by_month(monthly_filtered)
quarterly_fill_trend = fill_rate_by_month(quarterly_filtered)

# ==============================
# Column-level comparison
# ==============================

col = "your_column_name"  # change this

compare_trend = pd.DataFrame({
    "monthly": monthly_fill_trend[col],
    "quarterly": quarterly_fill_trend[col]
})

# MoM difference
compare_trend["mom_diff_monthly"] = compare_trend["monthly"].diff()
compare_trend["mom_diff_quarterly"] = compare_trend["quarterly"].diff()

# 3-month average
compare_trend["3m_avg_monthly"] = compare_trend["monthly"].rolling(3).mean()
compare_trend["3m_avg_quarterly"] = compare_trend["quarterly"].rolling(3).mean()

# ==============================
# Export to CSV
# ==============================

trend_df.to_csv("record_count_trend.csv", index=False)
compare_trend.to_csv(f"{col}_fill_rate_trend.csv")

# Optional: full fill trend
monthly_fill_trend.to_csv("monthly_fill_trend.csv")
quarterly_fill_trend.to_csv("quarterly_fill_trend.csv")

# ==============================
# Outputs
# ==============================

trend_df
compare_trend
