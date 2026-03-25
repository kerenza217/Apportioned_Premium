import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from collections import defaultdict

# Load Excel
df = pd.read_excel("PREMIUM.xlsx")

# Clean column names
df.columns = df.columns.str.strip()

# Convert date columns
df['START DATE'] = pd.to_datetime(df['START DATE'], errors='coerce')
df['END DATE'] = pd.to_datetime(df['END DATE'], errors='coerce')

# Drop rows with missing/invalid dates
df = df.dropna(subset=['START DATE', 'END DATE'])

# Get earliest START DATE and build monthly periods to Dec 2026
start_month = df['START DATE'].min().replace(day=1)
end_month = datetime(2026, 12, 31).replace(day=1)

# Build monthly periods
monthly_periods = []
current = start_month
while current <= end_month:
    monthly_periods.append(current)
    current += relativedelta(months=1)

# Format for output columns: 'Jan-2024', etc.
month_labels = [dt.strftime('%b-%Y') for dt in monthly_periods]

# Store results
output_rows = []

# Process each row in the original data
for _, row in df.iterrows():
    start_date = row['START DATE']
    end_date = row['END DATE']
    total_premium = row['PREMIUM AMOUNT']

    # Skip invalid periods
    if pd.isnull(start_date) or pd.isnull(end_date) or end_date < start_date:
        continue

    # Calculate total number of days in the period
    total_days = (end_date - start_date).days + 1
    daily_rate = total_premium / total_days

    # Build month allocation
    allocations = defaultdict(float)

    for month_start in monthly_periods:
        month_end = (month_start + relativedelta(months=1)) - timedelta(days=1)

        # Overlap with current policy period
        overlap_start = max(start_date, month_start)
        overlap_end = min(end_date, month_end)

        if overlap_start <= overlap_end:
            days_in_month = (overlap_end - overlap_start).days + 1
            month_label = month_start.strftime('%b-%Y')
            allocations[month_label] += round(days_in_month * daily_rate, 2)

    # Combine original row with month allocations
    result_row = row.to_dict()
    for label in month_labels:
        result_row[label] = allocations.get(label, 0.0)

    output_rows.append(result_row)

# Create final DataFrame
final_df = pd.DataFrame(output_rows)

# Save to Excel
final_df.to_excel("Apportioned_Premium1.xlsx", index=False)
print("✅ Apportioned premium saved to Apportioned_Premium.xlsx")
