# Apportioned_Premium
This repository contains a Python-based solution for automating the pro-rata distribution of insurance premiums across monthly financial periods

🚀 The Problem & Solution
The Manual Headache: Previously, calculating how much premium belongs to which month required manual spreadsheet manipulation. For a policy spanning two years, a user had to manually calculate the daily rate and then count the specific number of days falling into each calendar month (considering leap years and varying month lengths). This was:

Time-consuming: Taking hours of manual data entry.
Error-prone: High risk of "off-by-one" day errors.
Static: Hard to scale when dealing with hundreds of policies.

The Automated Fix: This script transforms a simple list of policy dates and amounts into a comprehensive monthly schedule. It calculates the precise daily rate for every policy and maps the cost to the correct "month-bucket" up until December 2026.

Key Impact:
-Efficiency: Reduced hours of manual work to seconds of execution.
-Accuracy: Uses relativedelta and timedelta to ensure calendar-perfect precision.
-Scalability: Processes hundreds of rows instantly

📊 How It WorksThe script follows a linear logic to ensure financial integrity where the sum of monthly apportionments equals the total premium.
Data Ingestion: Loads PREMIUM.xlsx and cleans column headers.
Date Normalization: Converts start/end strings into Python datetime objects.
Daily Rate Calculation:
  $$\text{Daily Rate} = \frac{\text{Total Premium}}{\text{Total Days in Policy Period}}
  
$$Overlap Logic: For every month in the timeline (Start Date to Dec 2026), the script identifies if the policy was active and calculates the specific number of days of coverage in that month.Output: Generates a new Excel file, Apportioned_Premium1.xlsx, featuring the original data plus new columns for every month
