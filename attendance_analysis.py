import pandas as pd
from datetime import datetime

# Mapping of attendance codes to work value (fraction of day)
CODE_WORK_VALUES = {
    'W': 1.0, 'WHF': 1.0, 'WHH': 0.5, 'H': 1.0, 'HD': 0.5,
    'CO': 0.0, 'S': 0.0, 'L': 0.0, 'U': 0.0, 'OFF': 0.0
}

# ----------------------------
# Utility Functions
# ----------------------------

def normalize_id(emp_id):
    """Normalize employee IDs consistently."""
    try:
        return str(int(str(emp_id).strip()))
    except ValueError:
        return str(emp_id).strip()

def expected_work_value(weekday, contractual_days):
    """Return expected work fraction for a day based on contractual days per week."""
    if contractual_days == 5.5:
        return 1.0 if weekday in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'] else 0.5 if weekday == 'Saturday' else 0.0
    elif contractual_days == 6:
        return 1.0 if weekday in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'] else 0.0
    # Assume 7-day contract
    return 1.0

# ----------------------------
# Main Analysis Function
# ----------------------------

def calculate_comp_off_and_leave(df_wide, contracts_df, date_list, leave_codes=None):
    if leave_codes is None:
        leave_codes = ['L', 'S', 'U']

    leave_codes = [c.upper() for c in leave_codes]

    # Normalize contract employee IDs
    contracts_df['Employee #'] = contracts_df['Employee #'].apply(normalize_id)
    contract_dict = dict(zip(contracts_df['Employee #'], contracts_df['Contractual Days Per Week']))

    # Generate mappings using the precise datetime objects passed from app.py
    valid_cols = [dt.strftime("%d-%b") for dt in date_list]
    weekday_map = {dt.strftime("%d-%b"): dt.strftime("%A") for dt in date_list}
    datetime_map = {dt.strftime("%d-%b"): dt for dt in date_list}

    comp_records = []
    leave_records = []
    missing_contracts = []

    # Iterate employees
    for _, row in df_wide.iterrows():
        emp_id = normalize_id(row['Employee #'])
        emp_name = row.get('Employee Name', '')
        designation = row.get('Designation', '')
        company = row.get('Company', '')

        contractual = contract_dict.get(emp_id)

        if contractual is None:
            missing_contracts.append(emp_id)
            leave_records.append({
                'Employee #': emp_id,
                'Employee Name': emp_name,
                'Designation': designation,
                'Company': company,
                'Contractual Days/Week': 'Unknown',
                'Total Leave Days': 0,
                'Leave Days (Dates)': 'None',
                'Comp-Off Earned (Days)': 0
            })
            continue

        weekly_actual = {}
        weekly_expected = {}
        leave_days = []

        # Day-level processing mapped perfectly across month boundaries
        for col_name in valid_cols:
            code = str(row.get(col_name, "")).upper().strip()
            actual = CODE_WORK_VALUES.get(code, 0.0)

            if code in leave_codes:
                leave_days.append(col_name) # Store the actual "20-May" formatted string

            weekday = weekday_map[col_name]
            expected = expected_work_value(weekday, contractual)

            try:
                dt = datetime_map[col_name]
                iso_year, iso_week, _ = dt.isocalendar()
                week_key = (iso_year, iso_week)
            except ValueError:
                continue

            weekly_actual[week_key] = weekly_actual.get(week_key, 0.0) + actual
            weekly_expected[week_key] = weekly_expected.get(week_key, 0.0) + expected

        # Weekly Excess Calculation
        total_comp = sum(max(0, actual - weekly_expected.get(wk, 0)) for wk, actual in weekly_actual.items())
        total_comp = round(total_comp, 1)

        # Record comp-off only if >0
        if total_comp > 0:
            comp_records.append({
                'Employee #': emp_id,
                'Employee Name': emp_name,
                'Designation': designation,
                'Company': company,
                'Contractual Days/Week': contractual,
                'Comp-Off Earned (Days)': total_comp
            })

        leave_records.append({
            'Employee #': emp_id,
            'Employee Name': emp_name,
            'Designation': designation,
            'Company': company,
            'Contractual Days/Week': contractual,
            'Total Leave Days': len(leave_days),
            'Leave Days (Dates)': ', '.join(leave_days) if leave_days else 'None',
            'Comp-Off Earned (Days)': total_comp
        })

    comp_report = pd.DataFrame(comp_records)
    leave_report = pd.DataFrame(leave_records)

    return comp_report, leave_report
