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

def get_weekday_name(day_num, month, year):
    """Return weekday name (e.g., 'Monday') for a given day in the month/year."""
    try:
        return datetime(year, month, day_num).strftime("%A")
    except ValueError:
        return None

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

def calculate_comp_off_and_leave(df_wide, contracts_df, month, year, leave_codes=None):
    if leave_codes is None:
        leave_codes = ['L', 'S', 'U']

    leave_codes = [c.upper() for c in leave_codes]

    # Normalize contract employee IDs
    contracts_df['Employee #'] = contracts_df['Employee #'].apply(normalize_id)
    contract_dict = dict(zip(contracts_df['Employee #'], contracts_df['Contractual Days Per Week']))

    # Identify valid day columns
    day_cols = [c for c in df_wide.columns if c.startswith('Day ')]
    day_nums = [int(c.split()[1]) for c in day_cols]

    weekday_map = {}
    valid_days = []

    for d in day_nums:
        wd = get_weekday_name(d, month, year)
        if wd:
            weekday_map[d] = wd
            valid_days.append(d)

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
                'Leave Days (Numbers)': 'None',
                'Comp-Off Earned (Days)': 0
            })
            continue

        weekly_actual = {}
        weekly_expected = {}
        leave_days = []

        # Day-level processing
        for day_num in valid_days:
            code = str(row[f'Day {day_num}']).upper().strip()
            actual = CODE_WORK_VALUES.get(code, 0.0)

            if code in leave_codes:
                leave_days.append(day_num)

            weekday = weekday_map[day_num]
            expected = expected_work_value(weekday, contractual)

            try:
                iso_year, iso_week, _ = datetime(year, month, day_num).isocalendar()
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
            'Leave Days (Numbers)': ', '.join(map(str, sorted(leave_days))) if leave_days else 'None',
            'Comp-Off Earned (Days)': total_comp
        })

    comp_report = pd.DataFrame(comp_records)
    leave_report = pd.DataFrame(leave_records)

    return comp_report, leave_report
