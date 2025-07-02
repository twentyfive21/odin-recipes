import pandas as pd

# Load your Excel file
df = pd.read_excel("cal_data.xlsx")

# Convert each column to sets of integers, dropping blanks
ucal_q1 = set(df['UCAL_Q1'].dropna().astype(int))
nonucal_q1 = set(df['nonUCAL_Q1'].dropna().astype(int))
ucal_q2 = set(df['UCAL_Q2'].dropna().astype(int))
nonucal_q2 = set(df['nonUCAL_Q2'].dropna().astype(int))

# Analysis
ucal_removed = ucal_q1 - (ucal_q2 | nonucal_q2)
nonucal_removed = nonucal_q1 - (ucal_q2 | nonucal_q2)

ucal_added = ucal_q2 - (ucal_q1 | nonucal_q1)
nonucal_added = nonucal_q2 - (ucal_q1 | nonucal_q1)

ucal_to_nonucal = ucal_q1 & nonucal_q2
nonucal_to_ucal = nonucal_q1 & ucal_q2

# Cleaner printing function
def print_clean(title, items):
    print(f"\n{title}:")
    if not items:
        print("  (none)")
        return
    for item in sorted(items):
        print(f"  {item}")

# Print all results with clean formatting
print_clean("Items moved from UCAL (Q1) → non-UCAL (Q2)", ucal_to_nonucal)
print_clean("Items moved from non-UCAL (Q1) → UCAL (Q2)", nonucal_to_ucal)
print_clean("Items removed from UCAL (present in Q1 but NOT in Q2)", ucal_removed)
print_clean("Items removed from non-UCAL (present in Q1 but NOT in Q2)", nonucal_removed)
print_clean("New items added to UCAL in Q2 (not in Q1)", ucal_added)
print_clean("New items added to non-UCAL in Q2 (not in Q1)", nonucal_added)