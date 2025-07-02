import pandas as pd

# Load your Excel file
df = pd.read_excel("cal_data.xlsx")

# Convert each column to a set, dropping blanks and converting to strings
ucal_q1 = set(df['UCAL_Q1'].dropna().astype(str))
nonucal_q1 = set(df['nonUCAL_Q1'].dropna().astype(str))
ucal_q2 = set(df['UCAL_Q2'].dropna().astype(str))
nonucal_q2 = set(df['nonUCAL_Q2'].dropna().astype(str))

# Analysis

# Items that were in Q1 UCAL but removed in Q2 (not present in either Q2 list)
ucal_removed = ucal_q1 - (ucal_q2 | nonucal_q2)
# Items that were in Q1 non-UCAL but removed in Q2
nonucal_removed = nonucal_q1 - (ucal_q2 | nonucal_q2)

# Items that are new in Q2 UCAL (not present in either Q1 list)
ucal_added = ucal_q2 - (ucal_q1 | nonucal_q1)
# Items that are new in Q2 non-UCAL
nonucal_added = nonucal_q2 - (ucal_q1 | nonucal_q1)

# Items that moved from Q1 UCAL to Q2 non-UCAL
ucal_to_nonucal = ucal_q1 & nonucal_q2
# Items that moved from Q1 non-UCAL to Q2 UCAL
nonucal_to_ucal = nonucal_q1 & ucal_q2

# Print results with clear labeling
print("=== Changes from Q1 to Q2 ===")
print(f"Items moved from UCAL (Q1) → non-UCAL (Q2): {ucal_to_nonucal}")
print(f"Items moved from non-UCAL (Q1) → UCAL (Q2): {nonucal_to_ucal}")
print(f"Items removed from UCAL (present in Q1 but NOT in Q2): {ucal_removed}")
print(f"Items removed from non-UCAL (present in Q1 but NOT in Q2): {nonucal_removed}")
print(f"New items added to UCAL in Q2 (not in Q1): {ucal_added}")
print(f"New items added to non-UCAL in Q2 (not in Q1): {nonucal_added}")