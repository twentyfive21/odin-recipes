import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- SETUP FILE PATHS ---
folder = "H:/report_gen"
aps_path = os.path.join(folder, "aps.xlsx")
cio_path = os.path.join(folder, "cio.xlsx")
lookup_path = os.path.join(folder, "s.xlsx")
output_path = os.path.join(folder, "attestation_report.xlsx")

# --- LOAD DATA ---
aps_df = pd.read_excel(aps_path)
cio_df = pd.read_excel(cio_path)
lookup1 = pd.read_excel(lookup_path, sheet_name="Sheet1")    # Changed to Sheet1
lookup2 = pd.read_excel(lookup_path, sheet_name="Sheet2")    # Changed to Sheet2

# --- CLEAN & RENAME APS FILE ---
aps_df.columns = ['contextid', 'context name', 'title', 'status', 'due date', 'completion date', 'accountable']
aps_df = aps_df.drop(columns=aps_df.columns[7:], errors='ignore')  # Drop extra cols if any

aps_df.rename(columns={
    'completion date': 'Application Owner',
    'accountable': 'APS Accountable'
}, inplace=True)

print("APS columns after rename:", aps_df.columns.tolist())

# --- CLEAN & RENAME CIO FILE ---
cio_df.columns = ['contextid', 'context name', 'title', 'status', 'due date', 'completion date', 'accountable']
cio_df = cio_df.drop(columns=['completion date'], errors='ignore')
cio_df = cio_df.drop(columns=cio_df.columns[6:], errors='ignore')

cio_df.rename(columns={
    'accountable': 'Application Owner'
}, inplace=True)

print("CIO columns after rename:", cio_df.columns.tolist())

# --- ADD NEW COLUMNS BEFORE & AFTER Due Date ---
for df in [aps_df, cio_df]:
    df.insert(4, 'CPPM Review Status', '')
    df.insert(5, 'PECM_Delegate', '')

print("APS columns after insert new cols:", aps_df.columns.tolist())
print("CIO columns after insert new cols:", cio_df.columns.tolist())

# --- RENAME COLUMNS TO FINAL NAMES ---
aps_df.columns = [
    'AIT #', 'AIT Name', 'Assessment Name', 'Trident status',
    'CPPM Review Status', 'Due Date', 'PECM_Delegate',
    'Application Owner', 'APS Accountable'
]

cio_df.columns = [
    'AIT #', 'AIT Name', 'Assessment Name', 'Trident status',
    'CPPM Review Status', 'Due Date', 'PECM_Delegate',
    'Application Owner'
]

print("APS columns after final rename:", aps_df.columns.tolist())
print("CIO columns after final rename:", cio_df.columns.tolist())

# --- ADD EMPTY FINAL COLUMNS ---
aps_df['Support Owner'] = ''
aps_df['APS SLT'] = ''
aps_df['Tech Exec'] = ''
aps_df['CIO Exec'] = ''

cio_df['APS Accountable'] = ''
cio_df['Support Owner'] = ''
cio_df['APS SLT'] = ''
cio_df['Tech Exec'] = ''
cio_df['CIO Exec'] = ''

print("APS final columns:", aps_df.columns.tolist())
print("CIO final columns:", cio_df.columns.tolist())

# The rest of your code (merging, pivoting, exporting, formatting) remains unchanged
# Just make sure to update your lookup sheets names and due date as you want