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
lookup1 = pd.read_excel(lookup_path, sheet_name="Sheet1")  # For Support Owner, APS SLT
lookup2 = pd.read_excel(lookup_path, sheet_name="Sheet2")  # For Tech Exec, CIO Exec

# --- CLEAN & RENAME APS FILE ---
aps_df.columns = ['contextid', 'context name', 'title', 'status', 'due date', 'completion date', 'accountable']
aps_df = aps_df.drop(columns=aps_df.columns[7:], errors='ignore')  # Drop extra cols if exist

# Print debug
print("APS DF after drop, before insert:")
print("Columns:", aps_df.columns.tolist())
print("Count:", len(aps_df.columns))

# Insert columns
aps_df.insert(4, 'CPPM Review Status', '')
aps_df.insert(5, 'PECM_Delegate', '')

# Rename columns after insert
aps_df.columns = [
    'AIT #', 'AIT Name', 'Assessment Name', 'Trident status',
    'CPPM Review Status', 'Due Date', 'PECM_Delegate',
    'Application Owner', 'APS Accountable'
]

# --- CLEAN & RENAME CIO FILE ---
cio_df.columns = ['contextid', 'context name', 'title', 'status', 'due date', 'completion date', 'accountable']
cio_df = cio_df.drop(columns=['completion date'], errors='ignore')
cio_df = cio_df.drop(columns=cio_df.columns[6:], errors='ignore')

# Print debug
print("CIO DF after drop, before insert:")
print("Columns:", cio_df.columns.tolist())
print("Count:", len(cio_df.columns))

# Insert columns
cio_df.insert(4, 'CPPM Review Status', '')
cio_df.insert(5, 'PECM_Delegate', '')

# Rename columns after insert
cio_df.columns = [
    'AIT #', 'AIT Name', 'Assessment Name', 'Trident status',
    'CPPM Review Status', 'Due Date', 'PECM_Delegate',
    'Application Owner'
]

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

# --- PERFORM VLOOKUPS (MERGE) ---
lookup1 = lookup1[['app_id', 'Support Owner', 'APS SLT']]
lookup2 = lookup2[['app_id', 'Tech Exec', 'CIO Exec']]

aps_df = aps_df.merge(lookup1, how='left', left_on='AIT #', right_on='app_id').drop(columns='app_id')
aps_df = aps_df.merge(lookup2, how='left', left_on='AIT #', right_on='app_id').drop(columns='app_id')

cio_df = cio_df.merge(lookup1, how='left', left_on='AIT #', right_on='app_id').drop(columns='app_id')
cio_df = cio_df.merge(lookup2, how='left', left_on='AIT #', right_on='app_id').drop(columns='app_id')

# --- CREATE PIVOT TABLES ---
aps_pivot = aps_df.pivot_table(
    index=['CIO Exec', 'Tech Exec'],
    columns='Trident status',
    values='AIT #',
    aggfunc='count'
).reset_index()

cio_pivot = cio_df.pivot_table(
    index=['CIO Exec', 'Tech Exec'],
    columns='Trident status',
    values='AIT #',
    aggfunc='count'
).reset_index()

# --- CREATE MODIFIED PIVOTS (CPPM Review Status) ---
aps_mod_pivot = aps_df.pivot_table(
    index=['CIO Exec', 'Tech Exec'],
    columns='CPPM Review Status',
    values='AIT #',
    aggfunc='count'
).reset_index()

cio_mod_pivot = cio_df.pivot_table(
    index=['CIO Exec', 'Tech Exec'],
    columns='CPPM Review Status',
    values='AIT #',
    aggfunc='count'
).reset_index()

# --- EXPORT ALL TO EXCEL ---
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    aps_df.to_excel(writer, sheet_name="APS Attestations", index=False)
    cio_df.to_excel(writer, sheet_name="CIO Attestations", index=False)
    aps_pivot.to_excel(writer, sheet_name="Apivot", index=False)
    cio_pivot.to_excel(writer, sheet_name="Cpivot", index=False)
    aps_mod_pivot.to_excel(writer, sheet_name="APS Summary", index=False)
    cio_mod_pivot.to_excel(writer, sheet_name="CIO Summary", index=False)

# --- APPLY FORMATTING TO FINAL SUMMARY SHEETS ---
wb = load_workbook(output_path)

green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
white_font = Font(color="FFFFFF", size=16, bold=True)
header_align = Alignment(horizontal="center", vertical="center")
border = Border(left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin'))

def format_sheet(ws, fill_color):
    ws.insert_rows(1, amount=3)

    ws['A4'] = "APS Attestation summary" if ws.title == "APS Summary" else "CIO Attestation summary"
    ws.merge_cells('A4:F4')
    ws['A4'].font = white_font
    ws['A4'].alignment = header_align
    ws['A4'].fill = fill_color

    ws['B1'] = "Due Date"
    ws['B2'] = "6/20/2025"
    for col in ['B', 'C', 'D', 'E', 'F']:
        ws[f'{col}1'].alignment = header_align

    ws['A3'] = "Count of AITs"
    ws['A3'].font = Font(bold=True)

    for row in [4, 9, 18, 20, 23, 29]:
        for col in 'ABCDEFG':
            cell = ws[f'{col}{row}']
            cell.fill = fill_color

    for row in ws.iter_rows(min_row=1, max_row=31, min_col=1, max_col=7):
        for cell in row:
            cell.border = border

    for col in 'ABCDEFG':
        cell = ws[f'{col}51']
        cell.font = Font(bold=True)

    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

format_sheet(wb["APS Summary"], blue_fill)
format_sheet(wb["CIO Summary"], green_fill)

# Save workbook
wb.save(output_path)
print(f"✅ Done! File saved to: {output_path}")