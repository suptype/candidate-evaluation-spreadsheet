import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
import os

# Create a new workbook
wb = openpyxl.Workbook()

# ==================== ARKUSZ 1: OCENA KANDYDATÓW ====================
ws1 = wb.active
ws1.title = "Ocena kandydatów"

# Define headers
headers = [
    "ID kandydata",
    "Imię",
    "Nazwisko",
    "Stanowisko",
    "Data rozmowy",
    "Wiedza techniczna",
    "Znajomość Microsoft 365",
    "Rozwiązywanie problemów",
    "Świadomość cyberbezpieczeństwa",
    "Umiejętności komunikacyjne",
    "Współpraca zespołowa",
    "Motywacja",
    "Wynik końcowy",
    "Rekomendacja",
    "Uwagi rekrutera"
]

# Add headers to the first row
for col_num, header in enumerate(headers, 1):
    cell = ws1.cell(row=1, column=col_num)
    cell.value = header
    cell.font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Set column widths
column_widths = [12, 12, 12, 15, 15, 12, 18, 18, 20, 18, 15, 12, 12, 15, 25]
for idx, width in enumerate(column_widths, 1):
    ws1.column_dimensions[get_column_letter(idx)].width = width

# Freeze the first row
ws1.freeze_panes = "A2"

# Add borders style
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

# Add 30 empty rows with borders and data validation
for row in range(2, 32):
    for col in range(1, len(headers) + 1):
        cell = ws1.cell(row=row, column=col)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # Add data validation for rating columns (6-12)
        if 6 <= col <= 12:
            cell.alignment = Alignment(horizontal="center", vertical="center")

# Create data validation for ratings (1-5)
dv_rating = DataValidation(type="list", formula1='"1,2,3,4,5"', allow_blank=True)
dv_rating.error = 'Wpisz wartość od 1 do 5'
dv_rating.errorTitle = 'Nieprawidłowa ocena'
ws1.add_data_validation(dv_rating)

# Apply data validation to rating columns (F to L - columns 6 to 12)
for row in range(2, 32):
    for col in range(6, 13):
        dv_rating.add(ws1.cell(row=row, column=col))

# Add formulas for "Wynik końcowy" (column M - column 13)
for row in range(2, 32):
    cell = ws1.cell(row=row, column=13)
    cell.value = f'=IF(COUNTA(F{row}:L{row})=0,"",ROUND(AVERAGE(F{row}:L{row}),2))'
    cell.number_format = '0.00'
    cell.font = Font(bold=True)

# Add formulas for "Rekomendacja" (column N - column 14)
for row in range(2, 32):
    cell = ws1.cell(row=row, column=14)
    cell.value = f'=IF(M{row}="","",IF(M{row}>=4.5,"Zatrudnić",IF(M{row}>=3.5,"Rozważyć","Odrzucić")))'
    cell.font = Font(bold=True)

# Conditional formatting for "Wynik końcowy" (column M)
green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

green_font = Font(color="006100")
yellow_font = Font(color="9C6500")
red_font = Font(color="9C0006")

# Green for >= 4.5
green_rule = CellIsRule(operator='greaterThanOrEqual', formula=['4.5'], fill=green_fill, font=green_font)
ws1.conditional_formatting.add(f'M2:M31', green_rule)

# Yellow for 3.5-4.49
yellow_rule = CellIsRule(operator='greaterThanOrEqual', formula=['3.5'], fill=yellow_fill, font=yellow_font)
ws1.conditional_formatting.add(f'M2:M31', yellow_rule)

# Red for < 3.5
red_rule = CellIsRule(operator='lessThan', formula=['3.5'], fill=red_fill, font=red_font)
ws1.conditional_formatting.add(f'M2:M31', red_rule)

# Conditional formatting for "Rekomendacja" column (N)
green_rule_rec = CellIsRule(operator='equal', formula=['"Zatrudnić"'], fill=green_fill, font=green_font)
yellow_rule_rec = CellIsRule(operator='equal', formula=['"Rozważyć"'], fill=yellow_fill, font=yellow_font)
red_rule_rec = CellIsRule(operator='equal', formula=['"Odrzucić"'], fill=red_fill, font=red_font)

ws1.conditional_formatting.add(f'N2:N31', green_rule_rec)
ws1.conditional_formatting.add(f'N2:N31', yellow_rule_rec)
ws1.conditional_formatting.add(f'N2:N31', red_rule_rec)

# Enable AutoFilter
ws1.auto_filter.ref = f'A1:{get_column_letter(len(headers))}31'

# ==================== ARKUSZ 2: PYTANIA REKRUTACYJNE ====================
ws2 = wb.create_sheet("Pytania rekrutacyjne")

# Title
title_cell = ws2['A1']
title_cell.value = "PYTANIA REKRUTACYJNE"
title_cell.font = Font(name='Calibri', size=14, bold=True, color="FFFFFF")
title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
title_cell.alignment = Alignment(horizontal="center", vertical="center")
ws2.merge_cells('A1:B1')
ws2.row_dimensions[1].height = 25

# Subtitle
subtitle_cell = ws2['A2']
subtitle_cell.value = "Sprawdzenie kompetencji"
subtitle_cell.font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
subtitle_cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
ws2.merge_cells('A2:B2')
ws2.row_dimensions[2].height = 20

# Category header
category_cell = ws2['A3']
category_cell.value = "KATEGORIA: KOMPETENCJE TECHNICZNE"
category_cell.font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
category_cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
ws2.merge_cells('A3:B3')
ws2.row_dimensions[3].height = 20

# Questions
questions = [
    "1. Jakimi usługami Microsoft 365 zarządzał(a) Pan/Pani w poprzednich miejscach pracy?",
    "2. W jaki sposób zabezpieczył(a)by Pan/Pani środowisko Microsoft 365?",
    "3. Jakie są różnice pomiędzy SharePoint, OneDrive i Teams?",
    "4. Jak diagnozować problemy z logowaniem użytkowników?",
    "5. Proszę opowiedzieć o najtrudniejszym incydencie technicznym, który udało się rozwiązać."
]

current_row = 4
for question in questions:
    q_cell = ws2.cell(row=current_row, column=1)
    q_cell.value = question
    q_cell.font = Font(name='Calibri', size=10, bold=False)
    q_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    q_cell.border = thin_border
    ws2.merge_cells(f'A{current_row}:B{current_row}')
    
    # Set row height for better readability
    ws2.row_dimensions[current_row].height = 35
    current_row += 1

# Add space for notes section
ws2.row_dimensions[current_row].height = 15
current_row += 1

notes_header = ws2.cell(row=current_row, column=1)
notes_header.value = "UWAGI I OBSERWACJE"
notes_header.font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
notes_header.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
ws2.merge_cells(f'A{current_row}:B{current_row}')
ws2.row_dimensions[current_row].height = 20

# Add space for notes
for i in range(5):
    current_row += 1
    note_cell = ws2.cell(row=current_row, column=1)
    note_cell.border = thin_border
    note_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws2.merge_cells(f'A{current_row}:B{current_row}')
    ws2.row_dimensions[current_row].height = 30

# Set column widths for sheet 2
ws2.column_dimensions['A'].width = 80
ws2.column_dimensions['B'].width = 2

# Save the workbook
output_path = "Ocena_Kandydatow.xlsx"
wb.save(output_path)
print(f"✓ Skoroszyt Excel został utworzony: {output_path}")
