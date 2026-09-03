#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import CellIsRule
except ImportError:
    print("Error: openpyxl not installed. Install with: pip install openpyxl")
    sys.exit(1)

def main():
    try:
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Ocena kandydatow"

        headers = ["ID", "Imie", "Nazwisko", "Stanowisko", "Data", "Wiedza", "Microsoft", "Rozwiazywanie", "Bezpieczenstwo", "Komunikacja", "Wspolpraca", "Motywacja", "Wynik", "Rekomendacja", "Uwagi"]

        header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        for col_num, header in enumerate(headers, 1):
            cell = ws1.cell(row=1, column=col_num)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border

        column_widths = [8, 10, 10, 12, 10, 10, 10, 12, 12, 10, 10, 8, 8, 12, 15]
        for idx, width in enumerate(column_widths, 1):
            ws1.column_dimensions[get_column_letter(idx)].width = width

        ws1.freeze_panes = "A2"

        for row in range(2, 32):
            for col in range(1, len(headers) + 1):
                cell = ws1.cell(row=row, column=col)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name='Arial', size=10)

        dv = DataValidation(type="list", formula1='"1,2,3,4,5"', allow_blank=True)
        ws1.add_data_validation(dv)

        for row in range(2, 32):
            for col in range(6, 13):
                dv.add(ws1.cell(row=row, column=col))

        for row in range(2, 32):
            cell = ws1.cell(row=row, column=13)
            cell.value = '=IF(COUNTA(F' + str(row) + ':L' + str(row) + ')=0,"",ROUND(AVERAGE(F' + str(row) + ':L' + str(row) + '),2))'
            cell.number_format = '0.00'
            cell.font = Font(name='Arial', size=10, bold=True)

        for row in range(2, 32):
            cell = ws1.cell(row=row, column=14)
            cell.value = '=IF(M' + str(row) + '="","",IF(M' + str(row) + '>=4.5,"TAK",IF(M' + str(row) + '>=3.5,"MOZE","NIE")))'
            cell.font = Font(name='Arial', size=10, bold=True)

        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        green_font = Font(name='Arial', size=10, color="006100", bold=True)
        yellow_font = Font(name='Arial', size=10, color="9C6500", bold=True)
        red_font = Font(name='Arial', size=10, color="9C0006", bold=True)

        ws1.conditional_formatting.add('M2:M31', CellIsRule(operator='greaterThanOrEqual', formula=['4.5'], fill=green_fill, font=green_font))
        ws1.conditional_formatting.add('M2:M31', CellIsRule(operator='greaterThanOrEqual', formula=['3.5'], fill=yellow_fill, font=yellow_font))
        ws1.conditional_formatting.add('M2:M31', CellIsRule(operator='lessThan', formula=['3.5'], fill=red_fill, font=red_font))

        ws1.conditional_formatting.add('N2:N31', CellIsRule(operator='equal', formula=['"TAK"'], fill=green_fill, font=green_font))
        ws1.conditional_formatting.add('N2:N31', CellIsRule(operator='equal', formula=['"MOZE"'], fill=yellow_fill, font=yellow_font))
        ws1.conditional_formatting.add('N2:N31', CellIsRule(operator='equal', formula=['"NIE"'], fill=red_fill, font=red_font))

        ws1.auto_filter.ref = 'A1:O31'

        ws2 = wb.create_sheet("Pytania")

        ws2.merge_cells('A1:B1')
        title = ws2['A1']
        title.value = "PYTANIA REKRUTACYJNE"
        title.font = Font(name='Arial', size=12, bold=True, color="FFFFFF")
        title.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        title.alignment = Alignment(horizontal="center", vertical="center")
        ws2.row_dimensions[1].height = 20

        current_row = 2

        for i in range(1, 6):
            ws2.merge_cells('A' + str(current_row) + ':B' + str(current_row))
            cell = ws2['A' + str(current_row)]
            cell.value = "Pytanie " + str(i)
            cell.font = Font(name='Arial', size=10, bold=True)
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.border = thin_border
            ws2.row_dimensions[current_row].height = 30
            current_row += 1

        ws2.column_dimensions['A'].width = 70
        ws2.column_dimensions['B'].width = 2

        output_file = "Ocena_Kandydatow.xlsx"
        wb.save(output_file)
        
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print("Success: " + output_file + " (" + str(file_size) + " bytes)")
            return 0
        else:
            print("Error: File was not created")
            return 1
            
    except Exception as e:
        print("Error: " + str(e))
        return 1

if __name__ == "__main__":
    sys.exit(main())
