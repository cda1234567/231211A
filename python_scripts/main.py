import xlwings as xw
import os
from pathlib import Path
import datetime

def find_last_non_empty_column_value_in_row(ws, row_index):
    # C# 是 1-based，xlwings 也是 1-based
    for col in range(ws.used_range.columns.count, 0, -1):
        val = ws.cells(row_index, col).value
        if val is not None:
            try:
                return int(val)
            except Exception:
                continue
    return 0

def get_last_eight_digits(val):
    s = str(val) if val is not None else ""
    return s[-8:] if len(s) >= 8 else s

def update_worksheet_cells(main_ws, ws, i, j):
    last = main_ws.used_range.columns.count
    for col in range(last, 0, -1):
        if main_ws.cells(1, col).value is not None:
            # mainWorksheet.Cells[1, col + 3] = dataArray[2, 3];
            try:
                v23 = ws.cells(2, 3).value
            except Exception:
                v23 = ""
            main_ws.cells(1, col + 3).value = v23
            cell = main_ws.cells(1, col + 3)
            cell.font.name = "Arial"
            cell.font.size = 9
            cell.api.WrapText = True

            # mainWorksheet.Cells[1, col + 2] = lastEightDigits;
            try:
                val7 = ws.cells(1, 7).value
            except Exception:
                val7 = ""
            g3 = str(val7) if val7 is not None else ""
            lastEightDigits = g3[-8:] if len(g3) >= 8 else g3
            main_ws.cells(1, col + 2).value = lastEightDigits
            cell2 = main_ws.cells(1, col + 2)
            cell2.font.name = "Arial"
            cell2.font.size = 9
            cell2.api.WrapText = True

            last1 = find_last_non_empty_column_value_in_row(main_ws, i)
            try:
                if ws.cells(j, 7).value is not None and str(ws.cells(j, 7).value).strip() == "-":
                    break
            except Exception:
                pass
            try:
                ws.cells(j, 7).value = last1
            except Exception:
                pass

            try:
                f2 = int(ws.cells(j, 6).value)
            except Exception:
                f2 = 0
            main_ws.cells(i, col + 2).value = f2

            try:
                original_value = ws.cells(j, 10).value
                rounded = int(round(original_value)) if original_value is not None else 0
            except Exception:
                rounded = 0
            main_ws.cells(i, col + 3).value = rounded

            try:
                if (last1 - f2) < 0:
                    main_ws.cells(i, col + 3).color = (255, 0, 0)
            except Exception:
                pass
            break

def main():
    file_list = [
        '初始/庚霖倉庫存-主檔250612.xlsx',
        '初始/PO#4500058712 TA7-9.xls',
        '初始/PO#4500058712 T3iu-A 板 REV-1.2.xls',
        '初始/PO#4500058712 T3iu-B 板 REV-1.0.xls',
        '初始/PO#4500058712 T356789iu_Main_Board (A) REV-3.41.xls',
        '初始/PO#4500058712 T356789iu_IDE_Board (B).xls',
        '初始/PO#4500058712 T356789iu_Main_Board (A) REV-3.41-2.xls',
        '初始/PO#4500058712 T356789iu_IDE_Board (B)-2.xls',
        '初始/PO#4500058712  T8u-REV-1.4.xls',
    ]

    timestamp = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    folder_path = Path(fr"\\St-nas\個人資料夾\Andy\excel\{timestamp}")
    folder_path.mkdir(parents=True, exist_ok=True)

    app = xw.App(visible=False)
    try:
        main_wb = app.books.open(file_list[0])
        main_ws = main_wb.sheets[0]

        for k in range(1, len(file_list)):
            wb = app.books.open(file_list[k])
            ws = wb.sheets[0]

            main_rows = main_ws.used_range.rows.count
            sec_rows = ws.used_range.rows.count

            for j in range(1, sec_rows + 1):
                for i in range(1, main_rows + 1):
                    try:
                        sec_key = str(ws.cells(j, 3).value).strip() if ws.cells(j, 3).value is not None else ""
                        main_key = str(main_ws.cells(i, 1).value).strip() if main_ws.cells(i, 1).value is not None else ""
                    except Exception:
                        continue
                    if sec_key == main_key:
                        update_worksheet_cells(main_ws, ws, i, j)
                        break

            wb.save(folder_path / Path(file_list[k]).name)
            wb.close()

        main_wb.save(folder_path / Path(file_list[0]).name)
        main_wb.close()
    finally:
        app.quit()

if __name__ == '__main__':
    main() 