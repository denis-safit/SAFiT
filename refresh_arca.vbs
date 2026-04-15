Set objExcel = CreateObject("Excel.Application")
' Nasconde Excel mentre lavora
objExcel.Visible = False
objExcel.DisplayAlerts = False

' ── File 1: righe_Ordini_ARCA.xlsx (motore principale) ──────────────────────
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files\righe_Ordini_ARCA.xlsx")
objWorkbook.RefreshAll
WScript.Sleep 10000
objWorkbook.Save
objWorkbook.Close

' ── File 2: righe_ordini_storico_con_date.xlsx (BTL, Atoplast, Cronistoria) ──
Set objWorkbook2 = objExcel.Workbooks.Open("C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files\righe_ordini_storico_con_date.xlsx")
objWorkbook2.RefreshAll
WScript.Sleep 15000
objWorkbook2.Save
objWorkbook2.Close

objExcel.Quit