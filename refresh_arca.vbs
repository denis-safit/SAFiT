Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
objExcel.AskToUpdateLinks = False

' ── File 1: righe_Ordini_ARCA.xlsx ──────────────────────────────────────────
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files\righe_Ordini_ARCA.xlsx", False, False)
objWorkbook.RefreshAll
WScript.Sleep 10000
objWorkbook.Save
objWorkbook.Close False

' Chiude Avanzamento se aperto come dipendenza
On Error Resume Next
objExcel.Workbooks("Avanzamento_access.xlsx").Close False
On Error GoTo 0

' ── File 2: righe_ordini_storico_con_date.xlsx ──────────────────────────────
Set objWorkbook2 = objExcel.Workbooks.Open("C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files\righe_ordini_storico_con_date.xlsx", False, False)
objWorkbook2.RefreshAll
WScript.Sleep 15000
objWorkbook2.Save
objWorkbook2.Close False

' Chiude Avanzamento se aperto come dipendenza
On Error Resume Next
objExcel.Workbooks("Avanzamento_access.xlsx").Close False
On Error GoTo 0

objExcel.Quit
Set objExcel = Nothing