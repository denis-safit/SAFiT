Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

' ── File 1: righe_Ordini_ARCA.xlsx ──────────────────────────────────────────
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files\righe_Ordini_ARCA.xlsx")
objWorkbook.RefreshAll
WScript.Sleep 10000
objWorkbook.Save
objWorkbook.Close

' ── File 2: righe_ordini_storico_con_date.xlsx ──────────────────────────────
Set objWorkbook2 = objExcel.Workbooks.Open("C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files\righe_ordini_storico_con_date.xlsx")
objWorkbook2.RefreshAll
WScript.Sleep 15000
objWorkbook2.Save
objWorkbook2.Close

' ── File 3: dettagli_consegne_CLI.xlsx ──────────────────────────────────────
Set objWorkbook3 = objExcel.Workbooks.Open("C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files\dettagli_consegne_CLI.xlsx")
objWorkbook3.RefreshAll
WScript.Sleep 12000
objWorkbook3.Save
objWorkbook3.Close

' ── Chiude tutti i workbook rimasti aperti come dipendenze ──────────────────
Dim wb
For Each wb In objExcel.Workbooks
    wb.Close False
Next

objExcel.Quit

Set objExcel = Nothing