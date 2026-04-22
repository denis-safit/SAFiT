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

' ── File 4: dettagli_consegne.xlsx ──────────────────────────────────────────
Set objWorkbook4 = objExcel.Workbooks.Open("C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files\dettagli_consegne.xlsx")
objWorkbook4.RefreshAll
WScript.Sleep 12000
objWorkbook4.Save
objWorkbook4.Close

' ── File 5: btl_istruzioni.xlsx ─────────────────────────────────────────────
Set objWorkbook5 = objExcel.Workbooks.Open("C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files\btl_istruzioni.xlsx")
objWorkbook5.RefreshAll
WScript.Sleep 8000
objWorkbook5.Save
objWorkbook5.Close

objExcel.Quit