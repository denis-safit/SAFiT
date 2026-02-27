Set objExcel = CreateObject("Excel.Application")
' Nasconde Excel mentre lavora
objExcel.Visible = False
objExcel.DisplayAlerts = False

' Cambia il percorso con quello reale del tuo file Arca
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files\righe_Ordini_ARCA.xlsx")

' Forza l'aggiornamento di tutte le query/pivot
objWorkbook.RefreshAll

' Aspetta 10 secondi per dare tempo al database di rispondere
WScript.Sleep 10000

' Salva e chiude
objWorkbook.Save
objWorkbook.Close
objExcel.Quit