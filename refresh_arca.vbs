Dim objExcel, objWB
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Dim BASE
BASE = "C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files\"

' ── Helper: apri, aggiorna, salva, chiudi ───────────────────────────────────
Sub RefreshFile(filename, waitSec)
    Dim path
    path = BASE & filename
    On Error Resume Next
    Set objWB = objExcel.Workbooks.Open(path)
    If Err.Number <> 0 Then
        WScript.Echo "ERRORE apertura: " & filename
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    objWB.RefreshAll
    WScript.Sleep waitSec * 1000
    objWB.Save
    objWB.Close False
    WScript.Echo "OK: " & filename
End Sub

' ── Aggiornamento file in ordine (più leggeri prima) ────────────────────────
WScript.Echo "=== Avvio refresh ARCA ==="

Call RefreshFile("btl_istruzioni.xlsx",                  8)
Call RefreshFile("dettagli_consegne_CLI.xlsx",           12)
Call RefreshFile("dettagli_consegne.xlsx",               12)
Call RefreshFile("righe_Ordini_ARCA.xlsx",               12)
Call RefreshFile("Avanzamento_access.xlsx",              10)
Call RefreshFile("righe_ordini_storico_con_date.xlsx",   18)

WScript.Echo "=== Tutti i file aggiornati ==="
objExcel.Quit