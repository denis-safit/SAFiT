' Chiude solo Avanzamento_access.xlsx senza toccare altri file Excel aperti
On Error Resume Next
Dim objExcel
Set objExcel = GetObject(, "Excel.Application")
If Not objExcel Is Nothing Then
    Dim wb
    For Each wb In objExcel.Workbooks
        If InStr(LCase(wb.Name), "avanzamento_access") > 0 Then
            wb.Close False
        End If
    Next
End If
On Error GoTo 0