Private Sub Workbook_Open()
Application.Visible = True
If Date < DateValue("14/03/2021") Then

Else
Application.DisplayAlerts = False
    MsgBox "Este archivo se daño"
ThisWorkbook.Close
End If

End Sub
