Sub TestConsolidateGRIR()
    
    Call HideSheet("Sheet2", True)

End Sub

Sub HideSheet(in_Sheet As String, in_Action As Boolean)

Sheets(in_Sheet).Visible = in_Action

End Sub
