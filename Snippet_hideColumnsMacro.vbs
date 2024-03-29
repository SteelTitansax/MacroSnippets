Sub TestHideColumn()

Call HideColumn(True)

End Sub


Sub HideColumn(in_Hide As Boolean)

Worksheets("Daily open balance").Activate
Columns("H:H").EntireColumn.Hidden = in_Hide

End Sub

