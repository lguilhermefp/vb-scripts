Private Sub Workbook_BeforeClose(Cancel As Boolean)

    if Me.Saved = False Then Me.Save

End Sub