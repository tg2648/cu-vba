Sub paste_then_delete_every_other_row()
'
' delete_every_other_row_in_selection Macro
'
'
Dim selectedRange As Range

    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True

    Set selectedRange = Selection
    
    Application.ScreenUpdating = False
    For i = selectedRange.Rows.Count To 1 Step -2
        selectedRange.Cells(i, 1).EntireRow.Delete
    Next
    Application.ScreenUpdating = True
    
End Sub