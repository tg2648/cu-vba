Sub dept_profile_copy_formulas()

    ' Group worksheets
    ' Make a selection
    ' Macro will paste the selection into each grouped worksheet
    
    Application.ScreenUpdating = False

    Dim selectedRange As Range
    Dim ws As Worksheet
    Dim rangeRow As Integer
    Dim rangeCol As Integer
    
    Set selectedRange = Selection
    selectedRange.Copy
    
    rangeRow = selectedRange.Row
    rangeCol = selectedRange.Column

    For Each ws In ActiveWindow.SelectedSheets

         ws.Cells(rangeRow, rangeCol).PasteSpecial xlPasteAll

    Next ws

End Sub