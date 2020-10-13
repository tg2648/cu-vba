Sub split_by_unique()
'
' Splits current table by unique values in the selected column
' Saves each split into a file with the filename pre-pended with the value
'

Application.ScreenUpdating = False

' https://stackoverflow.com/questions/5890257/populate-unique-values-into-a-vba-array-from-excel
Dim d As Object
Dim c As Range
Dim unique, tmp As String

Set d = CreateObject("Scripting.Dictionary")
For Each c In Selection
    tmp = Trim(c.Value)
    If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
Next c


Dim current_workbook, new_workbook As Workbook
Dim current_workbook_name, new_workbook_name, filepath, table_name As String
Dim column As Integer
Dim data_range As Range

' Get details of the current workbook
current_workbook_name = ActiveWorkbook.Name
Set current_workbook = ActiveWorkbook
filepath = current_workbook.Path & "\"

table_name = ActiveCell.ListObject.Name
column = Selection.column


For Each unique In d.keys
    
    ' Create a new workbook
    ' Append unique value to the filename
    new_workbook_name = filepath & unique & "_" & current_workbook_name
    Set new_workbook = Workbooks.Add
    Application.DisplayAlerts = False
    new_workbook.SaveAs Filename:=new_workbook_name
    Application.DisplayAlerts = True
    
    
    ' Select the range
    Set data_range = current_workbook.ActiveSheet.ListObjects(table_name).Range
    ' Filter by department
    data_range.AutoFilter Field:=column, Criteria1:=unique
    ' Add the sheet (same name as table name) to the new workbook
    new_workbook.Sheets.Add(After:=Sheets(Sheets.Count)).Name = table_name
    
    ' Copy filtered data
    ' data_range.SpecialCells(xlCellTypeVisible).Copy Destination:=new_workbook.Sheets(table_name).Cells(1, 1)
    
    ' Copy filtered data preserving column widths (kind of)
    data_range.SpecialCells(xlCellTypeVisible).Copy
    With new_workbook.Sheets(table_name).Cells(1, 1)
        .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
    
    ' Autofit columns in new workbook
    new_workbook.Sheets(table_name).Cells.Select
    new_workbook.Sheets(table_name).Cells.EntireColumn.AutoFit
    
    ' Remove the default sheet from the new workbook, save, and close
    Application.DisplayAlerts = False
    new_workbook.Sheets("Sheet1").Delete
    new_workbook.Save
    new_workbook.Close
    Application.DisplayAlerts = True
    
Next unique

MsgBox ("Done!")

End Sub
