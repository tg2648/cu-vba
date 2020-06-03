Sub split_by_dept()
'
' split_by_Dept Macro
'
' Select a table that has a column named "Dept"
' The table will be split into 27 individual based on the 27 department codes

Application.ScreenUpdating = False

Dim dept_list As Variant

' dept_list = Array("SPS", "ARTS", "AHAR", "CLAS", "EALC", "ENCL", "FRRP", "GERL", "ITAL", "LAIC", _
' "MESA", "MUSI", "PHIL", "RELI", "SLAL", "ASTR", "BIOL", "CHEM", "DEES", _
' "EEEB", "MATH", "PHYS", "PSYC", "STAT", "ANTH", "ECON", "HIST", "POLS", "SOCI")
' dept_list = Array("MUSI", "AHAR", "FRRP", "GERL", "SLAL", "UWRT", "ENCL", "FILM", "CLAS", "ITAL", "MATH", "SPPO", "HIST", "WRIT", "MELC", "PSYC", "PHYS")

dept_list = Array("AHAR", "CLAS", "EALC", "ENCL", "FRRP", "GERL", "ITAL", "SPPO", "MELC", "MUSI", "PHIL", "RELI", "SLAL", "ASTR", "BIOS", "CHEM", "EESC", "EEEB", "MATH", "PHYS", "PSYC", "STAT", "ANTH", "ECON", "HIST", "POLS", "SOCI", "AFAM")

Dim current_workbook As Workbook
Dim current_workbook_name As String
Dim filepath As String
Dim new_workbook As Workbook
Dim new_workbook_name As String
Dim table_name As String
Dim table_dept_column As Integer
Dim data_range As Range

' Get details of the current workbook
current_workbook_name = ActiveWorkbook.Name
Set current_workbook = ActiveWorkbook
filepath = current_workbook.Path & "\"

table_name = ActiveCell.ListObject.Name
table_dept_column = Range(table_name & "[[#Headers],[Department Code]]").Column

For Each dept In dept_list

    ' Create a new workbook with the same name as the department
    new_workbook_name = filepath & dept & "_" & current_workbook_name
    Set new_workbook = Workbooks.Add
    Application.DisplayAlerts = False
    new_workbook.SaveAs Filename:=new_workbook_name
    Application.DisplayAlerts = True
    
    
    ' Select the range
    Set data_range = current_workbook.ActiveSheet.ListObjects(table_name).Range
    ' Filter by department
    data_range.AutoFilter Field:=table_dept_column, Criteria1:=dept
    ' Add the sheet (same name as table name) to the new workbook
    new_workbook.Sheets.Add(After:=Sheets(Sheets.Count)).Name = table_name
    ' Copy the filtered data
    data_range.SpecialCells(xlCellTypeVisible).Copy Destination:=new_workbook.Sheets(table_name).Cells(1, 1)

    ' Autofit columns in new workbook
    new_workbook.Sheets(table_name).Cells.Select
    new_workbook.Sheets(table_name).Cells.EntireColumn.AutoFit
    
    ' Remove the default sheet from the new workbook, save, and close
    Application.DisplayAlerts = False
    new_workbook.Sheets("Sheet1").Delete
    new_workbook.Save
    new_workbook.Close
    Application.DisplayAlerts = True

Next dept

MsgBox ("Done!")

End Sub
