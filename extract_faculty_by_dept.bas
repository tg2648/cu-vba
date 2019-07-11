Sub Extract_faculty_by_dept()

Application.ScreenUpdating = False

Dim data_sheets As Variant
Dim data_sheet As Variant

Dim current_workbook As Workbook
Dim current_workbook_name As String

Dim dept_list As Range

Dim new_workbook As Workbook
Dim new_workbook_name As String
Dim data_range As Range

Dim filepath As String

current_workbook_name = "Department_Profiles_Faculty_List.xlsx"
Set current_workbook = Workbooks(current_workbook_name)

filepath = current_workbook.Path

' Name of sheets and tables (should be identical) where the data will be copied from
data_sheets = Array("Faculty_List")

' Dept_List has unique values of Departments
With current_workbook.Sheets("Dept_List")
    Set dept_list = .Range(.Cells(2, 1), .Cells(.Rows.Count, 1).End(xlUp))
End With

For Each Department In dept_list

    ' Create a new workbook with the same name as the department
    new_workbook_name = filepath & "\Department_Profiles_Faculty_" & Department & ".xlsx"
    Set new_workbook = Workbooks.Add
    Application.DisplayAlerts = False
    new_workbook.SaveAs Filename:=new_workbook_name
    Application.DisplayAlerts = True
    
    For Each data_sheet In data_sheets
    
        ' Select the range
        Set data_range = current_workbook.Sheets(data_sheet).ListObjects(data_sheet).Range
        ' Filter by department in the third column
        data_range.AutoFilter Field:=3, Criteria1:=Department
    
        ' Add the sheet to the new workbook
        new_workbook.Sheets.Add(After:=Sheets(Sheets.Count)).Name = data_sheet
        
        ' Copy the filtered data
        data_range.SpecialCells(xlCellTypeVisible).Copy Destination:=new_workbook.Sheets(data_sheet).Cells(1, 1)
        
    Next data_sheet
    
    ' Remove the temp sheet and the default sheet from the new workbook, save, and close
    Application.DisplayAlerts = False
    new_workbook.Sheets("Sheet1").Delete
    new_workbook.Save
    new_workbook.Close
    Application.DisplayAlerts = True

Next Department

' Remove all filters from the original file
For Each data_sheet In data_sheets

    current_workbook.Sheets(data_sheet).ListObjects(data_sheet).Range.AutoFilter Field:=3

Next data_sheet

Application.StatusBar = False

MsgBox "Done!"
 
End Sub