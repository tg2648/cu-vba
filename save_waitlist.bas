Sub Save_Waitlist()
'
' Save_Waitlist Macro
' Renames first sheet as 'Students', second as 'Classes'. Saves with current date and time.
'
' 1. Download student list from SSOL
' 2. Open file
' 3. Go to class list on SSOL
' 4. Select first few rows staring with table heading - "List of <> Wait Lists"
' 5. Ctrl+Shift+End to select entire table
' 6. Ctrl+C
' 7. Run macro, it will paste automatically

'
    Dim current_name As String
    Dim new_name As String
    Dim filepath As String
    Dim current_date As String
    Dim current_time As String
    
    ' Create new sheet and paste classes information.
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.ClearContents
    
    ' File from the SSOL is in the form "WaitLists_<term>.xls"
    current_name = Replace(ActiveWorkbook.Name, ".xls", "")
    filepath = ActiveWorkbook.Path
    
    current_date = Format(Now, "yyyy-mm-dd")
    current_time = Format(Now, "Medium Time")
    current_time = Replace(current_time, ":", "")
    current_time = Replace(current_time, " ", "")
    
    ' Append current time and date to the name
    new_name = filepath & "\" & current_date & "_" & current_name & "_" & current_time & ".xlsx"
    
    ' Save as default workbook because the SSOL file is just a text file with a .xls extention
    ActiveWorkbook.SaveAs Filename:=new_name, FileFormat:=xlWorkbookDefault
    
    Worksheets(1).Name = "Students" ' First from the left
    Worksheets(2).Name = "Classes"  ' Second from the left
    
    ActiveWorkbook.Save
    
    MsgBox "Saved as " & current_date & "_" & current_name & "_" & current_time & ".xlsx"
    
End Sub
