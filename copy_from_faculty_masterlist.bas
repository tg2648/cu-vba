Sub copy_from_faculty_masterlist()
'
' copy_from_faculty_masterlist Macro
'

' 1. Open faculty master list
' 2. Delete sheets that should not be transferred
' 3. Run macro
'
' Destination sheet should be "Faculty Sheet"


    Application.ScreenUpdating = False
    
    Dim emptyRow As Long
    Dim deptName As String
    Dim lastRow As Long
    Dim tempData As Range
    Dim sheet As Worksheet
    
    Dim FTE_workbook_name As String
    Dim FTE_workbook As Workbook
    
    ' Change name as needed
    
    FTE_workbook_name = "FY2019_FTE.xlsx"
    Set FTE_workbook = Workbooks(FTE_workbook_name)
    
    Dim masterlist_workbook_name As String
    Dim masterlist_workbook As Workbook
    
    ' Change name as needed
    
    masterlist_workbook_name = ActiveWorkbook.Name
    Set masterlist_workbook = Workbooks(masterlist_workbook_name)
    
    ' Loop through each department in the masterlist (irrelevant sheets need to be deleted)
    
    For Each sheet In masterlist_workbook.Worksheets
    
        deptName = sheet.Name
        
        Select Case deptName
        Case Is = "GERM"
            deptName = "GERL"
        Case Is = "MELC"
            deptName = "MESA"
        Case Is = "SPPO"
            deptName = "LAIC"
        Case Is = "CE"
            deptName = "SPS"
        End Select
        
        ' Copy entire sheet
        sheet.Cells.Copy
        
        ' Temporarily paste entire thing into a new sheet
        FTE_workbook.Sheets("Sheet1").Range("A1").PasteSpecial
    
        ' Loop through from the bottom. If first column is empty, delete entire row
        ' This will consolidate all types of faculty into one big table
        For i = FTE_workbook.Sheets("Sheet1").Range("C" & Rows.Count).End(xlUp).Row To 1 Step -1
            
            If IsEmpty(FTE_workbook.Sheets("Sheet1").Range("A" & i).Value) = True Then
                FTE_workbook.Sheets("Sheet1").Rows(i).Delete
            End If
                       
        Next
        
        ' Delete salary and empty columns
        FTE_workbook.Sheets("Sheet1").Range("E:E,F:F,L:L").EntireColumn.Delete
            
        ' Copy everything but the column headings and the sheet heading
        Set tempData = FTE_workbook.Sheets("Sheet1").Range("A1").CurrentRegion
        tempData.Offset(2, 0).Resize(tempData.Rows.Count - 1, tempData.Columns.Count).Copy
        
        ' Paste to the last empty row of the main worksheet
        FTE_workbook.Sheets("Faculty List").Activate
        emptyRow = FTE_workbook.Sheets("Faculty List").Range("E" & Rows.Count).End(xlUp).Row + 1
        FTE_workbook.Sheets("Faculty List").Range("E" & emptyRow).PasteSpecial
        
        ' Insert department name and fill down
        FTE_workbook.Sheets("Faculty List").Range("D" & emptyRow).Value = deptName
        lastRow = FTE_workbook.Sheets("Faculty List").Range("E" & Rows.Count).End(xlUp).Row
        FTE_workbook.Sheets("Faculty List").Range("D" & emptyRow).Select
        
        If emptyRow <> lastRow Then
            Selection.AutoFill Destination:=FTE_workbook.Sheets("Faculty List").Range("D" & emptyRow & ":D" & lastRow)
        End If
     
        ' Clear the temp worksheet
        FTE_workbook.Sheets("Sheet1").Cells.Delete Shift:=xlUp
    
    Next sheet
    
    ' Move the UNI column for all
    Range("P" & 2 & ":P" & lastRow).Cut Range("S2")
    
    ' Clean up data
    For j = lastRow To 2 Step -1
         
        ' Delete rows with blank names
        If IsEmpty(FTE_workbook.Sheets("Faculty List").Range("G" & j).Value) Then
            FTE_workbook.Sheets("Faculty List").Rows(j).Delete
        End If
        
        ' Shift non-renewable terms to a separate column
        ' Otherwise would end up in joint_interdisc column
        ' Also move lecturer language in a separate column
        Select Case FTE_workbook.Sheets("Faculty List").Range("E" & j).Value
            Case Is = "Other Full-Time: Term"
                Range("J" & j).Cut Range("P" & j)
                Range("K" & j).Cut Range("R" & j)
            Case Is = "Other Full-Time:Term"
                Range("J" & j).Cut Range("P" & j)
                Range("K" & j).Cut Range("R" & j)
            Case Is = "Other Full-Time: Term "
                Range("J" & j).Cut Range("P" & j)
                Range("K" & j).Cut Range("R" & j)
            Case Is = "Other Full-Time:Term "
                Range("J" & j).Cut Range("P" & j)
                Range("K" & j).Cut Range("R" & j)
            Case Is = "Other Full-Time: FOS"
                Range("J" & j).Cut Range("P" & j)
            Case Is = "Other Full-Time:FOS"
                Range("J" & j).Cut Range("P" & j)
            Case Is = "Professorial: Term"
                Range("J" & j).Cut Range("P" & j)
            Case Is = "Professorial:Term"
                Range("J" & j).Cut Range("P" & j)
            Case Is = "Professorial: Term "
                Range("J" & j).Cut Range("P" & j)
            Case Is = "Professorial:Term "
                Range("J" & j).Cut Range("P" & j)
            Case Is = "Other Full-Time"
                Range("K" & j).Cut Range("R" & j)
        End Select
    
    Next j
    
    
End Sub