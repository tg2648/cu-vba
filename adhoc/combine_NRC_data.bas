Sub Combine_NRC_Data()
'
' Takes all NRC data, sorts by program, sorts by rank, then copies the first five institutions to a different sheet
'

'

    Application.ScreenUpdating = False
    
    Dim programs As Range
    Dim program As Variant
    Dim copy_range As Range
    Dim emptyRow As Long
    
    Set data_sheet = Worksheets("Data All")
    Set target_sheet = Worksheets("Data CU and Top 10 by Field")
    Set temp_sheet = Worksheets("Temp")
    Set programs = Worksheets("Unique Fields").Range("A2:A42") ' List of programs/fields to go through
    
    For Each program In programs
    
        ' Filter by program
        data_sheet.UsedRange.AutoFilter Field:=3, _
                                        Criteria1:=program
        
        ' Sort by a column
        data_sheet.AutoFilter.Sort.SortFields.Add Key:=Range("J:J"), _
                                                  SortOn:=xlSortOnValues, _
                                                  Order:=xlAscending, _
                                                  DataOption:=xlSortNormal
        
        data_sheet.AutoFilter.Sort.SortFields.Clear
        data_sheet.AutoFilter.Sort.SortFields.Add Key:=Range("J1"), SortOn:=xlSortOnValues, _
                                                  Order:=xlAscending, DataOption:=xlSortNormal
        With data_sheet.AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        'Set copy_range = data_sheet.UsedRange.Offset(1, 0) ' Take out the header column
        'Set copy_range = copy_range.SpecialCells(xlCellTypeVisible)
        'Set copy_range = copy_range.Resize(5) ' First 5 rows
        
        ' Copy filtered data to a temp worksheet
        Set copy_range = data_sheet.AutoFilter.Range.Offset(1, 0)
        copy_range.Copy Destination:=temp_sheet.Range("A1")
        
        ' Copy the first five rows from the temp worksheet, then clear the temp worksheet
        emptyRow = target_sheet.Range("A" & target_sheet.Rows.Count).End(xlUp).Row + 1
        temp_sheet.Rows("1:10").Copy Destination:=target_sheet.Range("A" & emptyRow)
        temp_sheet.Cells.Delete Shift:=xlUp
    
    Next program
    
    Application.CutCopyMode = False

End Sub