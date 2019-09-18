Sub copy_from_faculty_masterlist()
'
' copy_from_faculty_masterlist Macro
'

' 1. Open faculty master list
' 2. Delete sheets that should not be transferred
' 3. Run macro
'
' Destination sheet should be "Faculty List"


    Application.ScreenUpdating = False
    
    Dim emptyRow As Long
    Dim deptName As String
    Dim divName As String
    Dim lastRow As Long
    Dim tempData As Range
    Dim sheet As Worksheet
    
    Dim FTE_workbook_name As String
    Dim FTE_workbook As Workbook
    
    ' Adjust name as needed
    FTE_workbook_name = "FY2020_FTE.xlsx"
    Set FTE_workbook = Workbooks(FTE_workbook_name)
    

    Dim masterlist_workbook As Workbook
    Dim masterlist_workbook_names As Variant
    Dim masterlist_workbook_name As Variant
    
    ' Adjust names as needed
    ' Workbooks need to be opened beforehand
    masterlist_workbook_names = Array( _
        "2019-08-21_19-20 ARTS_ALP_SCE_Faculty_Masterlist_WORKING.xls", _
        "2019-08-21_19-20 HUMANITIES_Faculty_Masterlist_WORKING.xls", _
        "2019-09-11_19-20 NATURAL_SCIENCES_Faculty_Masterlist_WORKING.xls", _
        "2019-09-12_19-20 SOCIAL_SCIENCES_Faculty_Masterlist_WORKING.xls" _
    )

    ' Loop through each masterlist
    For Each masterlist_workbook_name In masterlist_workbook_names
    
        Set masterlist_workbook = Workbooks(masterlist_workbook_name)
        
        ' Loop through each sheet (department) in the masterlist (irrelevant sheets need to be deleted)
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
                           
            Next i
            
            ' Delete salary and empty columns
            FTE_workbook.Sheets("Sheet1").Range("E:E,F:F,L:L,P:P").EntireColumn.Delete
                
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

            ' Insert division name and fill down
            Select Case deptName
                Case Is = "AHAR"
                    divName = "HUM"
                Case Is = "CLAS"
                    divName = "HUM"
                Case Is = "EALC"
                    divName = "HUM"
                Case Is = "ENCL"
                    divName = "HUM"
                Case Is = "FRRP"
                    divName = "HUM"
                Case Is = "GERL"
                    divName = "HUM"
                Case Is = "ITAL"
                    divName = "HUM"
                Case Is = "LAIC"
                    divName = "HUM"
                Case Is = "MESA"
                    divName = "HUM"
                Case Is = "MUSI"
                    divName = "HUM"
                Case Is = "PHIL"
                    divName = "HUM"
                Case Is = "RELI"
                    divName = "HUM"
                Case Is = "SLAL"
                    divName = "HUM"
                Case Is = "ASTR"
                    divName = "NS"
                Case Is = "BIOL"
                    divName = "NS"
                Case Is = "CHEM"
                    divName = "NS"
                Case Is = "DEES"
                    divName = "NS"
                Case Is = "EEEB"
                    divName = "NS"
                Case Is = "MATH"
                    divName = "NS"
                Case Is = "PHYS"
                    divName = "NS"
                Case Is = "PSYC"
                    divName = "NS"
                Case Is = "STAT"
                    divName = "NS"
                Case Is = "ANTH"
                    divName = "SS"
                Case Is = "ECON"
                    divName = "SS"
                Case Is = "HIST"
                    divName = "SS"
                Case Is = "POLS"
                    divName = "SS"
                Case Is = "SOCI"
                    divName = "SS"
                Case Is = "FILM"
                    divName = "ARTS"
                Case Is = "THEA"
                    divName = "ARTS"
                Case Is = "VIAR"
                    divName = "ARTS"
                Case Is = "WRIT"
                    divName = "ARTS"
                Case Is = "ALP"
                    divName = "SPS"
                Case Is = "SPS"
                    divName = "SPS"
            End Select
            
            FTE_workbook.Sheets("Faculty List").Range("C" & emptyRow).Value = divName
            lastRow = FTE_workbook.Sheets("Faculty List").Range("E" & Rows.Count).End(xlUp).Row
            FTE_workbook.Sheets("Faculty List").Range("C" & emptyRow).Select
            
            If emptyRow <> lastRow Then
                Selection.AutoFill Destination:=FTE_workbook.Sheets("Faculty List").Range("C" & emptyRow & ":C" & lastRow)
            End If

            ' Clear the temp worksheet
            FTE_workbook.Sheets("Sheet1").Cells.Delete Shift:=xlUp
        
        Next sheet
    
    Next masterlist_workbook_name
    
    
    ' Perform cleaning on the FTE sheet after all masterlists had been copied
    FTE_workbook.Activate
    
    ' Move the UNI column for all
    Range("P" & 2 & ":P" & lastRow).Cut Range("T2")
    ' Move the Research Funds column for all
    Range("Q" & 2 & ":Q" & lastRow).Cut Range("S2")
    
    For j = lastRow To 2 Step -1
         
        ' Delete rows with blank names
        If IsEmpty(FTE_workbook.Sheets("Faculty List").Range("G" & j).Value) Then
            FTE_workbook.Sheets("Faculty List").Rows(j).Delete
        End If


        ' Clean up tenure status
        Select Case FTE_workbook.Sheets("Faculty List").Range("E" & j).Value
            Case Is = "Other Full-Time:Term"
                Range("E" & j).Value = "Other Full-Time: Term"
            Case Is = "Other Full-Time:Term "
                Range("E" & j).Value = "Other Full-Time: Term"
            Case Is = "Other Full-Time: Term "
                Range("E" & j).Value = "Other Full-Time: Term"
            Case Is = "Other Full-Time:  Term"
                Range("E" & j).Value = "Other Full-Time: Term"
            Case Is = "Other Full-Time:  Term "
                Range("E" & j).Value = "Other Full-Time: Term"
                
            Case Is = "Other Full-Time:FOS"
                Range("E" & j).Value = "Other Full-Time: FOS"
            Case Is = "Other Full-Time:FOS "
                Range("E" & j).Value = "Other Full-Time: FOS"
            Case Is = "Other Full-Time: FOS "
                Range("E" & j).Value = "Other Full-Time: FOS"
            Case Is = "Other Full-Time:  FOS"
                Range("E" & j).Value = "Other Full-Time: FOS"
            Case Is = "Other Full-Time:  FOS "
                Range("E" & j).Value = "Other Full-Time: FOS"

            Case Is = "Other Full-Time:Visitors"
                Range("E" & j).Value = "Other Full-Time: Visitors"
            Case Is = "Other Full-Time:Visitors "
                Range("E" & j).Value = "Other Full-Time: Visitors"
            Case Is = "Other Full-Time: Visitors "
                Range("E" & j).Value = "Other Full-Time: Visitors"
            Case Is = "Other Full-Time:  Visitors"
                Range("E" & j).Value = "Other Full-Time: Visitors"
            Case Is = "Other Full-Time:  Visitors "
                Range("E" & j).Value = "Other Full-Time: Visitors"

            Case Is = "Professorial:Term"
                Range("E" & j).Value = "Professorial: Term"
            Case Is = "Professorial:Term "
                Range("E" & j).Value = "Professorial: Term"
            Case Is = "Professorial: Term "
                Range("E" & j).Value = "Professorial: Term"
            Case Is = "Professorial:  Term"
                Range("E" & j).Value = "Professorial: Term"
            Case Is = "Professorial:  Term "
                Range("E" & j).Value = "Professorial: Term"

            Case Is = "Non-Ten & Ten-Track"
                Range("E" & j).Value = "Non-Ten/Ten-Track"
                
        End Select
        
        
        ' Clean up rank
        Select Case FTE_workbook.Sheets("Faculty List").Range("F" & j).Value
            Case Is = "Prof of Prof Pract"
                Range("F" & j).Value = "Professor of Professional Practice"
            Case Is = "Assoc Prof of Prof Pract"
                Range("F" & j).Value = "Associate Professor of Professional Practice"
            Case Is = "Asst Prof of Prof Pract"
                Range("F" & j).Value = "Assistant Professor of Professional Practice"
            Case Is = "Assistant Prof of Prof Pract"
                Range("F" & j).Value = "Assistant Professor of Professional Practice"
            Case Is = "Professor of Prof Practice"
                Range("F" & j).Value = "Professor of Professional Practice"
            Case Is = "Associate Prof of Prof Practice"
                Range("F" & j).Value = "Associate Professor of Professional Practice"
            Case Is = "Sr.Lecturer in Discipline"
                Range("F" & j).Value = "Senior Lecturer in Discipline"
            Case Is = "Burke Post doc Teaching Fellow"
                Range("F" & j).Value = "Burke Post Doc Teaching Fellow"
            Case Is = "Senior Lect in Discipline"
                Range("F" & j).Value = "Senior Lecturer in Discipline"
            Case Is = "Society of Fellow"
                Range("F" & j).Value = "Society of Fellows"
            Case Is = "Assoc in Music Perf"
                Range("F" & j).Value = "Associate in Music Performance"
            Case Is = "Sr. Lecturer in Discipline"
                Range("F" & j).Value = "Senior Lecturer in Discipline"
            Case Is = "Mellon/ Heyman Ctr"
                Range("F" & j).Value = "Mellon/Heyman Ctr"
            Case Is = "Assoc Res Scholar"
                Range("F" & j).Value = "Associate Research Scholar"
        End Select
        
        
        ' Shift non-renewable terms to a separate column
        ' Otherwise would end up in joint_interdisc column
        ' Also move lecturer language in a separate column
        Select Case FTE_workbook.Sheets("Faculty List").Range("E" & j).Value
            Case Is = "Other Full-Time: Term"
                Range("J" & j).Cut Range("P" & j)
                Range("K" & j).Cut Range("R" & j)
            Case Is = "Other Full-Time: FOS"
                Range("J" & j).Cut Range("P" & j)
            Case Is = "Professorial: Term"
                Range("J" & j).Cut Range("P" & j)
            Case Is = "Other Full-Time"
                Range("K" & j).Cut Range("R" & j)
        End Select
    
    Next j
    
End Sub