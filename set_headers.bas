Sub set_headers()

    Dim datasource As String
    
    datasource = InputBox("Enter data source:")

    With ActiveSheet.PageSetup
        .LeftHeader = "&K00-049Arts && Sciences Planning and Analysis"
        .CenterHeader = ""
        .RightHeader = "&K00-049&P/&N"
        .LeftFooter = "&10&K00-014Source: " & datasource & Chr(10) & "&F" & Chr(10) & "&D"
        .CenterFooter = "&""-,Bold""&12&K00-049DRAFT"
        .RightFooter = "&K00-049For any questions, please contact" & Chr(10) & "Rose Razaghian (rr222) or" & Chr(10) & "Timur Gulyamov (tg2648)"
    End With
    
End Sub
