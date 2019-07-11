Sub copy_current_sheet_prompt_for_name()

    Dim newName As Variant
    
    ActiveSheet.Copy After:=Sheets(Sheets.Count)
    
    myValue = InputBox("New sheet name:")
    
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = myValue

End Sub