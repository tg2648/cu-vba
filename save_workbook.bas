Sub Save_copy_workbook()

Dim filepath As String
Dim filepath_archive As String
Dim new_filename As String
Dim filecount As Integer
Dim old_filename As String

filecount = 1

old_filename = Replace(ActiveWorkbook.Name, ".xlsx", "")
new_filename = old_filename & ".Old.v" & filecount & ".xlsx"

filepath = ActiveWorkbook.Path & "\Archive\"

'Loop through the counts and figure out which version number is the last
Do While ((Dir(filepath & new_filename)) <> Empty)
    filecount = filecount + 1
    new_filename = old_filename & ".Old.v" & filecount & ".xlsx"
Loop

'Check if archive/file_name folder exists, if not, create
If Dir(filepath, vbDirectory) = "" Then
        MkDir filepath
End If

If Len(filepath) + Len(new_filename) > 218 Then
    
    MsgBox filepath & new_filename & vbNewLine & vbNewLine & "The file path cannon be longer than 218 characters due to Excel limitations.", vbCritical
    End

End If

ActiveWorkbook.Save
ActiveWorkbook.SaveCopyAs Filename:=filepath & new_filename
    
MsgBox "New copy saved in Archive\" & new_filename

End Sub
