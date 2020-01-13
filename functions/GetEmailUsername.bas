Public Function GetEmailUsername(CellRef As Range)
' Extracts the username of an email address

Dim SymbolPos As Long
SymbolPos = InStr(CellRef.Value, "@")

GetEmailUsername = Left(CellRef.Value, SymbolPos - 1)

End Function
