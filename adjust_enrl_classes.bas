Sub Adj_Enrl_Classes()

    ' This macro adjusts enrollments and classes based on the "adjustment code"
    ' This only works if merged_id only repeats twice
    ' Adjustment code is entered in the first instance of merged_id, the second is left blank
    
    Application.ScreenUpdating = False
    
    Dim Adj_Codes As Range
    Dim Enrl_Adj As Range
    Dim Classes_Adj As Range
    Dim Enr As Range
    Dim current_row As Integer
    Dim a As Integer
    
    Set Adj_Codes = Range("Classroom_Data[Adj_Code]")
    Set Enrl_Adj = Range("Classroom_Data[Enrl_Adj]")
    Set Classes_Adj = Range("Classroom_Data[Classes_Adj]")
    Set Enr = Range("Classroom_Data[Enr]")
    
    Enrl_Adj.ClearContents
    Classes_Adj.ClearContents
    
    For Each Adj_Code In Adj_Codes
    
        If Not IsEmpty(Adj_Code) Then
            
            current_row = Adj_Code.Row - 1 ' Subtract 1 to discount the header row
                      
            Adj_Code.Offset(1, 0).Select ' Select the row after the row with adjustment code
            Do Until ActiveCell.EntireRow.Hidden = False ' Loop until a non-hidden row is selected
                ActiveCell.Offset(1, 0).Select
            Loop
            next_row = ActiveCell.Row - 1 ' This is the next row in a filtered view
                        
            Select Case Adj_Code
            
                Case Is = 1 ' Combined enrollments and classes to lower numbered course (same row as where the adj_code sits if data is also ordered by course number). The other gets zeros.
                    
                    Enrl_Adj(current_row) = Enr(current_row) + Enr(next_row)
                    Classes_Adj(current_row) = 1
                    
                    Enrl_Adj(next_row) = 0
                    Classes_Adj(next_row) = 0
                    
                Case Is = 2 ' Own enrollments and 0.5 classes in each
                    
                    Enrl_Adj(current_row) = Enr(current_row)
                    Classes_Adj(current_row) = 0.5
                    
                    Enrl_Adj(next_row) = Enr(next_row)
                    Classes_Adj(next_row) = 0.5
                
                Case Is = 3 ' Own enrollments and classes in each
                    
                    Enrl_Adj(current_row) = Enr(current_row)
                    Classes_Adj(current_row) = 1
                    
                    Enrl_Adj(next_row) = Enr(next_row)
                    Classes_Adj(next_row) = 1
            
            End Select
            
        End If
        
    Next Adj_Code

    Application.ScreenUpdating = True
    
End Sub