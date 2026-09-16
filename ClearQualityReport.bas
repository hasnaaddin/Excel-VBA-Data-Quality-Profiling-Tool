Attribute VB_Name = "ClearQualityReport"
Sub ClearFromD9RightAndBelow()
    ' Clears values & formulas from D9 to the last cell on the active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    With ws
        .Range(.Cells(9, "D"), .Cells(.Rows.Count, .Columns.Count)).ClearContents
    End With
End Sub

