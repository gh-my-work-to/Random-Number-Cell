Attribute VB_Name = "mod_randomNumberCell"

Sub randomNumberCell()
    If Selection.Count = 0 Then Exit Sub
    
    Dim max, min, width
    max = InputBox("max?", , 100)
    min = InputBox("min?", , 0)
    
    width = max - min + 1
    
    If Selection.Count = 1 Then
        Call randomNumberCell_cols(min, width)
    Else
        Call randomNumberCell_selected(min, width)
    End If
End Sub

Sub randomNumberCell_cols(min, width)
    Dim cosuu
    cosuu = InputBox("How many cells to be randomed?", , 10)
    
    Dim i, y, x
    x = ActiveCell.Column
    y = ActiveCell.Row
    
    For i = 1 To cosuu
        Cells(y, x).Value = Int(Rnd() * width) + min
        y = y + 1
    Next i
End Sub

Sub randomNumberCell_selected(min, width)
    Dim cl
    For Each cl In Selection
        cl.Value = Int(Rnd() * width) + min
    Next cl
End Sub

