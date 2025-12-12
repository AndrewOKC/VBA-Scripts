Sub FixLinksAndSlash()
    Dim rng As Range
    Dim dataArr As Variant
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    
    ' 1. Optimization: Turn off screen updating and auto calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 2. Define the range (Works on currently selected cells)
    ' Only process the intersection of selection and used range to save time
    On Error Resume Next
    Set rng = Intersect(Selection, Selection.Parent.UsedRange)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Please select a range containing data."
        GoTo Cleanup
    End If
    
    ' 3. Load data into an array for fast processing
    ' If only one cell is selected, we treat it separately to avoid array errors
    If rng.Cells.Count = 1 Then
        rng.Value = Replace(rng.Value, "%253A", "/")
        ActiveSheet.Hyperlinks.Add Anchor:=rng, Address:=rng.Value
    Else
        dataArr = rng.Value
        rowCount = UBound(dataArr, 1)
        colCount = UBound(dataArr, 2)
        
        ' Loop through the array in memory (Fast)
        For i = 1 To rowCount
            For j = 1 To colCount
                If Not IsError(dataArr(i, j)) Then
                    If Len(dataArr(i, j)) > 0 Then
                        ' REPLACE STEP: Swapping %253A with /
                        dataArr(i, j) = Replace(dataArr(i, j), "%253A", "/")
                    End If
                End If
            Next j
        Next i
        
        ' Dump array back to sheet
        rng.Value = dataArr
        
        ' 4. Convert to Hyperlinks
        ' This step must be done on cells, but the text is already clean
        Dim cell As Range
        For Each cell In rng
            If Len(cell.Value) > 0 Then
                ActiveSheet.Hyperlinks.Add Anchor:=cell, Address:=cell.Value
            End If
        Next cell
    End If
    
Cleanup:
    ' 5. Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Done! Links cleaned and activated."

End Sub