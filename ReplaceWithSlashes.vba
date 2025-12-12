Sub CleanURLs_Final()
    Dim rng As Range
    Dim dataArr As Variant
    Dim i As Long, j As Long
    Dim dict As Object
    Dim key As Variant
    
    ' 1. Optimization Settings
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 2. Create Dictionary (Full ASCII List)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' --- CUSTOM OVERRIDES (Top Priority) ---
    dict.Add "%253A", "/"
    dict.Add "+", " "
    
    ' --- CONTROL CHARACTERS ---
    dict.Add "%09", vbTab
    dict.Add "%0A", vbLf
    dict.Add "%0D", vbCr
    
    ' --- SYMBOLS (HEX 20-2F) ---
    dict.Add "%20", " "
    dict.Add "%21", "!"
    dict.Add "%22", """"
    dict.Add "%23", "#"
    dict.Add "%24", "$"
    dict.Add "%25", "%"
    dict.Add "%26", "&"
    dict.Add "%27", "'"
    dict.Add "%28", "("
    dict.Add "%29", ")"
    dict.Add "%2A", "*"
    dict.Add "%2B", "+"
    dict.Add "%2C", ","
    dict.Add "%2D", "-"
    dict.Add "%2E", "."
    dict.Add "%2F", "/"
    
    ' --- NUMBERS (HEX 30-39) ---
    dict.Add "%30", "0"
    dict.Add "%31", "1"
    dict.Add "%32", "2"
    dict.Add "%33", "3"
    dict.Add "%34", "4"
    dict.Add "%35", "5"
    dict.Add "%36", "6"
    dict.Add "%37", "7"
    dict.Add "%38", "8"
    dict.Add "%39", "9"
    
    ' --- SYMBOLS (HEX 3A-40) ---
    dict.Add "%3A", ":"
    dict.Add "%3B", ";"
    dict.Add "%3C", "<"
    dict.Add "%3D", "="
    dict.Add "%3E", ">"
    dict.Add "%3F", "?"
    dict.Add "%40", "@"
    
    ' --- SYMBOLS (HEX 5B-60) ---
    dict.Add "%5B", "["
    dict.Add "%5C", "\"
    dict.Add "%5D", "]"
    dict.Add "%5E", "^"
    dict.Add "%5F", "_"
    dict.Add "%60", "`"
    
    ' --- SYMBOLS (HEX 7B-7E) ---
    dict.Add "%7B", "{"
    dict.Add "%7C", "|"
    dict.Add "%7D", "}"
    dict.Add "%7E", "~"
    
    ' 3. Select Range
    On Error Resume Next
    Set rng = Intersect(Selection, Selection.Parent.UsedRange)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Please select cells containing URLs."
        GoTo Cleanup
    End If
    
    ' 4. Process Data (Unified Array Logic)
    ' If selection is 1 cell, rng.Value is not an array, so we force it.
    If IsArray(rng.Value) Then
        dataArr = rng.Value
    Else
        ReDim dataArr(1 To 1, 1 To 1)
        dataArr(1, 1) = rng.Value
    End If
    
    ' Loop through the array (Rows then Columns)
    For i = LBound(dataArr, 1) To UBound(dataArr, 1)
        For j = LBound(dataArr, 2) To UBound(dataArr, 2)
            If Not IsError(dataArr(i, j)) Then
                If Len(dataArr(i, j)) > 0 Then
                    ' Cycle through Dictionary
                    For Each key In dict.Keys
                        dataArr(i, j) = Replace(dataArr(i, j), key, dict(key))
                    Next key
                End If
            End If
        Next j
    Next i
    
    ' Dump clean text back to Excel
    rng.Value = dataArr
    
    ' 5. Create Hyperlinks (with "http" check)
    Dim cell As Range
    Dim val As String
    
    For Each cell In rng
        val = CStr(cell.Value)
        ' Check if not empty AND starts with "http" (Case Insensitive)
        If Len(val) > 0 And Left(LCase(val), 4) = "http" Then
            ActiveSheet.Hyperlinks.Add Anchor:=cell, Address:=val
        End If
    Next cell

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Processing complete."

End Sub