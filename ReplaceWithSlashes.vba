Sub CleanURLs_FullASCII()
    Dim rng As Range
    Dim dataArr As Variant
    Dim i As Long, j As Long
    Dim dict As Object
    Dim key As Variant
    
    ' 1. Optimization Settings
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 2. Create Dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' =========================================================
    ' SECTION 1: CUSTOM / LONG OVERRIDES (HIGHEST PRIORITY)
    ' =========================================================
    ' We place these first so they are caught before generic codes.
    dict.Add "%253A", "/"   ' Your specific requirement
    dict.Add "+", " "       ' Google/Query string space
    
    ' =========================================================
    ' SECTION 2: COMMON CONTROL CHARACTERS
    ' =========================================================
    dict.Add "%09", vbTab   ' Horizontal Tab
    dict.Add "%0A", vbLf    ' Line Feed
    dict.Add "%0D", vbCr    ' Carriage Return
    
    ' =========================================================
    ' SECTION 3: SYMBOLS & PUNCTUATION (HEX 20-2F)
    ' =========================================================
    dict.Add "%20", " "     ' Space
    dict.Add "%21", "!"     ' Exclamation mark
    dict.Add "%22", """"    ' Double quotes (escaped in VBA)
    dict.Add "%23", "#"     ' Number sign
    dict.Add "%24", "$"     ' Dollar sign
    dict.Add "%25", "%"     ' Percent
    dict.Add "%26", "&"     ' Ampersand
    dict.Add "%27", "'"     ' Single quote
    dict.Add "%28", "("     ' Left parenthesis
    dict.Add "%29", ")"     ' Right parenthesis
    dict.Add "%2A", "*"     ' Asterisk
    dict.Add "%2B", "+"     ' Plus
    dict.Add "%2C", ","     ' Comma
    dict.Add "%2D", "-"     ' Hyphen
    dict.Add "%2E", "."     ' Period
    dict.Add "%2F", "/"     ' Slash
    
    ' =========================================================
    ' SECTION 4: NUMBERS 0-9 (HEX 30-39)
    ' =========================================================
    ' (Included in case your data encodes numbers like %31 for 1)
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
    
    ' =========================================================
    ' SECTION 5: SYMBOLS (HEX 3A-40)
    ' =========================================================
    dict.Add "%3A", ":"     ' Colon
    dict.Add "%3B", ";"     ' Semicolon
    dict.Add "%3C", "<"     ' Less than
    dict.Add "%3D", "="     ' Equals
    dict.Add "%3E", ">"     ' Greater than
    dict.Add "%3F", "?"     ' Question mark
    dict.Add "%40", "@"     ' At sign
    
    ' =========================================================
    ' SECTION 6: SYMBOLS (HEX 5B-60)
    ' =========================================================
    dict.Add "%5B", "["     ' Left bracket
    dict.Add "%5C", "\"     ' Backslash
    dict.Add "%5D", "]"     ' Right bracket
    dict.Add "%5E", "^"     ' Caret
    dict.Add "%5F", "_"     ' Underscore
    dict.Add "%60", "`"     ' Grave accent
    
    ' =========================================================
    ' SECTION 7: SYMBOLS (HEX 7B-7E)
    ' =========================================================
    dict.Add "%7B", "{"     ' Left brace
    dict.Add "%7C", "|"     ' Vertical bar
    dict.Add "%7D", "}"     ' Right brace
    dict.Add "%7E", "~"     ' Tilde
    
    ' Note: We skipped A-Z (%41-%5A) and a-z (%61-%7A) 
    ' per your request.
    
    ' 3. Select Range
    On Error Resume Next
    Set rng = Intersect(Selection, Selection.Parent.UsedRange)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Please select cells containing URLs."
        GoTo Cleanup
    End If
    
    ' 4. Process Data
    If rng.Cells.Count = 1 Then
        Dim txt As String
        txt = rng.Value
        If Len(txt) > 0 Then
            For Each key In dict.Keys
                txt = Replace(txt, key, dict(key))
            Next key
            rng.Value = txt
            ActiveSheet.Hyperlinks.Add Anchor:=rng, Address:=txt
        End If
    Else
        dataArr = rng.Value
        For i = 1 To UBound(dataArr, 1)
            For j = 1 To UBound(dataArr, 2)
                If Not IsError(dataArr(i, j)) Then
                    If Len(dataArr(i, j)) > 0 Then
                        For Each key In dict.Keys
                            dataArr(i, j) = Replace(dataArr(i, j), key, dict(key))
                        Next key
                    End If
                End If
            Next j
        Next i
        
        rng.Value = dataArr
        
        Dim cell As Range
        For Each cell In rng
            If Len(cell.Value) > 0 Then
                ActiveSheet.Hyperlinks.Add Anchor:=cell, Address:=cell.Value
            End If
        Next cell
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Full ASCII decode complete."

End Sub