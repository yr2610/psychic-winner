Function jsArray2dToSafeArray2d_old(jsArray)
    Dim l1, l2, result
    Dim i, j
    l1 = jsArray.length
    l2 = Eval("jsArray.[0].length")
    ReDim result(l1 - 1, l2 - 1)

    For i = 0 to l1 - 1
        For j = 0 to l2 - 1
            On Error Resume Next
            result(i, j) = Eval("jsArray.[" & i & "].[" & j & "]")
            On Error Goto 0
        Next
    Next

    jsArray2dToSafeArray2d_old = result
End Function

Function jsArray2dToSafeArray2d(jsArray, columns)
    Dim l1, l2, result
    Dim i, j
    l1 = jsArray.length
    'l2 = jsArray.[0].length
    l2 = columns
    ReDim result(l1 - 1, l2 - 1)

    Dim row, column, v
    i = 0
    For Each row In jsArray
        j = 0
        For Each v In row
            result(i, j) = v
            If j >= l2 - 1 Then
                Exit For
            End If
            j = j + 1
        Next
'        If j < l2 - 1 Then
'            For j = j To l2 - 1
'                result(i, j) = ""
'            Next
'        End If
        If i >= l1 - 1 Then
            Exit For
        End If
        i = i + 1
    Next
    
    'For i = 0 to l1 - 1
    '    For j = 0 to l2 - 1
    '        On Error Resume Next
    '        result(i, j) = Eval("jsArray.[" & i & "].[" & j & "]")
    '        On Error Goto 0
    '    Next
    'Next

    jsArray2dToSafeArray2d = result
End Function


' JScript の column major な１次元配列を excel の range.value に代入できる配列に
'Function jsArray1DToExcelRangeArray(jsArray, rows)
'    Dim l1, l2, result
'    Dim i, j
'    l1 = rows
'    l2 = jsArray.length / rows
'    ReDim result(l1 - 1, l2 - 1)
'
'    ' excel は (r, c)
'    For i = 0 to l1 - 1
'        For j = 0 to l2 - 1
'            On Error Resume Next
'            result(i, j) = Eval("jsArray.[" & (rows * j + i) & "]")
'            On Error Goto 0
'        Next
'    Next
'
'    jsArray1DToExcelRangeArray = result
'End Function

Function jsArray1dRowMajorToSafeArray2d(jsArray, columns)
    Dim rows : rows = jsArray.length / columns
    Dim l1, l2, result
    Dim i, j
    Dim t
    l1 = rows - 1
    l2 = columns - 1
    ReDim result(l1, l2)

    i = 0
    j = 0
    For Each t In jsArray
        result(i, j) = t
        j = j + 1
        If j > l2 Then
            If i >= l1 Then
                Exit For
            End If
            i = i + 1
            j = 0
        End If
    Next

    'For i = 0 to l1
    '    For j = 0 to l2
    '        On Error Resume Next
    '        result(i, j) = Eval("jsArray.[" & t & "]")
    '        t = t + 1
    '        On Error Goto 0
    '    Next
    'Next

    jsArray1dRowMajorToSafeArray2d = result
End Function

' JScript の column major な１次元配列を excel の range.value に代入できる配列に
'Function jsArray1DToExcelRangeArray(jsArray, rows)
Function jsArray1dColumnMajorToSafeArray2d(jsArray, rows)
    Dim columns : columns = jsArray.length / rows
    Dim l1, l2, result
    Dim i, j
    Dim t
    l1 = rows - 1
    l2 = columns - 1
    ReDim result(l1, l2)

    ' excel は (r, c)
    i = 0
    j = 0
    For Each t In jsArray
        result(i, j) = t
        i = i + 1
        If i > l1 Then
            If j >= l2 Then
                Exit For
            End If
            j = j + 1
            i = 0
        End If
    Next
    
    'For i = 0 to l1 - 1
    '    For j = 0 to l2 - 1
    '        On Error Resume Next
    '        result(i, j) = Eval("jsArray.[" & (rows * j + i) & "]")
    '        On Error Goto 0
    '    Next
    'Next

    jsArray1dColumnMajorToSafeArray2d = result
End Function
