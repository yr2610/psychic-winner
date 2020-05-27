Function jsArray2dToSafeArray2d(jsArray)
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

    jsArray2dToSafeArray2d = result
End Function

' JScript の column major な１次元配列を excel の range.value に代入できる配列に
Function jsArray1DToExcelRangeArray(jsArray, rows)
    Dim l1, l2, result
    Dim i, j
    l1 = rows
    l2 = jsArray.length / rows
    ReDim result(l1 - 1, l2 - 1)

    ' excel は (r, c)
    For i = 0 to l1 - 1
        For j = 0 to l2 - 1
            On Error Resume Next
            result(i, j) = Eval("jsArray.[" & (rows * j + i) & "]")
            On Error Goto 0
        Next
    Next

    jsArray1DToExcelRangeArray = result
End Function
