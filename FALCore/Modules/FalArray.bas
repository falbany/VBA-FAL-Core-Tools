Attribute VB_Name = "ArrayX"
' **************************************************************************************
' Class     : ArrayX
' Author    : Forent ALBANY
' Website   :
' Purpose   : Manipulation of arrays
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2023-07-00              Initial Release
'---------------------------------------------------------------------------------------
' Revision Propositions:
' ~~~~~~~~~~~~~~~~
' **************************************************************************************

Option Explicit



'/////////////////////////////////////////////////////////////////////////////////////////////////'
'                            _   _ _____ _     ____  _____ ____                                   '
'                           | | | | ____| |   |  _ \| ____|  _ \                                  '
'                           | |_| |  _| | |   | |_) |  _| | |_) |                                 '
'                           |  _  | |___| |___|  __/| |___|  _ <                                  '
'                           |_| |_|_____|_____|_|   |_____|_| \_\                                 '
'/////////////////////////////////////////////////////////////////////////////////////////////////'


Public Enum RegressionType
    rtExponential = 0
    rtLinear = 1
    rtPolynomialDegree2 = 2
    rtPolynomialDegree3 = 3
    rtPolynomialDegree4 = 4
    rtPolynomialDegree5 = 5
    rtPolynomialDegree6 = 6
End Enum

Public Function Get_Regression_Coefficients(ByVal YValues As Variant, ByVal XValues As Variant, ByVal RegressionModel As RegressionType) As Variant
    ' @brief Calculates the coefficients for various types of regression models (trendlines).
    ' @param YValues A range or array of known Y-values.
    ' @param XValues A range or array of known X-values.
    ' @param RegressionModel The type of regression to perform, from the RegressionType enum.
    ' @return A 2D array containing the calculated coefficients and regression statistics.
    '         The structure of the returned array depends on the model and is consistent with Excel's LINEST/LOGEST functions.
    ' @details This function is a wrapper for Excel's built-in LINEST and LOGEST worksheet functions.
    '          - For rtExponential, it returns {m, b} for the model y = b * m^x.
    '          - For rtLinear, it returns {m, b} for the model y = mx + b.
    '          - For polynomial models, it returns {c_n, c_n-1, ..., c_1, c_0} for y = c_n*x^n + ... + c_0.
    
    Select Case RegressionModel
        Case rtExponential:         Get_Regression_Coefficients = Application.LogEst(YValues, XValues, True, True)
        Case rtLinear:              Get_Regression_Coefficients = Application.LinEst(YValues, XValues, True, True)
        Case rtPolynomialDegree2:   Get_Regression_Coefficients = Application.LinEst(YValues, Application.Power(XValues, Application.Transpose(Array(1, 2))), True, True)
        Case rtPolynomialDegree3:   Get_Regression_Coefficients = Application.LinEst(YValues, Application.Power(XValues, Application.Transpose(Array(1, 2, 3))), True, True)
        Case rtPolynomialDegree4:   Get_Regression_Coefficients = Application.LinEst(YValues, Application.Power(XValues, Application.Transpose(Array(1, 2, 3, 4))), True, True)
        Case rtPolynomialDegree5:   Get_Regression_Coefficients = Application.LinEst(YValues, Application.Power(XValues, Application.Transpose(Array(1, 2, 3, 4, 5))), True, True)
        Case rtPolynomialDegree6:   Get_Regression_Coefficients = Application.LinEst(YValues, Application.Power(XValues, Application.Transpose(Array(1, 2, 3, 4, 5, 6))), True, True)
    End Select
End Function

' NOTE: The TrendEst() function for compatibility.
Public Function TrendEst(ByVal YValues As Variant, ByVal XValues As Variant, ByVal Degree As Integer) As Variant
    ' @brief 
    TrendEst =Get_Regression_Coefficients(YValues, XValues, RegressionType(Degree))
End Function

Private Function JsonFormatValue(ByVal Value As Variant) As String
    ' @brief Formats a VBA value into a valid JSON value string.
    ' @param Value The VBA value to format.
    ' @return A string representing the value in JSON format.
    If IsEmpty(Value) Or IsNull(Value) Or Value = vbNullString Then
        JsonFormatValue = "null"
    ElseIf IsNumeric(Value) Then
        ' Use Str() to ensure a period is used as the decimal separator, regardless of locale.
        ' Trim to remove the leading space for positive numbers.
        JsonFormatValue = Trim(Str(Value))
    ElseIf TypeName(Value) = "Boolean" Then
        JsonFormatValue = IIf(Value, "true", "false")
    ElseIf IsDate(Value) Then
        ' ISO 8601 format is a common and safe standard for dates in JSON.
        JsonFormatValue = """" & Format(Value, "yyyy-mm-ddThh:nn:ss") & """"
    Else
        ' It's a string, so escape special characters and wrap in quotes.
        Dim tempStr As String
        tempStr = CStr(Value)
        ' 1. Escape backslashes
        tempStr = Replace(tempStr, "\", "\\")
        ' 2. Escape double quotes
        tempStr = Replace(tempStr, """", "\""")
        ' 3. Escape other common control characters
        tempStr = Replace(tempStr, vbCrLf, "\n")
        tempStr = Replace(tempStr, vbCr, "\r")
        tempStr = Replace(tempStr, vbLf, "\n")
        tempStr = Replace(tempStr, vbTab, "\t")
        
        JsonFormatValue = """" & tempStr & """"
    End If
End Function

Private Function pIs_Array_Of_Dimension(ByVal Arr As Variant, ByVal TargetDimension As Long) As Boolean
    ' @brief (Private Helper) Checks if a variant is an array of a specific dimension.
    ' @param Arr The variant to check.
    ' @param TargetDimension The dimension to check for (e.g., 1 for 1D, 2 for 2D).
    ' @return True if Arr is an array of exactly TargetDimension dimensions, False otherwise.
    
    pIs_Array_Of_Dimension = False
    
    ' Must be an array to begin with.
    If Not IsArray(Arr) Then Exit Function
    
    On Error Resume Next
    
    ' Check if it has the target dimension.
    Dim dummy As Long
    dummy = UBound(Arr, TargetDimension)
    If Err.Number <> 0 Then
        ' It doesn't even have the target dimension, so it can't be an array of that dimension.
        Err.Clear
        Exit Function
    End If
    
    ' Check if it has the *next* dimension.
    dummy = UBound(Arr, TargetDimension + 1)
    If Err.Number <> 0 Then
        ' It has the target dimension, but not the next one. This is what we want.
        pIs_Array_Of_Dimension = True
    End If
    
    Err.Clear
End Function

Public Function Is_a1D(ByVal Arr As Variant) As Boolean
    ' @brief Checks if a variant is a 1-dimensional array.
    ' @param Arr The variant to check.
    ' @return True if Arr is a 1D array, False otherwise.
    Is_a1D = pIs_Array_Of_Dimension(Arr, 1)
End Function

Public Function Is_a2D(ByVal Arr As Variant) As Boolean
    ' @brief Checks if a variant is a 2-dimensional array.
    ' @param Arr The variant to check.
    ' @return True if Arr is a 2D array, False otherwise.
    Is_a2D = pIs_Array_Of_Dimension(Arr, 2)
End Function

Public Function Is_a3D(ByVal Arr As Variant) As Boolean
    ' @brief Checks if a variant is a 3-dimensional array.
    ' @param Arr The variant to check.
    ' @return True if Arr is a 3D array, False otherwise.
    Is_a3D = pIs_Array_Of_Dimension(Arr, 3)
End Function

Public Function Is_a4D(ByVal Arr As Variant) As Boolean
    ' @brief Checks if a variant is a 4-dimensional array.
    ' @param Arr The variant to check.
    ' @return True if Arr is a 4D array, False otherwise.
    Is_a4D = pIs_Array_Of_Dimension(Arr, 4)
End Function






'/////////////////////////////////////////////////////////////////////////////////////////////////'
'                     ____  ____       _                                                          ' 
'                    |___ \|  _ \     / \   _ __ _ __ __ _ _   _                                  ' 
'                     __) | | | |   / _ \ | '__| '__/ _` | | | |                                  '
'                    / __/| |_| |  / ___ \| |  | | | (_| | |_| |                                  '
'                   |_____|____/  /_/   \_\_|  |_|  \__,_|\__, |                                  '
'                                                         |___/                                   '
'/////////////////////////////////////////////////////////////////////////////////////////////////'








Public Function a2D_find_LastNonEmptyRowInColumnFromLine(Arr2D As Variant, Optional StartRow As Long = 1, Optional colNumber As Long = 1) As Long
    Dim i As Long
    For i = StartRow To UBound(Arr2D, 1)
        If IsEmpty(Arr2D(i, colNumber)) Then a2D_find_LastNonEmptyRowInColumnFromLine = i: Exit Function
    Next
    ' Si aucune valeur non vide n'est trouv�e, renvoie 0.
    a2D_find_LastNonEmptyRowInColumnFromLine = 0
End Function

Public Function a2D_find_LastNonEmptyColumnInRowFromColumn(Arr2D As Variant, Optional startCol As Long = 1, Optional rowNumber As Long = 1) As Long
    Dim j As Long
    For j = startCol To UBound(Arr2D, 2)
        If IsEmpty(Arr2D(rowNumber, j)) Then
            a2D_find_LastNonEmptyColumnInRowFromColumn = j - 1
            Exit Function
        End If
    Next j
    a2D_find_LastNonEmptyColumnInRowFromColumn = UBound(Arr2D, 2)
End Function

Public Function a1D_count_Occurrences(Arr1D As Variant, search As Variant) As Long
    Dim cell As Variant
    Dim count As Long
    count = 0
    For Each cell In Arr1D
        If cell = search Then count = count + 1
    Next cell
    
    a1D_count_Occurrences = count
End Function

Public Function a2D_count_StringOccurrences(Arr2D As Variant, search As Variant) As Long
    Dim cell As Variant
    Dim count As Long
    count = 0
    For Each cell In Arr2D
        If cell = search Then count = count + 1
    Next cell
    
    a2D_count_StringOccurrences = count
End Function

Public Function a2D_math_Derivate(Arr2D As Variant, Optional X_Split As Long = 3, Optional Degree As Integer = 2) As Variant
    ' Arr2D(1 To X, 1 To 2)
    ' Degree = 1 : linear regression. /!\ firsts and lasts values.
    ' Degree = 2 : Polynomial regression of degree order.
    'On Error GoTo ArrayError
    On Error Resume Next
    Dim XTrendEst() As Variant
    Dim YTrendEst() As Variant
    Dim tempArray() As Variant
    Dim resultArray() As Variant
    Dim X_Offset As Long
    Dim XIndex As Long
    Dim index As Long
    
    ReDim resultArray(1 To UBound(Arr2D), 1 To 2)
    ' Error End Conditions.
    'If UBound(Arr2D) < X_Split + 1 Then Debug.Print "DerivateArray Function error : UBound(Arr2D) < X_Split + 1": GoTo ArrayError
    'If X_Split > UBound(Arr2D) / 2 Then Debug.Print "DerivateArray Function error : X_Split > UBound(Arr2D) / 2": GoTo ArrayError
    ' WARNINGS.
    If X_Split > UBound(Arr2D) Then DebugPrint "ArrayX", "a2D_Derivate", "WARN, desc = X_Split > UBound(Arr2D) => X_Split = UBound(Arr2D)"
    If UBound(Arr2D) < X_Split + 1 Then DebugPrint "ArrayX", "a2D_Derivate", "WARN, desc = UBound(Arr2D) < X_Split + 1"
    If X_Split > UBound(Arr2D) / 2 Then DebugPrint "ArrayX", "a2D_Derivate", "WARN, desc = X_Split > UBound(Arr2D) / 2"
    
    ' Part Derivative.
    ReDim XTrendEst(0 To X_Split - 1)
    ReDim YTrendEst(0 To X_Split - 1)
    For XIndex = LBound(Arr2D) To UBound(Arr2D)
        X_Offset = XIndex - CInt((X_Split - 1) / 2)
        ' Boundary condition handling.
        If X_Offset + LBound(XTrendEst) < LBound(Arr2D) Then X_Offset = LBound(Arr2D)
        If X_Offset + UBound(XTrendEst) > UBound(Arr2D) Then
            X_Offset = X_Offset - UBound(XTrendEst)
            If X_Offset + LBound(XTrendEst) < LBound(Arr2D) Then X_Offset = LBound(Arr2D)
        End If
        ' Sub Array Construction.
        For index = 0 To UBound(XTrendEst)
            XTrendEst(index) = Arr2D(X_Offset + index, 1)
            YTrendEst(index) = Arr2D(X_Offset + index, 2)
        Next
        ' Derivative functions.
        Select Case Degree
            Case 1
                tempArray = TrendEst(YTrendEst, XTrendEst, 1)   ' A*X+B with A = (1,1) / B = (1,2) / R� = (3,1).
                resultArray(XIndex, 1) = Arr2D(XIndex, 1)   ' Copy X Data in ResultArray.
                resultArray(XIndex, 2) = tempArray(1, 1)        ' Derivative Calculation.
            Case 2
                tempArray = TrendEst(YTrendEst, XTrendEst, 2)   ' Ax�+Bx+C=Y with A = (1,1) / B = (1,2) / C = (1,3)/ R� = (3,1).
                resultArray(XIndex, 1) = Arr2D(XIndex, 1)   ' Copy X Data in ResultArray.
                resultArray(XIndex, 2) = 2 * tempArray(1, 1) * Arr2D(XIndex, 1) + tempArray(1, 2)     ' Derivative calculation.
        End Select
    Next
    ' Result Array.
    a2D_math_Derivate = resultArray
    Exit Function
    
ArrayError:
    a2D_math_Derivate = resultArray

End Function

Public Function a1D_math_Abs(Arr1D As Variant, Optional d1_offset As Long = 0) As Variant
    ' * @brief Returns the absolute value of numeric elements in a 3-dimensional array.
    ' * @param Arr3D The input array to calculate the absolute values.
    ' * @param d1_offset The offset for the first dimension. Default is 0.
    ' * @return The array with absolute values of numeric elements.
    
    Dim i                   As Long
    Dim j                   As Long
    Dim xLB1                As Long
    Dim xUB1                As Long
    Dim aXD                 As Variant
    
    aXD = Arr1D
    xLB1 = LBound(aXD, 1) + d1_offset
    xUB1 = UBound(aXD, 1)
   
    For i = xLB1 To xUB1
        If IsNumeric(aXD(i)) Then aXD(i) = Abs(aXD(i))
    Next
    
    a1D_math_Abs = aXD
End Function

Public Function a2D_math_Abs(Arr2D As Variant, Optional d1_offset As Long = 0, Optional d2_offset As Long = 0) As Variant
    ' * @brief Returns the absolute value of numeric elements in a 3-dimensional array.
    ' * @param Arr3D The input array to calculate the absolute values.
    ' * @param d1_offset The offset for the first dimension. Default is 0.
    ' * @param d2_offset The offset for the second dimension. Default is 0.
    ' * @return The array with absolute values of numeric elements.
    
    Dim i                   As Long
    Dim j                   As Long
    Dim xLB1                As Long
    Dim xUB1                As Long
    Dim xLB2                As Long
    Dim xUB2                As Long
    Dim aXD                 As Variant
    
    ' Convert Range in Array and Check if Arr2D is an array
    If TypeName(Arr2D) = "Range" Then Arr2D = Arr2D.value
    If Not IsArray(Arr2D) Then Exit Function
    
    aXD = Arr2D
    xLB1 = LBound(aXD, 1) + d1_offset
    xUB1 = UBound(aXD, 1)
    xLB2 = LBound(aXD, 2) + d2_offset
    xUB2 = UBound(aXD, 2)
    
    For i = xLB1 To xUB1
        For j = xLB2 To xUB2
            If IsNumeric(aXD(i, j)) Then aXD(i, j) = Abs(aXD(i, j))
        Next
    Next
    
    a2D_math_Abs = aXD
End Function

Public Function a1D_math_Clip(Arr1D As Variant, Optional Low As Variant, Optional High As Variant, Optional d1_offset As Long = 0) As Variant
    ' * @brief Returns the clipped value of numeric elements in a 3-dimensional array.
    ' * @param Arr1D The input array to calculate the absolute values.
    ' * @param d1_offset The offset for the first dimension. Default is 0.
    ' * @return The array with clipped values of numeric elements.

    Dim i                   As Long
    Dim xLB1                As Long
    Dim xUB1                As Long
    Dim aXD                 As Variant

    aXD = Arr1D
    xLB1 = LBound(aXD, 1) + d1_offset
    xUB1 = UBound(aXD, 1)

    For i = xLB1 To xUB1
        If TypeName(aXD(i)) = TypeName(Low) Then
            If aXD(i) < Low Then aXD(i) = Low
        End If
        If TypeName(aXD(i, j)) = TypeName(High) Then
            If aXD(i) > Low Then aXD(i) = High
        End If
    Next

    a1D_math_Clip = aXD
End Function

Public Function a2D_math_Clip(Arr2D As Variant, Optional Low As Variant, Optional High As Variant, Optional d1_offset As Long = 0, Optional d2_offset As Long = 0) As Variant
    ' * @brief Returns the clipped value of numeric elements in a 3-dimensional array.
    ' * @param Arr2D The input array to calculate the absolute values.
    ' * @param d1_offset The offset for the first dimension. Default is 0.
    ' * @param d2_offset The offset for the second dimension. Default is 0.
    ' * @return The array with clipped values of numeric elements.

    Dim i                   As Long
    Dim j                   As Long
    Dim xLB1                As Long
    Dim xUB1                As Long
    Dim xLB2                As Long
    Dim xUB2                As Long
    Dim aXD                 As Variant

    aXD = Arr2D
    xLB1 = LBound(aXD, 1) + d1_offset
    xUB1 = UBound(aXD, 1)
    xLB2 = LBound(aXD, 2) + d2_offset
    xUB2 = UBound(aXD, 2)

    For i = xLB1 To xUB1
        For j = xLB2 To xUB2
            If TypeName(aXD(i, j)) = TypeName(Low) Then
                If aXD(i, j) < Low Then aXD(i, j) = Low
            End If
            If TypeName(aXD(i, j)) = TypeName(High) Then
                If aXD(i, j) > Low Then aXD(i, j) = High
            End If
        Next
    Next

    a2D_math_Clip = aXD
End Function


' ' # TODO: use a2D_Write() instead
' Public Function a2D_WriteToSpreadSheet(Arr2D As Variant, Optional TopLeftCellAddress As String = "A1", Optional Workbook_Name As String = "ActiveWorkbook", Optional Worksheet_Name As Variant = "ActiveSheet") As Boolean
' ' OBSOLETE use a2D_Write() instead
'     ' Write Arr2D  to Workbooks(Workbook_Name).Worksheets(wks_name).Range(TopLeftCellAddress...
'     On Error GoTo ifError
'     If Workbook_Name = "ActiveWorkbook" Then Workbook_Name = ActiveWorkbook.name
'     If Worksheet_Name = "ActiveSheet" Then Worksheet_Name = ActiveSheet.name
'     Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range(TopLeftCellAddress & ":" & LANG_MOD.col(Range(TopLeftCellAddress).column + UBound(Arr2D, 2) - LBound(Arr2D, 2)) & (Range(TopLeftCellAddress).Row + UBound(Arr2D, 1) - LBound(Arr2D, 1))).value = Arr2D
'     a2D_WriteToSpreadSheet = True
'     Exit Function
' ifError:
'     Debug.Print format(Now, "h:mm:ss    ") & "FUNCTION : [a2D_WriteToSpreadSheet], ERROR"
' End Function

Public Function a2D_Write(ByVal Arr2D As Variant, ByVal TopLeftCell As Range, Optional ByVal WriteAs As String = "Value") As Boolean
    ' @brief Writes a 2D array to a worksheet starting at a specified cell, as values or formulas.
    ' @param Arr2D The 2D array (or a Range object) containing the data to write.
    ' @param TopLeftCell The top-left cell of the destination range.
    ' @param WriteAs (Optional) The property to write. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return True if the write operation was successful, False otherwise.
    ' @details This is the most efficient method to write an array to a sheet, as it avoids cell-by-cell looping
    '          and uses Range.Resize for direct range manipulation.
    On Error GoTo ifError
    a2D_Write = False ' Default return value

    ' 1. Input Validation
    If TopLeftCell Is Nothing Then
        Debug.Print "a2D_Write Error: TopLeftCell cannot be Nothing."
        Exit Function
    End If

    ' Allow passing a Range object directly
    If TypeName(Arr2D) = "Range" Then Arr2D = Arr2D.Value

    ' If Not IsArray(Arr2D) Then
    '     Debug.Print "a2D_Write Error: Input data is not a valid array."
    '     Exit Function
    ' End If

    ' Check if it's a 2D array
    if Not Is_a2D(Arr2D) Then 
        Debug.Print "a2D_Write Error: Input array is not a 2D array."
        Exit Function
    End If

    ' 2. Calculate dimensions and define destination range
    Dim numRows As Long, numCols As Long
    numRows = UBound(Arr2D, 1) - LBound(Arr2D, 1) + 1
    numCols = UBound(Arr2D, 2) - LBound(Arr2D, 2) + 1
    Dim destRange As Range
    Set destRange = TopLeftCell.Resize(numRows, numCols)

    ' 3. Write to the range using the specified property
    Select Case LCase(WriteAs)
        Case "value", "values": destRange.Value = Arr2D
        Case "formula", "formulas": destRange.Formula = Arr2D
        Case "formular1c1": destRange.FormulaR1C1 = Arr2D
        Case Else
            Debug.Print "a2D_Write Warning: Invalid 'WriteAs' property '" & WriteAs & "'. Defaulting to 'Value'."
            destRange.Value = Arr2D
    End Select

    a2D_Write = True
    Exit Function
ifError:
    Debug.Print "An error occurred in a2D_Write on worksheet '" & TopLeftCell.Parent.Name & "'. " & vbCrLf & "Error: " & Err.Description
End Function


Public Function aXD_count_Occurrence(ArrXD As Variant, matchValue As Variant, Optional matchType As Boolean = True, Optional matchComplete As Boolean = False) As Long
    ' * @brief Counts the number of occurrences of a specific value in an array.
    ' * @param ArrXD The input array to search for occurrences.
    ' * @param matchValue The value to match and count occurrences.
    ' * @param matchType Indicates whether the match should consider the value type. Default is True.
    ' * @param matchComplete Indicates whether the match should be complete or partial. Default is False.
    ' * @return The count of occurrences of the specified value and value type in the array.

    Dim Element         As Variant
    Dim count           As Long
    
    count = 0
    For Each Element In ArrXD
        If matchType Then
            If TypeName(Element) = TypeName(matchValue) Then
                If Element = matchValue Then count = count + 1
            End If
        Else
            If matchComplete Then
                If CStr(Element) = CStr(matchValue) Then count = count + 1
            Else
                If InStr(CStr(Element), CStr(matchValue)) > 0 Then count = count + 1
            End If
        End If
    Next

    aXD_count_Occurrence = count
End Function

Public Function a2D_replace_string(Arr2D As Variant, string1 As String, string2 As String) As Variant
    Dim i                   As Long
    Dim j                   As Long
    Dim xLB1                As Long
    Dim xUB1                As Long
    Dim xLB2                As Long
    Dim xUB2                As Long
    Dim aXD                 As Variant
    
    aXD = Arr2D
    xLB1 = LBound(aXD, 1)
    xUB1 = UBound(aXD, 1)
    xLB2 = LBound(aXD, 2)
    xUB2 = UBound(aXD, 2)
    
    For i = xLB1 To xUB1
        For j = xLB2 To xUB2
            If TypeName(ArrXD(i, j)) = TypeName(string1) Then aXD(i, j) = Replace(aXD(i, j), string1, string2)
        Next
    Next
    
    a2D_replace_string = aXD
End Function


Public Function a2D_Find_To_Collection(ByVal SearchArray As Variant, ByVal WhatToFind As Variant, Optional ByVal LookAt As XlLookAt = xlPart, Optional ByVal MatchCase As Boolean = False, Optional ByVal StartRow As Long = -1, Optional ByVal MaxItems As Long = -1) As Collection
    ' @brief Finds all occurrences of a value within a 2D array and returns a collection of their indices.
    ' @param SearchArray The 2D array to search within.
    ' @param WhatToFind The value to search for.
    ' @param LookAt (Optional) xlPart to match a substring, xlWhole to match the entire cell content. Defaults to xlPart.
    ' @param MatchCase (Optional) True for a case-sensitive search, False for case-insensitive. Defaults to False.
    ' @param StartRow (Optional) The row index to begin the search from. Defaults to the array's lower bound.
    ' @param MaxItems (Optional) The maximum number of items to find. A value of -1 (default) means find all occurrences.
    ' @return A Collection object where each item is a 2-element array containing the [row, col] indices of a found item.
    '         Returns an empty Collection if no matches are found. Returns Nothing on input error.
    
    ' 1. Input Validation
    If Not IsArray(SearchArray) Then Set a2D_Find_To_Collection = Nothing: Exit Function
    On Error Resume Next
    Dim ub2 As Long: ub2 = UBound(SearchArray, 2)
    If Err.Number <> 0 Then Set a2D_Find_To_Collection = Nothing: Exit Function ' Exit if not a 2D array
    On Error GoTo 0

    ' 2. Set up search parameters and results collection
    Dim r As Long, c As Long
    Dim lbRow As Long, ubRow As Long
    Dim lbCol As Long, ubCol As Long
    Dim foundItems As New Collection
    
    lbRow = LBound(SearchArray, 1)
    ubRow = UBound(SearchArray, 1)
    lbCol = LBound(SearchArray, 2)
    ubCol = UBound(SearchArray, 2)
    
    ' Determine start position
    If StartRow < lbRow Then StartRow = lbRow

    ' 3. Perform the search
    Dim compareMethod As VbCompareMethod
    compareMethod = IIf(MatchCase, vbBinaryCompare, vbTextCompare)
    
    For r = StartRow To ubRow
        For c = lbCol To ubCol
            Dim isMatch As Boolean
            isMatch = False
            If LookAt = xlWhole Then
                If StrComp(CStr(SearchArray(r, c)), CStr(WhatToFind), compareMethod) = 0 Then isMatch = True
            Else ' xlPart
                If InStr(1, CStr(SearchArray(r, c)), CStr(WhatToFind), compareMethod) > 0 Then isMatch = True
            End If
            
            If isMatch Then
                ' Store the found coordinates as a simple 1D array in the collection
                foundItems.Add Array(r, c)
                ' If a limit is set and reached, exit the search early.
                If MaxItems <> -1 And foundItems.Count >= MaxItems Then GoTo ExitSearch
            End If
        Next c
    Next r
    
ExitSearch:
    ' Return the collection (will be empty if nothing was found)
    Set a2D_Find_To_Collection = foundItems
End Function

Public Function a2D_Find(ByVal SearchArray As Variant, ByVal WhatToFind As Variant, Optional ByVal LookAt As XlLookAt = xlPart, Optional ByVal MatchCase As Boolean = False, Optional ByVal StartRow As Long = -1, Optional ByVal MaxItems As Long = -1) As Variant
    ' @brief Finds all occurrences of a value within a 2D array and returns their indices as a 2D array.
    ' @param SearchArray The 2D array to search within.
    ' @param WhatToFind The value to search for.
    ' @param LookAt (Optional) xlPart to match a substring, xlWhole to match the entire cell content. Defaults to xlPart.
    ' @param MatchCase (Optional) True for a case-sensitive search, False for case-insensitive. Defaults to False.
    ' @param StartRow (Optional) The row index to begin the search from. Defaults to the array's lower bound.
    ' @param MaxItems (Optional) The maximum number of items to find. A value of -1 (default) means find all occurrences.
    ' @return A 1-based, 2D array where each row contains the [row, col] indices of a found item.
    '         Returns Empty if no matches are found or if the input is not a valid 2D array.

    a2D_Find = Empty ' Default return value
    
    ' 1. Perform the search using the collection-based function
    Dim foundItems As Collection
    Set foundItems = a2D_Find_To_Collection(SearchArray, WhatToFind, LookAt, MatchCase, StartRow, MaxItems)
    
    ' 2. Convert the collection of results to a 2D array
    If Not foundItems Is Nothing Then
        If foundItems.Count > 0 Then
            Dim resultArray As Variant
            ReDim resultArray(1 To foundItems.Count, 1 To 2)
            Dim i As Long
            For i = 1 To foundItems.Count
                resultArray(i, 1) = foundItems(i)(0) ' Row
                resultArray(i, 2) = foundItems(i)(1) ' Column
            Next i
            a2D_Find = resultArray
        End If
    End If
End Function

Public Function a2D_Find_First(ByVal SearchArray As Variant, ByVal WhatToFind As Variant, Optional ByVal LookAt As XlLookAt = xlPart, Optional ByVal MatchCase As Boolean = False, Optional ByVal StartRow As Long = -1) As Variant
    ' @brief Finds the first occurrence of a value within a 2D array and returns its indices.
    ' @param SearchArray The 2D array to search within.
    ' @param WhatToFind The value to search for.
    ' @param LookAt (Optional) xlPart to match a substring, xlWhole to match the entire cell content. Defaults to xlPart.
    ' @param MatchCase (Optional) True for a case-sensitive search, False for case-insensitive. Defaults to False.
    ' @param StartRow (Optional) The row index to begin the search from. Defaults to the array's lower bound.
    ' @return A 1-based, 2D array where each row contains the [row, col] indices of a found item.
    '         Returns Empty if no matches are found or if the input is not a valid 2D array.
    ' @dependencies a2D_Find_To_Collection

    a2D_Find_First = Empty ' Default return value
    
    ' 1. Perform the search using the collection-based function
    Dim foundItems As Collection
    Set foundItems = a2D_Find_To_Collection(SearchArray, WhatToFind, LookAt, MatchCase, StartRow, 1)
    
    ' 2. Convert the collection of results to a 1D array
    If Not foundItems Is Nothing Then
        If foundItems.Count > 0 Then
            Dim resultArray As Variant
            ReDim resultArray(1 To 2)
            resultArray(1) = foundItems(1)(0) ' Row
            resultArray(2) = foundItems(1)(1) ' Column
            a2D_Find = resultArray
        End If
    End If

End Function

' # TODO: replace by a2D_Find_First()
Function a2D_find_String(searchString As String, SearchArray As Variant, Optional caseSensitive = True, Optional skipLines As Long = 0) As String
    ' * brief search for searchString in searchArray
    ' * return the position of the first found string. "" else
    Dim i               As Long
    Dim j               As Long
    Dim LBRows          As Long
    Dim UBRows          As Long
    Dim LBCols          As Long
    Dim UBCols          As Long
    
    If Not IsArray(SearchArray) Then a2D_find_String = "Not An Array": Exit Function
    LBRows = LBound(SearchArray, 1)
    UBRows = UBound(SearchArray, 1)
    LBCols = LBound(SearchArray, 2)
    UBCols = UBound(SearchArray, 2)
    
    If Not caseSensitive Then searchString = UCase(searchString)
    
    For i = LBRows + skipLines To UBRows
        For j = LBCols To UBCols
            If caseSensitive Then
                If SearchArray(i, j) = searchString Then a2D_find_String = LANG_MOD.col(j + 1 - LBCols) & i + 1 - LBRows: Exit Function
            Else
                If UCase(SearchArray(i, j)) = searchString Then a2D_find_String = LANG_MOD.col(j + 1 - LBCols) & i + 1 - LBRows: Exit Function
            End If
        Next j
    Next i
    a2D_find_String = ""
End Function

Public Function a2D_Isolate_XY(Arr2D As Variant, X_Column As Variant, Y_Column As Variant) As Variant
    ' Convert 2DArray(x To X, y to Y) to 2DArray(x to X, X_Column & Y_Column)
    ' X_Column & Y_Column Inputs can be Integer or String. Example : Y_Column = 2 <=> Y_Column = "B".
    Dim XIndex As Long
    Dim resultArray() As Variant
    ReDim resultArray(LBound(Arr2D, 1) To UBound(Arr2D, 1), LBound(Arr2D, 2) To LBound(Arr2D, 2) + 1)
    ' Convert Column Address to Column Number.
    If VarType(X_Column) = vbString Then X_Column = Range(X_Column & 1).column
    If VarType(Y_Column) = vbString Then Y_Column = Range(Y_Column & 1).column
    ' ERROR.
    If X_Column < LBound(Arr2D, 2) Or X_Column > UBound(Arr2D, 2) Then Debug.Print "ERROR ! a2D_Isolate_XY() : X_Column = " & X_Column & " is out of Arr2D range": GoTo ifError
    If Y_Column < LBound(Arr2D, 2) Or Y_Column > UBound(Arr2D, 2) Then Debug.Print "ERROR ! a2D_Isolate_XY() : Y_Column = " & Y_Column & " is out of Arr2D range": GoTo ifError
    ' 2DArray construction.
    For XIndex = LBound(Arr2D) To UBound(Arr2D)
        resultArray(XIndex, 1) = Arr2D(XIndex, CInt(X_Column))
        resultArray(XIndex, 2) = Arr2D(XIndex, CInt(Y_Column))
    Next
    a2D_Isolate_XY = resultArray
    Exit Function
ifError:
    a2D_Isolate_XY = resultArray
End Function

Public Function a1D_BaseOptionTransposition(Arr1D As Variant, ArrayBaseOption As Integer) As Variant
    Dim XIndex          As Long
    Dim BoOffset        As Integer
    Dim resultArray()   As Variant
    
    BoOffset = ArrayBaseOption - LBound(Arr1D)
    If BoOffset <> 0 Then
        ReDim resultArray(ArrayBaseOption To UBound(Arr1D, 1) + BoOffset)
        For XIndex = ArrayBaseOption To UBound(Arr1D, 1) + BoOffset
            resultArray(XIndex) = Arr1D(XIndex - BoOffset)
        Next
        a1D_BaseOptionTransposition = resultArray
    Else
        a1D_BaseOptionTransposition = Arr1D
    End If
        
    Erase resultArray
    Exit Function
End Function

Public Function a2D_BaseOptionTransposition(ByVal Arr2D As Variant, ArrayBaseOption As Integer) As Variant
    Dim XIndex As Long
    Dim YIndex As Long
    Dim BoOffset As Integer
    Dim resultArray() As Variant
    
    BoOffset = ArrayBaseOption - LBound(Arr2D)
    If BoOffset <> 0 Then
        ReDim resultArray(ArrayBaseOption To UBound(Arr2D, 1) + BoOffset, ArrayBaseOption To UBound(Arr2D, 2) + BoOffset)
        For XIndex = ArrayBaseOption To UBound(Arr2D, 1) + BoOffset
            For YIndex = ArrayBaseOption To UBound(Arr2D, 2) + BoOffset
                resultArray(XIndex, YIndex) = Arr2D(XIndex - BoOffset, YIndex - BoOffset)
            Next
        Next
        a2D_BaseOptionTransposition = resultArray
    Else
        a2D_BaseOptionTransposition = Arr2D
    End If
    Erase resultArray
    Exit Function
End Function

' # TODO: replace by a2D_Get_Column()
Public Function a2D_Column_Isolate(Arr2D As Variant, OutColumn As Long) As Variant
    ' Convert 2D Array(1 To X, 1 to Y) to Array(1 To X, 1 to 1)
    Dim XIndex As Long
    Dim resultArray() As Variant
    ReDim resultArray(1 To UBound(Arr2D), 1 To 1)
    ' ERROR.
    If OutColumn > UBound(Arr2D, 2) Then GoTo ArrayError
    ' 2DArray construction.
    For XIndex = 1 To UBound(Arr2D)
        resultArray(XIndex, 1) = Arr2D(XIndex, OutColumn)
    Next
    ' Result Array.
    a2D_Column_Isolate = resultArray
    Exit Function
ArrayError:
    a2D_Column_Isolate = resultArray
End Function


Public Function a2D_Get_Column(ByVal Arr2D As Variant, ByVal ColumnIndex As Long, Optional ByVal As1DArray As Boolean = False) As Variant
    ' @brief Extracts a single column from a 2D array.
    ' @param Arr2D The source 2D array.
    ' @param ColumnIndex The index of the column to extract.
    ' @param As1DArray (Optional) If True, returns a 1D array. If False (default), returns a 2D array with one column.
    ' @return A Variant array (1D or 2D) containing the data from the specified column.
    '         Returns Empty if inputs are invalid or an error occurs.

    a2D_Get_Column = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(Arr2D) Then Exit Function
    On Error Resume Next
    Dim ub2 As Long: ub2 = UBound(Arr2D, 2)
    If Err.Number <> 0 Then Exit Function ' Not a 2D array
    On Error GoTo 0

    Dim lbCol As Long: lbCol = LBound(Arr2D, 2)
    If ColumnIndex < lbCol Or ColumnIndex > UBound(Arr2D, 2) Then Exit Function

    ' 2. Set up dimensions and extract data
    Dim r As Long
    Dim lbRow As Long: lbRow = LBound(Arr2D, 1)
    Dim ubRow As Long: ubRow = UBound(Arr2D, 1)
    Dim resultArray As Variant

    If As1DArray Then
        ReDim resultArray(lbRow To ubRow)
        For r = lbRow To ubRow
            resultArray(r) = Arr2D(r, ColumnIndex)
        Next r
    Else ' Return as 2D array (N rows, 1 column)
        ReDim resultArray(lbRow To ubRow, 1 To 1)
        For r = lbRow To ubRow
            resultArray(r, 1) = Arr2D(r, ColumnIndex)
        Next r
    End If

    ' 3. Return the result
    a2D_Get_Column = resultArray
End Function

Public Function a2D_Get_Row(ByVal Arr2D As Variant, ByVal RowIndex As Long, Optional ByVal As1DArray As Boolean = True) As Variant
    ' @brief Extracts a single row from a 2D array.
    ' @param Arr2D The source 2D array.
    ' @param RowIndex The index of the row to extract.
    ' @param As1DArray (Optional) If True (default), returns a 1D array. If False, returns a 2D array with one row.
    ' @return A Variant array (1D or 2D) containing the data from the specified row.
    '         Returns Empty if inputs are invalid or an error occurs.

    a2D_Get_Row = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(Arr2D) Then Exit Function
    On Error Resume Next
    Dim ub2 As Long: ub2 = UBound(Arr2D, 2)
    If Err.Number <> 0 Then Exit Function ' Not a 2D array
    On Error GoTo 0

    Dim lbRow As Long: lbRow = LBound(Arr2D, 1)
    If RowIndex < lbRow Or RowIndex > UBound(Arr2D, 1) Then Exit Function

    ' 2. Set up dimensions and extract data
    Dim c As Long
    Dim lbCol As Long: lbCol = LBound(Arr2D, 2)
    Dim ubCol As Long: ubCol = UBound(Arr2D, 2)
    Dim resultArray As Variant

    ReDim resultArray(lbCol To ubCol)
    For c = lbCol To ubCol
        resultArray(c) = Arr2D(RowIndex, c)
    Next c

    ' 3. Return the result
    a2D_Get_Row = resultArray
End Function

Public Function a2D_Create(ByVal NumRows As Long, ByVal NumCols As Long, Optional ByVal FillValue As Variant = vbNullString) As Variant
    ' @brief Creates a new 2D array of a specified size and initializes all its elements to a given value.
    ' @param NumRows The number of rows for the new array. Must be greater than 0.
    ' @param NumCols The number of columns for the new array. Must be greater than 0.
    ' @param FillValue (Optional) The value to fill every element of the array with. Defaults to an empty string.
    ' @return A 1-based 2D Variant array.
    '         Returns Empty if NumRows or NumCols are less than 1.

    a2D_Create = Empty ' Default return value

    ' 1. Input Validation
    If NumRows < 1 Or NumCols < 1 Then Exit Function

    ' 2. Create and initialize the array
    Dim resultArray As Variant
    ReDim resultArray(1 To NumRows, 1 To NumCols)

    Dim r As Long, c As Long
    For r = 1 To NumRows
        For c = 1 To NumCols
            resultArray(r, c) = FillValue
        Next c
    Next r

    ' 3. Return the result
    a2D_Create = resultArray
End Function



Public Function a2D_Create_NxN(ByVal Size As Long, Optional ByVal FillValue As Variant = vbNullString) As Variant
    ' @brief Creates a new square 2D array (NxN) and initializes all its elements to a given value.
    ' @param Size The number of rows and columns for the new square array. Must be greater than 0.
    ' @param FillValue (Optional) The value to fill every element of the array with. Defaults to an empty string.
    ' @return A 1-based 2D Variant array.
    '         Returns Empty if Size is less than 1.
    ' @dependencies a2D_Create
    a2D_Create_NxN = a2D_Create(Size, Size, FillValue)
End Function

' # TODO: replace by a2D_Create()
Public Function a2D_Create_MxN(m As Long, n As Long, value As Variant) As Variant
    '------------------------------------------------------------------------------
    ' @fn       Public Function a2D_Create_MxN
    ' @brief    Creates a 2D array with dimensions MxN and initializes it with a specified value.
    ' @param    M       Long      Number of rows in the 2D array.
    ' @param    N       Long      Number of columns in the 2D array.
    ' @param    value   Variant   Value with which the array will be initialized.
    ' @return   Variant           Initialized 2D array.
    '------------------------------------------------------------------------------
    Dim i                   As Integer
    Dim j                   As Integer
    Dim aXD                 As Variant
    
    ReDim aXD(1 To m, 1 To n)

    For i = 1 To m
        For j = 1 To n
            aXD(i, j) = value
        Next j
    Next i
    
    a2D_Create_MxN = aXD
End Function

Public Function a2D_Join(Arr2D As Variant, ByVal Delimiter As String, Optional JoinByLine As Boolean = True) As String
    '------------------------------------------------------------------------------
    ' @fn       Function a2D_Join
    ' @brief    Joins a 2D array into a string with a specified delimiter.
    ' @param    Arr2D      Variant   Input 2D array to be joined.
    ' @param    delimiter  String    Delimiter to be used for joining elements.
    ' @param    JoinByLine String   If elements have to be joined by lines or columns.
    ' @return   String              Joined string.
    '------------------------------------------------------------------------------
    a2D_Join = a2D_ToString(Arr2D, Delimiter, JoinByLine)
End Function

Public Function a2D_Join_ByLine(Arr2D As Variant, ByVal Delimiter As String) As String
    '------------------------------------------------------------------------------
    ' @fn       Function a2D_Join_ByLine
    ' @brief    Joins a 2D array into a string with a specified delimiter.
    ' @param    Arr2D      Variant   Input 2D array to be joined.
    ' @param    delimiter  String    Delimiter to be used for joining elements.
    ' @return   String              Joined string.
    '------------------------------------------------------------------------------
    a2D_Join_ByLine = a2D_ToString(Arr2D, Delimiter, True)
End Function

Public Function a2D_Join_ByColumn(Arr2D As Variant, ByVal Delimiter As String) As String
    '------------------------------------------------------------------------------
    ' @fn       Function a2D_Join_ByColumn
    ' @brief    Joins a 2D array into a string with a specified delimiter.
    ' @param    Arr2D      Variant   Input 2D array to be joined.
    ' @param    delimiter  String    Delimiter to be used for joining elements.
    ' @return   String              Joined string.
    '------------------------------------------------------------------------------
    a2D_Join_ByColumn = a2D_ToString(Arr2D, Delimiter, False)
End Function

Public Function a2D_ToString(Arr2D As Variant, ByVal Delimiter As String, Optional JoinByLine As Boolean = True) As String
    '------------------------------------------------------------------------------
    ' @fn       Function a2D_ToString
    ' @brief    Joins a 2D array into a string with a specified delimiter.
    ' @param    Arr2D      Variant   Input 2D array to be joined.
    ' @param    delimiter  String    Delimiter to be used for joining elements.
    ' @param    JoinByLine String   If elements have to be joined by lines or columns.
    ' @return   String              Joined string.
    '------------------------------------------------------------------------------
    Dim i               As Long
    Dim j               As Long
    Dim sbResult        As New StringBuilder
        
    If TypeName(Arr2D) = "Range" Then Arr2D = Arr2D.value
    If Not IsArray(Arr2D) Then Exit Function
        
    If JoinByLine Then
        For i = LBound(Arr2D, 1) To UBound(Arr2D, 1)
            For j = LBound(Arr2D, 2) To UBound(Arr2D, 2)
                sbResult.Append CStr(Arr2D(i, j)) & Delimiter
            Next j
        Next i
    Else
        For j = LBound(Arr2D, 2) To UBound(Arr2D, 2)
            For i = LBound(Arr2D, 1) To UBound(Arr2D, 1)
                sbResult.Append CStr(Arr2D(i, j)) & Delimiter
            Next i
        Next j
    End If

    a2D_ToString = sbResult.ToString
    Set sbResult = Nothing
End Function

Public Function a2D_to_a1D_Column(Arr2D As Variant, OutColumn As Long) As Variant
    ' Convert 2D Array(1 To X, 1 to Y) to Array(1 To X, 1 to 1)
    Dim XIndex As Long
    Dim resultArray() As Variant
    
    ReDim resultArray(1 To UBound(Arr2D))
    
    For XIndex = 1 To UBound(Arr2D)
        resultArray(XIndex) = Arr2D(XIndex, OutColumn)
    Next

    a2D_to_a1D_Column = resultArray
    Exit Function
ArrayError:
    a2D_to_a1D_Column = resultArray
End Function

Public Function a2D_Column_CopyTo(Src_Arr2D As Variant, Src_Column As Long, Des_Arr2D As Variant, Optional Des_Column As Variant = "Src_Column", Optional Des_StartLine As Variant = "LBound") As Variant
    Dim XIndex As Long
    Dim Des_LastLine As Long
    Dim Des_OffsetLine As Long
    Dim resultArray() As Variant
    ReDim resultArray(LBound(Des_Arr2D, 1) To UBound(Des_Arr2D, 1), LBound(Des_Arr2D, 2) To UBound(Des_Arr2D, 2))
    
    ' BOUNDARY.
    If Des_Column = "Src_Column" Then Des_Column = CLng(Src_Column)
    If Des_StartLine = "LBound" Then Des_StartLine = LBound(Des_Arr2D, 1)
    Des_OffsetLine = CLng(Des_StartLine) - CLng(LBound(Src_Arr2D))
    If UBound(Src_Arr2D, 1) + Abs(Des_OffsetLine) > UBound(Des_Arr2D, 1) Then Des_LastLine = UBound(Des_Arr2D, 1) Else Des_LastLine = UBound(Src_Arr2D, 1) + Abs(Des_OffsetLine)
    ' ERROR.
    If LBound(Des_Arr2D, 1) <> LBound(Src_Arr2D, 1) Then Debug.Print "ERROR ! a2D_Column_CopyTo() : LBound(Des_Arr2D, 1) <> LBound(Src_Arr2D, 1)": GoTo ErrorFunct
    If CLng(Des_StartLine) < LBound(Des_Arr2D, 1) Then Debug.Print "ERROR ! a2D_Column_CopyTo() : CLng(Des_StartLine) < LBound(Des_Arr2D, 1)": GoTo ErrorFunct
    ' WARNINGS.
    If UBound(Des_Arr2D, 1) <> UBound(Src_Arr2D, 1) Then Debug.Print "WARNING ! a2D_Column_CopyTo() : UBound(Des_Arr2D, 1) <> UBound(Src_Arr2D, 1)"
    ' COPY.
    resultArray = Des_Arr2D
    For XIndex = CLng(Des_StartLine) To Des_LastLine
        resultArray(XIndex, CLng(Des_Column)) = Src_Arr2D(XIndex - Des_OffsetLine, CLng(Src_Column))
    Next

    a2D_Column_CopyTo = resultArray
    Erase resultArray
    Exit Function
ErrorFunct:
    a2D_Column_CopyTo = resultArray
End Function


' # TODO: use a2D_Insert_Rows() instead
Public Function a2D_Row_Add(Arr2D As Variant, FirstRowPosition As Long, Optional NbRowsToAdd As Long = 1) As Variant
    ' Add empty Rows to 2DArray from FirstRowPosition.
' TO VALIDATE
    Dim YIndex As Long
    Dim XIndex As Long
    Dim XOffset As Long
    Dim XOffset2 As Long
    Dim resultArray() As Variant
    ReDim resultArray(LBound(Arr2D, 1) To UBound(Arr2D, 1) + NbRowsToAdd, LBound(Arr2D, 2) To UBound(Arr2D, 2))
    
    ' WARNINGS.
    If NbRowsToAdd < 1 Then Debug.Print "WARNING ! AddRowTo_2DArray() : NbRowsToAdd < 1 => NbRowsToAdd = 1"
    If FirstRowPosition < LBound(Arr2D, 1) Then Debug.Print "WARNING ! AddRowTo_2DArray() : FirstRowPosition < LBound(Arr2D, 1) => FirstRowPosition = LBound(Arr2D, 1)"
    If FirstRowPosition > UBound(Arr2D, 1) + 1 Then Debug.Print "WARNING ! AddRowTo_2DArray() : FirstRowPosition > UBound(Arr2D, 1) + 1 => FirstRowPosition = UBound(Arr2D, 1) + 1"
    ' BOUNDARY.
    If NbRowsToAdd < 1 Then NbRowsToAdd = 1
    If FirstRowPosition < LBound(Arr2D, 1) Then FirstRowPosition = LBound(Arr2D, 1)
    If FirstRowPosition > UBound(Arr2D, 1) + 1 Then FirstRowPosition = UBound(Arr2D, 1) + 1
    ' BUILD The 2DArray.
    XOffset = 0
    XOffset2 = 0
    For XIndex = LBound(Arr2D, 1) To UBound(resultArray, 1)
        If XIndex <> FirstRowPosition + Abs(XOffset2) Then
            For YIndex = LBound(Arr2D, 2) To UBound(Arr2D, 2)
                resultArray(XIndex, YIndex) = Arr2D(XIndex + XOffset, YIndex)
            Next
        Else
            XOffset = XOffset - 1
            If Abs(XOffset2) < NbRowsToAdd - 1 Then XOffset2 = XOffset2 + 1
        End If
    Next
    a2D_Row_Add = resultArray
    Erase resultArray
    Exit Function
End Function


Public Function a2D_Insert_Rows(ByVal Arr2D As Variant, ByVal InsertAtRow As Long, Optional ByVal RowCount As Long = 1) As Variant
    ' @brief Inserts one or more empty rows into a 2D array at a specified position.
    ' @param Arr2D The source 2D array.
    ' @param InsertAtRow The index where the new empty rows will be inserted.
    ' @param RowCount (Optional) The number of empty rows to insert. Defaults to 1.
    ' @return A new, larger 2D array with the empty rows inserted.
    '         Returns Empty if inputs are invalid or an error occurs.

    a2D_Insert_Rows = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(Arr2D) Then Exit Function
    On Error Resume Next
    Dim ub2 As Long: ub2 = UBound(Arr2D, 2)
    If Err.Number <> 0 Then Exit Function ' Not a 2D array
    On Error GoTo 0

    Dim lbRow As Long: lbRow = LBound(Arr2D, 1)
    Dim ubRow As Long: ubRow = UBound(Arr2D, 1)

    ' Check if insertion point is valid
    If InsertAtRow < lbRow Or InsertAtRow > ubRow + 1 Then Exit Function
    If RowCount < 1 Then RowCount = 1

    ' 2. Set up dimensions for the new array
    Dim lbCol As Long: lbCol = LBound(Arr2D, 2)
    Dim ubCol As Long: ubCol = UBound(Arr2D, 2)
    Dim newRowCount As Long: newRowCount = (ubRow - lbRow + 1) + RowCount

    Dim resultArray As Variant
    ReDim resultArray(lbRow To lbRow + newRowCount - 1, lbCol To ubCol)

    ' 3. Copy data to the new array
    Dim r As Long, c As Long
    Dim destRow As Long: destRow = lbRow

    ' Copy the part before the insertion point
    For r = lbRow To InsertAtRow - 1
        For c = lbCol To ubCol
            resultArray(destRow, c) = Arr2D(r, c)
        Next c
        destRow = destRow + 1
    Next r

    ' Skip the inserted rows (they are already Empty)
    destRow = destRow + RowCount

    ' Copy the part after the insertion point
    For r = InsertAtRow To ubRow
        For c = lbCol To ubCol
            resultArray(destRow, c) = Arr2D(r, c)
        Next c
        destRow = destRow + 1
    Next r

    ' 4. Return the result
    a2D_Insert_Rows = resultArray
End Function


Public Function a2D_Slice_Rows(ByVal Arr2D As Variant, ByVal StartRow As Long, ByVal EndRow As Long) As Variant
    ' @brief Extracts a horizontal slice (a range of rows) from a 2D array.
    ' @param Arr2D The source 2D array.
    ' @param StartRow The starting row index of the slice to extract.
    ' @param EndRow The ending row index of the slice to extract.
    ' @return A new 2D array containing the specified rows.
    '         The returned array is 1-based in its first dimension.
    '         Returns Empty if inputs are invalid or an error occurs.

    a2D_Slice_Rows = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(Arr2D) Then Exit Function
    On Error Resume Next
    Dim ub2 As Long: ub2 = UBound(Arr2D, 2)
    If Err.Number <> 0 Then Exit Function ' Not a 2D array
    On Error GoTo 0

    Dim lbRow As Long: lbRow = LBound(Arr2D, 1)
    Dim ubRow As Long: ubRow = UBound(Arr2D, 1)

    ' Check if row indices are valid and in the correct order
    If StartRow < lbRow Or StartRow > ubRow Then Exit Function
    If EndRow < lbRow Or EndRow > ubRow Then Exit Function
    If EndRow < StartRow Then Exit Function

    ' 2. Set up dimensions for the new array
    Dim numRowsToCopy As Long: numRowsToCopy = EndRow - StartRow + 1
    Dim lbCol As Long: lbCol = LBound(Arr2D, 2)
    Dim ubCol As Long: ubCol = UBound(Arr2D, 2)

    Dim resultArray As Variant
    ReDim resultArray(1 To numRowsToCopy, lbCol To ubCol)

    ' 3. Copy the slice data
    Dim destRow As Long, srcRow As Long, c As Long
    destRow = 1
    For srcRow = StartRow To EndRow
        For c = lbCol To ubCol
            resultArray(destRow, c) = Arr2D(srcRow, c)
        Next c
        destRow = destRow + 1
    Next srcRow

    ' 4. Return the result
    a2D_Slice_Rows = resultArray
End Function

' # TODO: use a2D_Slice_Rows() instead
Public Function a2D_Row_Cut(Arr2D As Variant, FirstArrayLine As Variant, LastArrayLine As Variant) As Variant
    Dim XIndex As Long
    Dim YIndex As Long
    Dim XOffset As Long
    Dim PointsNb As Long
    Dim resultArray() As Variant
    
    XOffset = LBound(Arr2D, 1) - FirstArrayLine
    PointsNb = LastArrayLine - FirstArrayLine
    ReDim resultArray(LBound(Arr2D, 1) To UBound(Arr2D, 1) + PointsNb, LBound(Arr2D, 2) To UBound(Arr2D, 2))
    ' ERROR.
    If FirstArrayLine < LBound(Arr2D, 1) Or FirstArrayLine > UBound(Arr2D, 1) Then Debug.Print "ERROR ! a2D_Row_Cut() : FirstArrayLine = " & FirstArrayLine & " is out of Arr2D row range": GoTo ErrorFunct
    If LastArrayLine < LBound(Arr2D, 1) Or LastArrayLine > UBound(Arr2D, 1) Then Debug.Print "ERROR ! a2D_Row_Cut() : LastArrayLine = " & LastArrayLine & " is out of Arr2D row range": GoTo ErrorFunct
    ' 2DArray construction.
    For XIndex = FirstArrayLine To LastArrayLine
        For YIndex = LBound(Arr2D, 2) To UBound(Arr2D, 2)
            resultArray(XIndex + XOffset, YIndex) = Arr2D(XIndex, YIndex)
        Next
    Next
    a2D_Row_Cut = resultArray
    Erase resultArray
    Exit Function
ErrorFunct:
    a2D_Row_Cut = resultArray
End Function

Public Function a2D_To_a1D(ByVal Arr2D As Variant, Optional ByVal ByRow As Boolean = True) As Variant
    ' @brief Converts a 2D array into a 1D array by concatenating rows or columns.
    ' @param Arr2D The source 2D array.
    ' @param ByRow (Optional) If True (default), flattens the array row by row. If False, flattens it column by column.
    ' @return A 1-based 1D Variant array.
    '         Returns Empty if the input is not a valid 2D array or is empty.

    a2D_To_a1D = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(Arr2D) Then Exit Function
    On Error Resume Next
    Dim ub2 As Long: ub2 = UBound(Arr2D, 2)
    If Err.Number <> 0 Then Exit Function ' Not a 2D array
    On Error GoTo 0

    ' 2. Calculate dimensions
    Dim lbRow As Long: lbRow = LBound(Arr2D, 1)
    Dim ubRow As Long: ubRow = UBound(Arr2D, 1)
    Dim lbCol As Long: lbCol = LBound(Arr2D, 2)
    Dim ubCol As Long: ubCol = UBound(Arr2D, 2)

    Dim totalElements As Long: totalElements = (ubRow - lbRow + 1) * (ubCol - lbCol + 1)
    If totalElements = 0 Then Exit Function ' Empty source array

    ' 3. Create and populate the result array
    Dim resultArray As Variant
    ReDim resultArray(1 To totalElements)
    Dim r As Long, c As Long
    Dim idx As Long: idx = 1

    If ByRow Then
        For r = lbRow To ubRow
            For c = lbCol To ubCol
                resultArray(idx) = Arr2D(r, c): idx = idx + 1
            Next c
        Next r
    Else
        For c = lbCol To ubCol
            For r = lbRow To ubRow
                resultArray(idx) = Arr2D(r, c): idx = idx + 1
            Next r
        Next c
    End If

    ' 4. Return the result
    a2D_To_a1D = resultArray
End Function

Public Function a2D_Merge_ByColumn(ByVal LeftArray As Variant, ByVal RightArray As Variant) As Variant
    ' @brief Merges two 2D arrays horizontally (side-by-side).
    ' @param LeftArray The first 2D array (placed on the left).
    ' @param RightArray The second 2D array (placed on the right).
    ' @return A new, 1-based 2D array containing the merged data.
    '         The resulting array will have the maximum number of rows from the two inputs.
    '         Returns Empty if inputs are invalid or an error occurs.

    a2D_Merge_ByColumn = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(LeftArray) Or Not IsArray(RightArray) Then Exit Function
    On Error Resume Next
    Dim ubL2 As Long: ubL2 = UBound(LeftArray, 2)
    Dim ubR2 As Long: ubR2 = UBound(RightArray, 2)
    If Err.Number <> 0 Then Exit Function ' One or both are not 2D arrays
    On Error GoTo 0

    ' 2. Calculate dimensions for the new array
    Dim l_lbRow As Long: l_lbRow = LBound(LeftArray, 1)
    Dim l_ubRow As Long: l_ubRow = UBound(LeftArray, 1)
    Dim l_lbCol As Long: l_lbCol = LBound(LeftArray, 2)
    Dim l_ubCol As Long: l_ubCol = UBound(LeftArray, 2)
    Dim r_lbRow As Long: r_lbRow = LBound(RightArray, 1)
    Dim r_ubRow As Long: r_ubRow = UBound(RightArray, 1)
    Dim r_lbCol As Long: r_lbCol = LBound(RightArray, 2)
    Dim r_ubCol As Long: r_ubCol = UBound(RightArray, 2)

    Dim l_numRows As Long: l_numRows = l_ubRow - l_lbRow + 1
    Dim l_numCols As Long: l_numCols = l_ubCol - l_lbCol + 1
    Dim r_numRows As Long: r_numRows = r_ubRow - r_lbRow + 1
    Dim r_numCols As Long: r_numCols = r_ubCol - r_lbCol + 1

    Dim newNumRows As Long: newNumRows = IIf(l_numRows > r_numRows, l_numRows, r_numRows)
    Dim newNumCols As Long: newNumCols = l_numCols + r_numCols

    Dim resultArray As Variant
    ReDim resultArray(1 To newNumRows, 1 To newNumCols)

    ' 3. Copy data into the new array
    Dim r As Long, c As Long
    ' Copy LeftArray
    For r = 1 To l_numRows
        For c = 1 To l_numCols
            resultArray(r, c) = LeftArray(r + l_lbRow - 1, c + l_lbCol - 1)
        Next c
    Next r
    ' Copy RightArray
    For r = 1 To r_numRows
        For c = 1 To r_numCols
            resultArray(r, c + l_numCols) = RightArray(r + r_lbRow - 1, c + r_lbCol - 1)
        Next c
    Next r

    ' 4. Return the result
    a2D_Merge_ByColumn = resultArray
End Function

Public Function a2D_Merge_ByRow(ByVal TopArray As Variant, ByVal BottomArray As Variant) As Variant
    ' @brief Merges two 2D arrays vertically (one below the other).
    ' @param TopArray The first 2D array (placed on top).
    ' @param BottomArray The second 2D array (placed at the bottom).
    ' @return A new, 1-based 2D array containing the merged data.
    '         The resulting array will have the maximum number of columns from the two inputs.
    '         Returns Empty if inputs are invalid or an error occurs.

    a2D_Merge_ByRow = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(TopArray) Or Not IsArray(BottomArray) Then Exit Function
    On Error Resume Next
    Dim ubT2 As Long: ubT2 = UBound(TopArray, 2)
    Dim ubB2 As Long: ubB2 = UBound(BottomArray, 2)
    If Err.Number <> 0 Then Exit Function ' One or both are not 2D arrays
    On Error GoTo 0

    ' 2. Calculate dimensions for the new array
    Dim t_lbRow As Long: t_lbRow = LBound(TopArray, 1)
    Dim t_ubRow As Long: t_ubRow = UBound(TopArray, 1)
    Dim t_lbCol As Long: t_lbCol = LBound(TopArray, 2)
    Dim t_ubCol As Long: t_ubCol = UBound(TopArray, 2)
    Dim b_lbRow As Long: b_lbRow = LBound(BottomArray, 1)
    Dim b_ubRow As Long: b_ubRow = UBound(BottomArray, 1)
    Dim b_lbCol As Long: b_lbCol = LBound(BottomArray, 2)
    Dim b_ubCol As Long: b_ubCol = UBound(BottomArray, 2)

    Dim t_numRows As Long: t_numRows = t_ubRow - t_lbRow + 1
    Dim t_numCols As Long: t_numCols = t_ubCol - t_lbCol + 1
    Dim b_numRows As Long: b_numRows = b_ubRow - b_lbRow + 1
    Dim b_numCols As Long: b_numCols = b_ubCol - b_lbCol + 1

    Dim newNumRows As Long: newNumRows = t_numRows + b_numRows
    Dim newNumCols As Long: newNumCols = IIf(t_numCols > b_numCols, t_numCols, b_numCols)

    Dim resultArray As Variant
    ReDim resultArray(1 To newNumRows, 1 To newNumCols)

    ' 3. Copy data into the new array
    Dim r As Long, c As Long
    ' Copy TopArray
    For r = 1 To t_numRows
        For c = 1 To t_numCols
            resultArray(r, c) = TopArray(r + t_lbRow - 1, c + t_lbCol - 1)
        Next c
    Next r
    ' Copy BottomArray
    For r = 1 To b_numRows
        For c = 1 To b_numCols
            resultArray(r + t_numRows, c) = BottomArray(r + b_lbRow - 1, c + b_lbCol - 1)
        Next c
    Next r

    ' 4. Return the result
    a2D_Merge_ByRow = resultArray
End Function

Public Function a2D_Insert(ByVal SourceArray As Variant, ByVal DestinationArray As Variant, Optional ByVal StartRow As Long = -1, Optional ByVal StartCol As Long = -1) As Variant
    ' @brief Overlays a source array onto a destination array at a specified starting position.
    ' @param SourceArray The 2D array containing the data to insert.
    ' @param DestinationArray The 2D array that will receive the data.
    ' @param StartRow (Optional) The row index in the destination array where the top-left of the source array will be placed. Defaults to the destination's lower bound.
    ' @param StartCol (Optional) The column index in the destination array where the top-left of the source array will be placed. Defaults to the destination's lower bound.
    ' @return A new 2D array with the source data overlaid.
    '         Returns Empty if inputs are invalid or an error occurs.
    ' @details This function does not resize the destination array. Data from the source that would fall outside
    '          the bounds of the destination array is ignored.

    a2D_Insert = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(SourceArray) Or Not IsArray(DestinationArray) Then Exit Function
    On Error Resume Next
    Dim ubS2 As Long: ubS2 = UBound(SourceArray, 2)
    Dim ubD2 As Long: ubD2 = UBound(DestinationArray, 2)
    If Err.Number <> 0 Then Exit Function ' One or both are not 2D arrays
    On Error GoTo 0

    ' 2. Set up dimensions and starting positions
    Dim d_lbRow As Long: d_lbRow = LBound(DestinationArray, 1)
    Dim d_ubRow As Long: d_ubRow = UBound(DestinationArray, 1)
    Dim d_lbCol As Long: d_lbCol = LBound(DestinationArray, 2)
    Dim d_ubCol As Long: d_ubCol = UBound(DestinationArray, 2)

    Dim s_lbRow As Long: s_lbRow = LBound(SourceArray, 1)
    Dim s_ubRow As Long: s_ubRow = UBound(SourceArray, 1)
    Dim s_lbCol As Long: s_lbCol = LBound(SourceArray, 2)
    Dim s_ubCol As Long: s_ubCol = UBound(SourceArray, 2)

    ' Default start positions to the destination's lower bounds if not provided or invalid
    If StartRow < d_lbRow Or StartRow > d_ubRow Then StartRow = d_lbRow
    If StartCol < d_lbCol Or StartCol > d_ubCol Then StartCol = d_lbCol

    ' 3. Create a copy of the destination array to modify
    Dim resultArray As Variant
    resultArray = DestinationArray

    ' 4. Loop through the source array and copy data to the result array
    Dim srcRow As Long, srcCol As Long
    Dim destRow As Long, destCol As Long

    For srcRow = s_lbRow To s_ubRow
        For srcCol = s_lbCol To s_ubCol
            ' Calculate the destination coordinates
            destRow = StartRow + (srcRow - s_lbRow)
            destCol = StartCol + (srcCol - s_lbCol)

            ' Check if the destination coordinates are within the bounds of the result array
            If destRow >= d_lbRow And destRow <= d_ubRow And destCol >= d_lbCol And destCol <= d_ubCol Then
                resultArray(destRow, destCol) = SourceArray(srcRow, srcCol)
            End If
        Next srcCol
    Next srcRow

    ' 5. Return the result
    a2D_Insert = resultArray
End Function

Public Function a1D_To_Linea2D(a1D As Variant) As Variant
    Dim i           As Long
    Dim a2D         As Variant
    
    ReDim a2D(1 To 1, LBound(a1D) To UBound(a1D))
    For i = LBound(a1D) To UBound(a1D)
        a2D(1, i) = Src1DArray(i)
    Next
    a1D_To_Linea2D = a1D_To_Linea2D
End Function

Public Function a1D_To_Columna2D(a1D As Variant) As Variant
    Dim i           As Long
    Dim a2D         As Variant
    
    ReDim a2D(LBound(a1D) To UBound(a1D), 1 To 1)
    For i = LBound(a1D) To UBound(a1D)
        a2D(i, 1) = a1D(i)
    Next
    a1D_To_Columna2D = a2D
End Function

Public Function a2D_Clear_Values(Arr2D As Variant) As Variant
    ' /!\ to validate
    ReDim a2D_Clear_Values(LBound(Arr2D, 1) To UBound(Arr2D, 1), LBound(Arr2D, 2) To UBound(Arr2D, 2))
End Function

Public Function a2D_Clear_RowValues(Arr2D As Variant, RowToClear As Long) As Variant
    ' /!\ to validate
    Dim YIndex As Long
    Dim resultArray() As Variant
    resultArray = Arr2D
    
    ' BOUNDARY.
    If Not (LBound(resultArray, 1) <= RowToClear <= UBound(resultArray, 1)) Then Debug.Print "ERROR ! a2D_Clear_RowValues() : RowToClear = " & RowToClear & " is out of Arr2D .row range": GoTo ErrorFunct
    
    ' ROW CLEAR.
    For YIndex = LBound(resultArray, 2) To UBound(resultArray, 2)
        resultArray(RowToClear, YIndex) = Nothing
    Next YIndex
    
    ' RESULT.
    a2D_Clear_RowValues = resultArray
    Exit Function
ErrorFunct:
    a2D_Clear_RowValues = resultArray
End Function

Public Function a2D_Clear_ColumnValues(Arr2D As Variant, ColumnToClear As Long) As Variant
    ' /!\ to validate
    Dim XIndex As Long
    Dim resultArray() As Variant
    resultArray = Arr2D
    
    ' BOUNDARY.
    If Not (LBound(resultArray, 2) <= ColumnToClear <= UBound(resultArray, 2)) Then Debug.Print "ERROR ! a2D_Clear_ColumnValues() : ColumnToClear = " & ColumnToClear & " is out of Arr2D .column range": GoTo ErrorFunct
    
    ' ROW CLEAR.
    For XIndex = LBound(resultArray, 1) To UBound(resultArray, 1)
        resultArray(XIndex, ColumnToClear) = Nothing
    Next XIndex
    
    ' RESULT.
    a2D_Clear_ColumnValues = resultArray
    Exit Function
ErrorFunct:
    a2D_Clear_ColumnValues = resultArray
End Function

Public Function a2D_Clear_Row(Arr2D As Variant, RowToClear As Long) As Variant
    ' /!\ to validate
    Dim XIndex As Long
    Dim YIndex As Long
    Dim XOffset As Long
    Dim resultArray() As Variant
    ReDim resultArray(LBound(Arr2D, 1) To UBound(Arr2D, 1) - 1, LBound(Arr2D, 2) To UBound(Arr2D, 2))
    
    ' ERROR.
    If Not (LBound(resultArray, 1) <= RowToClear <= UBound(resultArray, 1)) Then Debug.Print "ERROR ! ClearRow_2DArray() : RowToClear = " & RowToClear & " is out of Arr2D .row range": GoTo ErrorFunct
    
    ' ROW CLEAR.
    XOffset = 0
    For XIndex = LBound(resultArray, 1) To UBound(resultArray, 1)
        If XIndex = RowToClear Then XOffset = XOffset + 1
        For YIndex = LBound(resultArray, 2) To UBound(resultArray, 2)
             resultArray(XIndex, YIndex) = Arr2D(XIndex + XOffset, YIndex)
        Next YIndex
    Next XIndex
    
    ' RESULT.
    a2D_Clear_Row = resultArray
    Exit Function
ErrorFunct:
    a2D_Clear_Row = resultArray
End Function

Public Function a2D_Clear_Column(Arr2D As Variant, ColumnToClear As Long) As Variant
    ' /!\ to validate
    Dim XIndex As Long
    Dim YIndex As Long
    Dim YOffset As Long
    Dim resultArray() As Variant
    ReDim resultArray(LBound(Arr2D, 1) To UBound(Arr2D, 1), LBound(Arr2D, 2) To UBound(Arr2D, 2) - 1)
    
    ' ERROR.
    If Not (LBound(resultArray, 1) <= RowToClear <= UBound(resultArray, 1)) Then Debug.Print "ERROR ! ClearRow_2DArray() : ColumnToClear = " & ColumnToClear & " is out of Arr2D .column range": GoTo ErrorFunct
    
    ' ROW CLEAR.
    XOffset = 0
    For XIndex = LBound(resultArray, 2) To UBound(resultArray, 2)
        For YIndex = LBound(resultArray, 2) To UBound(resultArray, 2)
            If YIndex = RowToClear Then YOffset = YOffset + 1
            resultArray(XIndex, YIndex) = Arr2D(XIndex, YIndex + YOffset)
        Next YIndex
    Next XIndex
    
    ' RESULT.
    a2D_Clear_Column = resultArray
    Exit Function
ErrorFunct:
    a2D_Clear_Column = resultArray
End Function

Public Function a2D_Fill_With_Value(ByVal Arr2D As Variant, Optional ByVal FillValue As Variant = "") As Variant
    ' @brief Fills every element of a 2D array with a specified value.
    ' @param Arr2D The source 2D array.
    ' @param FillValue (Optional) The value to fill every element of the array with. Defaults to an empty string.
    ' @return A new 2D array of the same dimensions, with every element set to FillValue.
    '         Returns Empty if the input is not a valid 2D array.

    a2D_Fill_With_Value = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(Arr2D) Then Exit Function
    On Error Resume Next
    Dim ub2 As Long: ub2 = UBound(Arr2D, 2)
    If Err.Number <> 0 Then Exit Function ' Not a 2D array
    On Error GoTo 0

    ' 2. Create the result array
    Dim r As Long, c As Long
    Dim lbRow As Long: lbRow = LBound(Arr2D, 1)
    Dim ubRow As Long: ubRow = UBound(Arr2D, 1)
    Dim lbCol As Long: lbCol = LBound(Arr2D, 2)
    Dim ubCol As Long: ubCol = UBound(Arr2D, 2)
    Dim resultArray As Variant
    ReDim resultArray(lbRow To ubRow, lbCol To ubCol)

    ' 3. Fill the array
    For r = lbRow To ubRow
        For c = lbCol To ubCol
            resultArray(r, c) = FillValue
        Next c
    Next r

    ' 4. Return the result
    a2D_Fill_With_Value = resultArray
End Function

Public Function Csv_To_Range(sCsv As String, rng As Range, Optional Delimiter As String = ",", Optional QuoteChar As String = "'") As Boolean
    '* @brief Converts a CSV string to a range in Excel.
    '* @param sCsv The CSV string to convert.
    '* @param rng The destination range where the CSV data should be placed.
    '* @param Delimiter The delimiter used to separate CSV values (default is ",").
    '* @param QuoteChar The character used to enclose CSV values (default is "'").
    '* @return True if the conversion is successful, False otherwise.
    On Error GoTo ifError
    Dim Arr2D       As Variant
    
    Arr2D = Csv_To_a2D(sCsv, Delimiter, QuoteChar)
    If Not IsError(Arr2D) Then Csv_To_Range = ArrayX.a2D_Write(Arr2D, rng)
    Exit Function
ifError:
    Csv_To_Range = False
End Function

Public Function Csv_To_a2D(sCsv As String, Optional Delimiter As String = ",", Optional QuoteChar As String = "'") As Variant
    '*******************************************************************************
    '** Public Function: Csv_To_a2D
    '** Description: Converts a CSV formatted string to a 2D array.
    '**
    '** @param sCsv (String) - The CSV formatted string to be converted.
    '** @param Delimiter (String, optional) - The character used to separate values in CSV.
    '**        Default is ",".
    '** @param QuoteChar (String, optional) - The character used for quoting values in CSV.
    '**        Default is "'". for ", use chr(34)
    '**
    '** @return (Variant) - The resulting 2D array.
    '*******************************************************************************
'    On Error GoTo ifError
    Dim i               As Long
    Dim j               As Long
    Dim LBRows          As Long
    Dim UBRows          As Long
    Dim LBCols          As Long
    Dim UBCols          As Long
    Dim UBaField        As Long
    Dim Arr2D()         As String
    Dim aLine           As Variant
    Dim line            As Variant
    Dim aField          As Variant
    
    ' Replace any occurrences of quoted line breaks with plain line breaks
    sCsv = Replace(sCsv, vbCrLf & QuoteChar, vbCrLf)
    sCsv = Replace(sCsv, QuoteChar & vbCrLf, vbCrLf)
    
    ' Remove leading and trailing quotes if present
    If Mid(sCsv, 1, Len(QuoteChar)) = QuoteChar Then sCsv = Mid(sCsv, Len(QuoteChar) + 1)
'    If Right(sCsv, Len(QuoteChar) + 2) = QuoteChar Then sCsv = Left(sCsv, Len(sCsv) - Len(QuoteChar))

    ' Split the CSV string into an array of lines
    ' & Determine the number of rows based on the number of line breaks
    aLine = Split(sCsv, vbCrLf)
    UBRows = UBound(aLine)
    LBRows = LBound(aLine)
    
    ' Determine the number of fields in the first line for initialisation.
    aField = Split(aLine(LBRows), QuoteChar & Delimiter & QuoteChar)
    LBCols = LBound(aField)
    UBCols = UBound(aField)
    If UBCols < LBCols Then UBCols = LBCols
    
    ' Create a 2D array with the determined dimensions
    ReDim Arr2D(LBRows To UBRows, LBCols To UBCols)
    
    ' Populate the 2D array with values from the CSV string
    For i = LBRows To UBRows
        aField = Split(aLine(i), QuoteChar & Delimiter & QuoteChar)
        UBaField = UBound(aField)
        If UBaField > UBCols Then UBCols = UBaField: ReDim Preserve Arr2D(LBRows To UBRows, LBCols To UBCols)
        For j = LBCols To UBaField
            Arr2D(i, j) = CStr(aField(j))
        Next j
    Next i
    
    ' Return the resulting 2D array
    Csv_To_a2D = Arr2D
    Exit Function

ifError:
    Csv_To_a2D = CVErr(2001)
End Function


Public Function a2D_ToCsv(Arr2D As Variant, Optional Delimiter As String = ",", Optional QuoteChar As String = "'") As String
    '*******************************************************************************
    '** Public Function: a2D_ToCsv
    '** Description: Converts a 2D array to a CSV formatted string.
    '**
    '** @param arr2D (Variant) - The 2D array to be converted to CSV.
    '** @param Delimiter (String, optional) - The character used to separate values in CSV.
    '**        Default is ",".
    '** @param QuoteChar (String, optional) - The character used for quoting values in CSV.
    '**        Default is "'".
    '**
    '** @return (String) - The CSV formatted string.
    '*******************************************************************************
    Dim i           As Long
    Dim j           As Long
    Dim LBRows      As Long
    Dim UBRows      As Long
    Dim LBCols      As Long
    Dim UBCols      As Long
    Dim sbCSV       As New StringBuilder

    ' Convert Range in Array and Check if Arr2D is an array
    If TypeName(Arr2D) = "Range" Then Arr2D = Arr2D.value
    If Not IsArray(Arr2D) Then Exit Function
    
    ' Get the lower and upper bounds of the array.
    LBRows = LBound(Arr2D, 1)
    UBRows = UBound(Arr2D, 1)
    LBCols = LBound(Arr2D, 2)
    UBCols = UBound(Arr2D, 2)
    
    ' Loop through the rows and columns of the array.
    For i = LBRows To UBRows
        For j = LBCols To UBCols
            ' Append value to the CSV string, with appropriate delimiter and quotes.
            sbCSV.Append IIf(j > LBCols, Delimiter, "") & QuoteChar & CStr(Arr2D(i, j)) & QuoteChar & IIf(j >= UBCols, vbCrLf, "")
        Next j
    Next i

    ' Return the CSV formatted string.
    a2D_ToCsv = sbCSV.ToString
    Set sbCSV = Nothing
End Function

Sub Selection_To_CsvFile()
    Dim Arr2D           As Variant
    Dim sCsv            As String
    Dim path            As String
    Dim filePath        As String
    Dim fileName        As String
    Dim Delimiter       As String
    Dim QuoteChar       As String
       
    Delimiter = ","
    QuoteChar = Chr(34)
        
    ' R�cup�rer la s�lection en tant que tableau 2D
    Arr2D = a2D_From_Selection()
    If IsError(Arr2D) Then
        MsgBox "Aucune selection valide trouvee.", vbExclamation
        Exit Sub
    End If
    
    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "csv Files (*.csv),*csv")
    If filePath = "" Then Exit Sub
    
    ' Convertir le tableau 2D en cha�ne CSV
    sCsv = a2D_ToCsv(Arr2D, Delimiter, QuoteChar)
    
    ' Enregistrer le fichier CSV
    FileX.OverwriteTxt filePath, sCsv
    
    MsgBox "Fichier CSV cree avec succes : " & filePath, vbInformation
End Sub

Sub Sheet_To_CsvFile()
    Dim Arr2D           As Variant
    Dim sCsv            As String
    Dim path            As String
    Dim filePath        As String
    Dim fileName        As String
    Dim Delimiter       As String
    Dim QuoteChar       As String
       
    Delimiter = ","
    QuoteChar = Chr(34)
    
    
    ' R�cup�rer la s�lection en tant que tableau 2D
    Arr2D = Sheet_To_a2D(ActiveSheet)
    If IsError(Arr2D) Then
        MsgBox "Aucune feuille valide trouvee.", vbExclamation
        Exit Sub
    End If

    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "csv Files (*.csv),*csv")
    If filePath = "" Then Exit Sub
    
    ' Convertir le tableau 2D en cha�ne CSV
    sCsv = ArrayX.a2D_ToCsv(Arr2D, Delimiter, QuoteChar)
    
    ' Enregistrer le fichier CSV
    FileX.OverwriteTxt filePath, sCsv
    
    MsgBox "Fichier CSV cree avec succes : " & filePath, vbInformation
End Sub

Public Function a2D_From_Range(ByVal TargetRange As Range, Optional ByVal ReadProperty As String = "Value") As Variant
    ' @brief Converts a Range object into a 2D array, reading either values or formulas.
    ' @param TargetRange The Range to convert.
    ' @param ReadProperty (Optional) The property to read from cells. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return A 2D Variant array containing the data from the range.
    '         Returns an empty Variant on failure or if the range is invalid.
    ' @details This is the most efficient method to read a range's data into memory.
    '          It also handles the edge case where the range is only a single cell,
    '          ensuring a 2D array is always returned for consistency.
    On Error GoTo ifError
    a2D_From_Range = Empty ' Default return value

    Dim arrData As Variant

    If TargetRange Is Nothing Then
        Debug.Print "a2D_From_Range Error: TargetRange cannot be Nothing."
        Exit Function
    End If

    ' Read the specified property from the range
    Select Case LCase(ReadProperty)
        Case "value", "values": arrData = TargetRange.Value
        Case "formula", "formulas": arrData = TargetRange.Formula
        Case "formular1c1": arrData = TargetRange.FormulaR1C1
        Case Else
            Debug.Print "a2D_From_Range Warning: Invalid 'ReadProperty' specified: '" & ReadProperty & "'. Defaulting to 'Value'."
            arrData = TargetRange.Value
    End Select

    ' Handle the case where the Range is only a single cell, which returns a scalar value.
    If Not IsArray(arrData) Then
        Dim tempValue As Variant: tempValue = arrData ' Store the scalar value/formula
        ReDim arrData(1 To 1, 1 To 1): arrData(1, 1) = tempValue
    End If

    a2D_From_Range = arrData
    Exit Function

ifError:
    Debug.Print "An error occurred in a2D_From_Range. " & vbCrLf & "Error: " & Err.Description
    a2D_From_Range = Empty
End Function

' NOTE: Prefer using directly a2D_FROM_Range() function.
Public Function Range_To_a2D(ByVal TargetRange As Range, Optional ByVal ReadProperty As String = "Value") As Variant
    ' @brief (Alias for a2D_From_Range) Converts a Range object into a 2D array.
    ' @param TargetRange The Range to convert.
    ' @param ReadProperty (Optional) The property to read from cells. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return A 2D Variant array.
    Range_To_a2D = a2D_From_Range(TargetRange, ReadProperty)
End Function

Public Function a2D_From_Selection(Optional ByVal ReadProperty As String = "Value") As Variant
    ' @brief Converts the currently selected Range into a 2D array, reading either values or formulas.
    ' @param ReadProperty (Optional) The property to read from cells. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return A 2D Variant array containing the data from the selection.
    '         Returns Empty if the selection is not a valid range.
    ' @dependencies a2D_From_Range
    a2D_From_Selection = Empty ' Default return value
    If TypeName(Selection) = "Range" Then
        a2D_From_Selection = a2D_From_Range(Selection, ReadProperty)
    Else
        Debug.Print "a2D_From_Selection Error: The current selection is not a Range."
    End If
End Function

' NOTE: Prefer using directly a2D_FROM_Selection() function.
Public Function Selection_To_a2D(Optional ByVal ReadProperty As String = "Value") As Variant
    ' @brief (Alias for a2D_From_Selection) Converts the currently selected Range into a 2D array.
    ' @param ReadProperty (Optional) The property to read from cells. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return A 2D Variant array.
    Selection_To_a2D = a2D_From_Selection(ReadProperty)
End Function

Public Function a2D_From_Sheet(ByVal TargetSheet As Worksheet, Optional ByVal ReadProperty As String = "Value") As Variant
    ' @brief Converts the used range of a worksheet into a 2D array, reading either values or formulas.
    ' @param TargetSheet The worksheet to convert.
    ' @param ReadProperty (Optional) The property to read from cells. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return A 2D Variant array containing the data from the worksheet's used range.
    '         Returns an empty Variant on failure or if the sheet is empty.
    ' @dependencies a2D_From_Range
    a2D_From_Sheet = Empty ' Default return value
    If TargetSheet Is Nothing Then
        Debug.Print "a2D_From_Sheet Error: TargetSheet cannot be Nothing."
        Exit Function
    End If
    If Application.WorksheetFunction.CountA(TargetSheet.Cells) = 0 Then Exit Function ' Return Empty for a blank sheet
    a2D_From_Sheet = a2D_From_Range(TargetSheet.UsedRange, ReadProperty)
End Function

' NOTE: Prefer using directly a2D_FROM_Sheet() function.
Public Function Sheet_To_a2D(ByVal TargetSheet As Worksheet, Optional ByVal ReadProperty As String = "Value") As Variant
    ' @brief (Alias for a2D_From_Sheet) Converts the used range of a worksheet into a 2D array.
    ' @param TargetSheet The worksheet to convert.
    ' @param ReadProperty (Optional) The property to read from cells. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return A 2D Variant array.
    Sheet_To_a2D = a2D_From_Sheet(TargetSheet, ReadProperty)
End Function

Public Function a2D_To_Json(ByVal Arr2D As Variant, Optional ByVal Headers As Variant) As String
    ' @brief Converts a 2D array into a JSON string (an array of objects).
    ' @param Arr2D The source 2D array.
    ' @param Headers (Optional) A 1D array of strings to use as JSON object keys.
    '        If omitted, the first row of Arr2D is used as the headers.
    ' @return A string containing the data in JSON format.
    '         Returns an empty string if the input is invalid.
    ' @details Each row in the array is converted to a JSON object.
    '          Numeric and Boolean values are preserved. Strings are properly escaped.
    
    a2D_To_Json = "" ' Default return value

    ' 1. Input Validation
    If Not IsArray(Arr2D) Then Exit Function
    On Error Resume Next
    Dim ub2 As Long: ub2 = UBound(Arr2D, 2)
    If Err.Number <> 0 Then Exit Function ' Not a 2D array
    On Error GoTo 0

    ' 2. Get dimensions and header info
    Dim r As Long, c As Long
    Dim lbRow As Long, ubRow As Long, lbCol As Long, ubCol As Long
    Dim dataStartRow As Long, tempHeaders As Variant

    lbRow = LBound(Arr2D, 1): ubRow = UBound(Arr2D, 1)
    lbCol = LBound(Arr2D, 2): ubCol = UBound(Arr2D, 2)

    If ubRow < lbRow Then Exit Function ' Empty array

    ' Determine headers
    If IsMissing(Headers) Then
        ReDim tempHeaders(lbCol To ubCol)
        For c = lbCol To ubCol: tempHeaders(c) = CStr(Arr2D(lbRow, c)): Next c
        dataStartRow = lbRow + 1
    Else
        If Not IsArray(Headers) Then Exit Function ' Invalid Headers argument
        tempHeaders = Headers
        dataStartRow = lbRow
    End If
    
    If dataStartRow > ubRow Then Exit Function ' No data rows to process

    ' 3. Build the JSON string
    Dim numDataRows As Long: numDataRows = ubRow - dataStartRow + 1
    Dim jsonObjects() As String: ReDim jsonObjects(0 To numDataRows - 1)
    Dim i As Long: i = 0

    For r = dataStartRow To ubRow
        Dim rowPairs() As String: ReDim rowPairs(0 To ubCol - lbCol)
        Dim j As Long: j = 0
        For c = lbCol To ubCol
            Dim key As String: key = CStr(tempHeaders(c))
            rowPairs(j) = """" & key & """:" & JsonFormatValue(Arr2D(r, c))
            j = j + 1
        Next c
        jsonObjects(i) = "{" & Join(rowPairs, ",") & "}"
        i = i + 1
    Next r
    
    ' 4. Assemble the final JSON array string
    a2D_To_Json = "[" & Join(jsonObjects, ",") & "]"
End Function







'/////////////////////////////////////////////////////////////////////////////////////////////////'
'                       _____ ____       _                                                        '
'                      |___ /|  _ \     / \   _ __ _ __ __ _ _   _                                '
'                        |_ \| | | |   / _ \ | '__| '__/ _` | | | |                               '
'                       ___) | |_| |  / ___ \| |  | | | (_| | |_| |                               '
'                      |____/|____/  /_/   \_\_|  |_|  \__,_|\__, |                               '
'                                                            |___/                                '
'/////////////////////////////////////////////////////////////////////////////////////////////////'






Public Function a3D_Create(ByVal NumSlices As Long, ByVal NumRows As Long, ByVal NumCols As Long, Optional ByVal FillValue As Variant = vbNullString) As Variant
    ' @brief Creates a new 3D array of a specified size and initializes all its elements to a given value.
    ' @param NumSlices The number of slices (first dimension) for the new array. Must be greater than 0.
    ' @param NumRows The number of rows (second dimension) for the new array. Must be greater than 0.
    ' @param NumCols The number of columns (third dimension) for the new array. Must be greater than 0.
    ' @param FillValue (Optional) The value to fill every element of the array with. Defaults to an empty string.
    ' @return A 1-based 3D Variant array.
    '         Returns Empty if any dimension is less than 1.

    a3D_Create = Empty ' Default return value

    ' 1. Input Validation
    If NumSlices < 1 Or NumRows < 1 Or NumCols < 1 Then Exit Function

    ' 2. Create and initialize the array
    Dim resultArray As Variant
    ReDim resultArray(1 To NumSlices, 1 To NumRows, 1 To NumCols)

    Dim s As Long, r As Long, c As Long
    For s = 1 To NumSlices
        For r = 1 To NumRows
            For c = 1 To NumCols
                resultArray(s, r, c) = FillValue
            Next c
        Next r
    Next s

    ' 3. Return the result
    a3D_Create = resultArray
End Function

Public Function a3D_Write_Slice(ByVal Arr3D As Variant, ByVal SliceIndex As Long, ByVal TopLeftCell As Range, Optional ByVal WriteAs As String = "Value") As Boolean
    ' @brief Writes a specific 2D "slice" from a 3D array to a worksheet, as values or formulas.
    ' @param Arr3D The source 3D array.
    ' @param SliceIndex The index of the slice to extract from the first dimension.
    ' @param TopLeftCell The top-left cell of the destination range.
    ' @param WriteAs (Optional) The property to write. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return True if the write operation was successful, False otherwise.
    ' @dependencies a3D_To_a2D, a2D_Write
    On Error GoTo ifError
    a3D_Write_Slice = False ' Default return value

    ' 1. Input Validation
    If TopLeftCell Is Nothing Then
        Debug.Print "a3D_Write_Slice Error: TopLeftCell cannot be Nothing."
        Exit Function
    End If

    ' Check if it's a 3D array
    if Not Is_a3D(Arr3D) Then 
        Debug.Print "a3D_Write_Slice Error: Input array is not a 2D array."
        Exit Function
    End If

    ' 2. Extract the 2D slice from the 3D array
    Dim slice2D As Variant
    slice2D = a3D_To_a2D(Arr3D, SliceIndex)
    If Not IsArray(slice2D) Then
        Debug.Print "a3D_Write_Slice Error: Failed to extract slice " & SliceIndex & " from the 3D array."
        Exit Function
    End If

    ' 3. Write the extracted slice to the worksheet
    a3D_Write_Slice = a2D_Write(slice2D, TopLeftCell, WriteAs)
    Exit Function

ifError:
    Debug.Print "An error occurred in a3D_Write_Slice on worksheet '" & TopLeftCell.Parent.Name & "'. " & vbCrLf & "Error: " & Err.Description
End Function

Public Function a3D_Write(ByVal Arr3D As Variant, ByVal TargetWorkbook As Workbook, Optional ByVal SheetNames As Variant, Optional ByVal TopLeftAddress As String = "A1", Optional ByVal WriteAs As String = "Value") As Boolean
    ' @brief Writes each slice of a 3D array to a new, dedicated worksheet in a target workbook.
    ' @param Arr3D The source 3D array.
    ' @param TargetWorkbook The workbook where the new worksheets will be created.
    ' @param SheetNames (Optional) An array of names for the new worksheets. If omitted, default names ("Slice 1", "Slice 2", etc.) are used.
    ' @param TopLeftAddress (Optional) The A1-style address where the data will be written on each new sheet. Defaults to "A1".
    ' @param WriteAs (Optional) The property to write. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return True if all slices were written successfully, False otherwise.
    ' @dependencies WorkX.Create_Worksheet, a3D_Write_Slice
    On Error GoTo ifError
    a3D_Write = False ' Default return value

    ' 1. Input Validation
    If TargetWorkbook Is Nothing Then
        Debug.Print "a3D_Write Error: TargetWorkbook cannot be Nothing."
        Exit Function
    End If

    ' Check if it's a 3D array
    if Not Is_a3D(Arr3D) Then 
        Debug.Print "a3D_Write Error: Input array is not a 2D array."
        Exit Function
    End If

    ' 2. Loop through each slice and write to a new sheet
    Dim s As Long
    Dim sheetName As String
    Dim ws As Worksheet
    Dim overallSuccess As Boolean: overallSuccess = True
    Dim useSheetNamesArray As Boolean: useSheetNamesArray = IsArray(SheetNames)

    For s = LBound(Arr3D, 1) To UBound(Arr3D, 1)
        ' Determine the name for the new worksheet
        If useSheetNamesArray Then
            If s >= LBound(SheetNames) And s <= UBound(SheetNames) Then
                sheetName = CStr(SheetNames(s))
            Else
                sheetName = "Slice " & s ' Fallback if SheetNames array is too small
            End If
        Else
            sheetName = "Slice " & s
        End If

        ' Create the worksheet using the function from the WorkX module
        Set ws = WorkX.Create_Worksheet(sheetName, TargetWorkbook)

        If ws Is Nothing Then
            Debug.Print "a3D_Write Error: Failed to create worksheet '" & sheetName & "'."
            overallSuccess = False
        Else
            ' Write the slice to the new sheet using the helper function
            If Not a3D_Write_Slice(Arr3D, s, ws.Range(TopLeftAddress), WriteAs) Then
                overallSuccess = False
            End If
        End If
    Next s

    a3D_Write = overallSuccess
    Exit Function

ifError:
    Debug.Print "An unexpected error occurred in a3D_Write. " & vbCrLf & "Error: " & Err.Description
End Function

Public Function a3D_To_a2D(ByVal Arr3D As Variant, ByVal SliceIndex As Long) As Variant
    ' @brief Extracts a 2D "slice" from a 3D array at a specified index of the first dimension.
    ' @param Arr3D The source 3D array.
    ' @param SliceIndex The index of the slice to extract from the first dimension.
    ' @return A 2D Variant array representing the slice.
    '         Returns Empty if the input is not a valid 3D array or if the index is out of bounds.

    a3D_To_a2D = Empty ' Default return value for failure

    ' 1. Input Validation
    If Not IsArray(Arr3D) Then Exit Function

    On Error Resume Next
    Dim ub3 As Long: ub3 = UBound(Arr3D, 3)
    If Err.Number <> 0 Then Exit Function ' Not a 3D array
    On Error GoTo 0

    Dim lb1 As Long: lb1 = LBound(Arr3D, 1)
    Dim ub1 As Long: ub1 = UBound(Arr3D, 1)
    If SliceIndex < lb1 Or SliceIndex > ub1 Then Exit Function ' Index out of bounds

    ' 2. Set up dimensions for the new 2D array
    Dim r As Long, c As Long
    Dim lbRow As Long: lbRow = LBound(Arr3D, 2)
    Dim ubRow As Long: ubRow = UBound(Arr3D, 2)
    Dim lbCol As Long: lbCol = LBound(Arr3D, 3)
    Dim ubCol As Long: ubCol = UBound(Arr3D, 3)

    Dim resultArray As Variant
    ReDim resultArray(lbRow To ubRow, lbCol To ubCol)

    ' 3. Copy the slice data
    For r = lbRow To ubRow
        For c = lbCol To ubCol
            resultArray(r, c) = Arr3D(SliceIndex, r, c)
        Next c
    Next r

    ' 4. Return the result
    a3D_To_a2D = resultArray
End Function

Public Function a3D_replace_string(Arr3D As Variant, string1 As String, string2 As String) As Variant
    Dim i                   As Long
    Dim j                   As Long
    Dim k                   As Long
    Dim xLB1                As Long
    Dim xUB1                As Long
    Dim xLB2                As Long
    Dim xUB2                As Long
    Dim xLB3                As Long
    Dim xUB3                As Long
    Dim aXD                 As Variant
    
    aXD = Arr3D
    xLB1 = LBound(aXD, 1)
    xUB1 = UBound(aXD, 1)
    xLB2 = LBound(aXD, 2)
    xUB2 = UBound(aXD, 2)
    xLB3 = LBound(aXD, 3)
    xUB3 = UBound(aXD, 3)
    
    For i = xLB1 To xUB1
        For j = xLB2 To xUB2
            For k = xLB3 To xUB3
                If TypeName(aXD(i, j, k)) = TypeName(string1) Then aXD(i, j, k) = Replace(aXD(i, j, k), string1, string2)
            Next
        Next
    Next
    
    a3D_replace_string = aXD
End Function


    a3D_replace_string = aXD
End Function

Public Function a3D_math_Clip(Arr3D As Variant, Optional Low As Variant, Optional High As Variant, Optional d1_offset As Long = 0, Optional d2_offset As Long = 0, Optional d3_offset As Long = 0) As Variant
    ' * @brief Returns the clipped value of numeric elements in a 3-dimensional array.
    ' * @param Arr3D The input array to calculate the absolute values.
    ' * @param d1_offset The offset for the first dimension. Default is 0.
    ' * @param d2_offset The offset for the second dimension. Default is 0.
    ' * @param d3_offset The offset for the third dimension. Default is 0.
    ' * @return The array with clipped values of numeric elements.

    Dim i                   As Long
    Dim j                   As Long
    Dim k                   As Long
    Dim xLB1                As Long
    Dim xUB1                As Long
    Dim xLB2                As Long
    Dim xUB2                As Long
    Dim xLB3                As Long
    Dim xUB3                As Long
    Dim aXD                 As Variant

    aXD = Arr3D
    xLB1 = LBound(aXD, 1) + d1_offset
    xUB1 = UBound(aXD, 1)
    xLB2 = LBound(aXD, 2) + d2_offset
    xUB2 = UBound(aXD, 2)
    xLB3 = LBound(aXD, 3) + d3_offset
    xUB3 = UBound(aXD, 3)

    For i = xLB1 To xUB1
        For j = xLB2 To xUB2
            For k = xLB3 To xUB3
                If TypeName(aXD(i, j, k)) = TypeName(Low) Then
                    If aXD(i, j, k) < Low Then aXD(i, j, k) = Low
                End If
                If TypeName(aXD(i, j, k)) = TypeName(High) Then
                    If aXD(i, j, k) > Low Then aXD(i, j, k) = High
                End If
            Next
        Next
    Next

    a3D_math_Clip = aXD
End Function

Public Function a3D_math_Abs(Arr3D As Variant, Optional d1_offset As Long = 0, Optional d2_offset As Long = 0, Optional d3_offset As Long = 0) As Variant
    ' * @brief Returns the absolute value of numeric elements in a 3-dimensional array.
    ' * @param Arr3D The input array to calculate the absolute values.
    ' * @param d1_offset The offset for the first dimension. Default is 0.
    ' * @param d2_offset The offset for the second dimension. Default is 0.
    ' * @param d3_offset The offset for the third dimension. Default is 0.
    ' * @return The array with absolute values of numeric elements.

    Dim i                   As Long
    Dim j                   As Long
    Dim k                   As Long
    Dim xLB1                As Long
    Dim xUB1                As Long
    Dim xLB2                As Long
    Dim xUB2                As Long
    Dim xLB3                As Long
    Dim xUB3                As Long
    Dim aXD                 As Variant
    
    aXD = Arr3D
    xLB1 = LBound(aXD, 1) + d1_offset
    xUB1 = UBound(aXD, 1)
    xLB2 = LBound(aXD, 2) + d2_offset
    xUB2 = UBound(aXD, 2)
    xLB3 = LBound(aXD, 3) + d3_offset
    xUB3 = UBound(aXD, 3)
    
    For i = xLB1 To xUB1
        For j = xLB2 To xUB2
            For k = xLB3 To xUB3
                If IsNumeric(aXD(i, j, k)) Then
                    aXD(i, j, k) = Abs(aXD(i, j, k))
                End If
            Next
        Next
    Next
    
    a3D_math_Abs = aXD
End Function







'/////////////////////////////////////////////////////////////////////////////////////////////////'
'                       _  _   ____       _                                                       '
'                      | || | |  _ \     / \   _ __ _ __ __ _ _   _                               '
'                      | || |_| | | |   / _ \ | '__| '__/ _` | | | |                              '
'                      |__   _| |_| |  / ___ \| |  | | | (_| | |_| |                              '
'                         |_| |____/  /_/   \_\_|  |_|  \__,_|\__, |                              '
'                                                             |___/                               '
'/////////////////////////////////////////////////////////////////////////////////////////////////'





Public Function a4D_Write(ByVal Arr4D As Variant, ByVal BasePath As String, Optional ByVal WorkbookNames As Variant, Optional ByVal SheetNames As Variant, Optional ByVal TopLeftAddress As String = "A1", Optional ByVal WriteAs As String = "Value", Optional ByVal AutoClose As Boolean = True) As Boolean
    ' @brief Writes each 3D slice of a 4D array to a new, dedicated workbook, which is then saved.
    ' @param Arr4D The source 4D array.
    ' @param BasePath The directory where the new workbooks will be saved.
    ' @param WorkbookNames (Optional) A 1D array of names for the new workbooks. If omitted, default names ("Workbook 1", etc.) are used.
    ' @param SheetNames (Optional) A 2D array of names for the worksheets, where rows correspond to workbooks and columns to sheets. If omitted, default names ("Slice 1", etc.) are used.
    ' @param TopLeftAddress (Optional) The A1-style address where data will be written on each sheet. Defaults to "A1".
    ' @param WriteAs (Optional) The property to write. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @param AutoClose (Optional) If True, the created workbooks will be closed after saving. Defaults to True.
    ' @return True if all workbooks and slices were written successfully, False otherwise.
    ' @dependencies WorkX.Create_Workbook, WorkX.Save_Workbook_As, FileX.Combine_Paths, a4D_To_a3D, a3D_Write, a2D_Get_Row
    On Error GoTo ifError
    a4D_Write = False ' Default return value

    ' 1. Input Validation
    If Trim(BasePath) = "" Then Debug.Print "a4D_Write Error: BasePath cannot be empty.": Exit Function
    On Error Resume Next
    Dim ub4 As Long: ub4 = UBound(Arr4D, 4)
    If Err.Number <> 0 Then Debug.Print "a4D_Write Error: Input data is not a valid 4D array.": On Error GoTo ifError: Exit Function
    On Error GoTo ifError

    ' 2. Loop through each 3D slice (workbook) and write it
    Dim w As Long, wbkName As String, savePath As String
    Dim wbk As Workbook, arr3D As Variant, currentSheetNames As Variant
    Dim overallSuccess As Boolean: overallSuccess = True
    Dim useWbkNames As Boolean: useWbkNames = IsArray(WorkbookNames)
    Dim useSheetNames As Boolean: useSheetNames = IsArray(SheetNames)

    For w = LBound(Arr4D, 1) To UBound(Arr4D, 1)
        ' Determine workbook name
        If useWbkNames And w >= LBound(WorkbookNames) And w <= UBound(WorkbookNames) Then wbkName = CStr(WorkbookNames(w)) Else wbkName = "Workbook " & w

        ' Create the new workbook
        Set wbk = WorkX.Create_Workbook(False) ' Create hidden
        If wbk Is Nothing Then overallSuccess = False: GoTo NextWorkbook

        ' Extract the 3D slice and sheet names for this workbook
        arr3D = a4D_To_a3D(Arr4D, w)
        If useSheetNames Then currentSheetNames = a2D_Get_Row(SheetNames, w, True) Else Erase currentSheetNames

        ' Write the 3D data to the new workbook and save it
        If Not a3D_Write(arr3D, wbk, currentSheetNames, TopLeftAddress, WriteAs) Then overallSuccess = False
        savePath = FileX.Combine_Paths(BasePath, wbkName & ".xlsx")
        If Not WorkX.Save_Workbook_As(wbk, savePath, xlOpenXMLWorkbook, AutoClose) Then
            overallSuccess = False
            If Not wbk.Saved Then wbk.Close False
        End If
NextWorkbook:
    Next w

    a4D_Write = overallSuccess
    Exit Function

ifError:
    Debug.Print "An unexpected error occurred in a4D_Write. " & vbCrLf & "Error: " & Err.Description
End Function


Public Function a4D_To_a3D(ByVal Arr4D As Variant, ByVal SliceIndex As Long) As Variant
    ' @brief Extracts a 3D "slice" from a 4D array at a specified index of the first dimension.
    ' @param Arr4D The source 4D array.
    ' @param SliceIndex The index of the slice to extract from the first dimension.
    ' @return A 3D Variant array representing the slice.
    '         Returns Empty if the input is not a valid 4D array or if the index is out of bounds.

    a4D_To_a3D = Empty ' Default return value for failure

    ' 1. Input Validation
    If Not IsArray(Arr4D) Then Exit Function

    On Error Resume Next
    Dim ub4 As Long: ub4 = UBound(Arr4D, 4)
    If Err.Number <> 0 Then Exit Function ' Not a 4D array
    On Error GoTo 0

    Dim lb1 As Long: lb1 = LBound(Arr4D, 1)
    Dim ub1 As Long: ub1 = UBound(Arr4D, 1)
    If SliceIndex < lb1 Or SliceIndex > ub1 Then Exit Function ' Index out of bounds

    ' 2. Set up dimensions for the new 3D array
    Dim s As Long, r As Long, c As Long
    Dim lbSlice As Long: lbSlice = LBound(Arr4D, 2)
    Dim ubSlice As Long: ubSlice = UBound(Arr4D, 2)
    Dim lbRow As Long: lbRow = LBound(Arr4D, 3)
    Dim ubRow As Long: ubRow = UBound(Arr4D, 3)
    Dim lbCol As Long: lbCol = LBound(Arr4D, 4)
    Dim ubCol As Long: ubCol = UBound(Arr4D, 4)

    Dim resultArray As Variant
    ReDim resultArray(lbSlice To ubSlice, lbRow To ubRow, lbCol To ubCol)

    ' 3. Copy the slice data
    For s = lbSlice To ubSlice
        For r = lbRow To ubRow
            For c = lbCol To ubCol
                resultArray(s, r, c) = Arr4D(SliceIndex, s, r, c)
            Next c
        Next r
    Next s

    ' 4. Return the result
    a4D_To_a3D = resultArray
End Function