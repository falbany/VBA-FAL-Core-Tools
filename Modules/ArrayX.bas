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

Public Function a2D_WriteToSpreadSheet(Arr2D As Variant, Optional TopLeftCellAddress As String = "A1", Optional Workbook_Name As String = "ActiveWorkbook", Optional Worksheet_Name As Variant = "ActiveSheet") As Boolean
' OBSOLETE use a2D_Write() instead
    ' Write Arr2D  to Workbooks(Workbook_Name).Worksheets(wks_name).Range(TopLeftCellAddress...
    On Error GoTo ifError
    If Workbook_Name = "ActiveWorkbook" Then Workbook_Name = ActiveWorkbook.name
    If Worksheet_Name = "ActiveSheet" Then Worksheet_Name = ActiveSheet.name
    Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Range(TopLeftCellAddress & ":" & LANG_MOD.col(Range(TopLeftCellAddress).column + UBound(Arr2D, 2) - LBound(Arr2D, 2)) & (Range(TopLeftCellAddress).Row + UBound(Arr2D, 1) - LBound(Arr2D, 1))).value = Arr2D
    a2D_WriteToSpreadSheet = True
    Exit Function
ifError:
    Debug.Print format(Now, "h:mm:ss    ") & "FUNCTION : [a2D_WriteToSpreadSheet], ERROR"
End Function

Public Function a3D_WriteToSpreadSheet(Arr3D As Variant, firstDimensionIndex As Long, Optional TopLeftCellAddress As String = "A1", Optional Workbook_Name As String = "ActiveWorkbook", Optional wks_name As String = "ActiveSheet") As Boolean
' OBSOLETE use a3D_Write() instead
    ' Convert Arr3D to Arr2D and write to Workbooks(Workbook_Name).Worksheets(wks_name).Range(TopLeftCellAddress...
    On Error GoTo ifError
    a3D_WriteToSpreadSheet = a2D_WriteToSpreadSheet(a3D_To_a2D(Arr3D, firstDimensionIndex), TopLeftCellAddress, Workbook_Name, wks_name)
    Exit Function
ifError:
    Debug.Print format(Now, "h:mm:ss    ") & "FUNCTION : [a3D_WriteToSpreadSheet], ERROR"
End Function

Public Function a2D_Write(Arr2D As Variant, ByRef TopLeftCell As Range) As Boolean
    ' Write Arr2D  to Workbooks(Workbook_Name).Worksheets(wks_name).Range(TopLeftCellAddress...
    On Error GoTo ifError
    
    ' Convert Range in Array and Check if Arr2D is an array
    If TypeName(Arr2D) = "Range" Then Arr2D = Arr2D.value
    If Not IsArray(Arr2D) Then Exit Function
    
    TopLeftCell.parent.Range(TopLeftCell.Address & ":" & LANG_MOD.col(TopLeftCell.column + UBound(Arr2D, 2) - LBound(Arr2D, 2)) & (TopLeftCell.Row + UBound(Arr2D, 1) - LBound(Arr2D, 1))).value = Arr2D
    a2D_Write = True
    Exit Function
ifError:
    Debug.Print format(Now, "h:mm:ss    ") & "FUNCTION : [a2D_WriteToSpreadSheet], ERROR"
End Function

Public Function a3D_Write(Arr3D As Variant, firstDimensionIndex As Long, ByRef TopLeftCell As Range) As Boolean
    ' Convert Arr3D to Arr2D and write to Workbooks(Workbook_Name).Worksheets(wks_name).Range(TopLeftCellAddress...
    On Error GoTo ifError
    a3D_Write = a2D_Write(a3D_To_a2D(Arr3D, firstDimensionIndex), TopLeftCell)
    Exit Function
ifError:
    Debug.Print format(Now, "h:mm:ss    ") & "FUNCTION : [a3D_WriteToSpreadSheet], ERROR"
End Function

Public Function a3D_Write_All(Arr3D As Variant, ByRef TopLeftCell As Range, Optional byColumns As Boolean = True) As Range
    ' Convert Arr3D to Arr2D and write to Workbooks(Workbook_Name).Worksheets(wks_name).Range(TopLeftCellAddress...
    On Error GoTo ifError
    
    
    ' ENHANCEMENT : Use a2d_merge functions before writing operation
    
    
    Dim i                   As Long
    Dim lb1_Arr3D           As Long
    Dim ub1_Arr3D           As Long
    Dim lb2_Arr3D           As Long
    Dim ub2_Arr3D           As Long
    Dim lb3_Arr3D           As Long
    Dim ub3_Arr3D           As Long
    Dim nb2_Arr3D           As Long
    Dim nb3_Arr3D           As Long
    Dim tmp_TopLeftCell     As Range
    Dim wks                 As Worksheet
    Dim Arr2D               As Variant
    
    lb1_Arr3D = LBound(Arr3D, 1)
    ub1_Arr3D = UBound(Arr3D, 1)
    lb2_Arr3D = LBound(Arr3D, 2)
    ub2_Arr3D = UBound(Arr3D, 2)
    lb3_Arr3D = LBound(Arr3D, 3)
    ub3_Arr3D = UBound(Arr3D, 3)
    nb2_Arr3D = ub2_Arr3D - lb2_Arr3D + 1
    nb3_Arr3D = ub3_Arr3D - lb3_Arr3D + 1
    
    Set wks = TopLeftCell.parent
'    Set tmp_TopLeftCell = wks.Cells(TopLeftCell.row, TopLeftCell.column)
'
'    For i = lb1_Arr3D To ub1_Arr3D
'        If Not a2D_Write(a3D_To_a2D(Arr3D, i), tmp_TopLeftCell) Then GoTo ifError
'        If byColumns Then
'            Set tmp_TopLeftCell = wks.Cells(tmp_TopLeftCell.row, tmp_TopLeftCell.column + nb3_Arr3D)
'        Else
'            Set tmp_TopLeftCell = wks.Cells(tmp_TopLeftCell.row + nb2_Arr3D, tmp_TopLeftCell.column)
'        End If
'    Next
'
'    If byColumns Then
'        Set a3D_Write_All = wks.Range(TopLeftCell.Address & ":" & wks.Cells(tmp_TopLeftCell.row + nb2_Arr3D - 1, tmp_TopLeftCell.column).Address)
'    Else
'        Set a3D_Write_All = wks.Range(TopLeftCell.Address & ":" & wks.Cells(tmp_TopLeftCell.row, tmp_TopLeftCell.column + nb3_Arr3D - 1).Address)
'    End If
    
    Arr2D = a3D_To_a2D(Arr3D, lb1_Arr3D)
    For i = lb1_Arr3D + 1 To ub1_Arr3D
        If byColumns Then
            Arr2D = ArrayX.a2D_Merge_ByColumn(Arr2D, a3D_To_a2D(Arr3D, i))
        Else
            Arr2D = ArrayX.a2D_Merge_ByRow(Arr2D, a3D_To_a2D(Arr3D, i))
        End If
    Next
    If Not a2D_Write(Arr2D, TopLeftCell) Then GoTo ifError
    Set a3D_Write_All = wks.Range(TopLeftCell.Address & ":" & wks.Cells(TopLeftCell.Row + UBound(Arr2D, 1) - 1, TopLeftCell.column + UBound(Arr2D, 2) - 1).Address)
    

    Exit Function
ifError:
    Debug.Print format(Now, "h:mm:ss    ") & "FUNCTION : [a3D_WriteToSpreadSheet], ERROR"
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

Public Function a2D_Find_First(ByVal SearchArray As Variant, ByVal WhatToFind As Variant, Optional ByVal LookAt As XlLookAt = xlPart, Optional ByVal MatchCase As Boolean = False, Optional ByVal StartRow As Long = -1) As Variant
    ' @brief Finds the first occurrence of a value within a 2D array and returns its indices.
    ' @param SearchArray The 2D array to search within.
    ' @param WhatToFind The value to search for.
    ' @param LookAt (Optional) xlPart to match a substring, xlWhole to match the entire cell content. Defaults to xlPart.
    ' @param MatchCase (Optional) True for a case-sensitive search, False for case-insensitive. Defaults to False.
    ' @param StartRow (Optional) The row index to begin the search from. Defaults to the array's lower bound.
    ' @return A 1-based, 2-element array [row, col] containing the indices of the found item.
    '         Returns Empty if the value is not found or if the input is not a valid 2D array.
    
    a2D_Find_First = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(SearchArray) Then Exit Function
    On Error Resume Next
    Dim ub2 As Long: ub2 = UBound(SearchArray, 2)
    If Err.Number <> 0 Then Exit Function ' Exit if not a 2D array
    On Error GoTo 0

    ' 2. Set up search parameters
    Dim r As Long, c As Long
    Dim lbRow As Long, ubRow As Long
    Dim lbCol As Long, ubCol As Long

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
                Dim result(1 To 2) As Long
                result(1) = r
                result(2) = c
                a2D_Find_First = result
                Exit Function
            End If
        Next c
    Next r
End Function

# TODO: replace by a2D_Find_First()
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

# TODO: replace by a2D_Get_Column()
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

# TODO: replace by a2D_Create()
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


# TODO: use a2D_Insert_Rows() instead
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

# TODO: use a2D_Slice_Rows() instead
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

Public Function a2D_Fill_With_Value(Arr2D As Variant, Optional vValue As Variant = "") As Variant
    Dim XIndex As Long
    Dim YIndex As Long
    Dim resultArray() As Variant
    
    ReDim resultArray(LBound(Arr2D, 1) To UBound(Arr2D, 1), LBound(Arr2D, 2) To UBound(Arr2D, 2))
    For XIndex = LBound(resultArray, 2) To UBound(resultArray, 2)
        For YIndex = LBound(resultArray, 2) To UBound(resultArray, 2)
            If TypeName(Arr2D(XIndex, YIndex)) = TypeName(vValue) Then resultArray(XIndex, YIndex) = vValue Else resultArray(XIndex, YIndex) = Arr2D
        Next YIndex
    Next XIndex
    ' RESULT.
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

Function a2D_From_Sheet(ws As Worksheet) As Variant
    '* @brief Returns a 2D array filled with the UsedRange sheet values.
    '* @return Variant representing a 2D array
    ' D�finir le range contenant les donn�es
    a2D_From_Sheet = ws.UsedRange.value
End Function

Public Function a2D_From_Range(rng As Range) As Variant
    '* @brief Returns a 2D array filled with the range object values.
    '* @return Variant representing a 2D array
    a2D_From_Range = rng.value
End Function

Public Function a2D_From_Selection() As Variant
    '* @brief Returns a 2D array filled with the selected range object values.
    '* @return Variant representing a 2D array
    Select Case TypeName(Selection)
        Case "Range":  a2D_From_Selection = Selection.value
        Case Else: a2D_From_Selection = CVErr(2001)
    End Select
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
    Arr2D = Selection_To_a2D
    If IsError(Arr2D) Then
        MsgBox "Aucune s�lection valide trouv�e.", vbExclamation
        Exit Sub
    End If
    
    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "csv Files (*.csv),*csv")
    If filePath = "" Then Exit Sub
    
    ' Convertir le tableau 2D en cha�ne CSV
    sCsv = a2D_ToCsv(Arr2D, Delimiter, QuoteChar)
    
    ' Enregistrer le fichier CSV
    FileX.OverwriteTxt filePath, sCsv
    
    MsgBox "Fichier CSV cr�� avec succ�s : " & filePath, vbInformation
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
        MsgBox "Aucune feuille valide trouv�e.", vbExclamation
        Exit Sub
    End If

    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "csv Files (*.csv),*csv")
    If filePath = "" Then Exit Sub
    
    ' Convertir le tableau 2D en cha�ne CSV
    sCsv = ArrayX.a2D_ToCsv(Arr2D, Delimiter, QuoteChar)
    
    ' Enregistrer le fichier CSV
    FileX.OverwriteTxt filePath, sCsv
    
    MsgBox "Fichier CSV cr�� avec succ�s : " & filePath, vbInformation
End Sub

Public Function Selection_To_a2D() As Variant
    '* @brief Returns a 2D array filled with the selected range object values.
    '* @return Variant representing a 2D array
    Selection_To_a2D = ArrayX.a2D_From_Selection()
End Function

Public Function Range_To_a2D(rng As Range) As Variant
    '* @brief Returns a 2D array filled with the range object values.
    '* @return Variant representing a 2D array
    Range_To_a2D = ArrayX.a2D_From_Range(rng)
End Function

Function Sheet_To_a2D(ws As Worksheet) As Variant
    '* @brief Returns a 2D array filled with the UsedRange sheet values.
    '* @return Variant representing a 2D array
    ' D�finir le range contenant les donn�es
    Sheet_To_a2D = ArrayX.a2D_From_Sheet(ws)
End Function

'*******************************************************'
'                   PRIVATE FUNCTIONS                   '
'*******************************************************'

Private Function TrendEst(YValues As Variant, XValues As Variant, Degree As Integer) As Variant()
    ' Renvoie les Coefficient de regression polynomiale.
    ' Degree = 0 : Regression exponentielle.
    ' Degree = 1 : Regression lin�aire.
    ' Degree > 1 : Regression Polynomiale d'ordre Degree.
    
    Select Case Degree
        Case 0: TrendEst = Application.LogEst(YValues, XValues, True, True)      ' A*EXP(B*X) avec B = ln(m) : (1,2) : A /  (1,1) : m / (3,1) : R�.
        Case 1: TrendEst = Application.LinEst(YValues, XValues, True, True)      ' A*X+B avec A = (1,1) / B = (1,2) / R� = (3,1).
        Case 2: TrendEst = Application.LinEst(YValues, Application.Power(XValues, Application.Transpose(Array(1, 2))), True, True)     ' Ax�+Bx+C=Y avec A = (1,1) / B = (1,2) / C = (1,3)/ R� = (3,1).
        Case 3: TrendEst = Application.LinEst(YValues, Application.Power(XValues, Application.Transpose(Array(1, 2, 3))), True, True)
        Case 4: TrendEst = Application.LinEst(YValues, Application.Power(XValues, Application.Transpose(Array(1, 2, 3, 4))), True, True)
        Case 5: TrendEst = Application.LinEst(YValues, Application.Power(XValues, Application.Transpose(Array(1, 2, 3, 4, 5))), True, True)
        Case 6: TrendEst = Application.LinEst(YValues, Application.Power(XValues, Application.Transpose(Array(1, 2, 3, 4, 5, 6))), True, True)
    End Select

End Function
