Attribute VB_Name = "LANG_MOD"
Option Explicit

Public Function CStrToCDbl(str_value As String, Optional DecimalSeparator As String = "Auto") As Double
    Dim str_tmp As String
    
    ' Errors
    If str_value = "" Or str_value = "-" Then CStrToCDbl = 0: Exit Function
    
    ' Decimal Separator.
    If DecimalSeparator = "Auto" Then DecimalSeparator = IIf(CStr(val(1 / 10)) Like "*.*", ".", ",")
    
    ' Replace
    str_tmp = Replace(str_value, "a", "E-18")
    str_tmp = Replace(str_tmp, "f", "E-15")
    str_tmp = Replace(str_tmp, "p", "E-12")
    str_tmp = Replace(str_tmp, "n", "E-9")
    str_tmp = Replace(str_tmp, "u", "E-6")
    str_tmp = Replace(str_tmp, "m", "E-3")
    str_tmp = Replace(str_tmp, "k", "E3")
    str_tmp = Replace(str_tmp, "K", "E3")
    str_tmp = Replace(str_tmp, "M", "E6")
    str_tmp = Replace(str_tmp, "G", "E9")
    'Remove Unit A/V
    str_tmp = Replace(str_tmp, "A", "")
    str_tmp = Trim(Replace(str_tmp, "V", ""))
    
    ' Convert
    If DecimalSeparator = "." Then CStrToCDbl = CDbl(Replace(str_tmp, ",", ".")) Else CStrToCDbl = CDbl(Replace(str_tmp, ".", ","))
End Function

Public Function CStrToInt(str_value As String, Optional DecimalSeparator As String = "Auto") As Integer
    If str_value = "" Or str_value = "-" Then CStrToInt = 0: Exit Function
    If DecimalSeparator = "Auto" Then DecimalSeparator = IIf(CStr(val(1 / 10)) Like "*.*", ".", ",")    ' Decimal Separator.
    If DecimalSeparator = "." Then CStrToInt = CInt(CStr(Replace(str_value, ",", "."))) Else CStrToInt = CInt(CStr(Replace(str_value, ".", ",")))
End Function

Public Function CStrToLong(str_value As String, Optional DecimalSeparator As String = "Auto") As Long
    If str_value = "" Or str_value = "-" Then CStrToLong = 0: Exit Function
    If DecimalSeparator = "Auto" Then DecimalSeparator = IIf(CStr(val(1 / 10)) Like "*.*", ".", ",")    ' Decimal Separator.
    If DecimalSeparator = "." Then CStrToLong = val(CStr(Replace(str_value, ",", "."))) Else CStrToLong = val(CStr(Replace(str_value, ".", ",")))
End Function

Public Function CCDblToStr(cdbl_value As Double) As String
    CCDblToStr = Replace(CStr(CDbl(cdbl_value)), ",", ".")
End Function

Public Function CIntToStr(cint_value As Integer) As String
    CIntToStr = Replace(CStr(CInt(cint_value)), ",", ".")
End Function

Public Function CStrToStr(str_value As String) As String
    CStrToStr = Replace(CStr(str_value), ",", ".")
End Function

Function ConvertSecondsToDate(seconds As Double) As Date
    ConvertSecondsToDate = CDate(seconds / 86400)
End Function

Public Function AlphaToNumeric(Alpha As String) As Integer
    AlphaToNumeric = InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", Alpha, vbTextCompare)
End Function

Public Function NumericToAlpha(index As Integer) As String
    Dim alphabet As String
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    If index >= 1 And index <= Len(alphabet) Then NumericToAlpha = Mid(alphabet, index, 1) Else NumericToAlpha = ""
End Function

Public Function wks_AddressToColumn(CellAddress As String) As Long
    wks_AddressToColumn = ThisWorkbook.Sheets(1).Range(CellAddress & 1).column
End Function

Public Function wks_AddressToRow(CellAddress As String) As Long
    wks_AddressToRow = ThisWorkbook.Sheets(1).Range(CellAddress).Row
End Function

Public Function wks_ColumnToAddress(ColumnNb As Long) As String
    wks_ColumnToAddress = Split(ThisWorkbook.Sheets(1).Cells(1, ColumnNb).Address, "$")(1)
End Function

Public Function col(x As Long) As String
' FALBANY Function : ' Conversion : Numero -> Lettre de colone.
    col = Split(ThisWorkbook.Sheets(1).Cells(1, x).Address, "$")(1)
    'col = Chr(64 + x)
End Function

Public Function CleanStr(str As String, sCharToClear As String, Optional xLen As Long = -1) As String   ' To deprecate : use CleanResize_String() instead
    CleanStr = CleanResize_String(str, sCharToClear, xLen)
End Function

Public Function CleanResize_String(str As String, sCharToClear As String, Optional xLen As Long = -1, Optional caseSensitive As Boolean = True)
    Dim sOut    As String
    
    sOut = str
    sOut = Clear_CharInString(str, sCharToClear, caseSensitive)
    If xLen > 0 Then sOut = Resize_String(str, xLen)
    
    CleanResize_String = sOut
End Function

Public Function Clear_CharInString(str As String, sCharToClear As String, Optional caseSensitive As Boolean = True) As String
    ' * @brief Removes specified characters from a given string.
    ' * @param Str The input string to be cleaned.
    ' * @param sCharToClean The characters to be removed from the input string.
    ' * @return The cleaned string with specified characters removed.

    Dim i           As Integer
    Dim sCleaned    As String
    
    sCleaned = str
    For i = 1 To Len(sCharToClear)
        sCleaned = Replace(sCleaned, Mid(sCharToClear, i, 1), "", , , IIf(caseSensitive, vbBinaryCompare, vbTextCompare))
    Next
    
    Clear_CharInString = sCleaned
End Function

Public Function Clear_StringsInString(str As String, sStrToClear As String, Optional caseSensitive As Boolean = True) As String
    '* @brief Removes occurrences of specified strings in a main string.
    '* @param Str The main string to be cleaned.
    '* @param sStrToClear The strings to be removed from the main string. The strings should be separated by semicolons (;).
    '* @param caseSensitive Indicates whether the search should be case-sensitive. By default, the search is case-sensitive.
    '* @return The cleaned main string with the occurrences of the specified strings removed.
    
    Dim i               As Integer
    Dim sCleaned        As String
    Dim aStrToClean()   As String
    
    sCleaned = str
    aStrToClean = Split(sStrToClear & ";", ";")
    For i = LBound(aStrToClean) To UBound(aStrToClean) - 1
        sCleaned = Replace(sCleaned, aStrToClean(i), "", , , IIf(caseSensitive, vbBinaryCompare, vbTextCompare))
    Next
    
    Clear_StringsInString = sCleaned
End Function

Public Function Resize_String(str As String, xLen As Long) As String
    '* @brief Reduces the length of a string if it exceeds a specified value.
    '* @param Str The string to be resized.
    '* @param xLen The desired maximum length for the string.
    '* @return The resized string.
    If Len(str) > xLen Then Resize_String = Left(str, xLen) Else Resize_String = str
End Function

Function Count_sOccurrences(string1 As String, string2 As String, Optional caseSensitive As Boolean = False) As Integer
    ' @summary   Counts the occurrences of a substring within a string.
    ' @param     string1         The main string.
    ' @param     string2         The substring to count.
    ' @param     caseSensitive   (Optional) If true, performs a case-sensitive comparison.
    '                            Default is False (case-insensitive).
    ' @return    Returns the number of occurrences of string2 in string1.
    Dim pos As Integer
    Dim count As Integer
    
    pos = 1
    count = 0
    
    Do While pos > 0
        pos = InStr(pos, string1, string2, IIf(caseSensitive, vbBinaryCompare, vbTextCompare))
        If pos > 0 Then
            count = count + 1
            pos = pos + 1
        End If
    Loop
    
    Count_sOccurrences = count
End Function

