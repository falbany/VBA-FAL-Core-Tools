Attribute VB_Name = "FalUtils"
' **************************************************************************************
' Module    : FalUtils
' Author    : Florent ALBANY
' Website   :
' Purpose   : Provides miscellaneous helper functions for string manipulation,
'             type conversion, and address/coordinate conversion.
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2025-07-29              Initial Release as FalUtils, refactored from LANG_MOD.
'---------------------------------------------------------------------------------------
' Dependencies: Microsoft VBScript Regular Expressions 5.5 (for Slugify_String),
'               Microsoft Scripting Runtime (for some functions that might be added later)
' **************************************************************************************

Option Explicit

' --- API Declarations for GUID Generation ---
' Source: https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/
#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long
    Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (rguid As GUID, ByVal lpsz As LongPtr, ByVal cchMax As Long) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long
    Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As GUID, ByVal lpsz As String, ByVal cchMax As Long) As Long
#End If

' Type definition for a Globally Unique Identifier
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

' Enum for defining path styles for sanitization
Public Enum PathStyle
    psWindows
    psUnix
End Enum

Public Function Clear_SubChar(ByVal SourceString As String, ByVal CharsToRemove As String, Optional ByVal CaseSensitive As Boolean = True) As String
    ' @brief Removes specified characters from a given string.
    ' @param SourceString The input string to be cleaned.
    ' @param CharsToRemove A string containing all individual characters to be removed.
    ' @param CaseSensitive (Optional) If True, performs a case-sensitive replacement. Defaults to True.
    ' @return The cleaned string with specified characters removed.
    Dim i As Long
    Dim cleanedString As String
    
    cleanedString = SourceString
    For i = 1 To Len(CharsToRemove)
        cleanedString = Replace(cleanedString, Mid(CharsToRemove, i, 1), "", , , IIf(CaseSensitive, vbBinaryCompare, vbTextCompare))
    Next
    
    Clear_SubChar = cleanedString
End Function

Public Function Clear_SubString(ByVal SourceString As String, ByVal SubstringsToRemove As String, Optional ByVal Delimiter As String = ";", Optional ByVal CaseSensitive As Boolean = True) As String
    ' @brief Removes occurrences of specified substrings from a main string.
    ' @param SourceString The main string to be cleaned.
    ' @param SubstringsToRemove The substrings to be removed, separated by the specified delimiter.
    ' @param Delimiter (Optional) The delimiter used to separate the substrings in SubstringsToRemove. Defaults to ";".
    ' @param CaseSensitive (Optional) If True, performs a case-sensitive replacement. Defaults to True.
    ' @return The cleaned main string.
    Dim i As Long
    Dim cleanedString As String
    Dim arrSubstrings() As String
    
    cleanedString = SourceString
    arrSubstrings = Split(SubstringsToRemove, Delimiter)
    
    For i = LBound(arrSubstrings) To UBound(arrSubstrings)
        cleanedString = Replace(cleanedString, arrSubstrings(i), "", , , IIf(CaseSensitive, vbBinaryCompare, vbTextCompare))
    Next
    
    Clear_SubString = cleanedString
End Function

Public Function Resize_String(ByVal SourceString As String, ByVal MaxLength As Long) As String
    ' @brief Truncates a string if its length exceeds a specified maximum.
    ' @param SourceString The string to be resized.
    ' @param MaxLength The desired maximum length for the string.
    ' @return The resized string.
    If Len(SourceString) > MaxLength Then
        Resize_String = Left(SourceString, MaxLength)
    Else
        Resize_String = SourceString
    End If
End Function

Public Function Count_Occurrences(ByVal SourceString As String, ByVal Substring As String, Optional ByVal CaseSensitive As Boolean = False) As Long
    ' @brief Counts the occurrences of a substring within a string.
    ' @param SourceString The main string to search within.
    ' @param Substring The substring to count.
    ' @param CaseSensitive (Optional) If True, performs a case-sensitive comparison. Defaults to False.
    ' @return The number of occurrences of the substring.
    If Len(Substring) = 0 Then Exit Function
    Count_Occurrences = (Len(SourceString) - Len(Replace(SourceString, Substring, "", , , IIf(CaseSensitive, vbBinaryCompare, vbTextCompare)))) / Len(Substring)
End Function

Public Function Convert_String_To_Double(ByVal SourceString As String) As Double
    ' @brief Converts a string with potential engineering suffixes (k, m, u, n, etc.) to a Double.
    ' @param SourceString The string to convert (e.g., "1.5k", "200m", "50u").
    ' @return The numeric value as a Double.
    ' @details This function is locale-aware and uses the system's decimal separator.
    Dim tempStr As String

    If SourceString = "" Or SourceString = "-" Then Convert_String_To_Double = 9999: Exit Function

    ' First, remove common unit symbols
    tempStr = Remove_Common_Units(tempStr)

    ' Then, convert any engineering suffixes to scientific notation
    tempStr = Convert_Unit_To_Exponential(SourceString)
    
    ' CDbl is locale-aware, so it will correctly handle the system's decimal separator.
    Convert_String_To_Double = CDbl(tempStr)
End Function

Private Function Convert_Unit_To_Exponential(ByVal SourceString As String) As String
    ' @brief (Helper) Converts a string with engineering suffixes to scientific "E" notation.
    ' @param SourceString The string to convert (e.g., "1.5k", "200m").
    ' @return The string with suffixes replaced by their exponential equivalents (e.g., "1.5E3", "200E-3").
    Dim tempStr As String
    tempStr = Replace(SourceString, "a", "E-18", , , vbTextCompare)
    tempStr = Replace(tempStr, "f", "E-15", , , vbTextCompare)
    tempStr = Replace(tempStr, "p", "E-12", , , vbTextCompare)
    tempStr = Replace(tempStr, "a", "E-10", , , vbTextCompare)
    tempStr = Replace(tempStr, "n", "E-9", , , vbTextCompare)
    tempStr = Replace(tempStr, "u", "E-6", , , vbTextCompare)
    tempStr = Replace(tempStr, "m", "E-3", , , vbTextCompare)
    tempStr = Replace(tempStr, "k", "E3", , , vbTextCompare)
    tempStr = Replace(tempStr, "M", "E6", , , vbTextCompare)
    tempStr = Replace(tempStr, "G", "E9", , , vbTextCompare)
    Convert_Unit_To_Exponential = tempStr
End Function

Private Function Remove_Common_Units(ByVal SourceString As String) As String
    ' @brief (Helper) Removes common engineering unit symbols from a string.
    ' @param SourceString The string to clean.
    ' @return The string with unit symbols removed.
    Dim tempStr As String
    tempStr = SourceString
    tempStr = Replace(tempStr, "A", "", , , vbTextCompare)
    tempStr = Trim(Replace(tempStr, "V", "", , , vbTextCompare))
    tempStr = Trim(Replace(tempStr, "W", "", , , vbTextCompare))
    tempStr = Trim(Replace(tempStr, "F", "", , , vbTextCompare))
    tempStr = Trim(Replace(tempStr, "Ohm", "", , , vbTextCompare))
    tempStr = Trim(Replace(tempStr, "Ohms", "", , , vbTextCompare))
    tempStr = Trim(Replace(tempStr, "S", "", , , vbTextCompare))
    tempStr = Trim(Replace(tempStr, "H", "", , , vbTextCompare))
    tempStr = Trim(Replace(tempStr, "Hz", "", , , vbTextCompare))
    tempStr = Trim(Replace(tempStr, "T", "", , , vbTextCompare))
    tempStr = Trim(Replace(tempStr, "R", "", , , vbTextCompare))
    tempStr = Trim(Replace(tempStr, "L", "", , , vbTextCompare))
    Remove_Common_Units = tempStr
End Function

Public Function Convert_String_To_Long(ByVal SourceString As String) As Long
    ' @brief Converts a string with potential engineering suffixes to a Long integer.
    ' @param SourceString The string to convert (e.g., "1k", "2M").
    ' @return The numeric value as a Long.
    ' @details This function first converts the string to a Double to handle suffixes, then converts the result to a Long.
    ' @dependencies Convert_String_To_Double
    Convert_String_To_Long = CLng(Convert_String_To_Double(SourceString))
End Function

Public Function Convert_Double_To_String(ByVal DblValue As Double) As String
    ' @brief Converts a Double to a string using a period as the decimal separator, regardless of locale.
    ' @param DblValue The double-precision number to convert.
    ' @return A string representation of the number with a period decimal separator.
    Convert_Double_To_String = Replace(CStr(DblValue), Application.International(xlDecimalSeparator), ".")
End Function

Public Function Convert_Seconds_To_Date(ByVal Seconds As Double) As Date
    ' @brief Converts a total number of seconds into a VBA Date type.
    ' @param Seconds The total number of seconds.
    ' @return A Date value representing the duration.
    Convert_Seconds_To_Date = CDate(Seconds / 86400#) ' 86400 seconds in a day
End Function

Public Function Convert_ColumnLetter_To_Number(ByVal ColumnLetter As String) As Long
    ' @brief Converts a column letter (e.g., "A", "B", "AA") to its corresponding number.
    ' @param ColumnLetter The column letter(s) to convert.
    ' @return The column number (e.g., 1, 2, 27). Returns 0 on failure.
    On Error Resume Next
    Convert_ColumnLetter_To_Number = Range(ColumnLetter & "1").Column
    If Err.Number <> 0 Then Convert_ColumnLetter_To_Number = 0
    On Error GoTo 0
End Function

Public Function Convert_ColumnNumber_To_Letter(ByVal ColumnNumber As Long) As String
    ' @brief Converts a column number (e.g., 1, 2, 27) to its corresponding letter ("A", "B", "AA").
    ' @param ColumnNumber The column number to convert.
    ' @return The column letter(s). Returns an empty string on failure.
    ' @details This function uses a mathematical approach and does not rely on any worksheet cells.
    Dim s As String
    If ColumnNumber > 0 Then
        Do
            s = Chr(((ColumnNumber - 1) Mod 26) + 65) & s
            ColumnNumber = (ColumnNumber - 1) \ 26
        Loop While ColumnNumber > 0
    End If
    Convert_ColumnNumber_To_Letter = s
End Function

Public Function Create_GUID() As String
    ' @brief Generates a new Globally Unique Identifier (GUID).
    ' @return A string representing the new GUID in the format "{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}".
    '         Returns an empty string on failure.
    ' @details This function uses Windows API calls for reliable GUID generation.
    Dim udtGUID As GUID
    Dim sGUID As String
    Dim lResult As Long
    
    lResult = CoCreateGuid(udtGUID)
    If lResult = 0 Then ' S_OK
        sGUID = String$(39, vbNullChar)
        #If VBA7 Then
            lResult = StringFromGUID2(udtGUID, StrPtr(sGUID), 39)
        #Else
            lResult = StringFromGUID2(udtGUID, sGUID, 39)
        #End If
        
        If lResult > 0 Then
            Create_GUID = Left$(sGUID, lResult - 1)
        End If
    End If
End Function

Public Function Slugify_String(ByVal SourceString As String, Optional ByVal Separator As String = "-") As String
    ' @brief Converts a string into a URL-friendly or filename-friendly "slug".
    ' @param SourceString The string to convert.
    ' @param Separator (Optional) The character to use in place of spaces. Defaults to "-".
    ' @return A clean, lowercase, alphanumeric string.
    ' @dependencies Microsoft VBScript Regular Expressions 5.5
    
    Dim s As String
    s = LCase(Trim(SourceString))
    
    ' Replace spaces and common separators with the desired separator
    s = Replace(s, " ", Separator)
    s = Replace(s, "_", Separator)
    
    ' Remove any characters that are not alphanumeric or the separator
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .IgnoreCase = True
        .Pattern = "[^a-z0-9" & Separator & "]" ' Keep letters, numbers, and the separator
        s = .Replace(s, "")
    End With
    
    ' Replace multiple separators with a single one
    Do While InStr(s, Separator & Separator) > 0
        s = Replace(s, Separator & Separator, Separator)
    Loop
    
    ' Remove leading/trailing separators
    If Left(s, 1) = Separator Then s = Mid(s, 2)
    If Right(s, 1) = Separator Then s = Left(s, Len(s) - 1)
    
    Slugify_String = s
End Function

Public Function Sanitize_For_FileName(ByVal SourceString As String, Optional ByVal ReplacementChar As String = "") As String
    ' @brief Removes characters that are invalid in Windows filenames.
    ' @param SourceString The string to sanitize.
    ' @param ReplacementChar (Optional) The character to use as a replacement. Defaults to "" (removes the character).
    ' @return A string safe to be used as a filename.
    ' @details Invalid characters are: \ / : * ? " < > |
    Dim invalidChars As String
    Dim i As Long
    Dim tempChar As String
    
    invalidChars = "\/:*?""<>|"
    Sanitize_For_FileName = SourceString
    
    For i = 1 To Len(invalidChars)
        tempChar = Mid(invalidChars, i, 1)
        Sanitize_For_FileName = Replace(Sanitize_For_FileName, tempChar, ReplacementChar)
    Next i
End Function

Public Function Sanitize_For_WorksheetName(ByVal SourceString As String, Optional ByVal ReplacementChar As String = "") As String
    ' @brief Removes characters that are invalid in Excel worksheet names and truncates to 31 characters.
    ' @param SourceString The string to sanitize.
    ' @param ReplacementChar (Optional) The character to use as a replacement. Defaults to "" (removes the character).
    ' @return A string safe to be used as a worksheet name.
    ' @details Invalid characters are: \ / ? * [ ]
    Dim invalidChars As String
    Dim i As Long
    Dim tempChar As String
    
    invalidChars = "\/?*[]"
    Sanitize_For_WorksheetName = SourceString
    
    For i = 1 To Len(invalidChars)
        tempChar = Mid(invalidChars, i, 1)
        Sanitize_For_WorksheetName = Replace(Sanitize_For_WorksheetName, tempChar, ReplacementChar)
    Next i
    
    ' Worksheet names also have a length limit of 31 characters.
    If Len(Sanitize_For_WorksheetName) > 31 Then
        Sanitize_For_WorksheetName = Left(Sanitize_For_WorksheetName, 31)
    End If
End Function

Public Function Sanitize_For_FilePath(ByVal SourceString As String, Optional ByVal Style As PathStyle = psWindows, Optional ByVal ReplacementChar As String = "") As String
    ' @brief Sanitizes a path string by removing invalid characters and normalizing path separators.
    ' @param SourceString The path string to sanitize.
    ' @param Style (Optional) The path style to sanitize for (psWindows or psUnix). Defaults to psWindows.
    ' @param ReplacementChar (Optional) The character to use as a replacement. Defaults to "" (removes the character).
    ' @return A sanitized string with invalid characters removed and path separators normalized for the target style.
    ' @details This function first removes characters that are illegal within file/folder names,
    '          then converts all path separators to the correct style (`\` for Windows, `/` for Unix).
    Dim invalidChars As String
    Dim i As Long
    Dim tempChar As String
    Dim sanitizedString As String
    
    ' Define invalid characters based on the OS style
    Select Case Style
        Case psWindows: invalidChars = "<>:""|?*"           ' Separators (\, /) are handled separately
        Case psUnix:    invalidChars = "<>:""|?*" & Chr(0)  ' Be strict: remove characters that are problematic in shells, even if technically allowed.
    End Select
    
    sanitizedString = SourceString
    
    For i = 1 To Len(invalidChars)
        tempChar = Mid(invalidChars, i, 1)
        sanitizedString = Replace(sanitizedString, tempChar, ReplacementChar)
    Next i
    
    ' Normalize path separators
    If Style = psWindows Then
        sanitizedString = Replace(sanitizedString, "/", "\")
    Else ' psUnix
        sanitizedString = Replace(sanitizedString, "\", "/")
    End If
    
    Sanitize_For_FilePath = sanitizedString
End Function
