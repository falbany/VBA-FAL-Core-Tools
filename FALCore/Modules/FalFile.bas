Attribute VB_Name = "FalFile"
' **************************************************************************************
' Module    : FalFile
' Author    : Florent ALBANY
' Website   :
' Purpose   : Manipulation of Files
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2023-07-00              Initial Release
' 2         2025-08-25              Translated to English and refactored for clarity.
'---------------------------------------------------------------------------------------
' Revision Propositions:
' ~~~~~~~~~~~~~~~~
' Dependencies:
' ~~~~~~~~~~~~~~~~
'   FALCore
'   FalLang
'   FalWork
'   FalArray
' **************************************************************************************


Option Explicit

Private Sub HandleError(ByVal errorMessage As String)
    ' @brief Displays a standardized error message to the user.
    ' @param errorMessage The specific error message to display.
    MsgBox "An error occurred:" & vbCrLf & errorMessage, vbOKOnly + vbCritical, "Operation Failed"
End Sub


Public Function Logger(sFile As String, Optional sType As String = "", Optional sSource As String = "", Optional sDetails As String = "", Optional maxFileSize As Long = 20000) As Boolean
    ' * @brief Logs events to a specified file with optional archiving based on file size.
    ' * @param sFile The file path to log events to.
    ' * @param sType (Optional) The type of the event (e.g., INFO, WARNING, ERROR).
    ' * @param sSource (Optional) The source of the event (e.g., module, function).
    ' * @param sDetails (Optional) Additional details about the event.
    ' * @param maxFileSize (Optional) The maximum file size before archiving (default is 20,000 bytes).
    ' *        If set to 0 or not provided, no archiving based on file size will occur.
    ' * @return Boolean indicating if the logging operation was successful.
    ' *
    ' * @details
    ' * Usage Example:
    ' * @code{vba}
    ' * Logger "C:\Path\To\Log.txt", "INFO", "Module1", "Event details", 50000
    ' * @endcode
    ' *
    ' * This will log an INFO event with the specified details to the file "C:\Path\To\Log.txt".
    ' * If the file size exceeds 50,000 bytes, it will be archived.
    ' */
    Dim stext       As String
    Dim sExtension  As String

    sExtension = "." & Get_FileExtension(sFile)

    ' Archive file at certain size
    If maxFileSize <> 0 Then
        If FileExist(sFile) Then
            If FileLen(sFile) > maxFileSize Then
                FileCopy sFile, Replace(sFile, sExtension, Format(Now, "_ddmmyyyy-hhmmss") & sExtension)
                Kill sFile
            End If
        End If
    End If

    stext = "[" & Format(Now, "dd-mm-yyyy hh:mm:ss") & "]" & vbTab & _
            IIf(sType <> "", sType & vbTab, "") & _
            IIf(sSource <> "", sSource & vbTab, "") & _
            IIf(sDetails <> "", sDetails & vbTab, "") & _
            "(user: " & Application.UserName & vbTab & _
            "version: " & FalWork.FALCORE_VERSION & ")"

    Logger = AppendTxt(sFile, stext, True)

End Function

Function Get_SaveFilePath_WithDialog(Optional InitialFileName As String = "MyFile", Optional InitialFileExtension As String = ".txt", Optional FileFilter As String = "Text Files (*.txt), *.xlsx") As String
    Dim filePath As Variant

    ' Prompt user to choose a file path and name
    filePath = Application.GetSaveAsFilename( _
                InitialFileName:=InitialFileName & InitialFileExtension, _
                FileFilter:=FileFilter)

    ' Check if a file path was selected
    If filePath <> False Then Get_SaveFilePath_WithDialog = filePath Else Get_SaveFilePath_WithDialog = ""

End Function

Public Function Select_Files(Optional FileType As String = "", Optional AllowMultiSelect As Boolean = True) As Variant
    '* @brief Displays a file selection dialog for the user to choose one or more files.
    '* @param FileType The type of files to filter (e.g., "xls", "xlsx", "mdm", "csv"). Leave blank for all files.
    '* @param AllowMultiSelect True if multiple file selection is allowed, False otherwise.
    '* @return An array of selected file paths.
    On Error GoTo ifError
    Dim i               As Integer
    Dim count           As Integer
    Dim filePaths()     As String
    Dim FileDialog As FileDialog

    ' Selection de fichiers
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With FileDialog
        .Filters.Clear
        Select Case UCase(FileType)
            Case "":
            Case "XLS", "XLSX"
                .Filters.Add "Excel files", "*.xlsx; *.xls; *.xlsm", 1
                .Filters.Add "Excel Converted", "*.mdm.xlsx; *.mdm.xlsx", 2
            Case "MDM": .Filters.Add "ICCAP", "*.mdm", 1
            Case "CSV": .Filters.Add "CSV", "*.csv", 1
            Case Else: .Filters.Add "Specified", Replace(FileType, ".", "*."), 1
        End Select
        .Filters.Add "All files", "*.*"
        .Title = "File Selection"
        .AllowMultiSelect = AllowMultiSelect
        .InitialView = msoFileDialogViewDetails
        .ButtonName = "Select"
        .Show
        count = .SelectedItems.count
        ReDim filePaths(1 To count)
        For i = 1 To count
            filePaths(i) = .SelectedItems(i)
        Next
    End With

    Select_Files = filePaths
    Exit Function
ifError:
    Select_Files = CVErr(2001)
End Function

Public Function Select_Folder(Optional AllowMultiSelect As Boolean = False) As String
    On Error GoTo ifError
    Dim FileDialog As FileDialog
    ' Folder selection.
    Set FileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With FileDialog
        .Title = "Folder selection"
        .AllowMultiSelect = AllowMultiSelect
        .InitialView = msoFileDialogViewDetails
        .ButtonName = "Select"
        '.Show
        If .Show Then
            Select_Folder = .SelectedItems(1)
        Else
            Select_Folder = ""  ' if aborted.
        End If
    End With
    Exit Function
ifError:
    Select_Folder = ""
End Function

Public Function GetAllFiles(Directory As String, FileType As String, FileTypeNot As String, Filter1 As String, Filter2 As String, Filter3 As String, Filter4 As String, Filter5 As String, Filter6 As String, ByRef FileList() As String, ByRef FolderList() As String, Optional ByRef file_count As Long) As Boolean
    'On Error GoTo IfError
    ' Lists all files in a directory with filtering capabilities.
    ' Requires "Microsoft Scripting RunTime" reference to be enabled.
        'In the macro editor (Alt+F11):
        'Tools menu
        'References
        'Check the "Microsoft Scripting RunTime" line.
        'Click the OK button to validate.

    ' Filter1 & Filter2 & Filter3 -> .And
    ' Filter4 & Filter5 & Filter6 -> .NAnd

    Dim Fso As Scripting.FileSystemObject
    Dim SourceFolder As Scripting.Folder
    Dim SubFolder As Scripting.Folder
    Dim FileItem As Scripting.File
    Dim And1 As String, And2 As String, And3 As String
    Dim ExtensionFile As String
    Dim AndNot1 As String, AndNot2 As String, AndNot3 As String
    Dim ExtensionFileNot As String
    Dim File As String

    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = Fso.GetFolder(Directory)

    ' Filters: Formatting
    And1 = "*" & Filter1 & "*"
    And2 = "*" & Filter2 & "*"
    And3 = "*" & Filter3 & "*"
    ExtensionFile = "*" & FileType
    AndNot1 = "*" & Filter4 & "*"
    AndNot2 = "*" & Filter5 & "*"
    AndNot3 = "*" & Filter6 & "*"
    ExtensionFileNot = "*" & FileTypeNot & "*"

    ' Handling of no filter
    If Filter1 = "" Then And1 = "*.*"
    If Filter2 = "" Then And2 = "*.*"
    If Filter3 = "" Then And3 = "*.*"
    If FileType = "" Then ExtensionFile = "*.*"
    If Filter4 = "" Then AndNot1 = "*FAlbanY*"
    If Filter5 = "" Then AndNot2 = "*FAlbanY*"
    If Filter6 = "" Then AndNot3 = "*FAlbanY*"
    If FileTypeNot = "" Then ExtensionFileNot = "*FAlbanY*"

    'Loop through all files in the directory
    For Each FileItem In SourceFolder.Files
        File = FileItem.name
        ' Extension.
        If Not (File Like ExtensionFile) Then GoTo NextFile
        If (File Like ExtensionFileNot) Then GoTo NextFile
        ' And filters.
        If Not (File Like And1) Then GoTo NextFile
        If Not (File Like And2) Then GoTo NextFile
        If Not (File Like And3) Then GoTo NextFile
        ' NAnd filters.
        If (File Like AndNot1) Then GoTo NextFile
        If (File Like AndNot2) Then GoTo NextFile
        If (File Like AndNot3) Then GoTo NextFile
        ' If Match.
        ReDim Preserve FileList(file_count)
        ReDim Preserve FolderList(file_count)
        FileList(file_count) = FileItem.name                   ' Writes the file name to the array
        FolderList(file_count) = FileItem.ParentFolder         ' Writes the file path to the array
        ' Increment.
        file_count = file_count + 1
        ' Progression.
        Application.StatusBar = "Searching for files: " & file_count & " files found.     " & FolderList(file_count - 1) & "\" & FileList(file_count - 1)
NextFile:
    Next FileItem

    '--- Recursive call to list files in subdirectories ---.
    For Each SubFolder In SourceFolder.SubFolders
       Call GetAllFiles(SubFolder.path, FileType, FileTypeNot, Filter1, Filter2, Filter3, Filter4, Filter5, Filter6, FileList, FolderList, file_count)
    Next SubFolder

'    GetAllFiles = True
'    Exit Function
'IfError:
'    GetAllFiles = False
End Function

Public Function Sort_Files(ByVal FilePaths As Variant, ByVal SortKey As String, Optional ByVal SortOrder As XlSortOrder = xlAscending, Optional ByVal CustomPattern As String = "") As Variant
    ' @brief Sorts an array of file paths based on specified criteria.
    ' @param FilePaths An array of full file path strings.
    ' @param SortKey The criteria to sort by. Valid options: "Date", "FileName", "FileSize", "CustomNumber".
    ' @param SortOrder (Optional) The sort direction, xlAscending (default) or xlDescending.
    ' @param CustomPattern (Optional) A VBScript RegExp pattern used to extract a number from the filename when SortKey is "CustomNumber".
    ' @return A variant array of strings containing the sorted file paths. Returns Empty on failure.
    ' @dependencies Microsoft Scripting Runtime, Microsoft VBScript Regular Expressions 5.5
    ' @details
    ' For "CustomNumber" sorting, the function uses a regular expression to find a number in the filename.
    ' It is best practice to use a capturing group `()` in your `CustomPattern` to isolate the exact numeric part.
    ' If a capturing group is present, its value is used for sorting. Otherwise, the entire matched pattern is used.
    ' @example
    ' ' To sort files like "DataFile_10.csv", "DataFile_2.csv" numerically:
    ' Dim files As Variant
    ' files = Array("C:\Data\DataFile_10.csv", "C:\Data\DataFile_2.csv")
    '
    ' ' This pattern finds a number (\d+) surrounded by an underscore and a dot.
    ' ' The parentheses `()` capture just the number for sorting.
    ' Dim pattern As String
    ' pattern = "_(\d+)\."
    '
    ' Dim sortedFiles As Variant
    ' sortedFiles = Sort_Files(files, "CustomNumber", xlAscending, pattern)

    On Error GoTo ErrorHandler
    Sort_Files = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(FilePaths) Then
        Err.Raise vbObjectError, , "Invalid input for Sort_Files. An array of file paths is expected."
        Exit Function
    End If
    If LBound(FilePaths) > UBound(FilePaths) Then Exit Function ' Empty array, return Empty

    Dim lb As Long: lb = LBound(FilePaths)
    Dim ub As Long: ub = UBound(FilePaths)
    Dim i As Long, j As Long

    ' 2. Prepare data structure for sorting
    ' Using a 2D array to hold [SortValue, OriginalFilePath]
    Dim sortData() As Variant
    ReDim sortData(lb To ub, 1 To 2)

    Dim fso As Object ' Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim regEx As Object ' VBScript_RegExp_55.RegExp
    Dim regExMatches As Object

    If LCase(SortKey) = "customnumber" Then
        If CustomPattern = "" Then
            Err.Raise vbObjectError, , "A 'CustomPattern' must be provided when SortKey is 'CustomNumber'."
            Exit Function
        End If
        Set regEx = CreateObject("VBScript.RegExp")
        With regEx
            .Global = False
            .IgnoreCase = True
            .Pattern = CustomPattern
        End With
    End If

    ' 3. Populate the sort data array
    For i = lb To ub
        Dim file As Object
        Set file = fso.GetFile(FilePaths(i))

        Select Case LCase(SortKey)
            Case "date"
                sortData(i, 1) = file.DateLastModified
            Case "filename"
                sortData(i, 1) = file.Name
            Case "filesize"
                sortData(i, 1) = file.Size
            Case "customnumber"
                Set regExMatches = regEx.Execute(file.Name)
                If regExMatches.Count > 0 Then
                    If regExMatches(0).SubMatches.Count > 0 Then
                        ' A capturing group was used, extract the number from the first submatch.
                        sortData(i, 1) = Val(regExMatches(0).SubMatches(0))
                    Else
                        ' No capturing group, use the entire match.
                        sortData(i, 1) = Val(regExMatches(0).Value)
                    End If
                Else
                    sortData(i, 1) = -1 ' Default value for non-matching files
                End If
            Case Else
                Err.Raise vbObjectError, , "Invalid 'SortKey' provided: '" & SortKey & "'. Valid options are 'Date', 'FileName', 'FileSize', 'CustomNumber'."
                Exit Function
        End Select
        sortData(i, 2) = FilePaths(i)
    Next i

    ' 4. Sort the array
    For i = lb To ub - 1
        For j = i + 1 To ub
            If (SortOrder = xlAscending And sortData(i, 1) > sortData(j, 1)) Or (SortOrder = xlDescending And sortData(i, 1) < sortData(j, 1)) Then
                Dim tempSortValue As Variant
                Dim tempFilePath As String

                tempSortValue = sortData(i, 1)
                tempFilePath = sortData(i, 2)

                sortData(i, 1) = sortData(j, 1)
                sortData(i, 2) = sortData(j, 2)

                sortData(j, 1) = tempSortValue
                sortData(j, 2) = tempFilePath
            End If
        Next j
    Next i

    ' 5. Prepare the result array
    Dim result() As String
    ReDim result(lb To ub)

    For i = lb To ub
        result(i) = sortData(i, 2)
    Next i

    Sort_Files = result
    Exit Function

ErrorHandler:
    Sort_Files = Empty
End Function

Public Function Ini_ReadKeyVal(ByVal sIniFIle As String, _
                        ByVal sSection As String, _
                        ByVal sKey As String) As String
    On Error GoTo Error_Handler
    Dim sIniFileContent         As String
    Dim aIniLines()             As String
    Dim sLine                   As String
    Dim i                       As Long
    Dim bSectionExists          As Boolean
    Dim bKeyExists              As Boolean

    sIniFileContent = ""
    bSectionExists = False
    bKeyExists = False

    'Validate that the file actually exists
    If FileExist(sIniFIle) = False Then
        MsgBox "The specified ini file: " & vbCrLf & vbCrLf & _
               sIniFIle & vbCrLf & vbCrLf & _
               "could not be found.", vbCritical + vbOKOnly, "File not found"
        GoTo Error_Handler_Exit
    End If

    sIniFileContent = ReadFile(sIniFIle)    'Read the file into memory
    aIniLines = Split(sIniFileContent, vbCrLf)
    For i = 0 To UBound(aIniLines)
        sLine = Trim(aIniLines(i))
        If bSectionExists = True And Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
            Exit For    'Start of a new section
        End If
        If sLine = "[" & sSection & "]" Then
            bSectionExists = True
        End If
        If bSectionExists = True Then
            If Len(sLine) > Len(sKey) Then
                If Left(sLine, Len(sKey) + 1) = sKey & "=" Then
                    bKeyExists = True
                    Ini_ReadKeyVal = Mid(sLine, InStr(sLine, "=") + 1)
                End If
            End If
        End If
    Next i

Error_Handler_Exit:
    On Error Resume Next
    Exit Function

Error_Handler:
    'Err.Number = 75 'File does not exist, Permission issues to write is denied,
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Ini_ReadKeyVal" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Function Ini_WriteKeyVal(ByVal sIniFIle As String, _
                         ByVal sSection As String, _
                         ByVal sKey As String, _
                         ByVal sValue As String) As Boolean
    On Error GoTo Error_Handler
    Dim sIniFileContent         As String
    Dim aIniLines()             As String
    Dim sLine                   As String
    Dim sNewLine                As String
    Dim i                       As Long
    Dim bFileExist              As Boolean
    Dim bInSection              As Boolean
    Dim bKeyAdded               As Boolean
    Dim bSectionExists          As Boolean
    Dim bKeyExists              As Boolean


    sIniFileContent = ""
    bSectionExists = False
    bKeyExists = False

    'Validate that the file actually exists
    If FileExist(sIniFIle) = False Then GoTo SectionDoesNotExist Else bFileExist = True

    sIniFileContent = ReadFile(sIniFIle)    'Read the file into memory
    aIniLines = Split(sIniFileContent, vbCrLf)    'Break the content into individual lines
    sIniFileContent = ""    'Reset it
    For i = 0 To UBound(aIniLines)    'Loop through each line
        sNewLine = ""
        sLine = Trim(aIniLines(i))
        If sLine = "[" & sSection & "]" Then
            bSectionExists = True
            bInSection = True
        End If
        If bInSection = True Then
            If sLine <> "[" & sSection & "]" Then
                If Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
                    'Our section exists, but the key wasn't found, so append it
                    If bKeyAdded = False Then
                        sNewLine = sKey & "=" & sValue
                        i = i - 1
                        'bInSection = False
                        bKeyAdded = True
                    End If
                    bInSection = False
                End If
            End If
            If Len(sLine) > Len(sKey) Then
                If Split(sLine, "=")(0) = sKey Then
                    sNewLine = sKey & "=" & sValue
                    bKeyExists = True
                    bKeyAdded = True
                End If
            End If
        End If
        If Len(sIniFileContent) > 0 Then sIniFileContent = sIniFileContent & vbCrLf
        If sNewLine = "" Then sIniFileContent = sIniFileContent & sLine Else sIniFileContent = sIniFileContent & sNewLine
    Next i

SectionDoesNotExist:
    'if not found, add it to the end
    If bSectionExists = False Then
        If Len(sIniFileContent) > 0 Then sIniFileContent = sIniFileContent & vbCrLf
        sIniFileContent = sIniFileContent & "[" & sSection & "]"
    End If
    If bKeyAdded = False Then
        sIniFileContent = sIniFileContent & vbCrLf & sKey & "=" & sValue
    End If

    'Write to the ini file the new content
    Call OverwriteTxt(sIniFIle, sIniFileContent)
    Ini_WriteKeyVal = True

Error_Handler_Exit:
    On Error Resume Next
    Exit Function

Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Ini_WriteKeyVal" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Function DirExist(path As String) As Boolean
    On Error Resume Next
    DirExist = IIf(Dir(path, vbDirectory) <> "", True, False)
End Function

Public Function FileExist(strFile As String) As Boolean
    On Error GoTo err_handler

    strFile = Clean_FilePath(strFile)
    If Len(Dir(strFile)) > 0 Then FileExist = True Else FileExist = False

Exit_Err_Handler:
    Exit Function

err_handler:
    MsgBox "The following error has occurred." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: FileExist" & vbCrLf & _
            "Error Description: " & Err.Description, _
            vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

Public Function OverwriteTxt(sFile As String, stext As String) As Boolean
On Error GoTo err_handler
    Dim fileNumber As Integer

    sFile = Clean_FilePath(sFile)

    fileNumber = FreeFile()
    Open sFile For Output As #fileNumber
    Print #fileNumber, stext;
    Close #fileNumber
    OverwriteTxt = True
    Exit Function

Exit_Err_Handler:
    Exit Function

err_handler:
    MsgBox "The following error has occurred." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: OverwriteTxt" & vbCrLf & _
            "Error Description: " & Err.Description, _
            vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

Public Function AppendTxt(sFile As String, stext As String, Optional appendWithLineBreak As Boolean = False) As Boolean
    On Error GoTo err_handler
    Dim fileNumber As Integer

    sFile = Clean_FilePath(sFile)

    fileNumber = FreeFile()
    Open sFile For Append As #fileNumber

    If appendWithLineBreak Then
        If FileLen(sFile) > 0 Then Print #fileNumber, vbCrLf & stext;
    Else
        Print #fileNumber, stext;
    End If

    Close #fileNumber
    AppendTxt = True
    Exit Function

Exit_Err_Handler:
    Exit Function

err_handler:
    MsgBox "The following error has occurred." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: AppendTxt" & vbCrLf & _
            "Error Description: " & Err.Description, _
            vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

Public Function ReadFile(ByVal strFile As String) As String
On Error GoTo Error_Handler
    Dim fileNumber  As Integer
    Dim sFile       As String

    fileNumber = FreeFile
    Open strFile For Binary Access Read As fileNumber
    sFile = Space(LOF(fileNumber))
    Get #fileNumber, , sFile
    Close fileNumber

    ReadFile = sFile

Error_Handler_Exit:
    On Error Resume Next
    Exit Function

Error_Handler:
    MsgBox "The following error has occurred." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: ReadFile" & vbCrLf & _
            "Error Description: " & Err.Description, _
            vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Function Ini_ReadToDictionary(sIniFIle As String, Optional KeyCAseSensitive As Boolean = True, Optional sDelimiter As String = "=") As Dictionary
    Dim dict As New Dictionary
    Dim sSection As String
    Dim sKey As String
    Dim sValue As String
    Dim sIniFileContent As String
    Dim aIniLines() As String
    Dim i As Long

    If FileExist(sIniFIle) = False Then
        MsgBox "The specified ini file: " & vbCrLf & vbCrLf & _
               sIniFIle & vbCrLf & vbCrLf & _
               "could not be found.", vbCritical + vbOKOnly, "File not found"
        GoTo Error_Handler_Exit
    End If

    If KeyCAseSensitive Then dict.CompareMode = BinaryCompare Else dict.CompareMode = TextCompare
    dict("delimiter") = sDelimiter
    sIniFileContent = ReadFile(sIniFIle)
    aIniLines = Split(sIniFileContent, vbCrLf)
    For i = LBound(aIniLines) To UBound(aIniLines)
        If Left(aIniLines(i), 1) = "[" And Right(aIniLines(i), 1) = "]" Then sSection = "[" & Mid(aIniLines(i), 2, Len(aIniLines(i)) - 2) & "]": GoTo nextLine
        If aIniLines(i) = "" Then sSection = "": GoTo nextLine
        If Not aIniLines(i) Like "*" & sDelimiter & "*" Then GoTo nextLine
        sKey = Split(aIniLines(i), "=")(0)
        sValue = Split(aIniLines(i), "=")(1)
        dict(sSection & sKey) = sValue
nextLine:
    Next

    Set Ini_ReadToDictionary = dict
    Exit Function
Error_Handler_Exit:
End Function

Public Function Ini_ReadToCollection(sIniFIle As String, Optional sDelimiter As String = "=") As Collection
    Dim i               As Long
    Dim coll            As New Collection
    Dim sSection        As String
    Dim sKey            As String
    Dim sValue          As String
    Dim sIniFileContent As String
    Dim aIniLines()     As String


    If FileExist(sIniFIle) = False Then
        MsgBox "The specified ini file: " & vbCrLf & vbCrLf & _
               sIniFIle & vbCrLf & vbCrLf & _
               "could not be found.", vbCritical + vbOKOnly, "File not found"
        GoTo Error_Handler_Exit
    End If

    coll.Add Key:="delimiter", item:=sDelimiter
    sIniFileContent = ReadFile(sIniFIle)
    aIniLines = Split(sIniFileContent, vbCrLf)
    For i = LBound(aIniLines) To UBound(aIniLines)
        If Left(aIniLines(i), 1) = "[" And Right(aIniLines(i), 1) = "]" Then sSection = "[" & Mid(aIniLines(i), 2, Len(aIniLines(i)) - 2) & "]": GoTo nextLine
        If aIniLines(i) = "" Then sSection = "": GoTo nextLine
        If Not aIniLines(i) Like "*" & sDelimiter & "*" Then GoTo nextLine
        sKey = Split(aIniLines(i), "=")(0)
        sValue = Split(aIniLines(i), "=")(1)
        coll.Add Key:=sSection & sKey, item:=sValue

nextLine:
    Next

    Set Ini_ReadToCollection = coll
    Exit Function
Error_Handler_Exit:
End Function

Public Sub ClearTextFile(filePath As String)
    If FileExist(filePath) Then Open filePath For Output As #1: Close #1
End Sub

Public Sub Open_File(filePath As String)
    Dim objShell As Object
    If FileExist(filePath) Then
        Set objShell = CreateObject("Shell.Application")
        objShell.Open (filePath)
    End If
End Sub

Public Sub Open_Directory(path As String)
    Dim objShell As Object

    path = Clean_FilePath(path)
    If DirExist(path) Then
        Set objShell = CreateObject("Shell.Application")
        objShell.Open (path)
    Else
        MsgBox "The directory '" & path & "' does not exist.", vbExclamation, "Directory Not Found"
    End If

End Sub

Public Sub Import_CsvToSelection()
    On Error GoTo ifError
    Dim filePath        As Variant
    Dim rng             As Range
    Dim sCsv            As String

    Select Case TypeName(Selection)
        Case "Range": Set rng = Selection
        Case Else: MsgBox "You need to first select the destination cell", vbExclamation: GoTo ifError
    End Select

    filePath = Select_Files("csv", False)

    If IsError(filePath) Then MsgBox "No file selected.", vbExclamation: GoTo ifError

    sCsv = ReadFile(filePath(LBound(filePath)))

    If Trim(sCsv) = "" Then MsgBox Get_FileName(CStr(filePath(LBound(filePath)))) & "." & Get_FileExtension(CStr(filePath(LBound(filePath)))) & " file is empty.", vbExclamation: GoTo ifError

    If Not FalArray.Csv_To_Range(sCsv, rng, ",", Chr(34)) Then MsgBox "Write operation failed.", vbExclamation: GoTo ifError

    Exit Sub

ifError:
End Sub

Public Sub Import_CsvToNewSpreadsheet()
    On Error GoTo ifError
    Dim i               As Integer
    Dim filePath        As Variant
    Dim sCsv            As String
    Dim fileName        As String
    Dim SheetName       As String
    Dim wks             As Worksheet
    Dim wbk             As Workbook
    Dim Delimiter       As String
    Dim QuoteChar       As String

    Delimiter = ","
    QuoteChar = ""

    filePath = Select_Files("csv", True)

    If IsError(filePath) Then MsgBox "No file selected.", vbExclamation: GoTo ifError

    For i = LBound(filePath) To UBound(filePath)

        fileName = Get_FileName(CStr(filePath(i)))
        sCsv = ReadFile(filePath(i))

        If Trim(sCsv) = "" Then MsgBox fileName & "." & Get_FileExtension(CStr(filePath(i))) & " file is empty.", vbExclamation: GoTo ifError

        Set wbk = ActiveWorkbook
        SheetName = FalLang.Clear_SubChar(fileName, "_-:\/?*[]; ")
        SheetName = FalLang.Resize_String(SheetName, 31)
        If FalWork.wbk_SheetExist(wbk, SheetName) Then
            Set wks = wbk.Sheets(SheetName)
            wks.Cells.Clear
        Else
            Set wks = wbk.Sheets.Add
            wks.name = SheetName
            wks.Move After:=wbk.Sheets(wbk.Sheets.count)
        End If

        If Not FalArray.Csv_To_Range(sCsv, wks.Range("A1"), Delimiter, QuoteChar) Then MsgBox "Write operation failed.", vbExclamation: GoTo ifError
    Next i

    Exit Sub

ifError:
End Sub

Public Sub Export_SelectionToCsv()
    Dim Arr2D           As Variant
    Dim sCsv            As String
    Dim path            As String
    Dim filePath        As String
    Dim fileName        As String
    Dim Delimiter       As String
    Dim QuoteChar       As String

    Delimiter = ","
    QuoteChar = Chr(34)

    Arr2D = FalArray.a2D_From_Selection()
    If IsError(Arr2D) Then
        MsgBox "No valid selection found.", vbExclamation
        Exit Sub
    End If

    filePath = Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "csv Files (*.csv),*csv")
    If filePath = "" Then Exit Sub

    sCsv = FalArray.a2D_ToCsv(Arr2D, Delimiter, QuoteChar)

    OverwriteTxt filePath, sCsv

    MsgBox "CSV file created successfully: " & filePath, vbInformation
End Sub

Public Sub Export_SheetToCsv()
    Dim Arr2D           As Variant
    Dim sCsv            As String
    Dim path            As String
    Dim filePath        As String
    Dim fileName        As String
    Dim Delimiter       As String
    Dim QuoteChar       As String

    Delimiter = ","
    QuoteChar = Chr(34)

    Arr2D = FalArray.a2D_From_Selection()
    If IsError(Arr2D) Then
        MsgBox "No valid sheet found.", vbExclamation
        Exit Sub
    End If

    filePath = Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "csv Files (*.csv),*csv")
    If filePath = "" Then Exit Sub

    sCsv = FalArray.a2D_ToCsv(Arr2D, Delimiter, QuoteChar)

    OverwriteTxt filePath, sCsv

    MsgBox "CSV file created successfully: " & filePath, vbInformation
End Sub

Public Function ReadStringFromFile(ByVal path As String) As String
    On Error GoTo ErrHandler

    Dim fileNumber As Long
    Dim byteCount As Long
    Dim resultString As String

    If Dir(path, vbNormal + vbHidden + vbSystem + vbArchive + vbReadOnly) <> vbNullString Then
        fileNumber = FreeFile

        Open path For Binary Access Read Shared As #fileNumber
            If LOF(fileNumber) > 0 Then
                byteCount = LOF(fileNumber)
                resultString = Space(byteCount)
                Get #fileNumber, , resultString
            Else
                resultString = vbNullString
            End If
        Close #fileNumber
    Else
        Call HandleError("The specified file does not exist or the path is invalid: " & path)
        resultString = vbNullString
        GoTo CleanExit
    End If

    ReadStringFromFile = resultString

CleanExit:
    Exit Function

ErrHandler:
    If fileNumber <> 0 Then Close #fileNumber
    Call HandleError("Error reading file '" & path & "' : " & Err.Description & " (Code: " & Err.Number & ")")
    ReadStringFromFile = vbNullString
    Resume CleanExit
End Function

Public Function GetFileEncoding(ByVal filePath As String) As String
    On Error GoTo ErrHandler

    Dim fileNumber As Long
    Dim bytes() As Byte
    Dim actualBytesRead As Long
    Dim result As String
    Dim maxBytesToRead As Long

    If Dir(filePath) = vbNullString Then
        Call HandleError("The specified file does not exist: " & filePath)
        GetFileEncoding = "File not found"
        Exit Function
    End If

    fileNumber = FreeFile
    Open filePath For Binary Access Read As #fileNumber

    If LOF(fileNumber) > 0 Then
        maxBytesToRead = 4

        If LOF(fileNumber) < maxBytesToRead Then
            ReDim bytes(0 To LOF(fileNumber) - 1)
        Else
            ReDim bytes(0 To maxBytesToRead - 1)
        End If

        Get #fileNumber, , bytes
        actualBytesRead = UBound(bytes) - LBound(bytes) + 1

    Else
        GetFileEncoding = "Empty"
        Close #fileNumber
        Exit Function
    End If

    Close #fileNumber

    Select Case True
        Case actualBytesRead >= 3 And _
             bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF
            result = "UTF-8"

        Case actualBytesRead >= 2 And _
             bytes(0) = &HFF And bytes(1) = &HFE
            result = "UTF-16 LE"

        Case actualBytesRead >= 2 And _
             bytes(0) = &HFE And bytes(1) = &HFF
            result = "UTF-16 BE"

        Case actualBytesRead >= 4 And _
             bytes(0) = &HFF And bytes(1) = &HFE And bytes(2) = &H0 And bytes(3) = &H0
            result = "UTF-32 LE"

        Case actualBytesRead >= 4 And _
             bytes(0) = &H0 And bytes(1) = &H0 And bytes(2) = &HFE And bytes(3) = &HFF
            result = "UTF-32 BE"

        Case Else
            result = "ANSI"
    End Select

    GetFileEncoding = result

CleanExit:
    Exit Function

ErrHandler:
    If fileNumber <> 0 Then Close #fileNumber
    Call HandleError("Error detecting file encoding '" & filePath & "' : " & Err.Description & " (Code: " & Err.Number & ")")
    GetFileEncoding = "Error"
    Resume CleanExit
End Function

Public Function ReadFile_WithADO(ByVal strFile As String) As String
    On Error GoTo Error_Handler

    Dim adoStream As Object ' Late binding for ADODB.Stream
    Dim detectedEncoding As String
    Dim tempPath As String

    ReadFile_WithADO = ""

    tempPath = Dir$(strFile, vbNormal + vbHidden + vbSystem + vbArchive + vbReadOnly)
    If tempPath = vbNullString Then
        Call HandleError("The specified file does not exist or the path is invalid: " & strFile)
        GoTo Error_Handler_Exit
    End If

    detectedEncoding = GetFileEncoding(strFile)

    Select Case detectedEncoding
        Case "File not found", "Empty", "Error"
            GoTo Error_Handler_Exit
    End Select

    Set adoStream = CreateObject("ADODB.Stream")

    With adoStream
        .Type = 2 ' adTypeText
        .Open

        Select Case detectedEncoding
            Case "UTF-8": .Charset = "utf-8"
            Case "UTF-16 LE": .Charset = "utf-16"
            Case "UTF-16 BE": .Charset = "utf-16BE"
            Case "UTF-32 LE", "UTF-32 BE"
                Call HandleError("UTF-32 encoding detected for '" & strFile & "'. ADODB.Stream does not natively support direct text reading of UTF-32. Attempting with UTF-8 fallback.")
                .Charset = "utf-8"
            Case "ANSI"
                .Charset = "Windows-1252"
            Case Else
                Call HandleError("Unknown or unhandled encoding detected ('" & detectedEncoding & "') for '" & strFile & "'. Attempting with UTF-8 fallback.")
                .Charset = "utf-8"
        End Select

        On Error Resume Next
        .LoadFromFile strFile
        If Err.Number <> 0 Then
            Err.Clear
            If .State = 1 Then .Close
            .Open
            .Charset = "Windows-1252"
            .LoadFromFile strFile
            If Err.Number <> 0 Then
                Call HandleError("Irrecoverable error reading file '" & strFile & "' with any supported encoding. Error: " & Err.Description)
                .Close
                GoTo Error_Handler_Exit
            End If
        End If
        On Error GoTo Error_Handler

        ReadFile_WithADO = .ReadText
        .Close
    End With

Error_Handler_Exit:
    On Error Resume Next
    If Not adoStream Is Nothing Then
        If adoStream.State = 1 Then adoStream.Close
        Set adoStream = Nothing
    End If
    Exit Function

Error_Handler:
    MsgBox "An error occurred while reading the file." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: ReadFile_WithADO" & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf & _
            "File: " & strFile, _
            vbCritical, "File Reading Error"
    Resume Error_Handler_Exit
End Function

Public Function ReadFile_WithoutADO(ByVal strFile As String) As String
    On Error GoTo Error_Handler

    Dim fileNumber As Long
    Dim detectedEncoding As String
    Dim fileContent As String
    Dim byteCount As Long
    Dim initialBytes(0 To 3) As Byte
    Dim tempPath As String

    ReadFile_WithoutADO = ""

    tempPath = Dir$(strFile, vbNormal + vbHidden + vbSystem + vbArchive + vbReadOnly)
    If tempPath = vbNullString Then
        Call HandleError("The specified file does not exist or the path is invalid: " & strFile)
        GoTo Error_Handler_Exit
    End If

    detectedEncoding = GetFileEncoding(strFile)

    Select Case detectedEncoding
        Case "File not found", "Error", "Empty"
            GoTo Error_Handler_Exit
    End Select

    fileNumber = FreeFile

    Select Case detectedEncoding
        Case "UTF-8"
            Open strFile For Binary Access Read As #fileNumber

            If LOF(fileNumber) >= 3 Then
                Get #fileNumber, , initialBytes(0)
                Get #fileNumber, , initialBytes(1)
                Get #fileNumber, , initialBytes(2)
                If Not (initialBytes(0) = &HEF And initialBytes(1) = &HBB And initialBytes(2) = &HBF) Then
                    Seek #fileNumber, 1
                End If
            End If

            byteCount = LOF(fileNumber) - (Seek(fileNumber) - 1)
            If byteCount > 0 Then
                Dim utf8Bytes() As Byte
                ReDim utf8Bytes(0 To byteCount - 1)
                Get #fileNumber, , utf8Bytes

                fileContent = StrConv(utf8Bytes, vbUnicode)
            End If
            Close #fileNumber

        Case "UTF-16 LE", "UTF-16 BE"
            Open strFile For Input Access Read As #fileNumber

            If LOF(fileNumber) >= 2 Then
                Dim dummyChar As String * 1
                Input #fileNumber, dummyChar, dummyChar
            End If

            fileContent = Input$(LOF(fileNumber), #fileNumber)
            Close #fileNumber

        Case "ANSI"
            Open strFile For Input Access Read As #fileNumber
            fileContent = Input$(LOF(fileNumber), #fileNumber)
            Close #fileNumber

        Case Else
            Open strFile For Input Access Read As #fileNumber
            fileContent = Input$(LOF(fileNumber), #fileNumber)
            Close #fileNumber

    End Select

    ReadFile_WithoutADO = fileContent

Error_Handler_Exit:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    Exit Function

Error_Handler:
    If fileNumber <> 0 Then Close #fileNumber
    MsgBox "An error occurred while reading the file." & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: ReadFile_WithoutADO" & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf & _
            "File: " & strFile, _
            vbCritical, "File Reading Error"
    Resume Error_Handler_Exit
End Function


Public Function Copy_File(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Overwrite As Boolean = True) As Boolean
    On Error GoTo ifError

    Dim fso As Object
    Dim destFolder As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(SourcePath) Then
        Call HandleError("Source file not found: " & SourcePath)
        Exit Function
    End If

    destFolder = fso.GetParentFolderName(DestinationPath)
    If Not fso.FolderExists(destFolder) Then
        Call Create_Folder(destFolder)
    End If

    fso.CopyFile Source:=SourcePath, Destination:=DestinationPath, OverwriteFiles:=Overwrite
    Copy_File = True

    Exit Function
ifError:
    Call HandleError("Failed to copy file from '" & SourcePath & "' to '" & DestinationPath & "'. " & vbCrLf & "Error: " & Err.Description)
    Copy_File = False
End Function


Public Function Move_File(ByVal SourcePath As String, ByVal DestinationPath As String) As Boolean
    On Error GoTo ifError

    Dim fso As Object
    Dim destFolder As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(SourcePath) Then
        Call HandleError("Source file not found: " & SourcePath)
        Exit Function
    End If

    destFolder = fso.GetParentFolderName(DestinationPath)
    If Not fso.FolderExists(destFolder) Then
        Call Create_Folder(destFolder)
    End If

    fso.MoveFile Source:=SourcePath, Destination:=DestinationPath
    Move_File = True

    Exit Function
ifError:
    Call HandleError("Failed to move file from '" & SourcePath & "' to '" & DestinationPath & "'. " & vbCrLf & "Error: " & Err.Description)
    Move_File = False
End Function

Public Function Clean_FilePath(filePath As String) As String
    Clean_FilePath = Replace(Replace(filePath, "/", "\"), "\\", "\")
End Function

Public Function Get_DirectoryPath(filePath As String) As String
    Dim fileName As String
    fileName = Get_FileName(filePath)
    Get_DirectoryPath = Left(filePath, Len(filePath) - Len(fileName))
End Function

Public Function Get_FileName(filePath As String) As String
    Dim lastBackslashIndex  As Long
    Dim lastPointIndex      As Long
    Dim lenFileName         As Long
    lastBackslashIndex = InStrRev(filePath, "\")
    lastPointIndex = InStrRev(filePath, ".")

    lenFileName = IIf(lastBackslashIndex < lastPointIndex, lastPointIndex - lastBackslashIndex - 1, Len(filePath) - lastBackslashIndex)
    If lastBackslashIndex > 0 Then Get_FileName = Mid(filePath, lastBackslashIndex + 1, lenFileName) Else Get_FileName = filePath
End Function

Public Function Get_BaseFileName(ByVal FilePath As String) As String
    Dim fileName As String
    Dim dotPosition As Long

    fileName = Get_FileName(FilePath)
    dotPosition = InStrRev(fileName, ".")
    If dotPosition > 0 Then
        Get_BaseFileName = Left(fileName, dotPosition - 1)
    Else
        Get_BaseFileName = fileName
    End If

End Function

Public Function Get_FileExtension(filePath As String) As String
    Dim lastPointIndex As Long
    lastPointIndex = InStrRev(filePath, ".")
    If lastPointIndex > 0 Then Get_FileExtension = Mid(filePath, lastPointIndex + 1) Else Get_FileExtension = ""
End Function

Public Function Get_FileSize(ByVal FilePath As String) As Double
    On Error GoTo ifError

    If Not FileExist(FilePath) Then
        Get_FileSize = -1
        Exit Function
    End If

    Get_FileSize = FileLen(FilePath)

    Exit Function
ifError:
    Call HandleError("Could not get file size for '" & FilePath & "'. " & vbCrLf & "Error: " & Err.Description)
    Get_FileSize = -1
End Function

Public Function Combine_Paths(ByVal BasePath As String, ByVal RelativePath As String) As String
    Dim cleanBasePath As String
    Dim cleanRelativePath As String

    cleanBasePath = Clean_FilePath(BasePath)
    cleanRelativePath = Clean_FilePath(RelativePath)

    If InStrRev(cleanBasePath, "\") > 0 And InStrRev(cleanBasePath, ".") > InStrRev(cleanBasePath, "\") Then
        cleanBasePath = Left(cleanBasePath, InStrRev(cleanBasePath, "\") - 1)
    End If

    If Right(cleanBasePath, 1) = "\" Then
        cleanBasePath = Left(cleanBasePath, Len(cleanBasePath) - 1)
    End If

    If Left(cleanRelativePath, 1) = "\" Then
        cleanRelativePath = Mid(cleanRelativePath, 2)
    End If

    Combine_Paths = cleanBasePath & "\" & cleanRelativePath
End Function

Public Sub Zip_Files(ByVal FilePaths As Variant, ByVal ZipFilePath As String)
    On Error GoTo ifError

    Dim shellApp As Object
    Dim fileItem As Variant
    Dim i As Long

    If FileExist(ZipFilePath) Then Kill ZipFilePath
    Open ZipFilePath For Output As #1
    Print #1, "PK" & Chr(5) & Chr(6) & String(18, 0)
    Close #1

    Set shellApp = CreateObject("Shell.Application")

    If IsArray(FilePaths) Then
        For i = LBound(FilePaths) To UBound(FilePaths)
            If FileExist(CStr(FilePaths(i))) Then
                shellApp.Namespace(ZipFilePath).CopyHere CStr(FilePaths(i))
                Do While shellApp.Namespace(ZipFilePath).Items.Count <= i
                    Application.Wait Now + TimeValue("00:00:01")
                Loop
            Else
                Debug.Print "Zip_Files: Skipping non-existent file: " & FilePaths(i)
            End If
        Next i
    ElseIf VarType(FilePaths) = vbString Then
        If FileExist(CStr(FilePaths)) Then
            shellApp.Namespace(ZipFilePath).CopyHere CStr(FilePaths)
            Application.Wait Now + TimeValue("00:00:01")
        End If
    End If

    Set shellApp = Nothing
    Exit Sub

ifError:
    Call HandleError("Failed to create zip file '" & ZipFilePath & "'. " & vbCrLf & "Error: " & Err.Description)
    If Not shellApp Is Nothing Then Set shellApp = Nothing
End Sub

Public Sub Create_Folder(ByVal DirectoryPath As String)
    Dim i As Integer
    Dim strArray() As String
    Dim strSubFolder As String

    DirectoryPath = Clean_FilePath(DirectoryPath)
    DirectoryPath = Get_DirectoryPath(DirectoryPath)

    strArray = Split(DirectoryPath, "\")
    strSubFolder = Trim(strArray(0))
    For i = 1 To UBound(strArray)
        strSubFolder = strSubFolder & "\" & Trim(strArray(i))
        If Dir(strSubFolder, vbDirectory) = "" Then MkDir strSubFolder
    Next
End Sub

Public Function RemoveFileIfExist(FileToDelete As String) As Boolean
    If FileExist(FileToDelete) Then
        SetAttr FileToDelete, vbNormal
        Kill FileToDelete
        RemoveFileIfExist = True
    Else
        RemoveFileIfExist = False
    End If
End Function

Public Function SaveStringToFile(strText As String, strFile As String) As Boolean
    If Create_Folder_For_File(strFile) Then
        SaveStringToFile = OverwriteTxt(strFile, strText)
    Else
        SaveStringToFile = False
    End If
End Function

Public Function Create_Folder_For_File(ByVal FilePath As String) As Boolean
    Dim DirectoryPath As String
    DirectoryPath = Get_DirectoryPath(FilePath)
    Create_Folder (DirectoryPath)
    Create_Folder_For_File = DirExist(DirectoryPath)
End Function

Public Function AppendStringToFile(strText As String, FilePath As String, Optional AddLineBreak As Boolean = False) As Boolean
    Dim fileNumber As Integer
    fileNumber = FreeFile()
    Open FilePath For Append As #fileNumber
    If AddLineBreak Then
        Print #fileNumber, vbCrLf & strText
    Else
        Print #fileNumber, strText
    End If
    Close #fileNumber
    AppendStringToFile = True
End Function
