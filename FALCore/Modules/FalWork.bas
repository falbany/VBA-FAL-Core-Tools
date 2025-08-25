Attribute VB_Name = "FalWork"

' **************************************************************************************
' Module    : FalWork
' Author    : Florent ALBANY
' Website   :
' Purpose   : Manipulation of Workbooks and Worksheets
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2025-07-20              Initial creation with workbook/worksheet functions.
' 2         2025-07-28              Add several functions and comments.
' 3         2025-08-25              Translated to English and refactored for clarity.
'---------------------------------------------------------------------------------------
' Dependencies:
' ~~~~~~~~~~~~~~~~
'   Microsoft Scripting Runtime (for Dictionary object)
'   FalFile
' **************************************************************************************

Option Explicit

Private Sub Handle_Error(ByVal ErrMsg As String)
    ' @brief Displays a standardized error message to the user.
    ' @param ErrMsg The specific error message to display.
    MsgBox "An error occurred:" & vbCrLf & ErrMsg, vbOKOnly + vbCritical, "Operation Impossible"
End Sub


Private Function Clean_FilePath(filePath As String) As String
    ' @brief Cleans a file path string by replacing forward slashes with backslashes.
    Clean_FilePath = Replace(Replace(filePath, "/", "\"), "\\", "\")
End Function


Private Function Normalize_Object_Variant_To_Collection(ByVal InputVariant As Variant, Byval ExpectedTypeName As String, ByVal CallerName As String) As Collection
    ' @brief Normalizes a Dictionary, Collection, or Array of objects into a single Collection.
    ' @param InputVariant The input data structure (Dictionary, Collection, or Array).
    ' @param ExpectedTypeName The string name of the object type to filter for (e.g., "Worksheet", "Workbook").
    ' @param CallerName The name of the public function calling this helper, for better error messages.
    ' @return A Collection containing only the objects of the expected type. Returns Nothing on invalid input type.
    On Error GoTo ifError

    Dim outputColl As New Collection
    Dim item As Variant

    ' Check if the input is something we can iterate over.
    If IsObject(InputVariant) Then
        If TypeOf InputVariant Is Dictionary Then
            For Each item In InputVariant.Items
                If TypeName(item) = ExpectedTypeName Then outputColl.Add item
            Next item
        Else ' Assumes it's a Collection
            For Each item In InputVariant
                If TypeName(item) = ExpectedTypeName Then outputColl.Add item
            Next item
        End If
    ElseIf IsArray(InputVariant) Then
        For Each item In InputVariant
            If TypeName(item) = ExpectedTypeName Then outputColl.Add item
        Next item
    Else
        ' If the input is not a valid iterable type, return Nothing to signal an error.
        Call Handle_Error("Invalid input for " & CallerName & ". Expected a Dictionary, Collection, or Array of " & ExpectedTypeName & "s.")
        Set Normalize_Object_Variant_To_Collection = Nothing
        Exit Function
    End If

    Set Normalize_Object_Variant_To_Collection = outputColl
    Exit Function

ifError:
    Call Handle_Error("An unexpected error occurred in " & CallerName & " while normalizing the input collection. " & vbCrLf & "Error: " & Err.Description)
    Set Normalize_Object_Variant_To_Collection = Nothing
End Function


Private Function Convert_Collection_To_Array(ByVal coll As Collection) As Variant
    ' @brief Converts a Collection object into a 0-based array.
    ' @param coll The Collection to convert.
    ' @return A Variant array containing the items from the collection.
    '         Returns an uninitialized Variant if the collection is empty or invalid.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    ' Return an uninitialized variant for Nothing or empty collections.
    If coll Is Nothing Or coll.Count = 0 Then Exit Function

    Dim arr() As Variant
    ReDim arr(0 To coll.Count - 1)
    Dim i As Long
    Dim item As Variant

    i = 0
    For Each item In coll
        ' Handle both objects and simple values correctly.
        If IsObject(item) Then
            Set arr(i) = item
        Else
            arr(i) = item
        End If
        i = i + 1
    Next item

    Convert_Collection_To_Array = arr
    Exit Function

ifError:
    Call Handle_Error("Failed to convert Collection to Array. " & vbCrLf & "Error: " & Err.Description)
    ' The function will implicitly return an uninitialized Variant on error.
End Function


Public Function Get_Last_Row(ByVal TargetSheet As Worksheet, Optional ByVal Col As Long = 1) As Long
    ' @brief Finds the last used row in a specified column of a worksheet.
    ' @param TargetSheet The worksheet to search in.
    ' @param Col (Optional) The column number to search within. Defaults to column 1 (A).
    ' @return The last used row number. Returns 0 if the sheet is empty or an error occurs.
    ' @details This function uses the .Find method, which is more reliable than .End(xlUp) for sheets with filtered or hidden rows.
    On Error Resume Next

    If TargetSheet Is Nothing Then Exit Function ' Returns 0

    Get_Last_Row = TargetSheet.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlFormulas).Row
    
    If Err.Number <> 0 Then Get_Last_Row = 0
    Err.Clear
End Function

Public Function Convert_Address_To_Coords(ByVal Address As String, Optional ByVal TargetSheet As Worksheet) As Variant
    ' @brief Converts an A1-style cell address string into its row and column coordinates.
    ' @param Address The A1-style address to convert (e.g., "A1", "$B$5", "XFD1048576").
    ' @param TargetSheet (Optional) The worksheet context for the address. If omitted, the ActiveSheet is used.
    ' @return A 1-based, 2-element array [Row, Column] on success. Returns Empty on failure (e.g., invalid address).
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    Dim tempRange As Range
    Dim result(1 To 2) As Long
    Dim refSheet As Worksheet

    ' 1. Validate input
    If Trim(Address) = "" Then
        Call Handle_Error("Address cannot be empty in Convert_Address_To_Coords.")
        Convert_Address_To_Coords = Empty
        Exit Function
    End If

    ' 2. Determine reference sheet
    If TargetSheet Is Nothing Then
        Set refSheet = ActiveSheet
    Else
        Set refSheet = TargetSheet
    End If

    If refSheet Is Nothing Then
        Call Handle_Error("No valid worksheet context available for Convert_Address_To_Coords.")
        Convert_Address_To_Coords = Empty
        Exit Function
    End If

    ' 3. Convert address to range and get coordinates
    On Error Resume Next ' To catch invalid address strings
    Set tempRange = refSheet.Range(Address)
    If Err.Number <> 0 Then
        Call Handle_Error("Invalid address string provided: '" & Address & "'.")
        Convert_Address_To_Coords = Empty
        Exit Function
    End If
    On Error GoTo ifError

    result(1) = tempRange.Row
    result(2) = tempRange.Column

    Convert_Address_To_Coords = result
    Exit Function

ifError:
    Call Handle_Error("Failed to convert address '" & Address & "' to coordinates. " & vbCrLf & "Error: " & Err.Description)
    Convert_Address_To_Coords = Empty
End Function

Public Function Clear_Range(ByVal TargetRange As Range, Optional ByVal ClearType As String = "Contents") As Boolean
    ' @brief Clears a specified range using a given method.
    ' @param TargetRange The Range object to clear.
    ' @param ClearType (Optional) The method to use for clearing. Valid options: "All", "Contents", "Formats", "Comments", "Notes", "Outline". Defaults to "Contents".
    ' @return True if the range was cleared successfully, False otherwise.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    If TargetRange Is Nothing Then
        Call Handle_Error("Invalid Range object provided to Clear_Range.")
        Exit Function
    End If

    Select Case LCase(ClearType)
        Case "all": TargetRange.Clear
        Case "contents": TargetRange.ClearContents
        Case "formats": TargetRange.ClearFormats
        Case "comments": TargetRange.ClearComments
        Case "notes": TargetRange.ClearNotes
        Case "outline": TargetRange.ClearOutline
        Case Else
            Call Handle_Error("Invalid 'ClearType' provided: '" & ClearType & "'. Defaulting to 'Contents'.")
            TargetRange.ClearContents
    End Select

    Clear_Range = True
    Exit Function

ifError:
    Call Handle_Error("Failed to clear range on worksheet '" & TargetRange.Parent.Name & "'. " & vbCrLf & "Error: " & Err.Description)
    Clear_Range = False
End Function

Public Function Get_Last_Column(ByVal TargetSheet As Worksheet, Optional ByVal Row As Long = 1) As Long
    ' @brief Finds the last used column in a specified row of a worksheet.
    ' @param TargetSheet The worksheet to search in.
    ' @param Row (Optional) The row number to search within. Defaults to row 1.
    ' @return The last used column number. Returns 0 if the sheet is empty or an error occurs.
    ' @details This function uses the .Find method for reliability.
    On Error Resume Next

    If TargetSheet Is Nothing Then Exit Function ' Returns 0

    Get_Last_Column = TargetSheet.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlFormulas).Column

    If Err.Number <> 0 Then Get_Last_Column = 0
    Err.Clear
End Function


Public Function Convert_Coords_To_Address(ByVal RowIndex As Long, ByVal ColumnIndex As Long) As String
    ' @brief Converts row and column numbers to an A1-style cell address string.
    ' @param RowIndex The row number (must be > 0).
    ' @param ColumnIndex The column number (must be > 0).
    ' @return The A1-style address (e.g., "A1", "B10"). Returns an empty string if indices are invalid.
    ' @details This function provides a safe and reliable way to get a cell address without needing a specific worksheet object.
    On Error Resume Next ' In case of invalid indices that slip past the check

    ' 1. Input Validation
    If RowIndex < 1 Or ColumnIndex < 1 Then
        Convert_Coords_To_Address = ""
        Exit Function
    End If

    ' 2. Use the built-in Address property for reliable conversion
    Convert_Coords_To_Address = Application.Cells(RowIndex, ColumnIndex).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    If Err.Number <> 0 Then Convert_Coords_To_Address = ""
    Err.Clear
End Function

Public Function Open_Workbook(ByVal wbk_filepath As String) As Workbook
    ' @brief Opens a workbook from a specified path, or returns a reference if it's already open.
    ' @param wbk_filepath The full path to the workbook file.
    ' @return A Workbook object if successful, otherwise Nothing.
    ' @details
    ' This function first checks if a workbook with the same name is already open.
    ' If it is, it returns a reference to that workbook object.
    ' If not, it attempts to open the file from the specified path.
    ' If the file does not exist or an error occurs, it returns Nothing.
    On Error GoTo ifError

    Dim wbkName As String
    Dim wbk As Workbook

    ' Clean the file path for consistency and extract the filename
    wbk_filepath = Clean_FilePath(wbk_filepath)
    wbkName = Mid$(wbk_filepath, InStrRev(wbk_filepath, "\") + 1)

    ' Temporarily ignore errors to check if the workbook is already in the collection
    On Error Resume Next
    Set wbk = Application.Workbooks(wbkName)
    On Error GoTo ifError ' Restore the main error handler

    If wbk Is Nothing Then
        ' Workbook is not open, so open it. The main error handler will catch issues like file-not-found.
        Set Open_Workbook = Application.Workbooks.Open(wbk_filepath)
    Else
        ' Workbook is already open, return the existing object reference.
        Set Open_Workbook = wbk
    End If

    Exit Function
ifError:
    ' Use the module's existing error handler for consistent error reporting
    Call Handle_Error("Failed to open workbook '" & wbk_filepath & "'. " & vbCrLf & "Error: " & Err.Description)
    Set Open_Workbook = Nothing
End Function


Public Function Open_Workbooks_As_Dictionary(ByVal wbk_filepaths As Variant) As Dictionary
    ' @brief Opens a list of workbooks from specified paths.
    ' @param wbk_filepaths An array of strings, where each string is a full path to a workbook file.
    ' @return A Dictionary object containing the successfully opened Workbook objects, keyed by their full file path.
    '         Returns an empty Dictionary if the input is invalid or no workbooks could be opened.
    ' @dependencies
    ' - Microsoft Scripting Runtime (for Dictionary object)
    ' - Open_Workbook
    ' - Clean_FilePath
    ' - Handle_Error
    
    On Error GoTo ifError
    
    Dim dictWorkbooks As New Dictionary
    Dim wbk_filepath As Variant
    Dim wbk As Workbook
    Dim cleanedPath As String
    
    ' Set dictionary to be case-insensitive for file paths, which is standard for Windows.
    dictWorkbooks.CompareMode = TextCompare
    
    ' Check if the input is a valid array
    If Not IsArray(wbk_filepaths) Then
        Call Handle_Error("Input to Open_Workbooks_As_Dictionary was not a valid array of file paths.")
        Set Open_Workbooks_As_Dictionary = dictWorkbooks ' Return empty dictionary
        Exit Function
    End If
    
    ' Loop through the provided file paths
    For Each wbk_filepath In wbk_filepaths
        ' Ensure the element is a non-empty string before processing
        If VarType(wbk_filepath) = vbString And wbk_filepath <> "" Then
            ' Use the existing single-workbook-opener function which handles all checks
            Set wbk = Open_Workbook(CStr(wbk_filepath))
            
            ' If the workbook was opened successfully, add it to the dictionary
            If Not wbk Is Nothing Then
                ' Key the dictionary with the cleaned, full path for consistency.
                cleanedPath = Clean_FilePath(CStr(wbk_filepath))
                If Not dictWorkbooks.Exists(cleanedPath) Then
                    dictWorkbooks.Add Key:=cleanedPath, Item:=wbk
                End If
            End If
            ' Reset wbk object for the next loop iteration
            Set wbk = Nothing
        End If
    Next wbk_filepath
    
    ' Return the dictionary of opened workbooks
    Set Open_Workbooks_As_Dictionary = dictWorkbooks
    
    Exit Function
    
ifError:
    Call Handle_Error("An unexpected error occurred in Open_Workbooks_As_Dictionary function." & vbCrLf & "Error: " & Err.Description)
    ' On a fatal error, returning Nothing is appropriate to signal failure of the function itself.
    Set Open_Workbooks_As_Dictionary = Nothing
End Function


Public Function Open_Workbooks_As_Collection(ByVal wbk_filepaths As Variant) As Collection
    ' @brief Opens a list of workbooks and returns them in a Collection.
    ' @param wbk_filepaths An array of strings, where each string is a full path to a workbook file.
    ' @return A Collection object containing the successfully opened Workbook objects.
    '         The key for each item in the collection is its full file path.
    '         Returns an empty Collection if the input is invalid or no workbooks could be opened.
    ' @dependencies
    ' - Open_Workbook
    ' - Clean_FilePath
    ' - Handle_Error
    
    On Error GoTo ifError
    
    Dim collWorkbooks As New Collection
    Dim wbk_filepath As Variant
    Dim wbk As Workbook
    Dim cleanedPath As String
    
    ' Check if the input is a valid array
    If Not IsArray(wbk_filepaths) Then
        Call Handle_Error("Input to Open_Workbooks_As_Collection was not a valid array of file paths.")
        Set Open_Workbooks_As_Collection = collWorkbooks ' Return empty collection
        Exit Function
    End If
    
    ' Loop through the provided file paths
    For Each wbk_filepath In wbk_filepaths
        If VarType(wbk_filepath) = vbString And wbk_filepath <> "" Then
            Set wbk = Open_Workbook(CStr(wbk_filepath))
            
            If Not wbk Is Nothing Then
                cleanedPath = Clean_FilePath(CStr(wbk_filepath))
                ' Use path as key to prevent duplicates, ignore error if key already exists
                On Error Resume Next
                collWorkbooks.Add Item:=wbk, Key:=cleanedPath
                On Error GoTo ifError
            End If
            Set wbk = Nothing
        End If
    Next wbk_filepath
    
    Set Open_Workbooks_As_Collection = collWorkbooks
    
    Exit Function
    
ifError:
    Call Handle_Error("An unexpected error occurred in Open_Workbooks_As_Collection function." & vbCrLf & "Error: " & Err.Description)
    Set Open_Workbooks_As_Collection = Nothing
End Function


Public Function Open_Workbooks_As_Array(ByVal wbk_filepaths As Variant) As Workbook()
    ' @param wbk_filepaths An array of strings, where each string is a full path to a workbook file.
    ' @return An array of Workbook objects. Returns an uninitialized array if an error occurs or no workbooks are opened.
    ' @dependencies
    ' - Open_Workbook
    ' - Clean_FilePath
    ' - Handle_Error
    
    On Error GoTo ifError
    
    Dim collTempWbks As New Collection
    Dim openedWbks() As Workbook
    Dim wbk_filepath As Variant
    Dim wbk As Workbook
    Dim cleanedPath As String
    Dim i As Long
    
    ' Check if the input is a valid array
    If Not IsArray(wbk_filepaths) Then
        Call Handle_Error("Input to Open_Workbooks_As_Array was not a valid array of file paths.")
        Exit Function ' Returns an uninitialized array
    End If
    
    ' Loop through the provided file paths and open them, storing unique ones in a temporary collection
    For Each wbk_filepath In wbk_filepaths
        If VarType(wbk_filepath) = vbString And wbk_filepath <> "" Then
            Set wbk = Open_Workbook(CStr(wbk_filepath))
            
            If Not wbk Is Nothing Then
                cleanedPath = Clean_FilePath(CStr(wbk_filepath))
                ' Use path as key to prevent duplicates; ignore error if key already exists
                On Error Resume Next
                collTempWbks.Add Item:=wbk, Key:=cleanedPath
                On Error GoTo ifError ' Restore main error handler
            End If
            Set wbk = Nothing
        End If
    Next wbk_filepath
    
    ' If any workbooks were successfully opened, convert the collection to an array
    If collTempWbks.Count > 0 Then
        ReDim openedWbks(0 To collTempWbks.Count - 1)
        i = 0
        For Each wbk In collTempWbks
            Set openedWbks(i) = wbk
            i = i + 1
        Next wbk
        Open_Workbooks_As_Array = openedWbks
    End If
    
    Exit Function
ifError:
    Call Handle_Error("An unexpected error occurred in Open_Workbooks_As_Array function." & vbCrLf & "Error: " & Err.Description)
    ' The function will implicitly return an uninitialized array on error
End Function


Public Function Sort_Workbooks(ByVal WorkbooksToSort As Variant, ByVal SortKey As String, Optional ByVal SortOrder As XlSortOrder = xlAscending) As Workbook()
    ' @brief Sorts a collection of Workbook objects based on specified criteria.
    ' @param WorkbooksToSort A Dictionary, Collection, or Array of Workbook objects.
    ' @param SortKey The criteria to sort by. Valid options: "Name", "FullName", "Date".
    ' @param SortOrder (Optional) The sort direction, xlAscending (default) or xlDescending.
    ' @return An array of Workbook objects, sorted as specified.
    '         Returns an uninitialized array on failure or if no valid workbooks are provided.
    ' @dependencies Normalize_Object_Variant_To_Collection, Handle_Error (Internal)

    On Error GoTo ifError

    ' 1. Normalize the input into a single collection of workbook objects.
    Dim wbksToProcess As Collection
    Set wbksToProcess = Normalize_Object_Variant_To_Collection(WorkbooksToSort, "Workbook", "Sort_Workbooks")
    If wbksToProcess Is Nothing Or wbksToProcess.Count = 0 Then
        Exit Function ' Return uninitialized array
    End If

    Dim lb As Long: lb = 1
    Dim ub As Long: ub = wbksToProcess.Count
    Dim i As Long, j As Long

    ' 2. Prepare data structure for sorting: [SortValue, OriginalWorkbookObject]
    Dim sortData() As Variant
    ReDim sortData(lb To ub, 1 To 2)

    ' 3. Populate the sort data array
    Dim wbk As Workbook
    i = lb
    For Each wbk In wbksToProcess
        Set sortData(i, 2) = wbk ' Store the object itself

        Select Case LCase(SortKey)
            Case "name": sortData(i, 1) = wbk.Name
            Case "fullname": sortData(i, 1) = wbk.FullName
            Case "date"
                If wbk.path <> "" Then sortData(i, 1) = FileDateTime(wbk.FullName) Else sortData(i, 1) = 0
            Case "sheetcount"
                sortData(i, 1) = wbk.Worksheets.Count
            Case Else
                Call Handle_Error("Invalid 'SortKey' provided to Sort_Workbooks: '" & SortKey & "'. Valid options are 'Name', 'FullName', 'Date', 'SheetCount'.")
                Exit Function
        End Select
        i = i + 1
    Next wbk

    ' 4. Sort the array using a bubble sort
    For i = lb To ub - 1
        For j = i + 1 To ub
            If (SortOrder = xlAscending And sortData(i, 1) > sortData(j, 1)) Or _
               (SortOrder = xlDescending And sortData(i, 1) < sortData(j, 1)) Then
                
                Dim tempSortValue As Variant: tempSortValue = sortData(i, 1)
                Dim tempWorkbook As Workbook: Set tempWorkbook = sortData(i, 2)

                sortData(i, 1) = sortData(j, 1)
                Set sortData(i, 2) = sortData(j, 2)

                sortData(j, 1) = tempSortValue
                Set sortData(j, 2) = tempWorkbook
            End If
        Next j
    Next i

    ' 5. Prepare the result array
    Dim result() As Workbook: ReDim result(0 To ub - 1)
    For i = lb To ub: Set result(i - 1) = sortData(i, 2): Next i
    Sort_Workbooks = result
    Exit Function

ifError:
    Call Handle_Error("An unexpected error occurred in Sort_Workbooks. " & vbCrLf & "Error: " & Err.Description)
End Function



Public Function Close_Workbook(ByVal wbk As Workbook, Optional ByVal SaveChanges As Boolean = True) As Boolean
    ' @brief Safely closes a workbook.
    ' @param wbk The Workbook object to close.
    ' @param SaveChanges (Optional) True to save changes before closing, False to discard them. Default is True.
    ' @return True if the workbook was closed successfully or was already closed, False if an error occurred.
    ' @details
    ' This function checks if the provided workbook object is valid and if the workbook is still open
    ' before attempting to close it. This prevents runtime errors if the workbook has been closed
    ' by another process or if the object variable is invalid.

    On Error GoTo ifError

    Dim wbkName As String

    If wbk Is Nothing Then
        Close_Workbook = True
        Exit Function
    End If

    On Error Resume Next
    wbkName = wbk.Name ' Attempt to access a property to see if the object is valid.
    If Err.Number <> 0 Then
        Close_Workbook = True ' Success, as the workbook is already closed or the object is invalid.
        Exit Function
    End If
    On Error GoTo ifError ' Restore the main error handler.
    
    wbk.Close SaveChanges:=SaveChanges
    Close_Workbook = True

    Exit Function

ifError:
    Call Handle_Error("Failed to close workbook '" & wbkName & "'. " & vbCrLf & "Error: " & Err.Description)
    Close_Workbook = False
End Function


Public Function Close_Workbooks(ByVal WorkbooksToClose As Variant, Optional ByVal SaveChanges As Boolean = True) As Boolean
    ' @brief Safely closes a collection of workbooks.
    ' @param WorkbooksToClose A Dictionary, Collection, or Array of Workbook objects to close.
    ' @param SaveChanges (Optional) True to save changes before closing, False to discard them. Default is True.
    ' @return True if all workbooks were closed successfully, False otherwise.
    ' @dependencies
    ' - Close_Workbook
    ' - Handle_Error

    On Error GoTo ifError

    Dim wbkItem As Variant
    Dim overallSuccess As Boolean
    overallSuccess = True ' Assume success until a failure occurs

    ' Check if the input is something we can iterate over.
    If Not IsObject(WorkbooksToClose) And Not IsArray(WorkbooksToClose) Then
        Call Handle_Error("Invalid input to Close_Workbooks. Expected a Dictionary, Collection, or Array.")
        Close_Workbooks = False
        Exit Function
    End If

    ' Iterate through the provided collection/dictionary.
    If TypeOf WorkbooksToClose Is Dictionary Then
        For Each wbkItem In WorkbooksToClose.Items ' For Dictionaries, we iterate the .Items collection
            If TypeOf wbkItem Is Workbook Then
                If Not Close_Workbook(wbkItem, SaveChanges) Then overallSuccess = False
            End If
        Next wbkItem
    Else ' For Collections and Arrays
        For Each wbkItem In WorkbooksToClose
            If TypeOf wbkItem Is Workbook Then
                If Not Close_Workbook(wbkItem, SaveChanges) Then overallSuccess = False
            End If
        Next wbkItem
    End If

    Close_Workbooks = overallSuccess
    Exit Function

ifError:
    Call Handle_Error("An unexpected error occurred in Close_Workbooks function." & vbCrLf & "Error: " & Err.Description)
    Close_Workbooks = False
End Function


Public Function Save_Workbook_As(ByVal TargetWorkbook As Workbook, ByVal FilePath As String, Optional ByVal FileFormat As XlFileFormat = xlWorkbookDefault, Optional ByVal AutoClose As Boolean = False) As Boolean
    ' @brief Saves a workbook to a specified path with an optional file format.
    ' @param TargetWorkbook The Workbook object to save.
    ' @param FilePath The full path (including filename and extension) to save the workbook to.
    ' @param FileFormat (Optional) The Excel file format to use (e.g., xlOpenXMLWorkbook, xlExcel8). Defaults to xlWorkbookDefault.
	' @param AutoClose (Optional) If True, the workbook will be closed after saving. Defaults to False.
    ' @return True if the workbook was saved successfully, False otherwise.
    ' @details This function provides a way to explicitly save a workbook to a new location.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    Dim wbName As String

    If TargetWorkbook Is Nothing Then
        Call Handle_Error("Invalid workbook object provided to Save_Workbook_As.")
        Exit Function
    End If

    wbName = TargetWorkbook.Name

    ' Use the SaveAs method to save the workbook with the specified file format.
    TargetWorkbook.SaveAs Filename:=FilePath, FileFormat:=FileFormat, ConflictResolution:=xlLocalSessionChanges

    ' Check if the save operation was successful.
    If FalFile.FileExist(FilePath) Then
		If AutoClose Then
			TargetWorkbook.Close SaveChanges:=False
		End If
	
        Save_Workbook_As = True
    Else
        Save_Workbook_As = False
        Call Handle_Error("Failed to save workbook '" & wbName & "' to '" & FilePath & "'.")
        Exit Function
    End If

    Exit Function

ifError:
    Call Handle_Error("Failed to save workbook '" & wbName & "' to '" & FilePath & "'. " & vbCrLf & "Error: " & Err.Description)
    Save_Workbook_As = False
End Function


Public Function Save_Workbooks_As(ByVal WorkbooksToSave As Variant, ByVal BasePath As String, Optional ByVal FileFormat As XlFileFormat = xlWorkbookDefault, Optional ByVal AutoClose As Boolean = False) As Boolean
    ' @brief Saves a collection of workbooks to a specified directory, using the workbook name as the filename.
    ' @param WorkbooksToSave A Dictionary, Collection, or Array of Workbook objects to save.
    ' @param BasePath The base directory where the workbooks will be saved.
    ' @param FileFormat (Optional) The Excel file format to use (e.g., xlOpenXMLWorkbook, xlExcel8). Defaults to xlWorkbookDefault.
    ' @param AutoClose (Optional) If True, the workbook will be closed after saving. Defaults to False.
    ' @return True if all workbooks were saved successfully, False otherwise.
    ' @details This function iterates through the provided workbooks, saves each one to the specified directory
    '          using the workbook's name, and optionally closes them.
    ' @dependencies Normalize_Object_Variant_To_Collection, Save_Workbook_As, Handle_Error (Internal)
    On Error GoTo ifError

    Dim wbksToProcess As Collection
    Dim wbk As Workbook
    Dim overallSuccess As Boolean
    Dim saveFile As String
    Dim i As Long

    ' 1. Normalize the input into a single collection of workbook objects.
    Set wbksToProcess = Normalize_Object_Variant_To_Collection(WorkbooksToSave, "Workbook", "Save_Workbooks_As")
    If wbksToProcess Is Nothing Then
        Save_Workbooks_As = False
        Exit Function
    End If

    overallSuccess = True ' Assume success until a failure occurs

    ' 2. Iterate through the workbooks and save them.
    For Each wbk In wbksToProcess
        ' Create the full file path using the provided directory and workbook name.
        saveFile = FalFile.Combine_Paths(BasePath, wbk.Name)

        ' Use the Save_Workbook_As function to handle the actual saving.
        If Not Save_Workbook_As(wbk, saveFile, FileFormat, AutoClose) Then
            overallSuccess = False
        End If
    Next wbk

    ' 3. Return the result
    Save_Workbooks_As = overallSuccess
    Exit Function

ifError:
    Call Handle_Error("An unexpected error occurred in Save_Workbooks_As. " & vbCrLf & "Error: " & Err.Description)
    Save_Workbooks_As = False
End Function


Public Function Is_Workbook_Open(ByVal wbkIdentifier As String) As Boolean
    ' @brief Checks if a workbook is currently open in the Excel application.
    ' @param wbkIdentifier The name of the workbook (e.g., "Book1.xlsx") or its full path.
    ' @return True if the workbook is open, False otherwise.
    ' @details This function is case-insensitive.
    On Error Resume Next
    Dim wbkName As String

    ' If a full path is provided, extract just the filename.
    If InStrRev(wbkIdentifier, "\") > 0 Then
        wbkName = Mid$(wbkIdentifier, InStrRev(wbkIdentifier, "\") + 1)
    Else
        wbkName = wbkIdentifier
    End If

    ' The core of the check: try to get a reference to the workbook by name.
    ' If it succeeds, .Name will not raise an error. If it fails, Err.Number will be non-zero.
    Is_Workbook_Open = (Len(Application.Workbooks(wbkName).Name) > 0)
    Err.Clear ' Clear any error that occurred if the workbook was not found.
End Function


Public Function Get_Workbook(ByVal wbkIdentifier As String) As Workbook
    ' @brief Retrieves a reference to an already open workbook without opening it from disk.
    ' @param wbkIdentifier The name of the workbook (e.g., "Book1.xlsx") or its full path.
    ' @return A Workbook object if found and open, otherwise Nothing.
    ' @details This function does NOT open a workbook from disk; it only checks the collection of currently open workbooks.
    On Error Resume Next
    Dim wbkName As String
    
    ' If a full path is provided, extract just the filename.
    If InStrRev(wbkIdentifier, "\") > 0 Then
        wbkName = Mid$(wbkIdentifier, InStrRev(wbkIdentifier, "\") + 1)
    Else
        wbkName = wbkIdentifier
    End If
    
    Set Get_Workbook = Application.Workbooks(wbkName)
    Err.Clear ' Clear error if workbook not found, function will correctly return Nothing.
End Function


Public Function Create_Workbook(Optional ByVal MakeVisible As Boolean = True) As Workbook
    ' @brief Creates a new, empty workbook.
    ' @param MakeVisible (Optional) If True, the new workbook will be visible. Defaults to True.
    ' @return A Workbook object representing the newly created workbook. Returns Nothing on failure.
    ' @dependencies Handle_Error
    On Error GoTo ifError

    Dim wbk As Workbook
    Set wbk = Application.Workbooks.Add

    If Not wbk Is Nothing Then
        wbk.Windows(1).Visible = MakeVisible
    End If

    Set Create_Workbook = wbk
    Exit Function
    
ifError:
    Call Handle_Error("Failed to create a new workbook. " & vbCrLf & "Error: " & Err.Description)
    Set Create_Workbook = Nothing
End Function


Public Function Create_Workbooks(ByVal WorkbooksToCreate As Variant, Optional ByVal MakeVisible As Boolean = True) As Workbook()
    ' @brief Creates multiple new, empty workbooks.
    ' @param WorkbooksToCreate An array of strings for workbook titles, or a number indicating how many to create.
    ' @param MakeVisible (Optional) If True, the new workbooks will be visible. Defaults to True.
    ' @return An array of the newly created Workbook objects. Returns an uninitialized array on failure or if the input is empty/invalid.
    ' @details If an array of names is provided, the .Title property of each new workbook will be set. This title is then suggested as the filename upon saving.
    ' @dependencies Create_Workbook, Handle_Error (Internal)

    On Error GoTo ifError

    Dim tempColl As New Collection
    Dim createdWbks() As Workbook
    Dim item As Variant
    Dim wbk As Workbook
    Dim i As Long

    ' 1. Input Validation and processing
    If IsArray(WorkbooksToCreate) Then
        ' --- Input is an array of names ---
        For Each item In WorkbooksToCreate
            Set wbk = Create_Workbook(MakeVisible)
            If Not wbk Is Nothing Then
                ' Set the title property if a name is provided
                If VarType(item) = vbString And item <> "" Then
                    wbk.Title = CStr(item)
                    ' wbk.Name = CStr(item)
                End If
                tempColl.Add wbk
            End If
        Next item
    ElseIf IsNumeric(WorkbooksToCreate) Then
        ' --- Input is a number ---
        If CLng(WorkbooksToCreate) > 0 Then
            For i = 1 To CLng(WorkbooksToCreate)
                Set wbk = Create_Workbook(MakeVisible)
                If Not wbk Is Nothing Then
                    tempColl.Add wbk
                End If
            Next i
        End If
    Else
        ' --- Invalid input type ---
        Call Handle_Error("Input to Create_Workbooks must be an array of names or a number.")
        Exit Function ' Returns uninitialized array
    End If

    ' 2. Convert collection to array
    If tempColl.Count > 0 Then
        ReDim createdWbks(0 To tempColl.Count - 1)
        i = 0
        For Each wbk In tempColl
            Set createdWbks(i) = wbk
            i = i + 1
        Next wbk
        Create_Workbooks = createdWbks
    End If
    Exit Function

ifError:
    Call Handle_Error("An unexpected error occurred in Create_Workbooks: " & Err.Description)
    ' Implicitly returns an uninitialized array
End Function


Public Function Merge_Workbooks(ByVal SourceWorkbooks As Variant, Optional ByVal DestinationWorkbook As Workbook = Nothing) As Workbook
    ' @brief Merges all worksheets from multiple source workbooks into a single destination workbook.
    ' @param SourceWorkbooks A Dictionary, Collection, or Array of Workbook objects to merge.
    ' @param DestinationWorkbook (Optional) The workbook to merge into. If omitted, a new workbook is created.
    ' @return The destination Workbook object if successful, otherwise Nothing.
    ' @details Excel automatically handles sheet name collisions by appending a number (e.g., "Sheet1 (2)").
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    Dim workbooksToProcess As Collection
    Dim destWbk As Workbook
    Dim sourceWbk As Workbook
    Dim ws As Worksheet
    Dim createdNewWorkbook As Boolean
    Dim defaultSheetCount As Long

    ' 1. Normalize the input into a single collection.
    Set workbooksToProcess = Normalize_Object_Variant_To_Collection(SourceWorkbooks, "Workbook", "Merge_Workbooks")
    If workbooksToProcess Is Nothing Then Exit Function

    If workbooksToProcess.Count = 0 Then
        Set Merge_Workbooks = DestinationWorkbook
        Exit Function
    End If

    ' 2. Determine the destination workbook.
    If DestinationWorkbook Is Nothing Then
        Set destWbk = Application.Workbooks.Add
        createdNewWorkbook = True
        defaultSheetCount = destWbk.Worksheets.Count
    Else
        Set destWbk = DestinationWorkbook
    End If

    Application.ScreenUpdating = False

    ' 3. Loop through workbooks and copy sheets.
    For Each sourceWbk In workbooksToProcess
        For Each ws In sourceWbk.Worksheets
            ws.Copy After:=destWbk.Sheets(destWbk.Sheets.Count)
        Next ws
    Next sourceWbk

    ' 4. Clean up default sheets if a new workbook was created.
    If createdNewWorkbook Then
        Application.DisplayAlerts = False
        Dim i As Long
        For i = 1 To defaultSheetCount
            destWbk.Worksheets(1).Delete
        Next i
        Application.DisplayAlerts = True
    End If

    Set Merge_Workbooks = destWbk

ifExit:
    Application.ScreenUpdating = True
    Exit Function

ifError:
    Call Handle_Error("Failed to merge workbooks. " & vbCrLf & "Error: " & Err.Description)
    Set Merge_Workbooks = Nothing
    Resume ifExit
End Function

Public Function Create_Worksheet(ByVal SheetName As String, Optional ByVal TargetWorkbook As Workbook = Nothing) As Worksheet
    ' @brief Ensures a worksheet with a specific name exists in a workbook, creating it if necessary.
    ' @param SheetName The desired name for the worksheet.
    ' @param TargetWorkbook (Optional) The workbook to add the sheet to. If omitted, the ActiveWorkbook is used.
    ' @return The created or existing Worksheet object. Returns Nothing on failure.
    ' @details This function is a robust replacement for CheckCreate_Sheet. It correctly handles the target workbook,
    '          and includes error handling for invalid sheet names.
    ' @dependencies Handle_Error

    On Error GoTo ifError

    Dim wbk As Workbook
    Dim ws As Worksheet

    ' 1. Determine the target workbook
    If TargetWorkbook Is Nothing Then
        Set wbk = ActiveWorkbook
    Else
        Set wbk = TargetWorkbook
    End If

    If wbk Is Nothing Then
        Call Handle_Error("No valid workbook was specified or active.")
        Exit Function ' Returns Nothing
    End If

    ' 2. Check if the sheet already exists
    On Error Resume Next
    Set ws = wbk.Worksheets(SheetName)
    On Error GoTo ifError

    ' 3. If sheet doesn't exist, create and name it
    If ws Is Nothing Then
        Set ws = wbk.Worksheets.Add(After:=wbk.Worksheets(wbk.Worksheets.Count))
        On Error Resume Next ' Temporarily disable errors for the naming attempt
        ws.Name = SheetName
        If Err.Number <> 0 Then ' Naming failed (e.g., invalid characters, name too long)
            Application.DisplayAlerts = False
            ws.Delete ' Clean up the failed sheet
            Application.DisplayAlerts = True
            Call Handle_Error("Failed to create worksheet. The name '" & SheetName & "' is invalid.")
            Set Create_Worksheet = Nothing
            Exit Function
        End If
        On Error GoTo ifError ' Re-enable main error handler
    End If

    ' 4. Return the worksheet object
    Set Create_Worksheet = ws
    Exit Function

ifError:
    Call Handle_Error("An unexpected error occurred in Create_Worksheet: " & Err.Description)
    Set Create_Worksheet = Nothing
End Function


Public Function Create_Worksheets_As_Dictionary(ByVal SheetNames As Variant, Optional ByVal TargetWorkbook As Workbook = Nothing) As Dictionary
    ' @brief Creates multiple worksheets in a workbook and returns them as a Dictionary.
    ' @param SheetNames An array of strings, where each string is a desired worksheet name.
    ' @param TargetWorkbook (Optional) The workbook to add the sheets to. If omitted, the ActiveWorkbook is used.
    ' @return A Dictionary object containing the successfully created/retrieved Worksheet objects, keyed by their name.
    ' @dependencies Create_Worksheet, Handle_Error, Microsoft Scripting Runtime
    
    On Error GoTo ifError
    
    Dim dictSheets As New Dictionary
    Dim sheetName As Variant
    Dim ws As Worksheet
    
    dictSheets.CompareMode = TextCompare ' Sheet names are not case-sensitive
    
    If Not IsArray(SheetNames) Then
        Call Handle_Error("Input to Create_Worksheets_As_Dictionary was not a valid array of sheet names.")
        Set Create_Worksheets_As_Dictionary = dictSheets ' Return empty dictionary
        Exit Function
    End If
    
    For Each sheetName In SheetNames
        If VarType(sheetName) = vbString And sheetName <> "" Then
            Set ws = Create_Worksheet(CStr(sheetName), TargetWorkbook)
            
            If Not ws Is Nothing Then
                If Not dictSheets.Exists(ws.Name) Then
                    dictSheets.Add Key:=ws.Name, Item:=ws
                End If
            End If
            Set ws = Nothing
        End If
    Next sheetName
    
    Set Create_Worksheets_As_Dictionary = dictSheets
    
    Exit Function
    
ifError:
    Call Handle_Error("An unexpected error occurred in Create_Worksheets_As_Dictionary: " & Err.Description)
    Set Create_Worksheets_As_Dictionary = Nothing
End Function


Public Function Create_Worksheets_As_Collection(ByVal SheetNames As Variant, Optional ByVal TargetWorkbook As Workbook = Nothing) As Collection
    ' @brief Creates multiple worksheets in a workbook and returns them as a Collection.
    ' @param SheetNames An array of strings, where each string is a desired worksheet name.
    ' @param TargetWorkbook (Optional) The workbook to add the sheets to. If omitted, the ActiveWorkbook is used.
    ' @return A Collection object containing the successfully created/retrieved Worksheet objects, keyed by their name.
    ' @dependencies Create_Worksheet, Handle_Error
    
    On Error GoTo ifError
    
    Dim collSheets As New Collection
    Dim sheetName As Variant
    Dim ws As Worksheet
    
    If Not IsArray(SheetNames) Then
        Call Handle_Error("Input to Create_Worksheets_As_Collection was not a valid array of sheet names.")
        Set Create_Worksheets_As_Collection = collSheets ' Return empty collection
        Exit Function
    End If
    
    For Each sheetName In SheetNames
        If VarType(sheetName) = vbString And sheetName <> "" Then
            Set ws = Create_Worksheet(CStr(sheetName), TargetWorkbook)
            
            If Not ws Is Nothing Then
                On Error Resume Next
                collSheets.Add Item:=ws, Key:=ws.Name
                On Error GoTo ifError
            End If
            Set ws = Nothing
        End If
    Next sheetName
    
    Set Create_Worksheets_As_Collection = collSheets
    
    Exit Function
    
ifError:
    Call Handle_Error("An unexpected error occurred in Create_Worksheets_As_Collection: " & Err.Description)
    Set Create_Worksheets_As_Collection = Nothing
End Function


Public Function Create_Worksheets_As_Array(ByVal SheetNames As Variant, Optional ByVal TargetWorkbook As Workbook = Nothing) As Worksheet()
    ' @brief Creates multiple worksheets in a workbook and returns them as an array.
    ' @param SheetNames An array of strings, where each string is a desired worksheet name.
    ' @param TargetWorkbook (Optional) The workbook to add the sheets to. If omitted, the ActiveWorkbook is used.
    ' @return An array of Worksheet objects. Returns an uninitialized array if an error occurs or no sheets are created.
    ' @dependencies Create_Worksheet, Handle_Error
    
    On Error GoTo ifError
    
    Dim tempColl As New Collection
    Dim createdSheets() As Worksheet
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim i As Long
    
    If Not IsArray(SheetNames) Then
        Call Handle_Error("Input to Create_Worksheets_As_Array was not a valid array of sheet names.")
        Exit Function ' Returns uninitialized array
    End If
    
    ' Use a temporary collection to gather unique, successfully created sheets
    For Each sheetName In SheetNames
        If VarType(sheetName) = vbString And sheetName <> "" Then
            Set ws = Create_Worksheet(CStr(sheetName), TargetWorkbook)
            
            If Not ws Is Nothing Then
                On Error Resume Next
                tempColl.Add Item:=ws, Key:=ws.Name ' Use name as key to avoid duplicates
                On Error GoTo ifError
            End If
            Set ws = Nothing
        End If
    Next sheetName
    
    ' Convert the collection to an array
    If tempColl.Count > 0 Then
        ReDim createdSheets(0 To tempColl.Count - 1)
        i = 0
        For Each ws In tempColl
            Set createdSheets(i) = ws
            i = i + 1
        Next ws
        Create_Worksheets_As_Array = createdSheets
    End If
    
    Exit Function
    
ifError:
    Call Handle_Error("An unexpected error occurred in Create_Worksheets_As_Array: " & Err.Description)
    ' Implicitly returns an uninitialized array
End Function


Public Function Get_Worksheet(ByVal SheetIdentifier As Variant, Optional ByVal TargetWorkbook As Workbook = Nothing) As Worksheet
    ' @brief Retrieves a reference to a worksheet from a workbook without creating it.
    ' @param SheetIdentifier The name (String) or index (Long) of the worksheet to find.
    ' @param TargetWorkbook (Optional) The workbook to search within. If omitted, the ActiveWorkbook is used.
    ' @return A Worksheet object if found, otherwise Nothing.
    ' @details This function is useful for checking the existence of a sheet or getting a reference to it
    '          without the side-effect of creating it, unlike Create_Worksheet.
    On Error Resume Next ' Use Resume Next to handle the case where the sheet doesn't exist.

    Dim wbk As Workbook

    ' 1. Determine the target workbook
    If TargetWorkbook Is Nothing Then
        Set wbk = ActiveWorkbook
    Else
        Set wbk = TargetWorkbook
    End If

    If wbk Is Nothing Then
        ' No workbook to search in, so we can't find a sheet.
        Set Get_Worksheet = Nothing
        Exit Function
    End If

    ' 2. Attempt to get the worksheet object.
    ' If the sheet doesn't exist, this will set Get_Worksheet to Nothing without raising a fatal error.
    Set Get_Worksheet = wbk.Worksheets(SheetIdentifier)
    
    ' 3. Clear any error that occurred if the sheet was not found.
    Err.Clear
End Function


Public Function Copy_Worksheet(ByVal SourceSheet As Worksheet, ByVal DestinationWorkbook As Workbook, Optional ByVal NewSheetName As String = "") As Worksheet
    ' @brief Copies a worksheet to another (or the same) workbook.
    ' @param SourceSheet The worksheet object to copy.
    ' @param DestinationWorkbook The workbook where the sheet will be copied.
    ' @param NewSheetName (Optional) The name for the newly copied sheet. If omitted, Excel provides a default name.
    ' @return The newly created Worksheet object if successful, otherwise Nothing.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    Dim copiedSheet As Worksheet

    If SourceSheet Is Nothing Or DestinationWorkbook Is Nothing Then
        Call Handle_Error("Invalid source sheet or destination workbook provided to Copy_Worksheet.")
        Exit Function
    End If

    ' Perform the copy operation. The new sheet becomes the active sheet in the destination workbook.
    SourceSheet.Copy After:=DestinationWorkbook.Sheets(DestinationWorkbook.Sheets.Count)
    Set copiedSheet = DestinationWorkbook.Sheets(DestinationWorkbook.Sheets.Count)

    ' Rename the new sheet if a name is provided.
    If NewSheetName <> "" Then
        On Error Resume Next ' Handle potential errors with invalid names
        copiedSheet.Name = NewSheetName
        If Err.Number <> 0 Then
            Call Handle_Error("Failed to rename the copied sheet to '" & NewSheetName & "'. It may be an invalid or duplicate name. The sheet was copied with a default name.")
            ' We don't exit the function, as the copy itself was successful.
            Err.Clear
        End If
        On Error GoTo ifError
    End If

    Set Copy_Worksheet = copiedSheet
    Exit Function

ifError:
    Call Handle_Error("Failed to copy worksheet '" & SourceSheet.Name & "'. " & vbCrLf & "Error: " & Err.Description)
    Set Copy_Worksheet = Nothing
End Function


Public Function Delete_Worksheet(ByVal SheetToDelete As Worksheet) As Boolean
    ' @brief Safely deletes a worksheet, suppressing confirmation prompts.
    ' @param SheetToDelete The worksheet object to be deleted.
    ' @return True if the sheet was deleted successfully, False otherwise.
    ' @details The function will fail if you try to delete the last visible sheet in a workbook.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    Dim targetWorkbook As Workbook
    Dim visibleSheetCount As Long
    Dim ws As Worksheet

    If SheetToDelete Is Nothing Then
        Call Handle_Error("Invalid worksheet object provided to Delete_Worksheet.")
        Exit Function ' Returns False
    End If

    Set targetWorkbook = SheetToDelete.Parent

    ' Prevent deleting the last visible sheet in a visible workbook
    If targetWorkbook.Windows(1).Visible Then
        For Each ws In targetWorkbook.Worksheets
            If ws.Visible = xlSheetVisible Then visibleSheetCount = visibleSheetCount + 1
        Next ws
        If visibleSheetCount <= 1 And SheetToDelete.Visible = xlSheetVisible Then
            Call Handle_Error("Cannot delete the last visible worksheet.")
            Exit Function
        End If
    End If

    Application.DisplayAlerts = False
    SheetToDelete.Delete
    Application.DisplayAlerts = True
    Delete_Worksheet = True

    Exit Function

ifError:
    Application.DisplayAlerts = True ' Always restore alerts on error
    Call Handle_Error("Failed to delete worksheet. " & vbCrLf & "Error: " & Err.Description)
    Delete_Worksheet = False
End Function


Public Function Activate_Worksheet(ByVal TargetSheet As Worksheet) As Boolean
    ' @brief Safely activates a given worksheet, also activating its parent workbook if necessary.
    ' @param TargetSheet The worksheet object to activate.
    ' @return True if activation was successful, False otherwise.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    If TargetSheet Is Nothing Then
        Call Handle_Error("Invalid worksheet object provided to Activate_Worksheet.")
        Exit Function
    End If

    ' Activate the parent workbook first
    TargetSheet.Parent.Activate
    ' Then activate the sheet
    TargetSheet.Activate

    Activate_Worksheet = True
    Exit Function

ifError:
    Call Handle_Error("Failed to activate worksheet '" & TargetSheet.Name & "'. " & vbCrLf & "Error: " & Err.Description)
    Activate_Worksheet = False
End Function


Public Function Convert_Worksheet_To_Array(ByVal TargetSheet As Worksheet, Optional ByVal ReadProperty As String = "Value") As Variant
    ' @brief Converts the used range of a worksheet into a 2D array, reading either values or formulas.
    ' @param TargetSheet The worksheet to convert.
    ' @param ReadProperty (Optional) The property to read from cells. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return A 2D Variant array containing the data from the worksheet's used range.
    '         Returns an empty Variant on failure or if the sheet is empty.
    ' @details This is the most efficient method to read a sheet's data into memory.
    '          It also handles the edge case where the used range is only a single cell,
    '          ensuring a 2D array is always returned for consistency.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    Dim arrData As Variant

    If TargetSheet Is Nothing Then
        Call Handle_Error("Invalid worksheet object provided to Convert_Worksheet_To_Array.")
        Convert_Worksheet_To_Array = Empty
        Exit Function
    End If

    ' Check if the sheet has any data
    If Application.WorksheetFunction.CountA(TargetSheet.Cells) = 0 Then
        Convert_Worksheet_To_Array = Empty ' Return Empty for a blank sheet
        Exit Function
    End If

    ' Read the specified property from the range
    Select Case LCase(ReadProperty)
        Case "value", "values"
            arrData = TargetSheet.UsedRange.Value
        Case "formula", "formulas"
            arrData = TargetSheet.UsedRange.Formula
        Case "formular1c1"
            arrData = TargetSheet.UsedRange.FormulaR1C1
        Case Else
            Call Handle_Error("Invalid 'ReadProperty' specified: '" & ReadProperty & "'. Defaulting to 'Value'.")
            arrData = TargetSheet.UsedRange.Value
    End Select

    ' Handle the case where UsedRange is only a single cell, which returns a scalar value.
    If Not IsArray(arrData) Then
        Dim tempValue As Variant: tempValue = arrData ' Store the scalar value/formula
        ReDim arrData(1 To 1, 1 To 1): arrData(1, 1) = tempValue
    End If

    Convert_Worksheet_To_Array = arrData
    Exit Function

ifError:
    Call Handle_Error("Failed to convert worksheet '" & TargetSheet.Name & "' to an array. " & vbCrLf & "Error: " & Err.Description)
    Convert_Worksheet_To_Array = Empty
End Function


Public Function Convert_Worksheets_To_3DArray(ByVal TargetSheets As Variant, Optional ByVal ReadProperty As String = "Value") As Variant
    ' @brief Converts a collection of worksheets into a single 3D array, reading either values or formulas.
    ' @param TargetSheets A Dictionary, Collection, or Array of Worksheet objects.
    ' @param ReadProperty (Optional) The property to read from cells. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return A 1-based 3D Variant array where the first dimension represents the sheet,
    '         the second represents the row, and the third represents the column.
    '         Returns an empty Variant on failure or if no valid sheets are provided.
    ' @details The dimensions of the resulting 3D array are determined by the worksheet
    '          with the most rows and the one with the most columns among all sheets in the collection.
    ' @dependencies Convert_Worksheet_To_Array, Handle_Error (Internal)
    On Error GoTo ifError

    Dim sheetsToProcess As New Collection
    Dim tempSheetArrays As New Collection
    Dim item As Variant
    Dim ws As Worksheet
    Dim arr2D As Variant
    Dim arr3D As Variant
    Dim maxRows As Long, maxCols As Long
    Dim numSheets As Long
    Dim i As Long, r As Long, c As Long

    ' 1. Normalize the input into a single collection of worksheet objects.
    Set sheetsToProcess = Normalize_Object_Variant_To_Collection(TargetSheets, "Worksheet", "Convert_Worksheets_To_3DArray")
    If sheetsToProcess Is Nothing Or sheetsToProcess.Count = 0 Then
        Convert_Worksheets_To_3DArray = Empty ' Return empty array
        Exit Function
    End If

    ' 2. First pass: Determine max dimensions and store 2D arrays temporarily.
    maxRows = 0
    maxCols = 0
    For Each ws In sheetsToProcess
        arr2D = Convert_Worksheet_To_Array(ws, ReadProperty)
        If IsArray(arr2D) Then
            tempSheetArrays.Add arr2D
            If UBound(arr2D, 1) > maxRows Then maxRows = UBound(arr2D, 1)
            If UBound(arr2D, 2) > maxCols Then maxCols = UBound(arr2D, 2)
        End If
    Next ws

    numSheets = tempSheetArrays.Count
    If numSheets = 0 Then
        Convert_Worksheets_To_3DArray = Empty ' All sheets were empty
        Exit Function
    End If

    ' 3. Dimension the final 3D array.
    ReDim arr3D(1 To numSheets, 1 To maxRows, 1 To maxCols)

    ' 4. Second pass: Populate the 3D array from the temporary 2D arrays.
    i = 1
    For Each arr2D In tempSheetArrays
        For r = LBound(arr2D, 1) To UBound(arr2D, 1)
            For c = LBound(arr2D, 2) To UBound(arr2D, 2)
                arr3D(i, r, c) = arr2D(r, c)
            Next c
        Next r
        i = i + 1
    Next arr2D

    Convert_Worksheets_To_3DArray = arr3D
    Exit Function

ifError:
    Call Handle_Error("Failed to convert worksheets to a 3D array. " & vbCrLf & "Error: " & Err.Description)
    Convert_Worksheets_To_3DArray = Empty
End Function


Public Function Convert_Range_To_Array(ByVal TargetRange As Range, Optional ByVal ReadProperty As String = "Value") As Variant
    ' @brief Converts a Range object into a 2D array, reading either values or formulas.
    ' @param TargetRange The Range to convert.
    ' @param ReadProperty (Optional) The property to read from cells. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return A 2D Variant array containing the data from the range.
    '         Returns an empty Variant on failure or if the range is invalid.
    ' @details This is the most efficient method to read a range's data into memory.
    '          It also handles the edge case where the range is only a single cell,
    '          ensuring a 2D array is always returned for consistency.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    Dim arrData As Variant

    If TargetRange Is Nothing Then
        Call Handle_Error("Invalid Range object provided to Convert_Range_To_Array.")
        Convert_Range_To_Array = Empty
        Exit Function
    End If

    ' Read the specified property from the range
    Select Case LCase(ReadProperty)
        Case "value", "values"
            arrData = TargetRange.Value
        Case "formula", "formulas"
            arrData = TargetRange.Formula
        Case "formular1c1"
            arrData = TargetRange.FormulaR1C1
        Case Else
            Call Handle_Error("Invalid 'ReadProperty' specified: '" & ReadProperty & "'. Defaulting to 'Value'.")
            arrData = TargetRange.Value
    End Select

    ' Handle the case where the Range is only a single cell, which returns a scalar value.
    If Not IsArray(arrData) Then
        Dim tempValue As Variant: tempValue = arrData ' Store the scalar value/formula
        ReDim arrData(1 To 1, 1 To 1): arrData(1, 1) = tempValue
    End If

    Convert_Range_To_Array = arrData
    Exit Function

ifError:
    Call Handle_Error("Failed to convert Range to an array. " & vbCrLf & "Error: " & Err.Description)
    Convert_Range_To_Array = Empty
End Function


Public Function Convert_Worksheet_In_Workbooks_To_3DArray(ByVal TargetWorkbooks As Variant, ByVal SheetIdentifier As Variant, Optional ByVal ReadProperty As String = "Value") As Variant
    ' @brief Extracts a specific sheet from a collection of workbooks and combines them into a single 3D array, reading either values or formulas.
    ' @param TargetWorkbooks A Dictionary, Collection, or Array of Workbook objects.
    ' @param SheetIdentifier The name (String) or index (Long) of the worksheet to extract from each workbook.
    ' @param ReadProperty (Optional) The property to read from cells. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return A 1-based 3D Variant array where the first dimension represents the workbook/sheet,
    '         the second represents the row, and the third represents the column.
    '         Returns an empty Variant on failure or if no valid sheets are found.
    ' @details The dimensions of the resulting 3D array are determined by the worksheet
    '          with the most rows and the one with the most columns among all sheets found.
    '          Workbooks that do not contain the specified sheet are skipped.
    ' @dependencies Worksheets_To_3DArray, Handle_Error (Internal)
    On Error GoTo ifError

    Dim workbooksToProcess As Collection
    Dim sheetsToProcess As New Collection
    Dim wbk As Workbook
    Dim ws As Worksheet

    ' 1. Normalize the workbook input into a single collection.
    Set workbooksToProcess = Normalize_Object_Variant_To_Collection(TargetWorkbooks, "Workbook", "Convert_Worksheet_In_Workbooks_To_3DArray")
    If workbooksToProcess Is Nothing Then
        Convert_Worksheet_In_Workbooks_To_3DArray = Empty
        Exit Function
    End If

    If workbooksToProcess.Count = 0 Then
        Convert_Worksheet_In_Workbooks_To_3DArray = Empty
        Exit Function
    End If

    ' 2. Iterate through workbooks to find and collect the specified sheet.
    For Each wbk In workbooksToProcess
        On Error Resume Next ' Ignore errors if a sheet doesn't exist in a workbook
        Set ws = wbk.Worksheets(SheetIdentifier)
        On Error GoTo ifError ' Restore main error handler

        If Not ws Is Nothing Then
            sheetsToProcess.Add ws
            Set ws = Nothing ' Reset for next loop
        End If
    Next wbk

    ' 3. Pass the collected sheets to the existing function to create the 3D array.
    Convert_Worksheet_In_Workbooks_To_3DArray = Convert_Worksheets_To_3DArray(sheetsToProcess, ReadProperty)
    Exit Function

ifError:
    Call Handle_Error("Failed to convert workbook sheets to a 3D array. " & vbCrLf & "Error: " & Err.Description)
    Convert_Worksheet_In_Workbooks_To_3DArray = Empty
End Function


Public Function Get_Worksheets_From_Workbooks_As_Dictionary(ByVal TargetWorkbooks As Variant, ByVal SheetIdentifier As Variant) As Dictionary
    ' @brief Extracts a specific sheet from a collection of workbooks and returns them as a Dictionary.
    ' @param TargetWorkbooks A Dictionary, Collection, or Array of Workbook objects.
    ' @param SheetIdentifier The name (String) or index (Long) of the worksheet to extract from each workbook.
    ' @return A Dictionary object containing the successfully found Worksheet objects, keyed by a composite key ("WorkbookName|SheetName").
    '         Returns an empty Dictionary on failure or if no valid sheets are found.
    ' @details Workbooks that do not contain the specified sheet are skipped. The dictionary keys are a composite of the
    '          workbook name and sheet name (e.g., "Book1.xlsx|Sheet1") to ensure uniqueness.
    ' @dependencies Handle_Error (Internal), Microsoft Scripting Runtime
    On Error GoTo ifError

    Dim dictSheets As New Dictionary
    Dim workbooksToProcess As Collection
    Dim wbk As Workbook
    Dim ws As Worksheet
    Dim compositeKey As String

    dictSheets.CompareMode = TextCompare ' Sheet names are not case-sensitive

    ' 1. Normalize the workbook input into a single collection.
    Set workbooksToProcess = Normalize_Object_Variant_To_Collection(TargetWorkbooks, "Workbook", "WorkbooksSheet_To_Dictionary")
    If workbooksToProcess Is Nothing Then
        Set Get_Worksheets_From_Workbooks_As_Dictionary = dictSheets ' Return empty dictionary
        Exit Function
    End If

    If workbooksToProcess.Count = 0 Then
        Set Get_Worksheets_From_Workbooks_As_Dictionary = dictSheets
        Exit Function
    End If

    ' 2. Iterate through workbooks to find and collect the specified sheet.
    For Each wbk In workbooksToProcess
        On Error Resume Next ' Ignore errors if a sheet doesn't exist in a workbook
        Set ws = wbk.Worksheets(SheetIdentifier)
        On Error GoTo ifError ' Restore main error handler

        If Not ws Is Nothing Then
            ' Create a composite key from the workbook and worksheet names to guarantee uniqueness.
            compositeKey = wbk.Name & "|" & ws.Name
            If Not dictSheets.Exists(compositeKey) Then
                dictSheets.Add Key:=compositeKey, Item:=ws
            End If
            Set ws = Nothing ' Reset for next loop
        End If
    Next wbk

    ' 3. Return the dictionary of sheets.
    Set Get_Worksheets_From_Workbooks_As_Dictionary = dictSheets
    Exit Function

ifError:
    Call Handle_Error("Failed to create Dictionary from workbook sheets. " & vbCrLf & "Error: " & Err.Description)
    Set Get_Worksheets_From_Workbooks_As_Dictionary = Nothing
End Function


Public Function Get_Worksheets_From_Workbooks_As_Collection(ByVal TargetWorkbooks As Variant, ByVal SheetIdentifier As Variant) As Collection
    ' @brief Extracts a specific sheet from a collection of workbooks and returns them as a Collection.
    ' @param TargetWorkbooks A Dictionary, Collection, or Array of Workbook objects.
    ' @param SheetIdentifier The name (String) or index (Long) of the worksheet to extract from each workbook.
    ' @return A Collection object containing the successfully found Worksheet objects, keyed by their name to ensure uniqueness.
    '         Returns an empty Collection on failure or if no valid sheets are found.
    ' @details Workbooks that do not contain the specified sheet are skipped.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    Dim collSheets As New Collection
    Dim workbooksToProcess As Collection
    Dim wbk As Workbook
    Dim ws As Worksheet
    Dim compositeKey As String

    ' 1. Normalize the workbook input into a single collection.
    Set workbooksToProcess = Normalize_Object_Variant_To_Collection(TargetWorkbooks, "Workbook", "Get_Worksheets_From_Workbooks_As_Collection")
    If workbooksToProcess Is Nothing Then
        Set Get_Worksheets_From_Workbooks_As_Collection = collSheets ' Return empty collection
        Exit Function
    End If

    If workbooksToProcess.Count = 0 Then
        Set Get_Worksheets_From_Workbooks_As_Collection = collSheets
        Exit Function
    End If

    ' 2. Iterate through workbooks to find and collect the specified sheet.
    For Each wbk In workbooksToProcess
        On Error Resume Next ' Ignore errors if a sheet doesn't exist in a workbook
        Set ws = wbk.Worksheets(SheetIdentifier)
        On Error GoTo ifError ' Restore main error handler

        If Not ws Is Nothing Then
            compositeKey = wbk.Name & "|" & ws.Name
            'if not collSheets.Exists(compositeKey) then
            '    collSheets.Add Item:=ws, Key:=compositeKey
            'end if
            On Error GoTo ifError
            Set ws = Nothing ' Reset for next loop
        End If
    Next wbk

    ' 3. Return the collection of sheets.
    Set Get_Worksheets_From_Workbooks_As_Collection = collSheets
    Exit Function

ifError:
    Call Handle_Error("Failed to create Collection from workbook sheets. " & vbCrLf & "Error: " & Err.Description)
    Set Get_Worksheets_From_Workbooks_As_Collection = Nothing
End Function


Public Function Get_Range_From_Worksheet(ByVal TargetSheet As Worksheet, Optional ByVal TopLeft As String = "", Optional ByVal BottomRight As String = "") As Range
    ' @brief Extracts a specific Range object from a worksheet. If addresses are omitted, returns all cells.
    ' @param TargetSheet The worksheet from which to extract the range.
    ' @param TopLeft (Optional) The address of the top-left cell of the range (e.g., "A1").
    ' @param BottomRight (Optional) The address of the bottom-right cell of the range (e.g., "D10").
    ' @return A Range object. If TopLeft/BottomRight are specified, it's that area. If omitted, it's all cells. Returns Nothing on failure.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    ' 1. Input Validation
    If TargetSheet Is Nothing Then
        Call Handle_Error("Invalid worksheet object provided to Get_Range_From_Worksheet.")
        Set Get_Range_From_Worksheet = Nothing
        Exit Function
    End If

    ' 2. Determine which range to get
    If Trim(TopLeft) = "" Or Trim(BottomRight) = "" Then
        ' If either address is missing, return all cells in the sheet.
        Set Get_Range_From_Worksheet = TargetSheet.Cells
    Else
        ' If both addresses are provided, get the specific range.
        Set Get_Range_From_Worksheet = TargetSheet.Range(TopLeft, BottomRight)
    End If
    Exit Function

ifError:
    Call Handle_Error("Failed to get range from worksheet '" & TargetSheet.Name & "'. " & vbCrLf & "Error: " & Err.Description)
    Set Get_Range_From_Worksheet = Nothing
End Function


Public Function Write_Array_To_Worksheet(ByVal DataArray As Variant, ByVal TargetCell As Range, Optional ByVal WriteAs As String = "Value") As Boolean
    ' @brief Writes a 2D array to a worksheet starting at a specified cell, as values or formulas.
    ' @param DataArray The 2D array containing the data to write.
    ' @param TargetCell The top-left cell of the destination range.
    ' @param WriteAs (Optional) The property to write. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return True if the write operation was successful, False otherwise.
    ' @details This is the most efficient method to write an array to a sheet, as it avoids cell-by-cell looping.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    Dim destRange As Range
    Dim numRows As Long
    Dim numCols As Long

    If Not IsArray(DataArray) Then
        Call Handle_Error("Input 'DataArray' is not a valid array.")
        Exit Function
    End If

    If TargetCell Is Nothing Then
        Call Handle_Error("Invalid 'TargetCell' provided.")
        Exit Function
    End If

    numRows = UBound(DataArray, 1) - LBound(DataArray, 1) + 1
    numCols = UBound(DataArray, 2) - LBound(DataArray, 2) + 1

    ' Define the destination range based on the array size
    Set destRange = TargetCell.Resize(numRows, numCols)

    ' Write the array to the range in one operation using the specified property
    Select Case LCase(WriteAs)
        Case "value", "values"
            destRange.Value = DataArray
        Case "formula", "formulas"
            destRange.Formula = DataArray
        Case "formular1c1"
            destRange.FormulaR1C1 = DataArray
        Case Else
            Call Handle_Error("Invalid 'WriteAs' property '" & WriteAs & "'. Defaulting to 'Value'.")
            destRange.Value = DataArray
    End Select
    
    Write_Array_To_Worksheet = True
    Exit Function

ifError:
    Call Handle_Error("Failed to write array to worksheet '" & TargetCell.Parent.Name & "'. " & vbCrLf & "Error: " & Err.Description)
    Write_Array_To_Worksheet = False
End Function


Public Function Write_Range_To_Worksheet(ByVal SourceRange As Range, ByVal TargetCell As Range, Optional ByVal WriteAs As String = "Value") As Boolean
    ' @brief Writes data from a source range to a worksheet, as values or formulas.
    ' @param SourceRange The Range object containing the data to write.
    ' @param TargetCell The top-left cell of the destination range.
    ' @param WriteAs (Optional) The property to write. Valid options: "Value", "Formula", "FormulaR1C1". Defaults to "Value".
    ' @return True if the write operation was successful, False otherwise.
    ' @details This is an efficient method to copy data from one range to another, as it avoids cell-by-cell looping.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    Dim destRange As Range

    If SourceRange Is Nothing Then
        Call Handle_Error("Invalid 'SourceRange' provided.")
        Exit Function
    End If

    If TargetCell Is Nothing Then
        Call Handle_Error("Invalid 'TargetCell' provided.")
        Exit Function
    End If

    ' Define the destination range based on the source range size
    Set destRange = TargetCell.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)

    ' Write the source range's data to the destination range using the specified property
    Select Case LCase(WriteAs)
        Case "value", "values"
            destRange.Value = SourceRange.Value
        Case "formula", "formulas"
            destRange.Formula = SourceRange.Formula
        Case "formular1c1"
            destRange.FormulaR1C1 = SourceRange.FormulaR1C1
        Case Else
            Call Handle_Error("Invalid 'WriteAs' property '" & WriteAs & "'. Defaulting to 'Value'.")
            destRange.Value = SourceRange.Value
    End Select

    Write_Range_To_Worksheet = True
    Exit Function

ifError:
    Call Handle_Error("Failed to write range to worksheet '" & TargetCell.Parent.Name & "'. " & vbCrLf & "Error: " & Err.Description)
    Write_Range_To_Worksheet = False
End Function

Public Function Find_In_Worksheet(ByVal WhatToFind As Variant, ByVal TargetSheet As Worksheet, Optional LookIn As XlFindLookIn = xlValues, Optional LookAt As XlLookAt = xlWhole, Optional MatchCase As Boolean = False) As Range
    ' @brief Finds the first occurrence of a value within a worksheet and returns the cell as a Range object.
    ' @param WhatToFind The value to search for.
    ' @param TargetSheet The worksheet to search within.
    ' @param LookIn (Optional) Specifies whether to search in formulas, values, or comments. Defaults to xlValues.
    ' @param LookAt (Optional) Specifies whether to match the whole cell or part of it. Defaults to xlWhole.
    ' @param MatchCase (Optional) True for a case-sensitive search. Defaults to False.
    ' @return The Range object of the found cell. Returns Nothing if the value is not found.
    ' @details This is a robust wrapper for the Range.Find method.
    
    On Error Resume Next ' Use Resume Next because .Find returns an error if nothing is found, which we handle by checking for Nothing.

    If TargetSheet Is Nothing Then Exit Function ' Returns Nothing

    Set Find_In_Worksheet = TargetSheet.Cells.Find(What:=WhatToFind, _
                                               LookIn:=LookIn, _
                                               LookAt:=LookAt, _
                                               SearchOrder:=xlByRows, _
                                               SearchDirection:=xlNext, _
                                               MatchCase:=MatchCase)
    Err.Clear
End Function

Public Function Protect_Worksheet(ByVal TargetSheet As Worksheet, Optional ByVal Password As String) As Boolean
    ' @brief Protects a worksheet, optionally with a password.
    ' @param TargetSheet The worksheet to protect.
    ' @param Password (Optional) The password to use for protection.
    ' @return True if the sheet was protected successfully, False otherwise.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    If TargetSheet Is Nothing Then
        Call Handle_Error("Invalid worksheet object provided to Protect_Worksheet.")
        Exit Function
    End If

    TargetSheet.Protect Password:=Password
    Protect_Worksheet = True
    Exit Function

ifError:
    Call Handle_Error("Failed to protect worksheet '" & TargetSheet.Name & "'. " & vbCrLf & "Error: " & Err.Description)
    Protect_Worksheet = False
End Function


Public Function Unprotect_Worksheet(ByVal TargetSheet As Worksheet, Optional ByVal Password As String) As Boolean
    ' @brief Unprotects a worksheet, optionally with a password.
    ' @param TargetSheet The worksheet to unprotect.
    ' @param Password (Optional) The password required to unprotect the sheet.
    ' @return True if the sheet was unprotected successfully, False otherwise.
    ' @dependencies Handle_Error (Internal)
    On Error GoTo ifError

    If TargetSheet Is Nothing Then
        Call Handle_Error("Invalid worksheet object provided to Unprotect_Worksheet.")
        Exit Function
    End If

    TargetSheet.Unprotect Password:=Password
    Unprotect_Worksheet = True
    Exit Function

ifError:
    Call Handle_Error("Failed to unprotect worksheet '" & TargetSheet.Name & "'. " & vbCrLf & "Error: " & Err.Description)
    Unprotect_Worksheet = False
End Function


Public Function Merge_Worksheets(ByVal SourceSheets As Variant, ByVal DestinationSheet As Worksheet, Optional ByVal HasHeaders As Boolean = True) As Boolean
    ' @brief Merges data from multiple source worksheets into a single destination worksheet.
    ' @param SourceSheets A Dictionary, Collection, or Array of Worksheet objects to merge.
    ' @param DestinationSheet The worksheet where the merged data will be placed.
    ' @param HasHeaders (Optional) If True, the header from the first sheet is used, and headers from subsequent sheets are skipped. Defaults to True.
    ' @return True if the merge was successful, False otherwise.
    ' @dependencies Get_Last_Row, Handle_Error (Internal)
    On Error GoTo ifError

    Dim sheetsToProcess As New Collection
    Dim item As Variant
    Dim ws As Worksheet
    Dim sourceRange As Range
    Dim lastRow As Long
    Dim nextRow As Long
    Dim isFirstSheet As Boolean

    ' 1. Validate inputs and normalize into a single collection.
    If DestinationSheet Is Nothing Then
        Call Handle_Error("A valid DestinationSheet must be provided.")
        Exit Function
    End If

    If IsObject(SourceSheets) Then
        If TypeOf SourceSheets Is Dictionary Then
            For Each item In SourceSheets.Items
                If TypeOf item Is Worksheet Then sheetsToProcess.Add item
            Next item
        Else ' Assumes it's a Collection
            For Each item In SourceSheets
                If TypeOf item Is Worksheet Then sheetsToProcess.Add item
            Next item
        End If
    ElseIf IsArray(SourceSheets) Then
        For Each item In SourceSheets
            If TypeOf item Is Worksheet Then sheetsToProcess.Add item
        Next item
    Else
        Call Handle_Error("Invalid input for SourceSheets. Expected a Dictionary, Collection, or Array.")
        Exit Function
    End If

    If sheetsToProcess.Count = 0 Then
        Merge_Worksheets = True ' Nothing to do, so it's a success.
        Exit Function
    End If

    ' 2. Prepare destination sheet and start merging.
    DestinationSheet.Cells.Clear
    nextRow = 1
    isFirstSheet = True

    Application.ScreenUpdating = False

    For Each ws In sheetsToProcess
        lastRow = Get_Last_Row(ws)
        If lastRow > 0 Then
            Dim startRow As Long
            startRow = 1
            ' Adjust start row to skip header if necessary
            If HasHeaders And Not isFirstSheet Then
                startRow = 2
            End If

            ' If there are rows to copy (i.e., not just a header that we're skipping)
            If lastRow >= startRow Then
                Set sourceRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, ws.UsedRange.Columns.Count))
                sourceRange.Copy DestinationSheet.Cells(nextRow, 1)
                nextRow = nextRow + sourceRange.Rows.Count
            End If
        End If
        isFirstSheet = False
    Next ws

    Merge_Worksheets = True

ifExit:
    Application.ScreenUpdating = True
    Exit Function

ifError:
    Call Handle_Error("Failed to merge worksheets. " & vbCrLf & "Error: " & Err.Description)
    Merge_Worksheets = False
    Resume ifExit
End Function

Public Function SheetExist(ByVal wbk As Workbook, ByVal SheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wbk.Sheets(SheetName)
    SheetExist = (Not ws Is Nothing)
End Function
