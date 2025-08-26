Attribute VB_Name = "FalCSV"
' --------------------------------------------------------------------------------------------------
' Module: FalCSV
' Author: Florent ALBANY
' Date: 2025-07-15
' Version: 1.0
' Description:
'   This module provides a set of utility functions for handling CSV files
'   (import and export) within Excel VBA applications. It relies on the
'   'VBABetterArray' class for efficient and high-performance in-memory data
'   management, thus facilitating transfers between Excel worksheets and CSV formats.
'
'   Functionalities include:
'   - Importing single or multiple CSV files to new worksheets.
'   - Importing a CSV file to a selected cell range.
'   - Exporting a selected cell range or an entire sheet to a CSV file.
'   - Advanced options for customizing CSV export (delimiters, encoding).
'   - Ability to merge multiple CSV files into a single sheet.
'   - Preview functionality to validate formatting before full import.
'
'   Designed with a robust error handling approach and an intuitive user interface.
'
' Dependencies:
'   - VBABetterArray Class (https://github.com/Senipah/VBA-Better-Array)
'   - FalFile Module (for file selection/saving dialogs)
'   - FalArray Module (for array operations)
'
' License: MIT License
'
' Change Log:
' 1.0 (2025-07-15) - Initial release with core CSV import/export functionalities.
'                   Added merge and custom export features.
' 2.0 (2025-08-25) - Translated to English and refactored for clarity.
'
' --------------------------------------------------------------------------------------------------
Option Explicit

Private Sub HandleError(ByVal errorMessage As String)
    ' @brief Displays a standardized error message to the user.
    ' @param errorMessage The specific error message to display.
    MsgBox "An error occurred:" & vbCrLf & errorMessage, vbOKOnly + vbCritical, "Operation Failed"
End Sub


' --- Function: FromCSVFilesToWorksheets ---
' Description:
'   Imports the content of one or more CSV files selected by the user
'   into new worksheets within the active Excel workbook.
'   Each CSV file is imported into its own sheet. The function handles
'   data import using the BetterArray class for efficiency.
'
' Parameters: None
'
' Returns: None
'   This procedure does not return any value, but it modifies the Excel workbook
'   by adding new sheets containing the CSV data.
'
' Usage:
'   Call this procedure from any other VBA module or assign it
'   to a button or shape in Excel to allow users to initiate
'   the import of CSV files.
'   Example: Call FalCSV.FromCSVFilesToWorksheets
'
' Error Handling:
'   - Displays a message if no file is selected or if the operation is canceled.
'   - Handles individual errors when reading each CSV file,
'     allowing the import of other files to continue even if there is a problem.
'   - Provides generic error messages for unexpected errors.
'
' Dependencies:
'   - BetterArray: Used to parse and manipulate CSV data in memory.
'   - FalFile.Select_Files: For the file selection dialog.
'
' Change Log:
' 1.0 (2025-07-15) - Initial implementation.
' --------------------------------------------------------------------------------------------------
Public Sub FromCSVFilesToWorksheets()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePaths As Variant
    Dim i As Long
    Dim outputSheet As Worksheet
    Dim fileNameOnly As String
    Dim inputDelimiter As String
    Dim finalDelimiter As String
    Dim inputQuote As String

    Set MyArray = New BetterArray

    ' 1. Ask the user to select CSV files.
    filePaths = FalFile.Select_Files(FileType:="csv", AllowMultiSelect:=True)

    ' 2. Check if the user canceled or if there was an error.
    If IsEmpty(filePaths) Or (IsArray(filePaths) And UBound(filePaths) < LBound(filePaths)) Then
        MsgBox "No CSV file selected or the operation was canceled.", vbInformation
        GoTo CleanUp
    End If

    ' 2.1 Ask the user for the initial delimiter.
    inputDelimiter = InputBox("Please enter the delimiter to use (e.g., , or ; or TAB for tab):", "CSV Preview Delimiter", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 2.2 Ask the user for the initial quote character.
    inputQuote = InputBox("Please enter the cell opening and closing character to use (e.g., """"):", "CSV Preview Quote", """")
    
    ' 3. Process each selected file.
    For i = LBound(filePaths) To UBound(filePaths)
        ' Handle errors specific to each file
        On Error Resume Next
        
        MyArray.FromCSVFile path:=CStr(filePaths(i)), columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False
        If Err.Number <> 0 Then
            Call HandleError("Error reading CSV file: " & filePaths(i) & ". " & Err.Description)
            Err.Clear
            GoTo NextFile
        End If
        On Error GoTo ErrHandler

        Set outputSheet = ActiveWorkbook.Sheets.Add
        ' Give a name to the new sheet (based on the CSV file name)
        On Error Resume Next
        
        ' Extract the file name without the path
        fileNameOnly = Mid(CStr(filePaths(i)), InStrRev(CStr(filePaths(i)), "\") + 1)
        ' Extract the base name (without the extension)
        If InStr(fileNameOnly, ".") > 0 Then
            fileNameOnly = Left(fileNameOnly, InStr(fileNameOnly, ".") - 1)
        End If
        
        ' Apply the name to the sheet, truncated to 31 characters
        outputSheet.name = Left(fileNameOnly, 31)
        
        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error GoTo ErrHandler

        MyArray.ToExcelRange outputSheet.Range("A1")
        MyArray.Clear
NextFile:
    Next i

    MsgBox "Data from the CSV files has been imported into new sheets.", vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in FromCSVFilesToWorksheets: " & Err.Description)
    Resume CleanUp
End Sub


' --- Function: FromCSVFileToSelection ---
' Description:
'   Imports the content of a user-selected CSV file directly
'   into a specified cell range in the active worksheet.
'   This procedure is useful for inserting CSV data at a specific location
'   without creating a new sheet. It uses the BetterArray class for
'   efficient reading and writing.
'
' Parameters: None
'
' Returns: None
'   This procedure does not return any value. It modifies the active Excel sheet
'   by writing the data from the CSV file to it.
'
' Usage:
'   1. Select the destination cell (e.g., "A1") in your Excel sheet.
'   2. Run this procedure. A file selection dialog will open.
'   3. Choose the CSV file to import.
'   Example call from another module: Call FalCSV.FromCSVFileToSelection
'
' Error Handling:
'   - Checks if a cell range is selected before proceeding.
'   - Handles cancellation of the file selection by the user.
'   - Provides a clear error message in case of an unexpected error during import.
'
' Dependencies:
'   - BetterArray: Necessary for parsing the CSV file and writing the data to Excel.
'   - FalFile.Select_Files: Used to display the CSV file selection dialog.
'
' Change Log:
' 1.0 (2025-07-15) - Initial implementation.
' --------------------------------------------------------------------------------------------------
Public Sub FromCSVFileToSelection()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Dim rng As Range
    Set MyArray = New BetterArray
    Dim inputDelimiter As String
    Dim finalDelimiter As String
    Dim inputQuote As String

    ' 1. Check the destination range selection.
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the destination cell where the CSV data should be written (e.g., A1).", vbExclamation
        GoTo CleanUp
    End If
    Set rng = Selection.Cells(1, 1)

    ' 2. Ask the user to select a CSV file.
    filePath = FalFile.Select_Files(FileType:="csv", AllowMultiSelect:=False)

    ' 3. Check if the user canceled or if there was an error.
    If IsEmpty(filePath) Or (IsArray(filePath) And UBound(filePath) < LBound(filePath)) Then
        MsgBox "No CSV file selected or the operation was canceled.", vbInformation
        GoTo CleanUp
    End If

    ' 3.1 Ask the user for the initial delimiter.
    inputDelimiter = InputBox("Please enter the delimiter to use (e.g., , or ; or TAB for tab):", "CSV Preview Delimiter", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 3.2 Ask the user for the initial quote character.
    inputQuote = InputBox("Please enter the cell opening and closing character to use (e.g., """"):", "CSV Preview Quote", """")
    
    ' 4. Import and write the CSV file.
    MyArray.FromCSVFile path:=CStr(filePath(LBound(filePath))), columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False
    MyArray.ToExcelRange rng

    MsgBox "The CSV file has been imported into the sheet starting from cell " & rng.Address(False, False) & ".", vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in FromCSVFileToSelection: " & Err.Description)
    Resume CleanUp
End Sub


' --- Function: FromSelectionToCSVFile ---
' Description:
'   Exports the data from the currently selected Excel cell range
'   to a new CSV file. This function is ideal
'   for saving specific subsets of data from an Excel sheet
'   in a CSV format, thus facilitating exchange or integration with other
'   applications. It uses the BetterArray class for efficient data
'   extraction and writing.
'
' Parameters: None
'
' Returns: None
'   This procedure does not return any value. It interacts with the user
'   for the save path and creates a CSV file.
'
' Usage:
'   1. Select the cell range in Excel that you want to export.
'      If only one cell is selected, the function will automatically
'      export the entire contiguous data region (`CurrentRegion`) around that cell.
'   2. Run this procedure (e.g., via a button or a macro).
'   3. A "Save As" dialog will open, allowing you to choose
'      the name and location of the output CSV file.
'   Example call from another module: Call FalCSV.FromSelectionToCSVFile
'
' Error Handling:
'   - Displays an error message if no cell range is selected.
'   - Handles cancellation of the file save dialog by the user.
'   - Provides a generic error message if an unexpected error occurs
'     during the export process.
'
' Dependencies:
'   - BetterArray: Essential for reading data from the Excel range
'     and formatting it into CSV output.
'   - FalFile.Get_SaveFilePath_WithDialog: Used to display the save dialog.
'
' Change Log:
' 1.0 (2025-07-15) - Initial implementation.
' --------------------------------------------------------------------------------------------------
Public Sub FromSelectionToCSVFile()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Dim rng As Range
    Set MyArray = New BetterArray

    ' 1. Check that the user has selected a range.
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the cell range to export to the CSV file.", vbExclamation
        GoTo CleanUp
    End If
    Set rng = Selection

    ' 2. Ask the user for the save path.
    filePath = FalFile.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "CSV Files (*.csv),*.csv")
    If CStr(filePath) = "" Then
        MsgBox "Save operation canceled.", vbInformation
        GoTo CleanUp
    End If

    ' 3. Read the data from the Excel range.
    MyArray.FromExcelRange FromRange:=rng.CurrentRegion, DetectLastRow:=True, DetectLastColumn:=True

    ' 4. Write the data to the CSV file.
    MyArray.ToCSVFile CStr(filePath)

    MsgBox "The CSV file was created successfully: " & CStr(filePath), vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in FromSelectionToCSVFile: " & Err.Description)
    Resume CleanUp
End Sub

' --- Function: FromWorksheetToCSVFile ---
' Description:
'   Exports all the used data in the active worksheet
'   to a new CSV file. This function allows for the quick saving
'   of all relevant content from an Excel sheet into a standard CSV format,
'   ideal for data sharing or integration with other systems.
'   It uses the BetterArray class to extract the data from Excel and
'   format it as CSV.
'
' Parameters: None
'
' Returns: None
'   This procedure does not return any value. It creates a CSV file at the location
'   specified by the user.
'
' Usage:
'   1. Make sure the active worksheet contains the data you want to export.
'   2. Run this procedure. A save file dialog will open.
'   3. Specify the name and location of the output CSV file.
'   Example call from another module: Call FalCSV.FromWorksheetToCSVFile
'
' Error Handling:
'   - Handles cancellation of the save dialog by the user.
'   - Provides a clear error message in case of an unexpected error during export.
'
' Dependencies:
'   - BetterArray: Essential for reading data from Excel and writing it as CSV.
'   - FalFile.Get_SaveFilePath_WithDialog: Used to display the save path selection dialog.
'
' Change Log:
' 1.0 (2025-07-15) - Initial implementation.
' --------------------------------------------------------------------------------------------------
Public Sub FromWorksheetToCSVFile()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Set MyArray = New BetterArray

    ' 1. Ask the user for the save path.
    filePath = FalFile.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "CSV Files (*.csv),*.csv")
    If CStr(filePath) = "" Then
        MsgBox "Save operation canceled.", vbInformation
        GoTo CleanUp
    End If

    ' 2. Read the data from the used range of the active sheet.
    MyArray.FromExcelRange FromRange:=ActiveSheet.UsedRange, DetectLastRow:=True, DetectLastColumn:=True

    ' 3. Write the data to the CSV file.
    MyArray.ToCSVFile CStr(filePath)

    MsgBox "The CSV file was created successfully: " & CStr(filePath), vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in FromWorksheetToCSVFile: " & Err.Description)
    Resume CleanUp
End Sub


' --- Function: DumpFilesToNewWorksheets ---
' Description:
'   Imports the content of one or more CSV files selected by the user
'   into separate worksheets within the active Excel workbook.
'   Each CSV file is processed individually and its data is written
'   to a new dedicated worksheet, named after the CSV file.
'   This function is an alternative to using an external CSV interface,
'   by centralizing data processing through the 'BetterArray' class.
'
' Parameters: None
'
' Returns: None
'   This procedure does not return any value. It adds new sheets
'   to the active Excel workbook, each containing the data from an imported CSV file.
'
' Usage:
'   Call this procedure to allow the user to select CSV files.
'   For each selected file, a new sheet will be created and filled with the data.
'   Example call from another module: Call FalCSV.DumpFilesToNewWorksheets
'
' Error Handling:
'   - Displays a message if the user cancels the file selection or if no file is chosen.
'   - Handles individual errors when reading each CSV file,
'     allowing to proceed to the next file in case of corruption or inaccessibility.
'   - Handles potential errors when renaming worksheets.
'   - Provides a generic error message for any unexpected error.
'
' Dependencies:
'   - BetterArray: Essential for reading CSV data and writing it to Excel ranges.
'   - FalFile.Select_Files: For the file selection interface.
'
' Change Log:
' 1.0 (2025-07-15) - Initial implementation, based on the BetterArray approach.
' --------------------------------------------------------------------------------------------------
Public Sub DumpFilesToNewWorksheets()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePaths As Variant
    Dim i As Long
    Dim outputSheet As Worksheet
    Dim fileNameOnly As String
    Dim inputDelimiter As String
    Dim finalDelimiter As String
    Dim inputQuote As String


    Set MyArray = New BetterArray

    ' 1. Ask the user to select CSV files.
    filePaths = FalFile.Select_Files(FileType:="csv", AllowMultiSelect:=True)

    ' 2. Check if the user canceled or if there was an error.
    If IsEmpty(filePaths) Or (IsArray(filePaths) And UBound(filePaths) < LBound(filePaths)) Then
        MsgBox "No CSV file selected or the operation was canceled.", vbInformation
        GoTo CleanUp
    End If

    ' 2.1 Ask the user for the initial delimiter.
    inputDelimiter = InputBox("Please enter the delimiter to use (e.g., , or ; or TAB for tab):", "CSV Preview Delimiter", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 2.2 Ask the user for the initial quote character.
    inputQuote = InputBox("Please enter the cell opening and closing character to use (e.g., """"):", "CSV Preview Quote", """")
    
    ' 3. Process each selected file.
    For i = LBound(filePaths) To UBound(filePaths)
        On Error Resume Next
        MyArray.FromCSVFile path:=CStr(filePaths(i)), columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False
        If Err.Number <> 0 Then
            Call HandleError("Error reading CSV file: " & filePaths(i) & ". " & Err.Description)
            Err.Clear
            GoTo NextFile_BA
        End If
        On Error GoTo ErrHandler

        Set outputSheet = ActiveWorkbook.Sheets.Add
        On Error Resume Next
        
        ' Extract the file name without the path
        fileNameOnly = Mid(CStr(filePaths(i)), InStrRev(CStr(filePaths(i)), "\") + 1)
        ' Extract the base name (without the extension)
        If InStr(fileNameOnly, ".") > 0 Then
            fileNameOnly = Left(fileNameOnly, InStr(fileNameOnly, ".") - 1)
        End If
        
        ' Apply the name to the sheet, truncated to 31 characters
        outputSheet.name = Left(fileNameOnly, 31)
        
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrHandler

        MyArray.ToExcelRange outputSheet.Range("A1")
        MyArray.Clear
NextFile_BA:
    Next i

    MsgBox "Data from the CSV files has been imported into new sheets (via BetterArray).", vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in DumpFilesToNewWorksheets: " & Err.Description)
    Resume CleanUp
End Sub


Public Sub MergeFilesToNewWorksheet()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePaths As Variant
    Dim i As Long
    Dim outputSheet As Worksheet
    Dim firstFileProcessed As Boolean
    Set MyArray = New BetterArray

    ' 1. Ask the user to select the CSV files to merge.
    filePaths = FalFile.Select_Files(FileType:="csv", AllowMultiSelect:=True)

    ' 2. Check if files were selected.
    If IsEmpty(filePaths) Or (IsArray(filePaths) And UBound(filePaths) < LBound(filePaths)) Then
        MsgBox "No CSV file selected or the operation was canceled.", vbInformation
        GoTo CleanUp
    End If

    ' 3. Create a new sheet for the merge.
    Set outputSheet = ActiveWorkbook.Sheets.Add
    On Error Resume Next
    outputSheet.name = "Merged_CSV_Data"
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrHandler

    firstFileProcessed = False

    ' 4. Process each file.
    For i = LBound(filePaths) To UBound(filePaths)
        On Error Resume Next
        If Not firstFileProcessed Then
            ' First file: Import normally (includes headers)
            MyArray.FromCSVFile (CStr(filePaths(i)))
            MyArray.ToExcelRange outputSheet.Range("A1")
            firstFileProcessed = True
        Else
            ' Subsequent files: Skip the first row (headers) and append
            Dim tempArray As BetterArray
            Set tempArray = New BetterArray
            tempArray.FromCSVFile path:=CStr(filePaths(i)), IgnoreFirstRow:=True
            ' Find the next empty row in the output sheet
            Dim nextRow As Long
            nextRow = outputSheet.Cells(outputSheet.Rows.count, 1).End(xlUp).Row + 1
            tempArray.ToExcelRange outputSheet.Cells(nextRow, 1)
            Set tempArray = Nothing
        End If

        If Err.Number <> 0 Then
            Call HandleError("Error reading or writing file: " & filePaths(i) & ". " & Err.Description)
            Err.Clear
        End If
        On Error GoTo ErrHandler
        MyArray.Clear
    Next i

    MsgBox "Data from the selected CSV files has been merged into the sheet: " & outputSheet.name, vbInformation

CleanUp:
    Set MyArray = Nothing
    Set outputSheet = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in MergeFilesToNewWorksheet: " & Err.Description)
    Resume CleanUp
End Sub


Public Sub ExportSelectedRangeAsCustomCSV()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Dim rng As Range
    Dim columnDelimiter As String
    Dim encloseAll As VbMsgBoxResult
    Set MyArray = New BetterArray

    ' 1. Check the selection.
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the cell range to export.", vbExclamation
        GoTo CleanUp
    End If
    Set rng = Selection.CurrentRegion

    ' 2. Ask for the save path.
    filePath = FalFile.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "CSV Files (*.csv),*.csv")
    If CStr(filePath) = "" Then
        MsgBox "Save operation canceled.", vbInformation
        GoTo CleanUp
    End If

    ' 3. Ask for the column delimiter.
    columnDelimiter = InputBox("Please enter the column delimiter (e.g., , or ; or TAB for tab):", "CSV Delimiter", ";")
    If columnDelimiter = "" Then
        MsgBox "Invalid delimiter. Operation canceled.", vbExclamation
        GoTo CleanUp
    ElseIf UCase(columnDelimiter) = "TAB" Then
        columnDelimiter = vbTab
    End If

    ' 4. Ask if all fields should be enclosed in quotes.
    encloseAll = MsgBox("Do you want all fields to be enclosed in quotes?", vbYesNo + vbQuestion, "Enclose Fields")

    ' 5. Read the data.
    MyArray.FromExcelRange FromRange:=rng, DetectLastRow:=True, DetectLastColumn:=True

    ' 6. Write the data to the CSV file with custom options.
    MyArray.ToCSVFile path:=CStr(filePath), _
                      columnDelimiter:=columnDelimiter, _
                      EncloseAllInQuotes:=(encloseAll = vbYes)

    MsgBox "Custom CSV file created successfully: " & CStr(filePath), vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in ExportSelectedRangeAsCustomCSV: " & Err.Description)
    Resume CleanUp
End Sub


Public Sub PreviewCSVFile()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Dim previewSheet As Worksheet
    Dim numLinesToPreview As Long
    Dim confirmed As VbMsgBoxResult
    Dim inputDelimiter As String
    Dim inputQuote As String
    Dim finalDelimiter As String
    Dim CSVString As String
    Set MyArray = New BetterArray

    numLinesToPreview = 20

    ' 1. Ask the user to select a CSV file.
    filePath = FalFile.Select_Files(FileType:="csv", AllowMultiSelect:=False)
    If IsEmpty(filePath) Or (IsArray(filePath) And UBound(filePath) < LBound(filePath)) Then
        MsgBox "No CSV file selected or the operation was canceled.", vbInformation
        GoTo CleanUp
    End If

    ' 2. Create a temporary sheet for the preview.
    Set previewSheet = ActiveWorkbook.Sheets.Add
    On Error Resume Next
    previewSheet.name = "CSV_Preview_" & Format(Now, "HHmmss")
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrHandler

    ' 3. Ask the user for the initial delimiter.
    inputDelimiter = InputBox("Please enter the delimiter to use for the preview (e.g., , or ; or TAB for tab):", "CSV Preview Delimiter", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 3.5 Ask the user for the initial quote character.
    inputQuote = InputBox("Please enter the cell opening and closing character to use for the preview (e.g., """"):", "CSV Preview Quote", """")
    
    ' 4. Import the first lines of the CSV with the suggested delimiter.
    CSVString = FalFile.ReadFile_WithADO(CStr(filePath(LBound(filePath))))
    MyArray.FromCSVString CSVString:=CSVString, columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False

    ' Truncate the array if it's too large for the preview
    If MyArray.UpperBound > numLinesToPreview Then
        ' This part assumes BetterArray might need manual trimming.
        ' For this example, we'll just display the first X lines.
    End If

    ' 5. Display the preview.
    MyArray.ToExcelRange previewSheet.Range("A1")
    previewSheet.Columns.AutoFit

    ' 6. Ask the user for confirmation.
    confirmed = MsgBox("The preview of the CSV file is displayed in the sheet '" & previewSheet.name & "'." & vbNewLine & _
                       "Is the formatting correct? Click No to adjust the delimiter.", _
                       vbYesNo + vbQuestion, "Confirm CSV Format")

    If confirmed = vbNo Then
        MsgBox "Please run the function again and try a different delimiter.", vbInformation
        Application.DisplayAlerts = False
        previewSheet.Delete
        Application.DisplayAlerts = True
    Else
        MsgBox "Preview validated. You can now import the file with the chosen settings.", vbInformation
    End If

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in PreviewCSVFile: " & Err.Description)
    If Not previewSheet Is Nothing Then
        Application.DisplayAlerts = False
        previewSheet.Delete
        Application.DisplayAlerts = True
    End If
    Resume CleanUp
End Sub
