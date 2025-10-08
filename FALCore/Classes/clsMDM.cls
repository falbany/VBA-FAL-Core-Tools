VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private Member Variables
Private prvMdmContent As String
Private prvUserInputs As Dictionary
Private prvIccapInputs As Dictionary
Private prvIccapOutputs As Dictionary
Private prvIccapValues As Dictionary
Private prvDataBlocks As Collection
Private prvValidationErrors As Dictionary
Private prvIsValid As Boolean

' Public Properties
Public Property Get IsValid() As Boolean
    '@brief Returns True if the last validation check passed, False otherwise.
    IsValid = prvIsValid
End Property

Public Property Get ValidationErrors() As Dictionary
    '@brief Returns a dictionary of validation errors found during the last check.
    '@details The dictionary key is a description of the error, and the item is a detailed message.
    Set ValidationErrors = prvValidationErrors
End Property

Public Property Get UserInputs() As Dictionary
    Set UserInputs = prvUserInputs
End Property

Public Property Get IccapInputs() As Dictionary
    Set IccapInputs = prvIccapInputs
End Property

Public Property Get IccapOutputs() As Dictionary
    Set IccapOutputs = prvIccapOutputs
End Property

Public Property Get IccapValues() As Dictionary
    Set IccapValues = prvIccapValues
End Property

Public Property Get DataBlocks() As Collection
    Set DataBlocks = prvDataBlocks
End Property

' Initialization
Private Sub Class_Initialize()
    Set prvUserInputs = New Dictionary
    Set prvIccapInputs = New Dictionary
    Set prvIccapOutputs = New Dictionary
    Set prvIccapValues = New Dictionary
    Set prvDataBlocks = New Collection
    Set prvValidationErrors = New Dictionary
    prvIsValid = True ' Default to true until validation is run
End Sub

' Public Create Methods
Public Sub AddUserInput(ByVal key As String, ByVal value As String)
    '@brief Adds or updates a user input in the header.
    '@param key The name of the user input.
    '@param value The value of the user input.
    prvUserInputs(key) = value
End Sub

Public Sub AddIccapInput(ByVal param As clsMdmInputParameter)
    '@brief Adds or updates an ICCAP input parameter in the header.
    '@param param The structured input parameter object to add.
    prvIccapInputs(param.Name) = param
End Sub

Public Sub AddIccapOutput(ByVal param As clsMdmOutputParameter)
    '@brief Adds or updates an ICCAP output parameter in the header.
    '@param param The structured output parameter object to add.
    prvIccapOutputs(param.Name) = param
End Sub

Public Sub AddIccapValue(ByVal key As String, ByVal value As String)
    '@brief Adds or updates an ICCAP value in the header.
    '@param key The name of the ICCAP value.
    '@param value The value of the ICCAP value.
    prvIccapValues(key) = value
End Sub

Public Function AddDataBlock(Optional ByVal dataBlock As clsMdmDataBlock = Nothing) As clsMdmDataBlock
    '@brief Adds a new data block to the MDM object. If no data block is provided, a new one is created.
    '@param dataBlock (Optional) An existing data block object to add.
    '@return The data block object (either the one passed in or the newly created one).
    '@example
    ' Dim mdm As New clsMDM
    ' Dim newBlock As clsMdmDataBlock
    ' Set newBlock = mdm.AddDataBlock()
    ' newBlock.AddInputValue "Temp", 25
    Dim newBlock As clsMdmDataBlock

    If dataBlock Is Nothing Then
        Set newBlock = New clsMdmDataBlock
    Else
        Set newBlock = dataBlock
    End If

    prvDataBlocks.Add newBlock
    Set AddDataBlock = newBlock
End Function

' Other Public Methods
Public Function Validate() As Boolean
    '@brief Runs a series of validation checks against the MDM object's structure and data.
    '@details Populates the IsValid and ValidationErrors properties.
    '@return True if the object is valid, False otherwise.
    On Error GoTo ErrorHandler

    ' Reset validation state
    Set prvValidationErrors = New Dictionary
    prvIsValid = True

    ' --- Perform all validation checks ---
    prvValidateRowCountConsistency
    prvValidateHeaderToParameterMapping
    prvValidateSweepParameters

    ' Update final status
    If Me.ValidationErrors.Count > 0 Then
        prvIsValid = False
    End If

    Validate = prvIsValid
    Exit Function

ErrorHandler:
    prvIsValid = False
    Validate = False
End Function

Public Function ToMdmString() As String
    '@brief Generates a string representation of the MDM data in the official MDM file format.
    '@return A string formatted according to the MDM file specification.
    On Error GoTo ErrorHandler

    Dim sb As New StringBuilder
    sb.Append "BEGIN_HEADER" & vbCrLf

    Dim key As Variant

    ' USER_INPUTS, ICCAP_INPUTS, ICCAP_OUTPUTS, ICCAP_VALUES
    If prvUserInputs.Count > 0 Then
        sb.Append "USER_INPUTS" & vbCrLf
        For Each key In prvUserInputs.Keys: sb.Append key & " " & prvUserInputs(key) & vbCrLf: Next key
    End If
    If prvIccapInputs.Count > 0 Then
        sb.Append "ICCAP_INPUTS" & vbCrLf
        Dim inputParam As clsMdmInputParameter
        For Each key In prvIccapInputs.Keys
            Set inputParam = prvIccapInputs(key)
            Dim lineBuilder As New StringBuilder
            lineBuilder.Append inputParam.Name & " " & inputParam.Mode & " "
            For Each key In inputParam.ModeOptions.Keys: lineBuilder.Append inputParam.ModeOptions(key) & " ": Next key
            lineBuilder.Append inputParam.SweepType & " "
            For Each key In inputParam.SweepOptions.Keys: lineBuilder.Append inputParam.SweepOptions(key) & " ": Next key
            sb.Append Trim(lineBuilder.ToString) & vbCrLf
        Next key
    End If
    If prvIccapOutputs.Count > 0 Then
        sb.Append "ICCAP_OUTPUTS" & vbCrLf
        Dim outputParam As clsMdmOutputParameter
        For Each key In prvIccapOutputs.Keys
            Set outputParam = prvIccapOutputs(key)
            Dim lineBuilder As New StringBuilder
            lineBuilder.Append outputParam.Name & " " & outputParam.Mode & " "
            For Each key In outputParam.ModeOptions.Keys: lineBuilder.Append outputParam.ModeOptions(key) & " ": Next key
            lineBuilder.Append outputParam.Unit & " " & outputParam.Compliance & " " & outputParam.Type
            sb.Append Trim(lineBuilder.ToString) & vbCrLf
        Next key
    End If
    If prvIccapValues.Count > 0 Then
        sb.Append "ICCAP_VALUES" & vbCrLf
        For Each key In prvIccapValues.Keys: sb.Append key & " " & prvIccapValues(key) & vbCrLf: Next key
    End If

    sb.Append "END_HEADER" & vbCrLf & vbCrLf

    ' --- Build Data Blocks ---
    Dim dataBlock As clsMdmDataBlock
    For Each dataBlock In prvDataBlocks
        sb.Append "BEGIN_DB" & vbCrLf

        Dim blockInputs As New StringBuilder
        For Each key In dataBlock.InputValues.Keys
            blockInputs.Append key & " " & dataBlock.InputValues(key) & " "
        Next key
        If dataBlock.InputValues.Count > 0 Then sb.Append Trim(blockInputs.ToString) & vbCrLf

        ' Build header row, expanding complex numbers
        Dim headerSb As New StringBuilder
        Dim header As Variant
        For Each header In dataBlock.HeadersInOrder
            If TypeOf dataBlock.Data(header)(0) Is clsComplexNumber Then
                headerSb.Append "R:" & header & " " & "I:" & header & " "
            Else
                headerSb.Append header & " "
            End If
        Next
        sb.Append Trim(headerSb.ToString) & vbCrLf

        ' Build data rows
        Dim i As Long
        For i = 0 To dataBlock.RowCount - 1
            Dim rowSb As New StringBuilder
            For Each header In dataBlock.HeadersInOrder
                Dim val As Variant
                On Error Resume Next
                val = dataBlock.Data(header)(i)
                If Err.Number <> 0 Then val = ""
                On Error GoTo ErrorHandler

                If TypeOf val Is clsComplexNumber Then
                    rowSb.Append val.Real & " " & val.Imaginary & " "
                Else
                    rowSb.Append val & " "
                End If
            Next
            sb.Append Trim(rowSb.ToString) & vbCrLf
        Next

        sb.Append "END_DB" & vbCrLf & vbCrLf
    Next dataBlock

    ToMdmString = sb.ToString
    Exit Function

ErrorHandler:
    ToMdmString = ""
End Function

Public Function ParseMdmString(ByVal mdmContent As String) As Boolean
    '@brief Parses MDM data from a string.
    '@param mdmContent The string containing the MDM data.
    '@return True if the string was parsed successfully, False otherwise.
    On Error GoTo ErrorHandler

    prvMdmContent = mdmContent

    If Len(prvMdmContent) > 0 Then
        If prvParseHeader() Then
            ParseMdmString = prvParseDataBlocks()
        Else
            ParseMdmString = False
        End If
    Else
        ParseMdmString = False
    End If

    Exit Function

ErrorHandler:
    ParseMdmString = False
End Function

Public Function ToJSON(Optional ByVal PrettyPrint As Boolean = False) As String
    '@brief Converts the MDM object to a JSON string.
    '@return A JSON string representation of the MDM object.
    On Error GoTo ErrorHandler

    Dim mdmObject As New Dictionary

    Dim inputsDict As New Dictionary
    Dim key As Variant
    For Each key In Me.IccapInputs.Keys
        inputsDict(key) = prvParameterToDict(Me.IccapInputs(key))
    Next

    Dim outputsDict As New Dictionary
    For Each key In Me.IccapOutputs.Keys
        outputsDict(key) = prvParameterToDict(Me.IccapOutputs(key))
    Next

    mdmObject("USER_INPUTS") = Me.UserInputs
    mdmObject("ICCAP_INPUTS") = inputsDict
    mdmObject("ICCAP_OUTPUTS") = outputsDict
    mdmObject("ICCAP_VALUES") = Me.IccapValues

    Dim dataBlockCollection As New Collection
    Dim dataBlock As clsMdmDataBlock
    For Each dataBlock In prvDataBlocks
        Dim blockDict As New Dictionary
        blockDict("InputValues") = dataBlock.InputValues

        ' Reconstruct rows for clean JSON output
        Dim rowsCollection As New Collection
        Dim i As Long
        For i = 0 To dataBlock.RowCount - 1
            Dim rowDict As New Dictionary
            Dim header As Variant
            For Each header In dataBlock.HeadersInOrder
                On Error Resume Next
                Dim val As Variant: val = dataBlock.Data(header)(i)
                If Err.Number <> 0 Then val = ""
                On Error GoTo ErrorHandler

                If TypeOf val Is clsComplexNumber Then
                    rowDict(header) = prvParameterToDict(val)
                Else
                    rowDict(header) = val
                End If
            Next header
            rowsCollection.Add rowDict
        Next i

        blockDict("DataRows") = rowsCollection
        dataBlockCollection.Add blockDict
    Next dataBlock

    mdmObject("DataBlocks") = dataBlockCollection

    If PrettyPrint Then
        ToJSON = JsonConverter.ConvertToJson(mdmObject, Whitespace:=2)
    Else
        ToJSON = JsonConverter.ConvertToJson(mdmObject)
    End If

    Exit Function

ErrorHandler:
    ToJSON = ""
End Function

Private Function prvParameterToDict(ByVal param As Object) As Dictionary
    Set prvParameterToDict = New Dictionary
    If TypeOf param Is clsMdmInputParameter Then
        Dim inputP As clsMdmInputParameter
        Set inputP = param
        prvParameterToDict("Name") = inputP.Name
        prvParameterToDict("Mode") = inputP.Mode
        prvParameterToDict("ModeOptions") = inputP.ModeOptions
        prvParameterToDict("SweepType") = inputP.SweepType
        prvParameterToDict("SweepOptions") = inputP.SweepOptions
    ElseIf TypeOf param Is clsMdmOutputParameter Then
        Dim outputP As clsMdmOutputParameter
        Set outputP = param
        prvParameterToDict("Name") = outputP.Name
        prvParameterToDict("Mode") = outputP.Mode
        prvParameterToDict("ModeOptions") = outputP.ModeOptions
        prvParameterToDict("Unit") = outputP.Unit
        prvParameterToDict("Compliance") = outputP.Compliance
        prvParameterToDict("Type") = outputP.Type
    ElseIf TypeOf param Is clsComplexNumber Then
        Dim complexP As clsComplexNumber
        Set complexP = param
        prvParameterToDict("Real") = complexP.Real
        prvParameterToDict("Imaginary") = complexP.Imaginary
    End If
End Function

Public Function To2DArray() As Variant
    '@brief Converts the MDM data into a 2D variant array for easy export to Excel.
    '@return A 2D variant array containing the combined data from all data blocks.
    On Error GoTo ErrorHandler

    If prvDataBlocks.Count = 0 Then
        To2DArray = Empty
        Exit Function
    End If

    Dim colHeaders As New Collection
    Dim colHeaderDict As New Dictionary
    colHeaderDict.CompareMode = TextCompare

    Dim dataBlock As clsMdmDataBlock
    Dim totalDataRows As Long: totalDataRows = 0

    For Each dataBlock In prvDataBlocks
        Dim key As Variant
        For Each key In dataBlock.InputValues.Keys
            If Not colHeaderDict.Exists(key) Then
                colHeaders.Add key
                colHeaderDict.Add key, 1
            End If
        Next key

        Dim header As Variant
        For Each header In dataBlock.HeadersInOrder
            If dataBlock.Data(header).Count > 0 And TypeOf dataBlock.Data(header)(0) Is clsComplexNumber Then
                If Not colHeaderDict.Exists("R:" & header) Then
                    colHeaders.Add "R:" & header
                    colHeaderDict.Add "R:" & header, 1
                End If
                If Not colHeaderDict.Exists("I:" & header) Then
                    colHeaders.Add "I:" & header
                    colHeaderDict.Add "I:" & header, 1
                End If
            Else
                If Not colHeaderDict.Exists(header) Then
                    colHeaders.Add header
                    colHeaderDict.Add header, 1
                End If
            End If
        Next header

        totalDataRows = totalDataRows + dataBlock.RowCount
    Next dataBlock

    If colHeaders.Count = 0 Then
        To2DArray = Empty
        Exit Function
    End If

    Dim outputArray() As Variant
    ReDim outputArray(1 To totalDataRows + 1, 1 To colHeaders.Count)

    Dim i As Long
    For i = 1 To colHeaders.Count
        outputArray(1, i) = colHeaders(i)
    Next i

    Dim currentRow As Long: currentRow = 2
    For Each dataBlock In prvDataBlocks
        For i = 0 To dataBlock.RowCount - 1
            Dim colIdx As Long
            For colIdx = 1 To colHeaders.Count
                Dim colName As String: colName = colHeaders(colIdx)
                Dim val As Variant: val = ""

                If dataBlock.InputValues.Exists(colName) Then
                    val = dataBlock.InputValues(colName)
                ElseIf dataBlock.Data.Exists(colName) Then
                    On Error Resume Next
                    val = dataBlock.Data(colName)(i)
                    On Error GoTo ErrorHandler
                ElseIf UCase(Left(colName, 2)) = "R:" Then
                    Dim baseName As String: baseName = Mid(colName, 3)
                    If dataBlock.Data.Exists(baseName) Then
                        Dim cNum As clsComplexNumber
                        On Error Resume Next
                        Set cNum = dataBlock.Data(baseName)(i)
                        If Err.Number = 0 And Not cNum Is Nothing Then val = cNum.Real
                        On Error GoTo ErrorHandler
                    End If
                ElseIf UCase(Left(colName, 2)) = "I:" Then
                    baseName = Mid(colName, 3)
                    If dataBlock.Data.Exists(baseName) Then
                        Set cNum = Nothing
                        On Error Resume Next
                        Set cNum = dataBlock.Data(baseName)(i)
                        If Err.Number = 0 And Not cNum Is Nothing Then val = cNum.Imaginary
                        On Error GoTo ErrorHandler
                    End If
                End If
                outputArray(currentRow, colIdx) = val
            Next colIdx
            currentRow = currentRow + 1
        Next i
    Next dataBlock

    To2DArray = outputArray
    Exit Function

ErrorHandler:
    To2DArray = Empty
End Function

Public Sub ConvertDataTypes()
    '@brief Attempts to convert string data in the data blocks to numeric types.
    '@details Iterates through all data points and converts them to Double if they are numeric.
    On Error GoTo ErrorHandler

    Dim dataBlock As clsMdmDataBlock
    For Each dataBlock In prvDataBlocks
        Dim header As Variant
        For Each header In dataBlock.HeadersInOrder
            ' Skip if the column is already complex numbers
            If dataBlock.Data(header).Count > 0 Then
                If TypeOf dataBlock.Data(header)(0) Is clsComplexNumber Then GoTo NextHeader
            End If

            Dim i As Long
            For i = 0 To dataBlock.Data(header).Count - 1
                Dim val As Variant
                val = dataBlock.Data(header)(i)
                If IsNumeric(val) Then
                    dataBlock.Data(header)(i) = CDbl(val)
                End If
            Next i
NextHeader:
        Next header
    Next dataBlock

    Exit Sub
ErrorHandler:
    Debug.Print "Error in ConvertDataTypes: " & Err.Description
End Sub

Public Sub ExportToWorksheet(ByVal wks As Worksheet, Optional ByVal startCellAddress As String = "A1")
    '@brief Exports the entire MDM data set to a specified worksheet.
    '@param wks The worksheet object to export the data to.
    '@param startCellAddress (Optional) The starting cell for the data export (e.g., "A1").
    On Error GoTo ErrorHandler

    Dim dataArray As Variant
    dataArray = Me.To2DArray()

    If IsEmpty(dataArray) Then
        Debug.Print "No data to export."
        Exit Sub
    End If

    Dim startCell As Range
    Set startCell = wks.Range(startCellAddress)

    ' Resize the target range and write the array in one operation for efficiency
    startCell.Resize(UBound(dataArray, 1), UBound(dataArray, 2)).value = dataArray

    Exit Sub

ErrorHandler:
    Debug.Print "Error in ExportToWorksheet: " & Err.Description
End Sub

Public Function ReadMdmFile(ByVal filePath As String) As Boolean
    '@brief Reads the content of an MDM file and parses its header and data blocks.
    '@param filePath The full path to the MDM file.
    '@return True if the file was read and parsed successfully, False otherwise.
    On Error GoTo ErrorHandler

    If Dir(filePath) = "" Then
        ReadMdmFile = False
        Exit Function
    End If

    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    prvMdmContent = Input$(LOF(fileNumber), fileNumber)
    Close #fileNumber

    If Len(prvMdmContent) > 0 Then
        If prvParseHeader() Then
            ReadMdmFile = prvParseDataBlocks()
        Else
            ReadMdmFile = False
        End If
    Else
        ReadMdmFile = False
    End If

    Exit Function

ErrorHandler:
    ReadMdmFile = False
End Function

' Private Helper Functions
Private Sub prvValidateSweepParameters()
    '@brief Validates that the data for swept variables matches the sweep definitions.
    '@details Checks NumPoints, Start, and Stop values against the actual data.
    On Error GoTo ErrorHandler

    Dim key As Variant
    Dim inputParam As clsMdmInputParameter

    For Each key In prvIccapInputs.Keys
        Set inputParam = prvIccapInputs(key)

        ' Check only parameters that are defined as sweeps (e.g., LIN, LOG)
        If inputParam.SweepType <> "CON" And inputParam.SweepType <> "SYNC" Then
            Dim paramName As String: paramName = inputParam.Name

            ' Find the data block that contains this sweep as its innermost sweep
            Dim targetBlock As clsMdmDataBlock
            Dim blockIndex As Long: blockIndex = 0
            Dim innerMostSweepForBlock As String

            Dim tempBlock As clsMdmDataBlock
            For Each tempBlock In prvDataBlocks
                blockIndex = blockIndex + 1
                ' This logic assumes we can identify the innermost sweep for a block.
                ' For now, we'll just check if the column exists in the block. A more robust
                ' implementation might need a way to link blocks to specific sweep orders.
                If tempBlock.Data.Exists(paramName) Then
                    Set targetBlock = tempBlock
                    Exit For
                End If
            Next tempBlock

            If Not targetBlock Is Nothing Then
                ' 1. Validate number of points
                If inputParam.SweepOptions.Exists("NumPoints") Then
                    Dim expectedPoints As Long: expectedPoints = CLng(inputParam.SweepOptions("NumPoints"))
                    Dim actualPoints As Long: actualPoints = targetBlock.Data(paramName).Count

                    If actualPoints <> expectedPoints Then
                        prvValidationErrors("Sweep.NumPoints." & paramName) = "Sweep '" & paramName & "' in Block " & blockIndex & " has " & actualPoints & " data points, but header defines " & expectedPoints & "."
                    End If
                End If

                ' 2. Validate Start and Stop values (with tolerance for float comparison)
                Dim firstVal As Double, lastVal As Double
                Dim expectedStart As Double, expectedStop As Double
                Const TOLERANCE As Double = 0.00001

                If targetBlock.Data(paramName).Count > 0 Then
                    If inputParam.SweepOptions.Exists("Start") And IsNumeric(inputParam.SweepOptions("Start")) Then
                        expectedStart = CDbl(inputParam.SweepOptions("Start"))
                        firstVal = CDbl(targetBlock.Data(paramName)(0))
                        If Abs(firstVal - expectedStart) > TOLERANCE Then
                            prvValidationErrors("Sweep.Start." & paramName) = "Sweep '" & paramName & "' in Block " & blockIndex & " starts at " & firstVal & ", but header defines " & expectedStart & "."
                        End If
                    End If

                    If inputParam.SweepOptions.Exists("Stop") And IsNumeric(inputParam.SweepOptions("Stop")) Then
                        expectedStop = CDbl(inputParam.SweepOptions("Stop"))
                        lastVal = CDbl(targetBlock.Data(paramName)(targetBlock.Data(paramName).Count - 1))
                        If Abs(lastVal - expectedStop) > TOLERANCE Then
                            prvValidationErrors("Sweep.Stop." & paramName) = "Sweep '" & paramName & "' in Block " & blockIndex & " stops at " & lastVal & ", but header defines " & expectedStop & "."
                        End If
                    End If
                End If
            Else
                ' Error: Sweep defined in header but no corresponding data found
                prvValidationErrors("Sweep.MissingData." & paramName) = "Sweep parameter '" & paramName & "' is defined in the header, but no corresponding data column was found in any data block."
            End If
        End If
    Next key

    Exit Sub
ErrorHandler:
    prvValidationErrors("Validation.SweepParameters") = "An unexpected error occurred during sweep parameter validation: " & Err.Description
End Sub

Private Sub prvValidateRowCountConsistency()
    '@brief Checks that all data columns within each data block have the same number of rows.
    '@details Adds errors to the ValidationErrors dictionary if inconsistencies are found.
    On Error GoTo ErrorHandler

    Dim blockIndex As Long: blockIndex = 0
    Dim dataBlock As clsMdmDataBlock
    For Each dataBlock In prvDataBlocks
        blockIndex = blockIndex + 1
        If dataBlock.HeadersInOrder.Count = 0 Then GoTo NextBlock

        Dim expectedRows As Long
        expectedRows = dataBlock.RowCount

        Dim header As Variant
        For Each header In dataBlock.HeadersInOrder
            Dim currentRows As Long
            currentRows = dataBlock.Data(header).Count
            If currentRows <> expectedRows Then
                Dim errorKey As String
                errorKey = "Row Count Mismatch in Block " & blockIndex

                Dim errorMessage As String
                errorMessage = "In Data Block " & blockIndex & ", column '" & header & "' has " & currentRows & " rows, but the block's reference row count is " & expectedRows & "."

                prvValidationErrors(errorKey & " - " & header) = errorMessage
            End If
        Next header
NextBlock:
    Next dataBlock

    Exit Sub
ErrorHandler:
    prvValidationErrors("Validation.RowCountConsistency") = "An unexpected error occurred during row count consistency validation: " & Err.Description
End Sub

Private Sub prvValidateHeaderToParameterMapping()
    '@brief Checks that every data column header corresponds to a defined ICCAP_OUTPUT or ICCAP_INPUT parameter.
    '@details Adds errors to the ValidationErrors dictionary if a mapping is not found.
    On Error GoTo ErrorHandler

    Dim blockIndex As Long: blockIndex = 0
    Dim dataBlock As clsMdmDataBlock
    For Each dataBlock In prvDataBlocks
        blockIndex = blockIndex + 1

        Dim header As Variant
        For Each header In dataBlock.HeadersInOrder
            ' A header is valid if it's in either the INPUTS or OUTPUTS collection
            If Not prvIccapInputs.Exists(header) And Not prvIccapOutputs.Exists(header) Then
                 Dim errorKey As String
                 errorKey = "Undefined Header in Block " & blockIndex

                 Dim errorMessage As String
                 errorMessage = "In Data Block " & blockIndex & ", the data header '" & header & "' is not defined as an ICCAP_INPUT or ICCAP_OUTPUT in the main header."

                 prvValidationErrors(errorKey & " - " & header) = errorMessage
            End If
        Next header
    Next dataBlock

    Exit Sub
ErrorHandler:
    prvValidationErrors("Validation.HeaderMapping") = "An unexpected error occurred during header to parameter mapping validation: " & Err.Description
End Sub

Private Function prvParseDataBlocks() As Boolean
    '@brief Parses all data blocks (BEGIN_DB/END_DB) in the MDM file.
    '@return True if parsing is successful, False otherwise.
    On Error GoTo ErrorHandler

    Dim currentPos As Long
    currentPos = 1

    Do
        Dim startPos As Long
        Dim endPos As Long

        startPos = InStr(currentPos, prvMdmContent, "BEGIN_DB", vbTextCompare)

        If startPos = 0 Then
            Exit Do ' No more data blocks found
        End If

        endPos = InStr(startPos, prvMdmContent, "END_DB", vbTextCompare)

        If endPos = 0 Then
            ' Malformed file, BEGIN_DB without END_DB
            prvParseDataBlocks = False
            Exit Function
        End If

        Dim blockContent As String
        blockContent = Mid(prvMdmContent, startPos + Len("BEGIN_DB"), endPos - startPos - Len("BEGIN_DB"))

        Dim dataBlock As New clsMdmDataBlock
        If prvParseSingleDataBlock(blockContent, dataBlock) Then
            prvDataBlocks.Add dataBlock
        Else
            ' Failed to parse a block, abort
            prvParseDataBlocks = False
            Exit Function
        End If

        currentPos = endPos + Len("END_DB")
    Loop

    prvParseDataBlocks = True
    Exit Function

ErrorHandler:
    prvParseDataBlocks = False
End Function

Private Function prvParseSingleDataBlock(ByVal blockContent As String, ByRef dataBlock As clsMdmDataBlock) As Boolean
    '@brief Parses a single data block's content into a column-oriented data structure, handling complex numbers.
    '@param blockContent The string content between BEGIN_DB and END_DB.
    '@param dataBlock The clsMdmDataBlock object to populate.
    '@return True if parsing is successful, False otherwise.
    On Error GoTo ErrorHandler

    Dim innermostSweepName As String
    Dim inputParam As clsMdmInputParameter
    Dim key As Variant
    For Each key In prvIccapInputs.Keys
        Set inputParam = prvIccapInputs(key)
        If inputParam.SweepOptions.Exists("SweepOrder") Then
            If CStr(inputParam.SweepOptions("SweepOrder")) = "1" Then
                innermostSweepName = inputParam.Name
                Exit For
            End If
        End If
    Next key
    If innermostSweepName = "" Then Err.Raise 5, "prvParseSingleDataBlock", "Could not determine innermost sweep from header."

    Dim lines() As String, parts() As String
    lines = Split(blockContent, vbCrLf)
    Dim i As Long, j As Long, k As Long
    Dim line As String, cleanLine As String
    Dim dataStartIndex As Long: dataStartIndex = -1

    ' Find the start of the data table and parse block-level parameters
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))
        If line = "" Or Left(line, 1) = "!" Then GoTo NextLine

        cleanLine = Trim(line)
        If Left(cleanLine, 1) = "#" Then cleanLine = Trim(Mid(cleanLine, 2))

        parts = Split(cleanLine)
        If UBound(parts) < 0 Then GoTo NextLine

        If UCase(parts(0)) = UCase(innermostSweepName) Then
            dataStartIndex = i
            Exit For
        Else
            Dim currentIdx As Long: currentIdx = LBound(parts)
            If UCase(parts(currentIdx)) = "ICCAP_VAR" Then currentIdx = currentIdx + 1
            For k = currentIdx To UBound(parts) Step 2
                If (k + 1) <= UBound(parts) Then dataBlock.AddInputValue parts(k), parts(k + 1)
            Next k
        End If
NextLine:
    Next i

    If dataStartIndex = -1 Then
        prvParseSingleDataBlock = True ' Valid block, just no data table
        Exit Function
    End If

    ' --- Parse Headers and initialize data columns, handling complex pairs ---
    Dim headerLine As String: headerLine = Trim(lines(dataStartIndex))
    If Left(headerLine, 1) = "#" Then headerLine = Trim(Mid(headerLine, 2))
    Do While InStr(headerLine, "  ") > 0: headerLine = Replace(headerLine, "  ", " "): Loop

    Dim headerParts() As String: headerParts = Split(headerLine, " ")

    j = LBound(headerParts)
    Do While j <= UBound(headerParts)
        Dim currentHeader As String: currentHeader = headerParts(j)
        If UCase(Left(currentHeader, 2)) = "R:" And (j + 1) <= UBound(headerParts) Then
            Dim nextHeader As String: nextHeader = headerParts(j + 1)
            If UCase(Left(nextHeader, 2)) = "I:" And Mid(currentHeader, 3) = Mid(nextHeader, 3) Then
                ' This is a complex number pair
                Dim baseName As String: baseName = Mid(currentHeader, 3)
                dataBlock.HeadersInOrder.Add baseName
                dataBlock.Data.Add baseName, New Dictionary
                j = j + 2 ' Skip the next part
            Else
                ' Standalone R: or other header
                dataBlock.HeadersInOrder.Add currentHeader
                dataBlock.Data.Add currentHeader, New Dictionary
                j = j + 1
            End If
        Else
            ' Standard header
            dataBlock.HeadersInOrder.Add currentHeader
            dataBlock.Data.Add currentHeader, New Dictionary
            j = j + 1
        End If
    Loop

    ' --- Process Data Rows and populate the column dictionaries ---
    Dim rowIndex As Long: rowIndex = 0
    For i = dataStartIndex + 1 To UBound(lines)
        Dim dataLine As String: dataLine = Trim(lines(i))

        If dataLine <> "" And Left(dataLine, 1) <> "!" Then
            Do While InStr(dataLine, "  ") > 0: dataLine = Replace(dataLine, "  ", " "): Loop

            Dim dataParts() As String: dataParts = Split(dataLine, " ")

            Dim dataIdx As Long: dataIdx = LBound(dataParts)

            Dim header As Variant
            For Each header In dataBlock.HeadersInOrder
                If dataIdx > UBound(dataParts) Then Exit For

                ' Check if the original header was complex to know how many parts to consume
                Dim isComplex As Boolean: isComplex = False
                If prvIccapOutputs.Exists(header) Then
                    Dim outputParam As clsMdmOutputParameter
                    Set outputParam = prvIccapOutputs(header)
                    Select Case UCase(outputParam.Mode)
                        Case "S", "H", "Z", "Y", "K", "A"
                            isComplex = True
                    End Select
                End If

                If isComplex And (dataIdx + 1) <= UBound(dataParts) Then
                    Dim complexNum As New clsComplexNumber
                    complexNum.Real = CDbl(dataParts(dataIdx))
                    complexNum.Imaginary = CDbl(dataParts(dataIdx + 1))
                    dataBlock.Data(header)(rowIndex) = complexNum
                    dataIdx = dataIdx + 2
                Else
                    dataBlock.Data(header)(rowIndex) = dataParts(dataIdx)
                    dataIdx = dataIdx + 1
                End If
            Next header

            rowIndex = rowIndex + 1
        End If
    Next i

    prvParseSingleDataBlock = True
    Exit Function

ErrorHandler:
    prvParseSingleDataBlock = False
End Function

Private Function prvParseHeader() As Boolean
    '@brief Parses the header section of the MDM file into structured objects.
    '@return True if the header was parsed successfully, False otherwise.
    On Error GoTo ErrorHandler

    Dim headerContent As String
    Dim startPos As Long, endPos As Long

    startPos = InStr(1, prvMdmContent, "BEGIN_HEADER", vbTextCompare)
    endPos = InStr(1, prvMdmContent, "END_HEADER", vbTextCompare)

    If startPos = 0 Or endPos = 0 Then Exit Function

    headerContent = Mid(prvMdmContent, startPos + Len("BEGIN_HEADER"), endPos - startPos - Len("BEGIN_HEADER"))

    Dim lines() As String
    lines = Split(headerContent, vbCrLf)

    Dim currentSection As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim(lines(i))

        If line = "" Or Left(line, 1) = "!" Then GoTo NextLine

        Select Case UCase(line)
            Case "USER_INPUTS", "ICCAP_INPUTS", "ICCAP_OUTPUTS", "ICCAP_VALUES"
                currentSection = UCase(line)
            Case Else
                If currentSection <> "" Then
                    Dim parts() As String
                    parts = Split(line)

                    Select Case currentSection
                        Case "USER_INPUTS"
                            prvUserInputs.Add parts(0), Trim(Mid(line, Len(parts(0)) + 1))
                        Case "ICCAP_INPUTS"
                            Dim inputParam As clsMdmInputParameter
                            Set inputParam = prvParseInputParameter(line)
                            If Not inputParam Is Nothing Then prvIccapInputs.Add inputParam.Name, inputParam
                        Case "ICCAP_OUTPUTS"
                            Dim outputParam As clsMdmOutputParameter
                            Set outputParam = prvParseOutputParameter(line)
                            If Not outputParam Is Nothing Then prvIccapOutputs.Add outputParam.Name, outputParam
                        Case "ICCAP_VALUES"
                            prvIccapValues.Add parts(0), Trim(Mid(line, Len(parts(0)) + 1))
                    End Select
                End If
        End Select
NextLine:
    Next i

    prvParseHeader = True
    Exit Function

ErrorHandler:
    prvParseHeader = False
End Function

Private Function prvParseInputParameter(ByVal line As String) As clsMdmInputParameter
    On Error GoTo ErrorHandler
    Dim param As New clsMdmInputParameter

    Dim cleanLine As String
    cleanLine = Trim(line)
    Do While InStr(cleanLine, "  ") > 0
        cleanLine = Replace(cleanLine, "  ", " ")
    Loop

    Dim parts() As String
    parts = Split(cleanLine, " ")

    If UBound(parts) < 2 Then Exit Function ' Must have at least name and mode

    param.Name = parts(0)
    param.Mode = parts(1)

    ' Find sweep type to determine boundaries of mode_options and sweep_options
    Dim sweepTypeIdx As Long
    sweepTypeIdx = -1
    Dim i As Long
    For i = 2 To UBound(parts)
        Select Case UCase(parts(i))
            Case "LIN", "LOG", "SYNC", "LIST", "CON", "AC", "HB", "EXP", "PULSE", "PWL", "SFFM", "SIN", "TDR", "SEG"
                sweepTypeIdx = i
                Exit For
        End Select
    Next i

    If sweepTypeIdx = -1 Then Exit Function ' No valid sweep type found

    param.SweepType = parts(sweepTypeIdx)

    ' Extract mode options
    Dim modeOptionsParts() As String
    If sweepTypeIdx > 2 Then
        ReDim modeOptionsParts(0 To sweepTypeIdx - 3)
        For i = 2 To sweepTypeIdx - 1
            modeOptionsParts(i - 2) = parts(i)
        Next i
    End If

    ' Extract sweep options
    Dim sweepOptionsParts() As String
    If UBound(parts) > sweepTypeIdx Then
        ReDim sweepOptionsParts(0 To UBound(parts) - sweepTypeIdx - 1)
        For i = sweepTypeIdx + 1 To UBound(parts)
            sweepOptionsParts(i - sweepTypeIdx - 1) = parts(i)
        Next i
    End If

    ' Parse mode options based on mode
    i = 0
    Select Case UCase(param.Mode)
        Case "V", "U": param.ModeOptions("PlusNode") = modeOptionsParts(i): i = i + 1: param.ModeOptions("MinusNode") = modeOptionsParts(i): i = i + 1: param.ModeOptions("Unit") = modeOptionsParts(i): i = i + 1: param.ModeOptions("Compliance") = modeOptionsParts(i)
        Case "I": param.ModeOptions("ToNode") = modeOptionsParts(i): i = i + 1: param.ModeOptions("FromNode") = modeOptionsParts(i): i = i + 1: param.ModeOptions("Unit") = modeOptionsParts(i): i = i + 1: param.ModeOptions("Compliance") = modeOptionsParts(i)
        Case "P": param.ModeOptions("ParamName") = modeOptionsParts(i): i = i + 1: param.ModeOptions("Unit") = modeOptionsParts(i)
    End Select

    ' Parse sweep options based on sweep type
    i = 0
    Select Case UCase(param.SweepType)
        Case "LIN": param.SweepOptions("SweepOrder") = sweepOptionsParts(i): i = i + 1: param.SweepOptions("Start") = sweepOptionsParts(i): i = i + 1: param.SweepOptions("Stop") = sweepOptionsParts(i): i = i + 1: param.SweepOptions("NumPoints") = sweepOptionsParts(i): i = i + 1: param.SweepOptions("StepSize") = sweepOptionsParts(i)
        Case "SYNC": param.SweepOptions("Ratio") = sweepOptionsParts(i): i = i + 1: param.SweepOptions("Offset") = sweepOptionsParts(i): i = i + 1: param.SweepOptions("MasterSweep") = sweepOptionsParts(i)
        Case "CON": param.SweepOptions("Value") = sweepOptionsParts(i)
    End Select

    Set prvParseInputParameter = param
    Exit Function
ErrorHandler:
    Set prvParseInputParameter = Nothing
End Function

Private Function prvParseOutputParameter(ByVal line As String) As clsMdmOutputParameter
    On Error GoTo ErrorHandler
    Dim param As New clsMdmOutputParameter

    Dim cleanLine As String
    cleanLine = Trim(line)
    Do While InStr(cleanLine, "  ") > 0
        cleanLine = Replace(cleanLine, "  ", " ")
    Loop

    Dim parts() As String
    parts = Split(cleanLine, " ")

    If UBound(parts) < 1 Then Exit Function

    param.Name = parts(0)
    param.Mode = parts(1)

    Dim modeOptionsCount As Long
    Select Case UCase(param.Mode)
        Case "V", "N", "U": modeOptionsCount = 2
        Case "I": modeOptionsCount = 2
        Case "C", "G": modeOptionsCount = 2
        Case "T": modeOptionsCount = 2
        Case "S", "H", "Z", "K", "A", "Y": modeOptionsCount = 3
        Case Else: modeOptionsCount = 0 ' For modes like U in the example
    End Select

    Dim currentIdx As Long
    currentIdx = 2

    If modeOptionsCount > 0 Then
        Dim modeOptionsParts() As String
        ReDim modeOptionsParts(0 To modeOptionsCount - 1)
        Dim i As Long
        For i = 0 To modeOptionsCount - 1
            modeOptionsParts(i) = parts(currentIdx + i)
        Next i
        currentIdx = currentIdx + modeOptionsCount

        i = 0
        Select Case UCase(param.Mode)
            Case "V", "N", "U": param.ModeOptions("PlusNode") = modeOptionsParts(i): i = i + 1: param.ModeOptions("MinusNode") = modeOptionsParts(i)
            Case "I": param.ModeOptions("ToNode") = modeOptionsParts(i): i = i + 1: param.ModeOptions("FromNode") = modeOptionsParts(i)
            Case "C", "G": param.ModeOptions("HighNode") = modeOptionsParts(i): i = i + 1: param.ModeOptions("LowNode") = modeOptionsParts(i)
            Case "T": param.ModeOptions("Node") = modeOptionsParts(i): i = i + 1: param.ModeOptions("PulseParam") = modeOptionsParts(i)
            Case "S", "H", "Z", "K", "A", "Y": param.ModeOptions("Port1") = modeOptionsParts(i): i = i + 1: param.ModeOptions("Port2") = modeOptionsParts(i): i = i + 1: param.ModeOptions("ACGround") = modeOptionsParts(i)
        End Select
    End If

    If currentIdx <= UBound(parts) Then param.Unit = parts(currentIdx): currentIdx = currentIdx + 1
    If currentIdx <= UBound(parts) Then param.Compliance = parts(currentIdx): currentIdx = currentIdx + 1
    If currentIdx <= UBound(parts) Then param.Type = parts(currentIdx)

    Set prvParseOutputParameter = param
    Exit Function
ErrorHandler:
    Set prvParseOutputParameter = Nothing
End Function