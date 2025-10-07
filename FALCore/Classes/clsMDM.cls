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

' Public Properties
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
Public Function ToMdmString() As String
    '@brief Generates a string representation of the MDM data in the official MDM file format.
    '@return A string formatted according to the MDM file specification.
    On Error GoTo ErrorHandler

    Dim sb As New StringBuilder
    sb.Append "BEGIN_HEADER" & vbCrLf

    Dim key As Variant

    ' USER_INPUTS
    If prvUserInputs.Count > 0 Then
        sb.Append "USER_INPUTS" & vbCrLf
        For Each key In prvUserInputs.Keys
            sb.Append key & " " & prvUserInputs(key) & vbCrLf
        Next key
    End If

    ' ICCAP_INPUTS
    If prvIccapInputs.Count > 0 Then
        sb.Append "ICCAP_INPUTS" & vbCrLf
        Dim inputParam As clsMdmInputParameter
        For Each key In prvIccapInputs.Keys
            Set inputParam = prvIccapInputs(key)
            Dim lineBuilder As New StringBuilder
            lineBuilder.Append inputParam.Name & " " & inputParam.Mode & " "
            Dim opt As Variant
            For Each opt In inputParam.ModeOptions.Keys
                lineBuilder.Append inputParam.ModeOptions(opt) & " "
            Next opt
            lineBuilder.Append inputParam.SweepType & " "
            For Each opt In inputParam.SweepOptions.Keys
                lineBuilder.Append inputParam.SweepOptions(opt) & " "
            Next opt
            sb.Append Trim(lineBuilder.ToString) & vbCrLf
        Next key
    End If

    ' ICCAP_OUTPUTS
    If prvIccapOutputs.Count > 0 Then
        sb.Append "ICCAP_OUTPUTS" & vbCrLf
        Dim outputParam As clsMdmOutputParameter
        For Each key In prvIccapOutputs.Keys
            Set outputParam = prvIccapOutputs(key)
            Dim lineBuilder As New StringBuilder
            lineBuilder.Append outputParam.Name & " " & outputParam.Mode & " "
            Dim opt As Variant
            For Each opt In outputParam.ModeOptions.Keys
                lineBuilder.Append outputParam.ModeOptions(opt) & " "
            Next opt
            lineBuilder.Append outputParam.Unit & " " & outputParam.Compliance & " " & outputParam.Type
            sb.Append Trim(lineBuilder.ToString) & vbCrLf
        Next key
    End If

    ' ICCAP_VALUES
    If prvIccapValues.Count > 0 Then
        sb.Append "ICCAP_VALUES" & vbCrLf
        For Each key In prvIccapValues.Keys
            sb.Append key & " " & prvIccapValues(key) & vbCrLf
        Next key
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
        If dataBlock.InputValues.Count > 0 Then
            sb.Append Trim(blockInputs.ToString) & vbCrLf
        End If

        Dim headerSb As New StringBuilder
        Dim header As Variant
        For Each header In dataBlock.HeadersInOrder
            headerSb.Append header & " "
        Next
        sb.Append Trim(headerSb.ToString) & vbCrLf

        Dim i As Long
        For i = 0 To dataBlock.RowCount - 1
            Dim rowSb As New StringBuilder
            For Each header In dataBlock.HeadersInOrder
                Dim val As Variant
                On Error Resume Next
                val = dataBlock.Data(header)(i)
                If Err.Number <> 0 Then val = "" ' Handle missing values gracefully
                On Error GoTo ErrorHandler
                rowSb.Append val & " "
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
        blockDict("HeadersInOrder") = dataBlock.HeadersInOrder

        ' Reconstruct rows for clean JSON output
        Dim rowsCollection As New Collection
        Dim i As Long
        For i = 0 To dataBlock.RowCount - 1
            Dim rowDict As New Dictionary
            Dim header As Variant
            For Each header In dataBlock.HeadersInOrder
                On Error Resume Next
                rowDict(header) = dataBlock.Data(header)(i)
                If Err.Number <> 0 Then rowDict(header) = "" ' Handle missing values
                On Error GoTo ErrorHandler
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

    Dim colHeaders As New Dictionary
    colHeaders.CompareMode = TextCompare

    Dim dataBlock As clsMdmDataBlock
    Dim totalDataRows As Long: totalDataRows = 0

    For Each dataBlock In prvDataBlocks
        Dim key As Variant
        For Each key In dataBlock.InputValues.Keys
            If Not colHeaders.Exists(key) Then colHeaders.Add key, colHeaders.Count + 1
        Next key

        Dim header As Variant
        For Each header In dataBlock.HeadersInOrder
            If Not colHeaders.Exists(header) Then colHeaders.Add header, colHeaders.Count + 1
        Next header

        totalDataRows = totalDataRows + dataBlock.RowCount
    Next dataBlock

    If colHeaders.Count = 0 Then
        To2DArray = Empty
        Exit Function
    End If

    Dim outputArray() As Variant
    ReDim outputArray(1 To totalDataRows + 1, 1 To colHeaders.Count)

    Dim colKey As Variant
    For Each colKey In colHeaders.Keys
        outputArray(1, colHeaders(colKey)) = colKey
    Next colKey

    Dim currentRow As Long: currentRow = 2
    For Each dataBlock In prvDataBlocks
        Dim i As Long
        For i = 0 To dataBlock.RowCount - 1
            For Each colKey In dataBlock.InputValues.Keys
                outputArray(currentRow, colHeaders(colKey)) = dataBlock.InputValues(colKey)
            Next colKey

            For Each colKey In dataBlock.HeadersInOrder
                Dim val As Variant
                On Error Resume Next
                val = dataBlock.Data(colKey)(i)
                If Err.Number <> 0 Then val = "" ' Handle empty cells
                On Error GoTo ErrorHandler
                outputArray(currentRow, colHeaders(colKey)) = val
            Next colKey

            currentRow = currentRow + 1
        Next i
    Next dataBlock

    To2DArray = outputArray
    Exit Function

ErrorHandler:
    To2DArray = Empty
End Function

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
    '@brief Parses a single data block's content into a Dictionary of Dictionaries structure.
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

    ' Parse Headers and initialize the Data dictionary
    Dim headerLine As String: headerLine = Trim(lines(dataStartIndex))
    If Left(headerLine, 1) = "#" Then headerLine = Trim(Mid(headerLine, 2))
    Do While InStr(headerLine, "  ") > 0: headerLine = Replace(headerLine, "  ", " "): Loop

    Dim headerParts() As String: headerParts = Split(headerLine, " ")
    For j = LBound(headerParts) To UBound(headerParts)
        dataBlock.HeadersInOrder.Add headerParts(j)
        dataBlock.Data.Add headerParts(j), New Dictionary
    Next j

    ' Process Data Rows and populate the column dictionaries
    Dim rowIndex As Long: rowIndex = 0
    For i = dataStartIndex + 1 To UBound(lines)
        Dim dataLine As String: dataLine = Trim(lines(i))

        If dataLine <> "" And Left(dataLine, 1) <> "!" Then
            Do While InStr(dataLine, "  ") > 0: dataLine = Replace(dataLine, "  ", " "): Loop

            Dim dataParts() As String: dataParts = Split(dataLine, " ")

            For j = LBound(dataParts) To UBound(dataParts)
                If (j + 1) <= dataBlock.HeadersInOrder.Count Then
                    dataBlock.Data(dataBlock.HeadersInOrder(j + 1))(rowIndex) = dataParts(j)
                End If
            Next j
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