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
Private prvValidationErrors As Collection
Private prvIsValid As Boolean

' Public Properties
Public Property Get IsValid() As Boolean
    '@brief Returns True if the last validation check passed, False otherwise.
    IsValid = prvIsValid
End Property

Public Property Get ValidationErrors() As Collection
    '@brief Returns a collection of validation errors (clsMdmValidationError objects) found during the last check.
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
    Set prvValidationErrors = New Collection
    prvIsValid = True ' Default to true until validation is run
End Sub

' Public Create Methods
Public Function AddUserInput(ByVal key As String, ByVal value As String) As clsMDM
    '@brief Adds or updates a user input in the header.
    '@param key The name of the user input.
    '@param value The value of the user input.
    prvUserInputs(key) = value
    Set AddUserInput = Me
End Function

Public Function AddIccapInput(ByVal param As clsMdmInputParameter) As clsMDM
    '@brief Adds or updates an ICCAP input parameter in the header.
    '@param param The structured input parameter object to add.
    prvIccapInputs(param.Name) = param
    Set AddIccapInput = Me
End Function

Public Function AddIccapOutput(ByVal param As clsMdmOutputParameter) As clsMDM
    '@brief Adds or updates an ICCAP output parameter in the header.
    '@param param The structured output parameter object to add.
    prvIccapOutputs(param.Name) = param
    Set AddIccapOutput = Me
End Function

Public Function AddIccapValue(ByVal key As String, ByVal value As String) As clsMDM
    '@brief Adds or updates an ICCAP value in the header.
    '@param key The name of the ICCAP value.
    '@param value The value of the ICCAP value.
    prvIccapValues(key) = value
    Set AddIccapValue = Me
End Function

Public Function DefineLinSweepInput(ByVal Name As String, ByVal SweepOrder As Long, ByVal StartValue As Double, ByVal StopValue As Double, ByVal NumPoints As Long, Optional ByVal Mode As MdmInputMode = mdmModeV, Optional ByVal PlusNode As String = "1", Optional ByVal MinusNode As String = "0", Optional ByVal Unit As String = "V", Optional ByVal Compliance As String = "100m") As clsMDM
    '@brief Defines and adds a linear sweep input parameter. This is a convenience factory method.
    '@param Name The name of the input parameter (e.g., "Vd").
    '@param SweepOrder The order of the sweep (1 for innermost).
    '@param StartValue The starting value for the sweep.
    '@param StopValue The stopping value for the sweep.
    '@param NumPoints The number of points in the sweep.
    '@param Mode (Optional) The mode of the input (e.g., mdmModeV for voltage). Defaults to mdmModeV.
    '@param PlusNode (Optional) The positive node. Defaults to "1".
    '@param MinusNode (Optional) The negative node. Defaults to "0".
    '@param Unit (Optional) The unit for the parameter. Defaults to "V".
    '@param Compliance (Optional) The compliance limit. Defaults to "100m".
    '@return The clsMDM instance for method chaining.
    On Error GoTo ErrorHandler

    Dim param As New clsMdmInputParameter
    param.Name = Name
    param.Mode = Mode
    param.SweepType = mdmSweepLIN

    ' Set mode-specific options
    param.PlusNode = PlusNode
    param.MinusNode = MinusNode
    param.Unit = Unit
    param.Compliance = Compliance

    ' Set sweep-specific options
    param.SweepOrder = SweepOrder
    param.Start = StartValue
    param.Stop = StopValue
    param.NumPoints = NumPoints
    If NumPoints > 1 Then
        param.StepSize = (StopValue - StartValue) / (NumPoints - 1)
    Else
        param.StepSize = 0
    End If

    ' Add the parameter to the collection
    prvIccapInputs(Name) = param

    Set DefineLinSweepInput = Me
    Exit Function

ErrorHandler:
    Set DefineLinSweepInput = Me ' Still return instance on error
End Function

Public Function DefineConstantInput(ByVal Name As String, ByVal Value As Variant, Optional ByVal Mode As MdmInputMode = mdmModeV, Optional ByVal PlusNode As String = "1", Optional ByVal MinusNode As String = "0", Optional ByVal Unit As String = "V", Optional ByVal Compliance As String = "100m") As clsMDM
    '@brief Defines and adds a constant input parameter. This is a convenience factory method.
    '@param Name The name of the input parameter (e.g., "V_Gate").
    '@param Value The constant value for the parameter.
    '@param Mode (Optional) The mode of the input (e.g., mdmModeV for voltage). Defaults to mdmModeV.
    '@param PlusNode (Optional) The positive node. Defaults to "1".
    '@param MinusNode (Optional) The negative node. Defaults to "0".
    '@param Unit (Optional) The unit for the parameter. Defaults to "V".
    '@param Compliance (Optional) The compliance limit. Defaults to "100m".
    '@return The clsMDM instance for method chaining.
    On Error GoTo ErrorHandler

    Dim param As New clsMdmInputParameter
    param.Name = Name
    param.Mode = Mode
    param.SweepType = mdmSweepCON

    ' Set mode-specific options
    param.PlusNode = PlusNode
    param.MinusNode = MinusNode
    param.Unit = Unit
    param.Compliance = Compliance

    ' Set sweep-specific options
    param.Value = Value

    ' Add the parameter to the collection
    prvIccapInputs(Name) = param

    Set DefineConstantInput = Me
    Exit Function

ErrorHandler:
    Set DefineConstantInput = Me ' Still return instance on error
End Function

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
Public Function AddValueToBlock(ByVal blockIndex As Long, ByVal Header As String, ByVal RowIndex As Long, ByVal Value As Variant) As clsMDM
    '@brief Adds or updates a single value in a specified data block. This is a convenience method that abstracts block retrieval.
    '@param blockIndex The 0-based index of the data block to modify.
    '@param Header The name of the column header.
    '@param RowIndex The 0-based index of the row.
    '@param Value The value to set for the cell.
    '@return The clsMDM instance for method chaining.
    On Error GoTo ErrorHandler

    If blockIndex >= 0 And blockIndex < prvDataBlocks.Count Then
        Dim dataBlock As clsMdmDataBlock
        Set dataBlock = prvDataBlocks(blockIndex + 1) ' Collection is 1-based
        dataBlock.AddValue Header, RowIndex, Value
    End If

    Set AddValueToBlock = Me
    Exit Function

ErrorHandler:
    ' Optionally log error
    Set AddValueToBlock = Me
End Function

Public Function SetValueInBlock(ByVal blockIndex As Long, ByVal Header As String, ByVal RowIndex As Long, ByVal Value As Variant) As clsMDM
    '@brief Adds or updates a single value in a specified data block. Alias for AddValueToBlock.
    '@param blockIndex The 0-based index of the data block to modify.
    '@param Header The name of the column header.
    '@param RowIndex The 0-based index of the row.
    '@param Value The value to set for the cell.
    '@return The clsMDM instance for method chaining.
    Set SetValueInBlock = AddValueToBlock(blockIndex, Header, RowIndex, Value)
End Function

Public Function GetValue(ByVal blockIndex As Long, ByVal header As String, ByVal rowIndex As Long) As Variant
    '@brief Retrieves a single value from a specified data block.
    '@param blockIndex The 0-based index of the data block.
    '@param header The name of the data column (header).
    '@param rowIndex The 0-based index of the row.
    '@return The value at the specified location, or CVErr(xlErrNA) if not found.
    On Error GoTo ErrorHandler

    If blockIndex < 0 Or blockIndex >= prvDataBlocks.Count Then
        GetValue = CVErr(xlErrNA)
        Exit Function
    End If

    Dim dataBlock As clsMdmDataBlock
    Set dataBlock = prvDataBlocks(blockIndex + 1) ' Collection is 1-based

    If Not dataBlock.Data.Exists(header) Then
        GetValue = CVErr(xlErrNA)
        Exit Function
    End If

    Dim columnData As Dictionary
    Set columnData = dataBlock.Data(header)

    If Not columnData.Exists(rowIndex) Then
        GetValue = CVErr(xlErrNA)
        Exit Function
    End If

    GetValue = columnData(rowIndex)
    Exit Function

ErrorHandler:
    GetValue = CVErr(xlErrNA)
End Function

Public Function GetColumn(ByVal blockIndex As Long, ByVal header As String) As Variant
    '@brief Retrieves an entire data column from a specified data block as a 1D array.
    '@param blockIndex The 0-based index of the data block.
    '@param header The name of the data column (header).
    '@return A 1D variant array containing the column data, or CVErr(xlErrNA) if not found.
    On Error GoTo ErrorHandler

    If blockIndex < 0 Or blockIndex >= prvDataBlocks.Count Then
        GetColumn = CVErr(xlErrNA)
        Exit Function
    End If

    Dim dataBlock As clsMdmDataBlock
    Set dataBlock = prvDataBlocks(blockIndex + 1) ' Collection is 1-based

    If Not dataBlock.Data.Exists(header) Then
        GetColumn = CVErr(xlErrNA)
        Exit Function
    End If

    Dim columnData As Dictionary
    Set columnData = dataBlock.Data(header)

    If columnData.Count = 0 Then
        GetColumn = Array() ' Return empty array if column exists but has no data
        Exit Function
    End If

    ' Assuming row indices are contiguous from 0 to N-1
    Dim result() As Variant
    ReDim result(0 To columnData.Count - 1)

    Dim i As Long
    For i = 0 To columnData.Count - 1
        result(i) = columnData(i)
    Next i

    GetColumn = result
    Exit Function

ErrorHandler:
    GetColumn = CVErr(xlErrNA)
End Function

Public Function Validate() As Boolean
    '@brief Runs a series of validation checks against the MDM object's structure and data.
    '@details Populates the IsValid and ValidationErrors properties.
    '@return True if the object is valid, False otherwise.
    On Error GoTo ErrorHandler

    ' Reset validation state
    Set prvValidationErrors = New Collection
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
            lineBuilder.Append inputParam.Name & " " & prvInputModeToString(inputParam.Mode) & " "

            ' --- Append Mode Options ---
            Select Case inputParam.Mode
                Case mdmModeV, mdmModeU
                    lineBuilder.Append inputParam.PlusNode & " "
                    lineBuilder.Append inputParam.MinusNode & " "
                    lineBuilder.Append inputParam.Unit & " "
                    lineBuilder.Append inputParam.Compliance & " "
                Case mdmModeI
                    lineBuilder.Append inputParam.ToNode & " "
                    lineBuilder.Append inputParam.FromNode & " "
                    lineBuilder.Append inputParam.Unit & " "
                    lineBuilder.Append inputParam.Compliance & " "
                Case mdmModeP
                    lineBuilder.Append inputParam.ParamName & " "
                    lineBuilder.Append inputParam.Unit & " "
                Case mdmModeW
                    lineBuilder.Append inputParam.PlusNode & " "
                    lineBuilder.Append inputParam.MinusNode & " "
                    lineBuilder.Append inputParam.PowerUnit & " "
                    lineBuilder.Append inputParam.Resistance & " "
                    lineBuilder.Append inputParam.Fund & " "
                    lineBuilder.Append inputParam.Unit & " "
                    lineBuilder.Append inputParam.Compliance & " "
                Case mdmModeF, mdmModeT
                    ' No options
            End Select

            lineBuilder.Append prvSweepTypeToString(inputParam.SweepType) & " "

            ' --- Append Sweep Options ---
            Select Case inputParam.SweepType
                Case mdmSweepLIN
                    lineBuilder.Append inputParam.SweepOrder & " "
                    lineBuilder.Append inputParam.Start & " "
                    lineBuilder.Append inputParam.Stop & " "
                    lineBuilder.Append inputParam.NumPoints & " "
                    lineBuilder.Append inputParam.StepSize & " "
                Case mdmSweepLOG
                    lineBuilder.Append inputParam.SweepOrder & " "
                    lineBuilder.Append inputParam.Start & " "
                    lineBuilder.Append inputParam.Stop & " "
                    lineBuilder.Append inputParam.NumPoints & " "
                    lineBuilder.Append inputParam.SweepScale & " "
                    lineBuilder.Append inputParam.TotalNumPoints & " "
                Case mdmSweepSYNC
                    lineBuilder.Append inputParam.Ratio & " "
                    lineBuilder.Append inputParam.Offset & " "
                    lineBuilder.Append inputParam.MasterSweep & " "
                Case mdmSweepLIST
                    lineBuilder.Append inputParam.SweepOrder & " "
                    If IsArray(inputParam.ListValues) Then
                        lineBuilder.Append (UBound(inputParam.ListValues) - LBound(inputParam.ListValues) + 1) & " "
                        Dim val As Variant
                        For Each val In inputParam.ListValues
                            lineBuilder.Append val & " "
                        Next val
                    End If
                Case mdmSweepCON
                    lineBuilder.Append inputParam.Value & " "
                Case mdmSweepAC
                    lineBuilder.Append inputParam.Magnitude & " "
                    lineBuilder.Append inputParam.Phase & " "
                Case mdmSweepHB
                    lineBuilder.Append inputParam.SweepOrder & " "
                    lineBuilder.Append inputParam.Value & " "
                    lineBuilder.Append inputParam.Order & " "
                Case mdmSweepEXP
                    lineBuilder.Append inputParam.InitialValue & " "
                    lineBuilder.Append inputParam.PulsedValue & " "
                    lineBuilder.Append inputParam.RiseDelay & " "
                    lineBuilder.Append inputParam.RiseConst & " "
                    lineBuilder.Append inputParam.FallDelay & " "
                    lineBuilder.Append inputParam.FallConst & " "
                Case mdmSweepPULSE
                    lineBuilder.Append inputParam.InitialValue & " "
                    lineBuilder.Append inputParam.PulsedValue & " "
                    lineBuilder.Append inputParam.DelayTime & " "
                    lineBuilder.Append inputParam.RiseTime & " "
                    lineBuilder.Append inputParam.FallTime & " "
                    lineBuilder.Append inputParam.PulseWidth & " "
                    lineBuilder.Append inputParam.Period & " "
                Case mdmSweepPWL
                    If Not inputParam.TimeValuePairs Is Nothing Then
                        If inputParam.TimeValuePairs.Count > 0 Then
                            lineBuilder.Append inputParam.TimeValuePairs.Count & " "
                            Dim tKey As Variant
                            For Each tKey In inputParam.TimeValuePairs.Keys
                                lineBuilder.Append tKey & " " & inputParam.TimeValuePairs(tKey) & " "
                            Next tKey
                        End If
                    End If
                    lineBuilder.Append inputParam.StartTime & " "
                    lineBuilder.Append inputParam.RepeatTimes & " "
                Case mdmSweepSFFM
                    lineBuilder.Append inputParam.Offset & " "
                    lineBuilder.Append inputParam.Amplitude & " "
                    lineBuilder.Append inputParam.CarrierFrequency & " "
                    lineBuilder.Append inputParam.ModulIndex & " "
                    lineBuilder.Append inputParam.SignalFrequency & " "
                Case mdmSweepSIN
                    lineBuilder.Append inputParam.Offset & " "
                    lineBuilder.Append inputParam.Amplitude & " "
                    lineBuilder.Append inputParam.Frequency & " "
                    lineBuilder.Append inputParam.DelayTime & " "
                    lineBuilder.Append inputParam.DampFactor & " "
                    lineBuilder.Append inputParam.Phase & " "
                Case mdmSweepTDR
                    lineBuilder.Append inputParam.InitialValue & " "
                    lineBuilder.Append inputParam.PulsedValue & " "
                    lineBuilder.Append inputParam.DelayTime & " "
                    lineBuilder.Append inputParam.RiseTime & " "
                    lineBuilder.Append inputParam.FallTime & " "
                    lineBuilder.Append inputParam.PulseWidth & " "
                    lineBuilder.Append inputParam.Period & " "
                    lineBuilder.Append inputParam.Resistance & " "
                Case mdmSweepSEG
                    lineBuilder.Append inputParam.SweepOrder & " "
                    lineBuilder.Append inputParam.NumSegments & " "
                    lineBuilder.Append inputParam.NumPoints & " "
                    If IsArray(inputParam.SegmentValues) Then
                        Dim segVal As Variant
                        For Each segVal In inputParam.SegmentValues
                            If Not IsEmpty(segVal) And Not IsNull(segVal) Then lineBuilder.Append segVal & " "
                        Next segVal
                    End If
            End Select

            sb.Append Trim(lineBuilder.ToString) & vbCrLf
        Next key
    End If
    If prvIccapOutputs.Count > 0 Then
        sb.Append "ICCAP_OUTPUTS" & vbCrLf
        Dim outputParam As clsMdmOutputParameter
        For Each key In prvIccapOutputs.Keys
            Set outputParam = prvIccapOutputs(key)
            Dim lineBuilder As New StringBuilder
            lineBuilder.Append outputParam.Name & " " & prvOutputTypeToString(outputParam.Mode) & " "
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

Public Function ParseMdmString(ByVal mdmContent As String, Optional AutoConvert As Boolean = True) As Boolean
    '@brief Parses MDM data from a string.
    '@param mdmContent The string containing the MDM data.
    '@param AutoConvert (Optional) If True, automatically converts numeric strings to numeric types after parsing.
    '@return True if the string was parsed successfully, False otherwise.
    On Error GoTo ErrorHandler

    prvMdmContent = mdmContent
    Dim parsedOk As Boolean

    ' Reset object state before parsing to ensure a clean slate
    Class_Initialize

    If Len(prvMdmContent) > 0 Then
        If prvParseHeader() Then
            parsedOk = prvParseDataBlocks()
        Else
            parsedOk = False
        End If
    Else
        parsedOk = False
    End If

    If parsedOk And AutoConvert Then
        Me.ConvertDataTypes
    End If

    ParseMdmString = parsedOk
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

Public Function ReadMdmFile(ByVal filePath As String, Optional AutoConvert As Boolean = True) As Boolean
    '@brief Reads the content of an MDM file and parses its header and data blocks.
    '@param filePath The full path to the MDM file.
    '@param AutoConvert (Optional) If True, automatically converts numeric strings to numeric types after parsing.
    '@return True if the file was read and parsed successfully, False otherwise.
    On Error GoTo ErrorHandler

    If Dir(filePath) = "" Then
        ReadMdmFile = False
        Exit Function
    End If

    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    Dim mdmContent As String
    mdmContent = Input$(LOF(fileNumber), fileNumber)
    Close #fileNumber

    ' Delegate parsing to the ParseMdmString method
    ReadMdmFile = Me.ParseMdmString(mdmContent, AutoConvert)

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
        If inputParam.SweepType <> mdmSweepCON And inputParam.SweepType <> mdmSweepSYNC Then
            Dim paramName As String: paramName = inputParam.Name

            ' Find the data block that contains this sweep
            Dim targetBlock As clsMdmDataBlock
            Dim blockIndex As Long: blockIndex = 0

            Dim tempBlock As clsMdmDataBlock
            For Each tempBlock In prvDataBlocks
                If tempBlock.Data.Exists(paramName) Then
                    Set targetBlock = tempBlock
                    Exit For
                End If
                blockIndex = blockIndex + 1
            Next tempBlock

            If Not targetBlock Is Nothing Then
                ' 1. Validate number of points
                If inputParam.NumPoints > 0 Then
                    Dim expectedPoints As Long: expectedPoints = inputParam.NumPoints
                    Dim actualPoints As Long: actualPoints = targetBlock.Data(paramName).Count

                    If actualPoints <> expectedPoints Then
                        prvAddValidationError errTypeSweepNumPointsMismatch, "Sweep '" & paramName & "' in Block " & blockIndex & " has " & actualPoints & " data points, but header defines " & expectedPoints & ".", blockIndex, paramName
                    End If
                End If

                ' 2. Validate Start and Stop values (with tolerance for float comparison)
                Dim firstVal As Double, lastVal As Double
                Dim expectedStart As Double, expectedStop As Double
                Const TOLERANCE As Double = 0.00001

                If targetBlock.Data(paramName).Count > 0 Then
                    If Not IsEmpty(inputParam.Start) Then
                        expectedStart = inputParam.Start
                        firstVal = CDbl(targetBlock.Data(paramName)(0))
                        If Abs(firstVal - expectedStart) > TOLERANCE Then
                             prvAddValidationError errTypeSweepStartMismatch, "Sweep '" & paramName & "' in Block " & blockIndex & " starts at " & firstVal & ", but header defines " & expectedStart & ".", blockIndex, paramName
                        End If
                    End If

                    If Not IsEmpty(inputParam.Stop) Then
                        expectedStop = inputParam.Stop
                        lastVal = CDbl(targetBlock.Data(paramName)(targetBlock.Data(paramName).Count - 1))
                        If Abs(lastVal - expectedStop) > TOLERANCE Then
                            prvAddValidationError errTypeSweepStopMismatch, "Sweep '" & paramName & "' in Block " & blockIndex & " stops at " & lastVal & ", but header defines " & expectedStop & ".", blockIndex, paramName
                        End If
                    End If
                End If
            Else
                ' Error: Sweep defined in header but no corresponding data found
                prvAddValidationError errTypeMissingSweepData, "Sweep parameter '" & paramName & "' is defined in the header, but no corresponding data column was found in any data block.", -1, paramName
            End If
        End If
    Next key

    Exit Sub
ErrorHandler:
    prvAddValidationError errTypeGeneric, "An unexpected error occurred during sweep parameter validation: " & Err.Description
End Sub

Private Sub prvValidateRowCountConsistency()
    '@brief Checks that all data columns within each data block have the same number of rows.
    '@details Adds errors to the ValidationErrors collection if inconsistencies are found.
    On Error GoTo ErrorHandler

    Dim blockIndex As Long: blockIndex = 0
    Dim dataBlock As clsMdmDataBlock
    For Each dataBlock In prvDataBlocks
        If dataBlock.HeadersInOrder.Count > 0 Then
            Dim expectedRows As Long
            expectedRows = dataBlock.RowCount

            Dim header As Variant
            For Each header In dataBlock.HeadersInOrder
                Dim currentRows As Long
                currentRows = dataBlock.Data(header).Count
                If currentRows <> expectedRows Then
                    prvAddValidationError errTypeRowCountMismatch, "In Data Block " & blockIndex & ", column '" & header & "' has " & currentRows & " rows, but the block's reference row count is " & expectedRows & ".", blockIndex, CStr(header)
                End If
            Next header
        End If
        blockIndex = blockIndex + 1
    Next dataBlock

    Exit Sub
ErrorHandler:
    prvAddValidationError errTypeGeneric, "An unexpected error occurred during row count consistency validation: " & Err.Description
End Sub

Private Sub prvValidateHeaderToParameterMapping()
    '@brief Checks that every data column header corresponds to a defined ICCAP_OUTPUT or ICCAP_INPUT parameter.
    '@details Adds errors to the ValidationErrors collection if a mapping is not found.
    On Error GoTo ErrorHandler

    Dim blockIndex As Long: blockIndex = 0
    Dim dataBlock As clsMdmDataBlock
    For Each dataBlock In prvDataBlocks
        Dim header As Variant
        For Each header In dataBlock.HeadersInOrder
            ' A header is valid if it's in either the INPUTS or OUTPUTS collection
            If Not prvIccapInputs.Exists(header) And Not prvIccapOutputs.Exists(header) Then
                 prvAddValidationError errTypeUndefinedHeader, "In Data Block " & blockIndex & ", the data header '" & header & "' is not defined as an ICCAP_INPUT or ICCAP_OUTPUT in the main header.", blockIndex, CStr(header)
            End If
        Next header
        blockIndex = blockIndex + 1
    Next dataBlock

    Exit Sub
ErrorHandler:
    prvAddValidationError errTypeGeneric, "An unexpected error occurred during header to parameter mapping validation: " & Err.Description
End Sub

Private Sub prvAddValidationError(ByVal errType As MdmValidationErrorType, ByVal msg As String, Optional ByVal blockIdx As Long = -1, Optional ByVal item As String = "")
    '@brief Creates a new validation error object and adds it to the collection.
    Dim err As New clsMdmValidationError
    err.ErrorType = errType
    err.Message = msg
    err.BlockIndex = blockIdx
    err.OffendingItem = item
    prvValidationErrors.Add err
End Sub

' --- Enum Conversion Helpers ---
Private Function prvStringToInputMode(ByVal s As String) As MdmInputMode
    Select Case UCase(s)
        Case "V": prvStringToInputMode = mdmModeV
        Case "I": prvStringToInputMode = mdmModeI
        Case "F": prvStringToInputMode = mdmModeF
        Case "T": prvStringToInputMode = mdmModeT
        Case "P": prvStringToInputMode = mdmModeP
        Case "U": prvStringToInputMode = mdmModeU
        Case "W": prvStringToInputMode = mdmModeW
        Case Else: prvStringToInputMode = mdmModeUnknown
    End Select
End Function

Private Function prvInputModeToString(ByVal e As MdmInputMode) As String
    Select Case e
        Case mdmModeV: prvInputModeToString = "V"
        Case mdmModeI: prvInputModeToString = "I"
        Case mdmModeF: prvInputModeToString = "F"
        Case mdmModeT: prvInputModeToString = "T"
        Case mdmModeP: prvInputModeToString = "P"
        Case mdmModeU: prvInputModeToString = "U"
        Case mdmModeW: prvInputModeToString = "W"
        Case Else: prvInputModeToString = ""
    End Select
End Function

Private Function prvStringToOutputType(ByVal s As String) As MdmOutputType
    Select Case UCase(s)
        Case "V": prvStringToOutputType = mdmTypeV
        Case "I": prvStringToOutputType = mdmTypeI
        Case "C": prvStringToOutputType = mdmTypeC
        Case "G": prvStringToOutputType = mdmTypeG
        Case "T": prvStringToOutputType = mdmTypeT
        Case "S": prvStringToOutputType = mdmTypeS
        Case "H": prvStringToOutputType = mdmTypeH
        Case "Z": prvStringToOutputType = mdmTypeZ
        Case "Y": prvStringToOutputType = mdmTypeY
        Case "K": prvStringToOutputType = mdmTypeK
        Case "A": prvStringToOutputType = mdmTypeA
        Case "N": prvStringToOutputType = mdmTypeN
        Case "U": prvStringToOutputType = mdmTypeU
        Case Else: prvStringToOutputType = mdmTypeUnknown
    End Select
End Function

Private Function prvOutputTypeToString(ByVal e As MdmOutputType) As String
    Select Case e
        Case mdmTypeV: prvOutputTypeToString = "V"
        Case mdmTypeI: prvOutputTypeToString = "I"
        Case mdmTypeC: prvOutputTypeToString = "C"
        Case mdmTypeG: prvOutputTypeToString = "G"
        Case mdmTypeT: prvOutputTypeToString = "T"
        Case mdmTypeS: prvOutputTypeToString = "S"
        Case mdmTypeH: prvOutputTypeToString = "H"
        Case mdmTypeZ: prvOutputTypeToString = "Z"
        Case mdmTypeY: prvOutputTypeToString = "Y"
        Case mdmTypeK: prvOutputTypeToString = "K"
        Case mdmTypeA: prvOutputTypeToString = "A"
        Case mdmTypeN: prvOutputTypeToString = "N"
        Case mdmTypeU: prvOutputTypeToString = "U"
        Case Else: prvOutputTypeToString = ""
    End Select
End Function

Private Function prvStringToSweepType(ByVal s As String) As MdmSweepType
    Select Case UCase(s)
        Case "LIN": prvStringToSweepType = mdmSweepLIN
        Case "LOG": prvStringToSweepType = mdmSweepLOG
        Case "SYNC": prvStringToSweepType = mdmSweepSYNC
        Case "LIST": prvStringToSweepType = mdmSweepLIST
        Case "CON": prvStringToSweepType = mdmSweepCON
        Case "AC": prvStringToSweepType = mdmSweepAC
        Case "HB": prvStringToSweepType = mdmSweepHB
        Case "EXP": prvStringToSweepType = mdmSweepEXP
        Case "PULSE": prvStringToSweepType = mdmSweepPULSE
        Case "PWL": prvStringToSweepType = mdmSweepPWL
        Case "SFFM": prvStringToSweepType = mdmSweepSFFM
        Case "SIN": prvStringToSweepType = mdmSweepSIN
        Case "TDR": prvStringToSweepType = mdmSweepTDR
        Case "SEG": prvStringToSweepType = mdmSweepSEG
        Case Else: prvStringToSweepType = mdmSweepUnknown
    End Select
End Function

Private Function prvSweepTypeToString(ByVal e As MdmSweepType) As String
    Select Case e
        Case mdmSweepLIN: prvSweepTypeToString = "LIN"
        Case mdmSweepLOG: prvSweepTypeToString = "LOG"
        Case mdmSweepSYNC: prvSweepTypeToString = "SYNC"
        Case mdmSweepLIST: prvSweepTypeToString = "LIST"
        Case mdmSweepCON: prvSweepTypeToString = "CON"
        Case mdmSweepAC: prvSweepTypeToString = "AC"
        Case mdmSweepHB: prvSweepTypeToString = "HB"
        Case mdmSweepEXP: prvSweepTypeToString = "EXP"
        Case mdmSweepPULSE: prvSweepTypeToString = "PULSE"
        Case mdmSweepPWL: prvSweepTypeToString = "PWL"
        Case mdmSweepSFFM: prvSweepTypeToString = "SFFM"
        Case mdmSweepSIN: prvSweepTypeToString = "SIN"
        Case mdmSweepTDR: prvSweepTypeToString = "TDR"
        Case mdmSweepSEG: prvSweepTypeToString = "SEG"
        Case Else: prvSweepTypeToString = ""
    End Select
End Function

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

Private Function prvGetInnermostSweepName() As String
    '@brief Finds the name of the innermost sweep variable (SweepOrder=1) from the header definitions.
    On Error GoTo ErrorHandler

    Dim key As Variant
    Dim inputParam As clsMdmInputParameter
    For Each key In prvIccapInputs.Keys
        Set inputParam = prvIccapInputs(key)
        ' SweepOrder is a direct property now
        If inputParam.SweepOrder = 1 Then
            prvGetInnermostSweepName = inputParam.Name
            Exit Function
        End If
    Next key

    ' If no sweep order is explicitly 1, we might need a fallback, but for now, this is an error.
    Err.Raise 5, "prvGetInnermostSweepName", "Could not determine innermost sweep from header."

    Exit Function
ErrorHandler:
    prvGetInnermostSweepName = ""
End Function

Private Sub prvParseDataBlockInputs(ByRef lines() As String, ByVal dataStartIndex As Long, ByRef dataBlock As clsMdmDataBlock)
    '@brief Parses the block-specific input variables (e.g., ICCAP_VAR lines) that appear before the main data table.
    On Error GoTo ErrorHandler

    Dim i As Long, k As Long
    Dim line As String, cleanLine As String
    Dim parts() As String

    For i = LBound(lines) To dataStartIndex - 1
        line = Trim(lines(i))
        If line = "" Or Left(line, 1) = "!" Then GoTo NextLine

        cleanLine = Trim(line)
        If Left(cleanLine, 1) = "#" Then cleanLine = Trim(Mid(cleanLine, 2))

        parts = Split(cleanLine)
        If UBound(parts) < 0 Then GoTo NextLine

        Dim currentIdx As Long: currentIdx = LBound(parts)
        If UCase(parts(currentIdx)) = "ICCAP_VAR" Then currentIdx = currentIdx + 1
        For k = currentIdx To UBound(parts) Step 2
            If (k + 1) <= UBound(parts) Then dataBlock.AddInputValue parts(k), parts(k + 1)
        Next k
NextLine:
    Next i

    Exit Sub
ErrorHandler:
    ' Handle or log error if necessary
End Sub

Private Sub prvParseDataBlockHeaders(ByVal headerLine As String, ByRef dataBlock As clsMdmDataBlock)
    '@brief Parses the data table header line, handling complex number pairs (e.g., R:Z, I:Z).
    On Error GoTo ErrorHandler

    Dim cleanHeaderLine As String
    cleanHeaderLine = Trim(headerLine)
    If Left(cleanHeaderLine, 1) = "#" Then cleanHeaderLine = Trim(Mid(cleanHeaderLine, 2))
    Do While InStr(cleanHeaderLine, "  ") > 0: cleanHeaderLine = Replace(cleanHeaderLine, "  ", " "): Loop

    Dim headerParts() As String: headerParts = Split(cleanHeaderLine, " ")

    Dim j As Long: j = LBound(headerParts)
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

    Exit Sub
ErrorHandler:
    ' Handle or log error if necessary
End Sub

Private Sub prvParseDataBlockRows(ByRef lines() As String, ByVal dataStartIndex As Long, ByRef dataBlock As clsMdmDataBlock)
    '@brief Parses the data rows of a block and populates the column-oriented data structure.
    On Error GoTo ErrorHandler

    Dim rowIndex As Long: rowIndex = 0
    Dim i As Long
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
                    Select Case outputParam.Mode
                        Case mdmTypeS, mdmTypeH, mdmTypeZ, mdmTypeY, mdmTypeK, mdmTypeA
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

    Exit Sub
ErrorHandler:
    ' Handle or log error if necessary
End Sub

Private Function prvParseSingleDataBlock(ByVal blockContent As String, ByRef dataBlock As clsMdmDataBlock) As Boolean
    '@brief Parses a single data block's content by coordinating helper functions.
    On Error GoTo ErrorHandler

    ' 1. Find the name of the innermost sweep, which identifies the header row
    Dim innermostSweepName As String
    innermostSweepName = prvGetInnermostSweepName()
    If innermostSweepName = "" Then GoTo ErrorHandler ' Error is raised inside the helper

    ' 2. Find the start of the data table (the header line)
    Dim lines() As String: lines = Split(blockContent, vbCrLf)
    Dim dataStartIndex As Long: dataStartIndex = -1
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim cleanLine As String: cleanLine = Trim(lines(i))
        If cleanLine <> "" And Left(cleanLine, 1) <> "!" Then
            If Left(cleanLine, 1) = "#" Then cleanLine = Trim(Mid(cleanLine, 2))

            Dim firstWord As String
            On Error Resume Next ' Handle lines with no spaces
            firstWord = Split(cleanLine)(0)
            On Error GoTo ErrorHandler

            If UCase(firstWord) = UCase(innermostSweepName) Then
                dataStartIndex = i
                Exit For
            End If
        End If
    Next i

    ' If no data table, it could be a block with only input values, which is valid.
    ' We parse all lines as input values in this case.
    If dataStartIndex = -1 Then
        prvParseDataBlockInputs lines, UBound(lines) + 1, dataBlock
        prvParseSingleDataBlock = True
        Exit Function
    End If

    ' 3. Parse the different parts of the block
    prvParseDataBlockInputs lines, dataStartIndex, dataBlock
    prvParseDataBlockHeaders lines(dataStartIndex), dataBlock
    prvParseDataBlockRows lines, dataStartIndex, dataBlock

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
    param.Mode = prvStringToInputMode(parts(1))

    ' Find sweep type to determine boundaries of mode_options and sweep_options
    Dim sweepTypeIdx As Long
    sweepTypeIdx = -1
    Dim i As Long
    For i = 2 To UBound(parts)
        If prvStringToSweepType(parts(i)) <> mdmSweepUnknown Then
            sweepTypeIdx = i
            Exit For
        End If
    Next i

    If sweepTypeIdx = -1 Then Exit Function ' No valid sweep type found
    param.SweepType = prvStringToSweepType(parts(sweepTypeIdx))

    ' --- Parse Mode Options ---
    Dim currentIdx As Long
    currentIdx = 2
    Select Case param.Mode
        Case mdmModeV, mdmModeU
            param.PlusNode = parts(currentIdx): currentIdx = currentIdx + 1
            param.MinusNode = parts(currentIdx): currentIdx = currentIdx + 1
            param.Unit = parts(currentIdx): currentIdx = currentIdx + 1
            param.Compliance = parts(currentIdx): currentIdx = currentIdx + 1
        Case mdmModeI
            param.ToNode = parts(currentIdx): currentIdx = currentIdx + 1
            param.FromNode = parts(currentIdx): currentIdx = currentIdx + 1
            param.Unit = parts(currentIdx): currentIdx = currentIdx + 1
            param.Compliance = parts(currentIdx): currentIdx = currentIdx + 1
        Case mdmModeP
            param.ParamName = parts(currentIdx): currentIdx = currentIdx + 1
            param.Unit = parts(currentIdx): currentIdx = currentIdx + 1
        Case mdmModeW
            param.PlusNode = parts(currentIdx): currentIdx = currentIdx + 1
            param.MinusNode = parts(currentIdx): currentIdx = currentIdx + 1
            param.PowerUnit = parts(currentIdx): currentIdx = currentIdx + 1
            param.Resistance = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Fund = parts(currentIdx): currentIdx = currentIdx + 1
            param.Unit = parts(currentIdx): currentIdx = currentIdx + 1
            param.Compliance = parts(currentIdx): currentIdx = currentIdx + 1
        Case mdmModeF, mdmModeT
            ' No options
    End Select

    ' --- Parse Sweep Options ---
    currentIdx = sweepTypeIdx + 1
    Select Case param.SweepType
        Case mdmSweepLIN
            param.SweepOrder = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Start = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Stop = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.NumPoints = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
            param.StepSize = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
        Case mdmSweepLOG
            param.SweepOrder = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Start = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Stop = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.NumPoints = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
            param.SweepScale = parts(currentIdx): currentIdx = currentIdx + 1
            param.TotalNumPoints = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
        Case mdmSweepSYNC
            param.Ratio = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Offset = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.MasterSweep = parts(currentIdx): currentIdx = currentIdx + 1
        Case mdmSweepLIST
            param.SweepOrder = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
            Dim numValues As Long: numValues = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
            Dim values() As Variant
            ReDim values(0 To numValues - 1)
            For i = 0 To numValues - 1
                values(i) = parts(currentIdx): currentIdx = currentIdx + 1
            Next i
            param.ListValues = values
        Case mdmSweepCON
            param.Value = parts(currentIdx): currentIdx = currentIdx + 1
        Case mdmSweepAC
            param.Magnitude = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Phase = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
        Case mdmSweepHB
            param.SweepOrder = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Value = parts(currentIdx): currentIdx = currentIdx + 1
            param.Order = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
        Case mdmSweepEXP
            param.InitialValue = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.PulsedValue = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.RiseDelay = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.RiseConst = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.FallDelay = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.FallConst = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
        Case mdmSweepPULSE
            param.InitialValue = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.PulsedValue = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.DelayTime = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.RiseTime = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.FallTime = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.PulseWidth = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Period = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
        Case mdmSweepPWL
            Dim numPairs As Long: numPairs = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
            For i = 1 To numPairs
                param.TimeValuePairs(parts(currentIdx)) = parts(currentIdx + 1)
                currentIdx = currentIdx + 2
            Next i
            param.StartTime = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.RepeatTimes = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
        Case mdmSweepSFFM
            param.Offset = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Amplitude = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.CarrierFrequency = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.ModulIndex = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.SignalFrequency = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
        Case mdmSweepSIN
            param.Offset = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Amplitude = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Frequency = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.DelayTime = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.DampFactor = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Phase = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
        Case mdmSweepTDR
            param.InitialValue = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.PulsedValue = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.DelayTime = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.RiseTime = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.FallTime = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.PulseWidth = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Period = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
            param.Resistance = CDbl(parts(currentIdx)): currentIdx = currentIdx + 1
        Case mdmSweepSEG
            param.SweepOrder = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
            param.NumSegments = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
            param.NumPoints = CLng(parts(currentIdx)): currentIdx = currentIdx + 1
            Dim segValues() As Variant
            ReDim segValues(0 To 9) ' Max 10 values
            For i = 0 To UBound(segValues)
                If currentIdx <= UBound(parts) Then
                    segValues(i) = parts(currentIdx): currentIdx = currentIdx + 1
                Else
                    Exit For
                End If
            Next i
            param.SegmentValues = segValues
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
    param.Mode = prvStringToOutputType(parts(1))

    Dim modeOptionsCount As Long
    Select Case param.Mode
        Case mdmTypeV, mdmTypeN, mdmTypeU: modeOptionsCount = 2
        Case mdmTypeI: modeOptionsCount = 2
        Case mdmTypeC, mdmTypeG: modeOptionsCount = 2
        Case mdmTypeT: modeOptionsCount = 2
        Case mdmTypeS, mdmTypeH, mdmTypeZ, mdmTypeK, mdmTypeA, mdmTypeY: modeOptionsCount = 3
        Case Else: modeOptionsCount = 0
    End Select

    Dim currentIdx As Long
    currentIdx = 2

    If modeOptionsCount > 0 Then
        Dim i As Long
        Select Case param.Mode
            Case mdmTypeV, mdmTypeN, mdmTypeU
                param.ModeOptions("PlusNode") = parts(currentIdx): currentIdx = currentIdx + 1
                param.ModeOptions("MinusNode") = parts(currentIdx): currentIdx = currentIdx + 1
            Case mdmTypeI
                param.ModeOptions("ToNode") = parts(currentIdx): currentIdx = currentIdx + 1
                param.ModeOptions("FromNode") = parts(currentIdx): currentIdx = currentIdx + 1
            Case mdmTypeC, mdmTypeG
                param.ModeOptions("HighNode") = parts(currentIdx): currentIdx = currentIdx + 1
                param.ModeOptions("LowNode") = parts(currentIdx): currentIdx = currentIdx + 1
            Case mdmTypeT
                param.ModeOptions("Node") = parts(currentIdx): currentIdx = currentIdx + 1
                param.ModeOptions("PulseParam") = parts(currentIdx): currentIdx = currentIdx + 1
            Case mdmTypeS, mdmTypeH, mdmTypeZ, mdmTypeK, mdmTypeA, mdmTypeY
                param.ModeOptions("Port1") = parts(currentIdx): currentIdx = currentIdx + 1
                param.ModeOptions("Port2") = parts(currentIdx): currentIdx = currentIdx + 1
                param.ModeOptions("ACGround") = parts(currentIdx): currentIdx = currentIdx + 1
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