VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Public Sub AddIccapInput(ByVal key As String, ByVal value As String)
    '@brief Adds or updates an ICCAP input in the header.
    '@param key The name of the ICCAP input.
    '@param value The value of the ICCAP input.
    prvIccapInputs(key) = value
End Sub

Public Sub AddIccapOutput(ByVal key As String, ByVal value As String)
    '@brief Adds or updates an ICCAP output in the header.
    '@param key The name of the ICCAP output.
    '@param value The value of the ICCAP output.
    prvIccapOutputs(key) = value
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
    '@example
    ' Dim mdm As New clsMDM
    ' ' ... populate mdm object ...
    ' Dim mdmFileContent As String
    ' mdmFileContent = mdm.ToMdmString()
    ' FalFile.OverwriteTxt "C:\temp\mydata.mdm", mdmFileContent
    On Error GoTo ErrorHandler

    Dim sb As StringBuilder
    Set sb = New StringBuilder

    ' --- Build Header ---
    sb.Append "BEGIN_HEADER" & vbCrLf

    Dim key As Variant

    If prvUserInputs.Count > 0 Then
        sb.Append "USER_INPUTS" & vbCrLf
        For Each key In prvUserInputs.Keys
            sb.Append key & " " & prvUserInputs(key) & vbCrLf
        Next key
    End If

    If prvIccapInputs.Count > 0 Then
        sb.Append "ICCAP_INPUTS" & vbCrLf
        For Each key In prvIccapInputs.Keys
            sb.Append key & " " & prvIccapInputs(key) & vbCrLf
        Next key
    End If

    If prvIccapOutputs.Count > 0 Then
        sb.Append "ICCAP_OUTPUTS" & vbCrLf
        For Each key In prvIccapOutputs.Keys
            sb.Append key & " " & prvIccapOutputs(key) & vbCrLf
        Next key
    End If

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

        ' Data block input values
        Dim blockInputs As StringBuilder
        Set blockInputs = New StringBuilder
        For Each key In dataBlock.InputValues.Keys
            blockInputs.Append key & " " & dataBlock.InputValues(key) & " "
        Next key
        If dataBlock.InputValues.Count > 0 Then
            sb.Append Trim(blockInputs.ToString) & vbCrLf
        End If

        ' Data table
        If Not IsEmpty(dataBlock.Data) Then
            Dim i As Long, j As Long
            Dim dataArray As Variant
            dataArray = dataBlock.Data

            For i = LBound(dataArray, 1) To UBound(dataArray, 1)
                Dim rowSb As StringBuilder
                Set rowSb = New StringBuilder
                For j = LBound(dataArray, 2) To UBound(dataArray, 2)
                    rowSb.Append dataArray(i, j) & " "
                Next j
                sb.Append Trim(rowSb.ToString) & vbCrLf
            Next i
        End If

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
    '@param PrettyPrint (Optional) If True, the output JSON will be formatted with whitespace for readability.
    '@return A JSON string representation of the MDM object.
    '@example
    ' Dim mdm As New clsMDM
    ' mdm.ReadMdmFile "C:\temp\mydata.mdm"
    ' Dim jsonString As String
    ' jsonString = mdm.ToJSON(PrettyPrint:=True)
    ' Debug.Print jsonString
    On Error GoTo ErrorHandler

    Dim mdmObject As Dictionary
    Set mdmObject = New Dictionary

    ' Add header sections
    mdmObject("USER_INPUTS") = Me.UserInputs
    mdmObject("ICCAP_INPUTS") = Me.IccapInputs
    mdmObject("ICCAP_OUTPUTS") = Me.IccapOutputs
    mdmObject("ICCAP_VALUES") = Me.IccapValues

    ' Process data blocks
    Dim dataBlockCollection As Collection
    Set dataBlockCollection = New Collection

    Dim dataBlock As clsMdmDataBlock
    For Each dataBlock In prvDataBlocks
        Dim blockDict As Dictionary
        Set blockDict = New Dictionary

        blockDict("InputValues") = dataBlock.InputValues

        ' Convert tabular data to a collection of dictionaries
        If Not IsEmpty(dataBlock.Data) Then
            Dim tableDataCollection As Collection
            Set tableDataCollection = New Collection

            Dim dataArray As Variant
            dataArray = dataBlock.Data

            Dim headers() As String
            Dim numCols As Long
            numCols = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
            ReDim headers(1 To numCols)

            Dim i As Long, j As Long
            For j = 1 To numCols
                headers(j) = CStr(dataArray(LBound(dataArray, 1), LBound(dataArray, 2) + j - 1))
            Next j

            For i = LBound(dataArray, 1) + 1 To UBound(dataArray, 1)
                Dim rowDict As Dictionary
                Set rowDict = New Dictionary
                For j = 1 To numCols
                    rowDict(headers(j)) = dataArray(i, LBound(dataArray, 2) + j - 1)
                Next j
                tableDataCollection.Add rowDict
            Next i

            blockDict("Data") = tableDataCollection
        End If

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

Public Function To2DArray() As Variant
    '@brief Converts the MDM data into a 2D variant array for easy export to Excel.
    '@return A 2D variant array containing the combined data from all data blocks.
    '@example
    ' Dim mdm As New clsMDM
    ' mdm.ReadMdmFile "C:\temp\mydata.mdm"
    ' Dim dataArray As Variant
    ' dataArray = mdm.To2DArray()
    ' ' Now you can paste dataArray to an Excel sheet
    ' ThisWorkbook.Worksheets("Sheet1").Range("A1").Resize(UBound(dataArray, 1), UBound(dataArray, 2)).value = dataArray
    On Error GoTo ErrorHandler

    If prvDataBlocks.Count = 0 Then
        To2DArray = Empty
        Exit Function
    End If

    Dim colHeaders As Dictionary
    Set colHeaders = New Dictionary
    colHeaders.CompareMode = TextCompare

    Dim dataBlock As clsMdmDataBlock
    Dim totalDataRows As Long
    totalDataRows = 0

    For Each dataBlock In prvDataBlocks
        Dim key As Variant
        For Each key In dataBlock.InputValues.Keys
            If Not colHeaders.Exists(key) Then
                colHeaders.Add key, colHeaders.Count + 1
            End If
        Next key

        If Not IsEmpty(dataBlock.Data) Then
            Dim dataArray As Variant
            dataArray = dataBlock.Data

            Dim c As Long
            For c = LBound(dataArray, 2) To UBound(dataArray, 2)
                Dim headerName As String
                headerName = CStr(dataArray(LBound(dataArray, 1), c))
                If Not colHeaders.Exists(headerName) Then
                    colHeaders.Add headerName, colHeaders.Count + 1
                End If
            Next c

            totalDataRows = totalDataRows + (UBound(dataArray, 1) - LBound(dataArray, 1))
        End If
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

    Dim currentRow As Long
    currentRow = 2

    For Each dataBlock In prvDataBlocks
        If Not IsEmpty(dataBlock.Data) Then
            Dim dataArray As Variant
            dataArray = dataBlock.Data

            Dim r As Long
            For r = LBound(dataArray, 1) + 1 To UBound(dataArray, 1)
                For Each key In dataBlock.InputValues.Keys
                    outputArray(currentRow, colHeaders(key)) = dataBlock.InputValues(key)
                Next key

                Dim c As Long
                For c = LBound(dataArray, 2) To UBound(dataArray, 2)
                    Dim headerName As String
                    headerName = CStr(dataArray(LBound(dataArray, 1), c))
                    outputArray(currentRow, colHeaders(headerName)) = dataArray(r, c)
                Next c

                currentRow = currentRow + 1
            Next r
        End If
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

    If Not FalFile.FileExist(filePath) Then
        ReadMdmFile = False
        Exit Function
    End If

    prvMdmContent = FalFile.ReadFile(filePath)

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
    '@brief Parses a single data block's content.
    '@param blockContent The string content between BEGIN_DB and END_DB.
    '@param dataBlock The clsMdmDataBlock object to populate.
    '@return True if parsing is successful, False otherwise.
    On Error GoTo ErrorHandler

    Dim lines() As String
    lines = Split(blockContent, vbCrLf)

    Dim i As Long
    Dim line As String
    Dim dataStartIndex As Long
    dataStartIndex = -1

    Dim allInputNames As Dictionary
    Set allInputNames = New Dictionary

    Dim key As Variant
    For Each key In prvIccapInputs.Keys
        allInputNames(key) = 1
    Next key
    For Each key In prvUserInputs.Keys
        allInputNames(key) = 1
    Next key

    ' Find the start of the data table by looking for a line starting with a known input name.
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))
        If line <> "" Then
            Dim firstWord As String
            firstWord = Split(line)(0)
            If allInputNames.Exists(firstWord) Then
                dataStartIndex = i
                Exit For
            End If
        End If
    Next i

    If dataStartIndex = -1 Then
        If Trim(blockContent) = "" Then
            prvParseSingleDataBlock = True
            Exit Function
        End If
        prvParseSingleDataBlock = False
        Exit Function
    End If

    ' Lines before dataStartIndex contain the input values for this block.
    For i = LBound(lines) To dataStartIndex - 1
        line = Trim(lines(i))
        If line <> "" Then
            Dim parts() As String
            parts = Split(line)
            Dim k As Long
            For k = LBound(parts) To UBound(parts) Step 2
                If k + 1 <= UBound(parts) Then
                    dataBlock.InputValues.Add parts(k), parts(k + 1)
                End If
            Next k
        End If
    Next i

    ' Lines from dataStartIndex onwards are the data table.
    Dim dataRows As New Collection
    For i = dataStartIndex To UBound(lines)
        line = Trim(lines(i))
        If line <> "" Then
            dataRows.Add line
        End If
    Next i

    If dataRows.Count > 0 Then
        Dim tempLine As String
        tempLine = Trim(dataRows(1))
        Do While InStr(tempLine, "  ") > 0
            tempLine = Replace(tempLine, "  ", " ")
        Loop
        Dim numCols As Long
        numCols = UBound(Split(tempLine)) + 1

        Dim dataArray() As Variant
        ReDim dataArray(1 To dataRows.Count, 1 To numCols)

        Dim r As Long, c As Long
        For r = 1 To dataRows.Count
            tempLine = Trim(dataRows(r))
            Do While InStr(tempLine, "  ") > 0
                tempLine = Replace(tempLine, "  ", " ")
            Loop
            Dim rowParts() As String
            rowParts = Split(tempLine)

            For c = 1 To numCols
                If c - 1 <= UBound(rowParts) Then
                   dataArray(r, c) = rowParts(c - 1)
                Else
                   dataArray(r, c) = "" ' Handle ragged data
                End If
            Next c
        Next r

        dataBlock.Data = dataArray
    End If

    prvParseSingleDataBlock = True
    Exit Function

ErrorHandler:
    prvParseSingleDataBlock = False
End Function

Private Function prvParseHeader() As Boolean
    '@brief Parses the header section of the MDM file.
    '@return True if the header was parsed successfully, False otherwise.
    On Error GoTo ErrorHandler

    Dim headerContent As String
    Dim startPos As Long
    Dim endPos As Long

    startPos = InStr(1, prvMdmContent, "BEGIN_HEADER", vbTextCompare)
    endPos = InStr(1, prvMdmContent, "END_HEADER", vbTextCompare)

    If startPos = 0 Or endPos = 0 Then
        prvParseHeader = False
        Exit Function
    End If

    headerContent = Mid(prvMdmContent, startPos + Len("BEGIN_HEADER"), endPos - startPos - Len("BEGIN_HEADER"))

    Dim lines() As String
    lines = Split(headerContent, vbCrLf)

    Dim currentSection As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim(lines(i))

        If line = "" Or Left(line, 1) = "!" Then ' Skip empty lines and comments
            GoTo NextLine
        End If

        Select Case UCase(line)
            Case "USER_INPUTS"
                currentSection = "USER_INPUTS"
            Case "ICCAP_INPUTS"
                currentSection = "ICCAP_INPUTS"
            Case "ICCAP_OUTPUTS"
                currentSection = "ICCAP_OUTPUTS"
            Case "ICCAP_VALUES"
                currentSection = "ICCAP_VALUES"
            Case Else
                If currentSection <> "" Then
                    Dim parts() As String
                    parts = Split(line)

                    If UBound(parts) >= 0 Then
                        Dim key As String
                        key = parts(0)

                        Dim value As String
                        value = Trim(Mid(line, Len(key) + 1))

                        Select Case currentSection
                            Case "USER_INPUTS"
                                prvUserInputs.Add key, value
                            Case "ICCAP_INPUTS"
                                prvIccapInputs.Add key, value
                            Case "ICCAP_OUTPUTS"
                                prvIccapOutputs.Add key, value
                            Case "ICCAP_VALUES"
                                prvIccapValues.Add key, value
                        End Select
                    End If
                End If
        End Select
NextLine:
    Next i

    prvParseHeader = True
    Exit Function

ErrorHandler:
    prvParseHeader = False
End Function