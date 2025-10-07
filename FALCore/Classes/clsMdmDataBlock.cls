VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMdmDataBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private Member Variables
Private prvInputValues As Dictionary
Private prvHeaders As Collection
Private prvDataRows As Collection

' Public Properties
Public Property Get InputValues() As Dictionary
    Set InputValues = prvInputValues
End Property

Public Property Get Headers() As Collection
    Set Headers = prvHeaders
End Property

Public Property Get DataRows() As Collection
    Set DataRows = prvDataRows
End Property

' Initialization
Private Sub Class_Initialize()
    Set prvInputValues = New Dictionary
    Set prvHeaders = New Collection
    Set prvDataRows = New Collection
End Sub

' Public Methods
Public Sub AddInputValue(ByVal key As String, ByVal value As Variant)
    '@brief Adds a key-value pair to the input values dictionary.
    '@param key The name of the input parameter.
    '@param value The value of the input parameter.
    If Not prvInputValues.Exists(key) Then
        prvInputValues.Add key, value
    Else
        prvInputValues(key) = value
    End If
End Sub

Public Sub AddHeader(ByVal headerName As String)
    '@brief Adds a column header, preserving its order.
    '@param headerName The name of the header.
    prvHeaders.Add headerName
End Sub

Public Sub AddRow(ByVal rowData As Dictionary)
    '@brief Adds a data row to the block.
    '@param rowData A dictionary representing a single row (header -> value).
    prvDataRows.Add rowData
End Sub

Public Sub AddRowFromArray(ByVal values As Variant)
    '@brief Adds a data row from a 1D array of values.
    '@param values A 1D array where each element corresponds to a header in order.
    On Error GoTo ErrorHandler

    If UBound(values) - LBound(values) + 1 <> prvHeaders.Count Then
        Err.Raise 5, "AddRowFromArray", "The number of values in the array must match the number of headers."
    End If

    Dim rowDict As New Dictionary
    Dim i As Long
    For i = 1 To prvHeaders.Count
        rowDict.Add prvHeaders(i), values(LBound(values) + i - 1)
    Next i

    Me.AddRow rowDict
    Exit Sub

ErrorHandler:
    Debug.Print "Error in AddRowFromArray: " & Err.Description
End Sub

Public Sub AddRowFromValues(ParamArray values() As Variant)
    '@brief Adds a data row from a comma-separated list of values.
    '@param values A comma-separated list of values that corresponds to the headers in order.
    '@example db.AddRowFromValues 0.1, 0.005
    On Error GoTo ErrorHandler

    ' The ParamArray is already a 0-based array, which is compatible with AddRowFromArray
    Me.AddRowFromArray values
    Exit Sub

ErrorHandler:
    Debug.Print "Error in AddRowFromValues: " & Err.Description
End Sub

Public Sub AddRowsFrom2DArray(ByVal dataArray As Variant)
    '@brief Adds multiple data rows from a 2D array.
    '@param dataArray A 2D array where each row represents a data row to be added.
    '@example
    ' Dim data(1 To 2, 1 To 2) As Variant
    ' data(1, 1) = 0: data(1, 2) = 0.01
    ' data(2, 1) = 0.5: data(2, 2) = 0.05
    ' db.AddRowsFrom2DArray data
    On Error GoTo ErrorHandler

    If Not IsArray(dataArray) Then
        Err.Raise 5, "AddRowsFrom2DArray", "Input must be an array."
    End If

    ' Check array dimensions
    Dim numDims As Integer
    On Error Resume Next
    numDims = UBound(dataArray, 2)
    If Err.Number <> 0 Then
        numDims = 1 ' It's a 1D array
        Err.Clear
    Else
        numDims = 2 ' It's a 2D array
    End If
    On Error GoTo ErrorHandler

    If numDims = 1 Then
        ' Handle 1D array as a single row
        Me.AddRowFromArray dataArray
    Else
        ' Handle 2D array
        Dim i As Long, j As Long
        For i = LBound(dataArray, 1) To UBound(dataArray, 1)
            Dim rowValues() As Variant
            ReDim rowValues(LBound(dataArray, 2) To UBound(dataArray, 2))

            For j = LBound(dataArray, 2) To UBound(dataArray, 2)
                rowValues(j) = dataArray(i, j)
            Next j

            Me.AddRowFromArray rowValues
        Next i
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Error in AddRowsFrom2DArray: " & Err.Description
End Sub