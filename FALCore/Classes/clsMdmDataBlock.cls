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

' --- Private Member Variables ---
Private prvInputValues As Dictionary
Private prvData As Dictionary
Private prvHeadersInOrder As Collection

' --- Public Properties ---
Public Property Get InputValues() As Dictionary
    Set InputValues = prvInputValues
End Property

Public Property Get Data() As Dictionary
    Set Data = prvData
End Property

Public Property Get HeadersInOrder() As Collection
    Set HeadersInOrder = prvHeadersInOrder
End Property

Public Property Get RowCount() As Long
    '@brief Returns the total number of rows in the data block.
    Dim max As Long
    max = 0
    Dim h As Variant
    If prvHeadersInOrder.Count > 0 Then
        For Each h In prvHeadersInOrder
            If prvData(h).Count > max Then
                max = prvData(h).Count
            End If
        Next h
    End If
    RowCount = max
End Property

' --- Initialization ---
Private Sub Class_Initialize()
    Set prvInputValues = New Dictionary
    Set prvData = New Dictionary
    Set prvHeadersInOrder = New Collection
End Sub

' --- Public Methods ---
Public Sub AddInputValue(ByVal key As String, ByVal value As Variant)
    '@brief Adds a key-value pair to the input values dictionary.
    '@param key The name of the input parameter.
    '@param value The value of the input parameter.
    prvInputValues(key) = value
End Sub

Public Sub AddValue(ByVal Header As String, ByVal RowIndex As Long, ByVal Value As Variant)
    '@brief Adds or updates a single value in the data table at a specific row and column (header).
    '@param Header The name of the column header. If the header is new, it will be added.
    '@param RowIndex The 0-based index of the row to modify.
    '@param Value The value to set for the cell.
    '@example db.AddValue "Vd", 0, 0.1
    On Error GoTo ErrorHandler

    ' --- 1. Ensure Header Exists ---
    If Not prvData.Exists(Header) Then
        prvHeadersInOrder.Add Header
        prvData.Add Header, New Collection
    End If

    ' --- 2. Get the target column ---
    Dim column As Collection
    Set column = prvData(Header)

    ' --- 3. Pad column if RowIndex is beyond current size ---
    Dim i As Long
    For i = column.Count To RowIndex - 1
        column.Add Empty ' Pad with Empty
    Next i

    ' --- 4. Insert or Replace Value ---
    If RowIndex < column.Count Then
        ' Replace existing item. This is the awkward part of VBA collections.
        column.Remove RowIndex + 1
        column.Add Value, After:=RowIndex ' After is 0-based index of item *before* insertion point
    Else
        ' This is a new item at the end
        column.Add Value
    End If

    ' --- 5. Ensure all other columns are padded to match the new max length ---
    Dim maxRows As Long
    maxRows = Me.RowCount

    Dim h As Variant
    For Each h In prvHeadersInOrder
        Dim otherColumn As Collection
        Set otherColumn = prvData(h)
        For i = otherColumn.Count To maxRows - 1
            otherColumn.Add Empty
        Next i
    Next h

    Exit Sub
ErrorHandler:
    Debug.Print "Error in AddValue: " & Err.Description
End Sub

Public Sub AddRowFromArray(ByVal values As Variant)
    '@brief Adds a data row from a 1D array of values.
    '@param values A 1D array where each element corresponds to a header in order.
    On Error GoTo ErrorHandler

    If Me.HeadersInOrder.Count = 0 Then
        Err.Raise 5, "AddRowFromArray", "Headers must be defined before adding rows. Use AddValue to add the first row."
    End If
    If UBound(values) - LBound(values) + 1 <> Me.HeadersInOrder.Count Then
        Err.Raise 5, "AddRowFromArray", "The number of values in the array must match the number of headers."
    End If

    Dim newRowIndex As Long
    newRowIndex = Me.RowCount

    Dim i As Long
    For i = 1 To Me.HeadersInOrder.Count
        Me.AddValue Me.HeadersInOrder(i), newRowIndex, values(LBound(values) + i - 1)
    Next i

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

    Dim i As Long, j As Long
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        Dim rowValues() As Variant
        ReDim rowValues(LBound(dataArray, 2) To UBound(dataArray, 2))

        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            rowValues(j) = dataArray(i, j)
        Next j

        Me.AddRowFromArray rowValues
    Next i

    Exit Sub

ErrorHandler:
    Debug.Print "Error in AddRowsFrom2DArray: " & Err.Description
End Sub