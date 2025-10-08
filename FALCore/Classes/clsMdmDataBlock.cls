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
' Stores the column data. Outer key is Header, Inner key is RowIndex.
Private prvData As Dictionary
' Stores the headers in the order they appear to preserve column order.
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
    '@brief Returns the total number of rows, calculated by the highest row index used.
    Dim maxIndex As Long
    maxIndex = -1

    Dim h As Variant
    Dim rowIndex As Variant

    If prvData.Count > 0 Then
        For Each h In prvData.Keys
            Dim innerDict As Dictionary
            Set innerDict = prvData(h)
            For Each rowIndex In innerDict.Keys
                If CLng(rowIndex) > maxIndex Then
                    maxIndex = CLng(rowIndex)
                End If
            Next rowIndex
        Next h
    End If

    RowCount = maxIndex + 1
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

    ' --- 1. Ensure Header Exists in both lookup collections ---
    Dim headerExists As Boolean
    Dim h As Variant
    For Each h In prvHeadersInOrder
        If CStr(h) = Header Then
            headerExists = True
            Exit For
        End If
    Next h

    If Not headerExists Then
        prvHeadersInOrder.Add Header
        prvData.Add Header, New Dictionary
    End If

    ' --- 2. Get the target column dictionary and set the value ---
    Dim columnDict As Dictionary
    Set columnDict = prvData(Header)
    columnDict(RowIndex) = Value ' Add or update the value for the given row index

    Exit Sub

ErrorHandler:
    Debug.Print "Error in AddValue: " & Err.Description
End Sub

' Note: Other data entry methods (AddRow, etc.) will be re-implemented in subsequent steps.