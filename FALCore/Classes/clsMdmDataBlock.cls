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