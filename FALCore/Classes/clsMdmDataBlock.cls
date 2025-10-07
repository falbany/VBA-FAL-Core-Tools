VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMdmDataBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private Member Variables
Private prvInputValues As Object ' Dictionary
Private prvData As Variant ' 2D Array

' Public Properties
Public Property Get InputValues() As Object
    Set InputValues = prvInputValues
End Property

Public Property Let InputValues(ByVal value As Object)
    Set prvInputValues = value
End Property

Public Property Get Data() As Variant
    Data = prvData
End Property

Public Property Let Data(ByVal value As Variant)
    prvData = value
End Property

' Initialization
Private Sub Class_Initialize()
    Set prvInputValues = CreateObject("Scripting.Dictionary")
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