VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMdmOutputParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private Member Variables
Private pName As String
Private pMode As String
Private pModeOptions As Dictionary
Private pUnit As String
Private pCompliance As String
Private pType As String

' Public Properties
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal value As String)
    pName = value
End Property

Public Property Get Mode() As String
    Mode = pMode
End Property

Public Property Let Mode(ByVal value As String)
    pMode = value
End Property

Public Property Get ModeOptions() As Dictionary
    Set ModeOptions = pModeOptions
End Property

Public Property Get Unit() As String
    Unit = pUnit
End Property

Public Property Let Unit(ByVal value As String)
    pUnit = value
End Property

Public Property Get Compliance() As String
    Compliance = pCompliance
End Property

Public Property Let Compliance(ByVal value As String)
    pCompliance = value
End Property

Public Property Get Type() As String
    Type = pType
End Property

Public Property Let Type(ByVal value As String)
    pType = value
End Property

' Initialization
Private Sub Class_Initialize()
    Set pModeOptions = New Dictionary
End Sub