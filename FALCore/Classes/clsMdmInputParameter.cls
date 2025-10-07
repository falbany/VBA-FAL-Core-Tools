VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMdmInputParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private Member Variables
Private pName As String
Private pMode As String
Private pModeOptions As Dictionary
Private pSweepType As String
Private pSweepOptions As Dictionary

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

Public Property Get SweepType() As String
    SweepType = pSweepType
End Property

Public Property Let SweepType(ByVal value As String)
    pSweepType = value
End Property

Public Property Get SweepOptions() As Dictionary
    Set SweepOptions = pSweepOptions
End Property

' Initialization
Private Sub Class_Initialize()
    Set pModeOptions = New Dictionary
    Set pSweepOptions = New Dictionary
End Sub