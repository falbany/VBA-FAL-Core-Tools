VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMdmValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@ModuleDescription("A class to hold structured information about a single validation error found in a clsMDM object.")

' --- Private Member Variables ---
Private pErrorType As MdmValidationErrorType
Private pMessage As String
Private pBlockIndex As Long
Private pOffendingItem As String

' --- Public Properties ---
Public Property Get ErrorType() As MdmValidationErrorType
    '@brief The category of the validation error.
    ErrorType = pErrorType
End Property
Public Property Let ErrorType(ByVal Value As MdmValidationErrorType)
    pErrorType = Value
End Property

Public Property Get Message() As String
    '@brief A detailed, user-friendly description of the error.
    Message = pMessage
End Property
Public Property Let Message(ByVal Value As String)
    pMessage = Value
End Property

Public Property Get BlockIndex() As Long
    '@brief The 0-based index of the data block where the error occurred. Returns -1 if not applicable.
    BlockIndex = pBlockIndex
End Property
Public Property Let BlockIndex(ByVal Value As Long)
    pBlockIndex = Value
End Property

Public Property Get OffendingItem() As String
    '@brief The name of the item (e.g., header, parameter) that caused the error.
    OffendingItem = pOffendingItem
End Property
Public Property Let OffendingItem(ByVal Value As String)
    pOffendingItem = Value
End Property

' --- Initialization ---
Private Sub Class_Initialize()
    ' Default to -1 to indicate not applicable
    pBlockIndex = -1
End Sub

' --- Public Methods ---
Public Function ToString() As String
    '@brief Returns a string representation of the validation error.
    Dim str As String
    str = "Error Type: " & pErrorType & ", Message: '" & pMessage & "'"
    If pBlockIndex <> -1 Then
        str = str & ", Block Index: " & pBlockIndex
    End If
    If pOffendingItem <> "" Then
        str = str & ", Item: '" & pOffendingItem & "'"
    End If
    ToString = str
End Function