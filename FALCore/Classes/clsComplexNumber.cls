VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComplexNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --- Private Member Variables ---
Private pReal As Double
Private pImaginary As Double

' --- Public Properties ---
Public Property Get Real() As Double
    Real = pReal
End Property

Public Property Let Real(ByVal value As Double)
    pReal = value
End Property

Public Property Get Imaginary() As Double
    Imaginary = pImaginary
End Property

Public Property Let Imaginary(ByVal value As Double)
    pImaginary = value
End Property

' --- Initialization ---
Private Sub Class_Initialize()
    pReal = 0
    pImaginary = 0
End Sub

' --- Public Methods ---
Public Function Create(ByVal realPart As Double, ByVal imaginaryPart As Double) As clsComplexNumber
    '@brief Factory method to create and initialize a new clsComplexNumber object.
    '@param realPart The real component of the number.
    '@param imaginaryPart The imaginary component of the number.
    '@return A new, initialized clsComplexNumber object.
    Dim newComplex As New clsComplexNumber
    newComplex.Real = realPart
    newComplex.Imaginary = imaginaryPart
    Set Create = newComplex
End Function

Public Function ToString() As String
    '@brief Returns a string representation of the complex number.
    '@return A string in the format "a + bi".
    If pImaginary >= 0 Then
        ToString = CStr(pReal) & " + " & CStr(pImaginary) & "i"
    Else
        ToString = CStr(pReal) & " - " & CStr(Abs(pImaginary)) & "i"
    End If
End Function