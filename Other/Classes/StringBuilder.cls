VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' Class Name: StringBuilder
' Purpose: Provides a way to efficiently build and manipulate strings.
' Inputs: None
' Outputs: None
' https://codereview.stackexchange.com/questions/196076/bringing-the-system-text-StringBuilder-up-to-lightning-speed-in-the-vba
' Based off vba_StringBuilder:Class - http://blog.livedoor.jp/midorityo/archives/50744656.html
' https://github.com/retailcoder/VBA-StringBuilder
' https://learn.microsoft.com/en-us/dotnet/api/system.text.StringBuilder?view=net-7.0&redirectedfrom=MSDN
'**************************************************************************************************

Option Explicit





'TO USE'
'Sub TestStringBuilderAppendFormat()
'    Dim sb As New clsStringBuilder
'    'Test AppendFormat Numbers: https://docs.microsoft.com/en-us/dotnet/standard/base-types/standard-numeric-format-strings
'    sb.AppendLine "Using Standard Numeric Format Strings - Currency (""C"") Format Specifier"
'    sb.AppendFormat "{0:C2}", 123.456
'    sb.AppendLine "Currency With Alignment Arguments"
'    sb.AppendLine "   Beginning Balance           Ending Balance"
'    sb.AppendFormat "   {0,-28:C2}{1,14:C2}", 16305.32, 18794.16
'    sb.AppendLine
'    sb.AppendLine "The Decimal (""D"") Format Specifier"
'    sb.AppendFormat "{0:D}", 12345
'    sb.AppendLine
'    sb.AppendLine "8 Digit Format Specifier"
'    sb.AppendFormat "{0:D8}", 12345
'    sb.AppendLine
'    sb.AppendLine "The Percent (""P"") Format Specifier"
'    sb.AppendFormat "{0:P}", 0.2468013
'    sb.AppendLine "8 Digit Format Specifier"
'    sb.AppendFormat "{0:P8}", 0.2468013
'    sb.AppendLine
'    sb.AppendLine "Custom Tests" & vbNewLine & String(50, "*")
'
'    sb.AppendLine "AppendFormat: Dates"
'    sb.AppendFormat "Day ## {0:dd}, Day Name {0:dddd}, Month ## {0:MM}, Month Name {0:MMMM}, YYYY {0:yyyy}", Date
'    sb.InsertFormat "Date {0}", 0, 0, "Formats: "
'
'    sb.AppendLine "AppendFormat: ParamArray"
'    sb.AppendFormat "Jack {0} Jill {1} up {2} hill {3} fetch {4} pail {5} water {6} fell {7} And {8} his {9} And {10} came {11} after", "and", "Went", "the", "To", "a", "of", "Jack", "down", "broke", "crown,", "Jill", "tumbling"
'
'    sb.AppendLine "AppendFormat: Array"
'    sb.AppendFormat "Jack {0} Jill {1} up {2} hill {3} fetch {4} pail {5} water {6} fell {7} And {8} his {9} And {10} came {11} after", Array("and", "Went", "the", "To", "a", "of", "Jack", "down", "broke", "crown,", "Jill", "tumbling")
'    Debug.Print sb.ToString
'
'End Sub

Private Type Members
    AppendsBufferSize As Long   ' Size of the buffer for storing appended strings
    Builder As Object           ' String builder object
    index As Long               ' Index for tracking the current position in the buffer
End Type

Private AppendsBuffer() As String ' Array to store appended strings
Private m As Members ' Instance of the Members type

Private Sub Class_Initialize()
    '**************************************************************************************************
    ' Sub Name: Class_Initialize
    ' Purpose: Initializes the StringBuilder class
    ' Inputs: None
    ' Outputs: None
    '**************************************************************************************************
    Const DEFAULT_APPENDSBUFFER_SIZE = 10000
    Set m.Builder = CreateObject("System.Text.StringBuilder")
    setAppendsBufferSize DEFAULT_APPENDSBUFFER_SIZE
    m.index = -1
End Sub

Private Sub Class_Terminate()
    '**************************************************************************************************
    ' Sub Name: Class_Terminate
    ' Purpose: Cleans up resources when the StringBuilder class is terminated
    ' Inputs: None
    ' Outputs: None
    '**************************************************************************************************
    Me.Clear
'    m.Builder.Clear
    Set m.Builder = Nothing
End Sub

Public Property Get Capacity() As Long
    '**************************************************************************************************
    ' Property Name: Capacity
    ' Purpose: Gets the current capacity of the string builder
    ' Inputs: None
    ' Outputs: Long - The current capacity
    '**************************************************************************************************
    Capacity = m.Builder.Capacity
End Property

Public Property Let Capacity(ByVal Length As Long)
    '**************************************************************************************************
    ' Property Name: Capacity
    ' Purpose: Sets the capacity of the string builder
    ' Inputs: Length - The new capacity
    ' Outputs: None
    '**************************************************************************************************
    m.Builder.Capacity = Length
End Property

Public Property Get Length() As Long
    '**************************************************************************************************
    ' Property Name: Length
    ' Purpose: Gets the current length of the string builder
    ' Inputs: None
    ' Outputs: Long - The current length
    '**************************************************************************************************
    Length = m.Builder.Length
End Property

Public Property Let Length(ByVal lLength As Long)
    '**************************************************************************************************
    ' Property Name: Length
    ' Purpose: Sets the length of the string builder
    ' Inputs: lLength - The new length
    ' Outputs: None
    '**************************************************************************************************
    m.Builder.Length = lLength
End Property

Public Function MaxCapacity() As Long
    '**************************************************************************************************
    ' Function Name: MaxCapacity
    ' Purpose: Gets the maximum capacity of the string builder
    ' Inputs: None
    ' Outputs: Long - The maximum capacity
    '**************************************************************************************************
    MaxCapacity = m.Builder.MaxCapacity
End Function

Public Function Append(ByRef Text As String, Optional repeatCount As Long = 0) As StringBuilder
    '**************************************************************************************************
    ' Function Name: Append
    ' Purpose: Appends a string to the string builder
    ' Inputs: Text - The string to append
    '         repeatCount - The number of times to append the string (optional, default is 0)
    ' Outputs: StringBuilder - The updated string builder
    '**************************************************************************************************
    m.index = m.index + 1
    AppendsBuffer(m.index) = Text
    If repeatCount > 0 Then Me.Append Text, (repeatCount - 1)
    If m.index = m.AppendsBufferSize Then Compact
    Set Append = Me
End Function

Public Function AppendFormat(ByVal format As Variant, ParamArray args() As Variant) As StringBuilder
    '**************************************************************************************************
    ' Function Name: AppendFormat
    ' Purpose: Appends a formatted string to the string builder
    ' Inputs: format - The format string
    '         Args() - The arguments to format (optional)
    ' Outputs: StringBuilder - The updated string builder
    '**************************************************************************************************
    Dim newArgs() As Variant
    Dim j As Long
    Me.Compact

    If IsArray(args(0)) Then
        newArgs = args(0)
    Else
        ReDim newArgs(UBound(args))
        For j = 0 To UBound(args)
            newArgs(j) = args(j)
        Next
    End If

    m.Builder.AppendFormat_4 format, newArgs
    Set AppendFormat = Me
End Function

Public Function InsertFormat(ByVal format As Variant, startIndex As Long, repeatCount As Long, ParamArray args() As Variant) As StringBuilder
    '**************************************************************************************************
    ' Function Name: InsertFormat
    ' Purpose: Inserts a formatted string into the string builder at the specified index
    ' Inputs: format - The format string
    '         startIndex - The index at which to insert the formatted string
    '         repeatCount - The number of times to insert the formatted string
    '         Args() - The arguments to format (optional)
    ' Outputs: StringBuilder - The updated string builder
    '**************************************************************************************************
    Dim newArgs() As Variant, sb As New StringBuilder
    Dim j As Long
    If IsArray(args(0)) Then
        newArgs = args(0)
    Else
        ReDim newArgs(UBound(args))
        For j = 0 To UBound(args)
            newArgs(j) = args(j)
        Next
    End If

    sb.AppendFormat format, newArgs
    Me.Insert startIndex, sb.ToString
    Set InsertFormat = Me
End Function

Public Function AppendLine(Optional ByRef Text As String, Optional repeatCount As Long = 0) As StringBuilder
    '**************************************************************************************************
    ' Function Name: AppendLine
    ' Purpose: Appends a string with a line break to the string builder
    ' Inputs: Text - The string to append
    '         repeatCount - The number of times to append the string (optional, default is 0)
    ' Outputs: StringBuilder - The updated string builder
    '**************************************************************************************************
    m.index = m.index + 1
    AppendsBuffer(m.index) = Text & vbNewLine
    If repeatCount > 0 Then Me.AppendLine Text, (repeatCount - 1)
    If m.index = m.AppendsBufferSize Then Compact
    Set AppendLine = Me
End Function

Public Sub Clear()
    '**************************************************************************************************
    ' Sub Name: Clear
    ' Purpose: Clears the string builder
    ' Inputs: None
    ' Outputs: None
    '**************************************************************************************************
    Me.Length = 0
End Sub

Public Sub Compact()
    '**************************************************************************************************
    ' Sub Name: Compact
    ' Purpose: Appends the buffered strings to the string builder and resets the buffer
    ' Inputs: None
    ' Outputs: None
    '**************************************************************************************************
    If m.index > -1 Then
        m.Builder.Append_3 Join(AppendsBuffer, vbNullString)
        m.index = -1
    End If
    ReDim AppendsBuffer(m.AppendsBufferSize)
End Sub

Public Function EnsureCapacity(ByVal Length As Long) As StringBuilder
    '**************************************************************************************************
    ' Function Name: EnsureCapacity
    ' Purpose: Ensures that the string builder has the specified capacity
    ' Inputs: Length - The desired capacity
    ' Outputs: StringBuilder - The updated string builder
    '**************************************************************************************************
    m.Builder.EnsureCapacity Length
    Set EnsureCapacity = Me
End Function

Public Function Insert(startIndex As Long, Text As String, Optional repeatCount As Long = 0) As StringBuilder
    '**************************************************************************************************
    ' Function Name: Insert
    ' Purpose: Inserts a string into the string builder at the specified index
    ' Inputs: startIndex - The index at which to insert the string
    '         Text - The string to insert
    '         repeatCount - The number of times to insert the string (optional, default is 0)
    ' Outputs: StringBuilder - The updated string builder
    '**************************************************************************************************
    Me.Compact
    m.Builder.Insert_2 startIndex, Text
    If repeatCount > 0 Then Me.Insert startIndex, Text, (repeatCount - 1)
    Set Insert = Me
End Function

Public Function Remove(ByVal startIndex As Long, ByVal Length As Long) As StringBuilder
    '**************************************************************************************************
    ' Function Name: Remove
    ' Purpose: Removes a specified number of characters from the string builder starting at the specified index
    ' Inputs: startIndex - The index at which to start removing characters
    '         Length - The number of characters to remove
    ' Outputs: StringBuilder - The updated string builder
    '**************************************************************************************************
    Me.Compact
    m.Builder.Remove startIndex, Length
    Set Remove = Me
End Function

Public Function Replace(oldValue As String, newValue As String, Optional startIndex As Long = -1, Optional repeatCount As Long = 0) As StringBuilder
    '**************************************************************************************************
    ' Function Name: Replace
    ' Purpose: Replaces occurrences of a string within the string builder with another string
    ' Inputs: oldValue - The string to replace
    '         newValue - The replacement string
    '         startIndex - The index at which to start replacing (optional, default is -1)
    '         repeatCount - The number of occurrences to replace (optional, default is 0)
    ' Outputs: StringBuilder - The updated string builder
    '**************************************************************************************************
    Me.Compact

    If startIndex > -1 And repeatCount > 0 Then
        m.Builder.Replace_2 oldValue, newValue, startIndex, repeatCount
    Else
        m.Builder.Replace oldValue, newValue
    End If
    Set Replace = Me
End Function

Public Sub setAppendsBufferSize(ByVal Length As Long)
    '**************************************************************************************************
    ' Sub Name: setAppendsBufferSize
    ' Purpose: Sets the size of the buffer for storing appended strings
    ' Inputs: Length - The size of the buffer
    ' Outputs: None
    '**************************************************************************************************
    m.AppendsBufferSize = Length
    Me.Compact
End Sub

Function ToString(Optional ByVal startIndex As Long = -1, Optional ByVal Length As Long = 0) As String
    '**************************************************************************************************
    ' Function Name: ToString
    ' Purpose: Converts the string builder to a string
    ' Inputs: startIndex - The index at which to start extracting the string (optional, default is -1)
    '         Length - The length of the extracted string (optional, default is 0)
    ' Outputs: String - The extracted string
    '**************************************************************************************************
    Me.Compact

    If startIndex > -1 And Length > 0 Then
        ToString = m.Builder.ToString_2(startIndex, Length)
    Else
        ToString = m.Builder.ToString
    End If
End Function

