VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BetterArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An array class for VBA providing features found in more modern languages "
'@Folder("VBABetterArray")
'@ModuleDescription("An array class for VBA providing features found in more modern languages ")
' Copyright (c) 2021 Nick Morgan. See License: https://github.com/Senipah/VBA-Better-Array/blob/master/LICENSE
'@IgnoreModule IIfSideEffect, AssignmentNotUsed, FunctionReturnValueDiscarded, ProcedureNotUsed, FunctionReturnValueNotUsed, FunctionReturnValueAlwaysDiscarded

Option Explicit

'''''''''''''''''''
' Constant Fields '
'''''''''''''''''''

Private Const DEFAULT_CAPACITY As Long = 4
' NOTE: placeholder
Private Const MAX_ARRAY_LENGTH As Long = &H7FEFFFFF
Private Const OBJECT_REPR As String = "OBJECT"
Private Const MISSING_LONG As Long = -9999
Private Const CHR_QUOTE As String = """"
Private Const CHR_COMMA As String = ","

Public Enum ArrayTypes
    BA_UNDEFINED
    BA_UNALLOCATED
    BA_ONEDIMENSION
    BA_MULTIDIMENSION
    BA_JAGGED
End Enum

' Public scope as used in unit tests
Public Enum ErrorCodes
    EC_START = vbObjectError + 512
    EC_EXPECTED_RANGE_OBJECT
    EC_EXPECTED_COLLECTION_OBJECT
    EC_MAX_DIMENSIONS_LIMIT
    EC_EXCEEDS_MAX_SORT_DEPTH
    EC_EXPECTED_JAGGED_ARRAY
    EC_EXPECTED_MULTIDIMENSION_ARRAY
    EC_EXPECTED_ARRAY
    EC_NULL_STRING
    EC_UNALLOCATED_ARRAY
    EC_UNDEFINED_ARRAY
    EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION
    EC_EXPECTED_VARIANT_ARRAY
    EC_EXCEEDS_MAX_ARRAY_LENGTH
    EC_STRING_TYPE_EXPECTED
    EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE
    EC_CANNOT_SORT_OBJECTS
    EC_END
End Enum

Public Enum ComparisonType
    CT_EQUALITY
    CT_LIKENESS
End Enum

Public Enum SortMethods
    SM_TIMSORT
    SM_QUICKSORT_RECURSIVE
    SM_QUICKSORT_ITERATIVE
End Enum

Private Type ErrorDefinition
    Number As Long
    Source As String
    Description As String
End Type


''''''''''
' Fields '
''''''''''

Private Type TFields
    Capacity As Long
    Length As Long
    LowerBound As Long
    items() As Variant
    ArrayType As ArrayTypes
    ErrorDefinitions(EC_START To EC_END) As ErrorDefinition
    LowerBoundSet As Boolean
    SortMethod As SortMethods
End Type

Private Type TString
    Text As String
    Length As Long
    ByteLength As Long
End Type

Private This As TFields

'@Description("Constructor")
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Constructor"
    Me.Capacity = DEFAULT_CAPACITY
    This.ArrayType = BA_UNALLOCATED
    PopulateErrorDefinitions
End Sub

'''''''''''''''''''''
' Public Properties '
'''''''''''''''''''''

'@Description("Returns the capacity of the internal array.")
Public Property Get Capacity() As Long
Attribute Capacity.VB_Description = "Returns the capacity of the internal array."
    Capacity = This.Capacity
End Property

'@Description("Sets the capacity of the internal array.")
'@Ignore ProcedureNotUsed
Public Property Let Capacity(ByVal value As Long)
Attribute Capacity.VB_Description = "Sets the capacity of the internal array."
    '@Ignore FunctionReturnValueDiscarded
    If value < 0 Then err.Raise 9
    If value <> This.Capacity Then
        If This.Capacity > 0 Then
            If GetArrayLength(This.items) <> value Then
                Dim NewItems() As Variant
                NewItems = This.items
                ReDim Preserve NewItems(This.LowerBound To (value + This.LowerBound - 1))
                InternalItems = NewItems
            End If
        Else
            ReDim This.items(This.LowerBound To (DEFAULT_CAPACITY + This.LowerBound - 1))
        End If
        This.Capacity = UBound(This.items) - This.LowerBound + 1
    End If
End Property

'@Description("Returns the number of entries stored in the array.")
Public Property Get Length() As Long
Attribute Length.VB_Description = "Returns the number of entries stored in the array."
    Length = This.Length
End Property

'@Description("Returns largest used index in the array.")
Public Property Get UpperBound() As Long
Attribute UpperBound.VB_Description = "Returns largest used index in the array."
    If This.ArrayType = BA_UNALLOCATED Then
        UpperBound = -1
    Else
        UpperBound = This.Length + This.LowerBound - 1
    End If
End Property

'@Description("Returns smallest used index in the array.")
Public Property Get LowerBound() As Long
Attribute LowerBound.VB_Description = "Returns smallest used index in the array."
    LowerBound = This.LowerBound
End Property

'@Description("Sets the starting index of the array. If an array is already stored then it will be re-indexed to match the new value")
Public Property Let LowerBound(ByVal value As Long)
Attribute LowerBound.VB_Description = "Sets the starting index of the array. If an array is already stored then it will be re-indexed to match the new value"
    This.LowerBoundSet = True
    If value <> This.LowerBound Then
        This.LowerBound = value
        InternalItems = Rebase()
        This.Capacity = GetArrayLength(This.items)
    End If
End Property

'@DefaultMember
'@Description("Returns the element in the array stored at the specified index.")
Public Property Get item(ByVal index As Long) As Variant
Attribute item.VB_Description = "Returns the element in the array stored at the specified index."
Attribute item.VB_UserMemId = 0
    If index <= This.Length Then
        ' not using letOrSetElement here as it might try to access the Items letter
        If IsObject(This.items(index)) Then
            Set item = This.items(index)
        Else
            item = This.items(index)
        End If
    Else
        '@Ignore FunctionReturnValueDiscarded
        err.Raise 9
    End If
End Property

'@Description("Assigns the passed element to the array at the specified index.")
Public Property Let item(ByVal index As Long, ByVal Element As Variant)
Attribute item.VB_Description = "Assigns the passed element to the array at the specified index."
    If Me.UpperBound >= index Then
        If index < This.LowerBound Then
            '@Ignore FunctionReturnValueDiscarded
            Me.Unshift Element
        Else
            LetOrSetElement This.items(index), Element
        End If
    Else
        '@Ignore FunctionReturnValueDiscarded
        Me.Push Element
    End If
End Property

'@Description("Returns the stored array.")
Public Property Get items() As Variant
Attribute items.VB_Description = "Returns the stored array."
    Dim result() As Variant
    result = InternalItems
    If This.ArrayType = ArrayTypes.BA_MULTIDIMENSION Then
        If IsJaggedArray(result) Then result = JaggedToMulti(result)
    End If
    items = result
End Property

'@Description("Stores the passed array.")
Public Property Let items(ByVal values As Variant)
Attribute items.VB_Description = "Stores the passed array."
    Const CONVERT_MD_TO_JAGGED As Boolean = True
    Const CONVERT_NESTED_JAGGED As Boolean = True
    Dim LocalLowerBound As Long
    Dim LocalValues() As Variant
    Dim TypeSet As Boolean
    
    If TypeName(values) = TypeName(Me) Then
        LocalValues = values.items
    ElseIf IsArray(values) Then
        This.ArrayType = GetArrayType(values)
        TypeSet = True
        If This.ArrayType = BA_UNALLOCATED Then
            LocalValues = GetEmptyArray
        Else
            LocalLowerBound = LBound(values)
            If Not This.LowerBoundSet Then
                This.LowerBound = LocalLowerBound
            End If
            LocalValues = ConvertArrayForStorage( _
                values, _
                This.ArrayType, _
                CONVERT_MD_TO_JAGGED, _
                CONVERT_NESTED_JAGGED _
            )
        End If
    Else
        ' If Values is not an array then clear our internal array and push the Values arg
        ' NOTE: this possibly violates PoLA - assess for refactor
        '@Ignore FunctionReturnValueDiscarded
        If IsEmpty(values) And This.ArrayType = ArrayTypes.BA_UNDEFINED Then
            RaiseError EC_EXPECTED_ARRAY, "Items", "Values"
        Else
            Me.Clear.Push values
        End If
        Exit Property
    End If
    
    If Not TypeSet Then
        This.ArrayType = GetArrayType(LocalValues)
    End If
    
    InternalItems = LocalValues
    This.Length = GetArrayLength(LocalValues)
    This.Capacity = This.Length
    If This.Capacity < DEFAULT_CAPACITY Then
        Me.Capacity = DEFAULT_CAPACITY
    End If

End Property

'@Description("Retrieves the stored array type.")
Public Property Get ArrayType() As ArrayTypes
Attribute ArrayType.VB_Description = "Retrieves the stored array type."
    ArrayType = This.ArrayType
End Property

'@Description("Sets the stored array type.")
Public Property Let ArrayType(ByVal NewType As ArrayTypes)
Attribute ArrayType.VB_Description = "Sets the stored array type."
    Select Case NewType
    Case ArrayTypes.BA_UNDEFINED
        RaiseError EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE, "ArrayType", "NewType"
    Case ArrayTypes.BA_UNALLOCATED
        If This.ArrayType <> BA_UNALLOCATED Then
            Me.ResetToDefault
        End If
    Case ArrayTypes.BA_ONEDIMENSION
        Select Case This.ArrayType
        Case ArrayTypes.BA_MULTIDIMENSION, ArrayTypes.BA_JAGGED
            RaiseError EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE, "ArrayType", "NewType"
        Case Else
            This.ArrayType = NewType
        End Select
    Case Else
        This.ArrayType = NewType
    End Select
End Property

'@Description("Retrieves the sorting algorithm used by the Sort() method.")
Public Property Get SortMethod() As SortMethods
Attribute SortMethod.VB_Description = "Retrieves the sorting algorithm used by the Sort() method."
    SortMethod = This.SortMethod
End Property

'@Description("Specifies the sorting algorithm used by the Sort() method.")
Public Property Let SortMethod(ByVal Method As SortMethods)
Attribute SortMethod.VB_Description = "Specifies the sorting algorithm used by the Sort() method."
    This.SortMethod = Method
End Property


''''''''''''''''''''''
' Private Properties '
''''''''''''''''''''''

'@Description("Retrieves the stored items trimmed to the correct length. Does not change structure (i.e jagged to multi).")
Private Property Get InternalItems() As Variant()
Attribute InternalItems.VB_Description = "Retrieves the stored items trimmed to the correct length. Does not change structure (i.e jagged to multi)."
    Dim result() As Variant
    If This.ArrayType <> BA_UNALLOCATED And This.ArrayType <> BA_UNDEFINED Then
        result = This.items
        If This.Capacity > This.Length Then
            If This.Length = 0 Then
                ReDim Preserve result(This.LowerBound To This.LowerBound)
            ElseIf This.Length > 0 Then
                ReDim Preserve result(This.LowerBound To Me.UpperBound)
            End If
        End If
    Else
        result = GetEmptyArray
    End If
    InternalItems = result
End Property

'@Description("Assigns the passed array to this.items")
'@Ignore MisleadingByRefParameter
Private Property Let InternalItems(ByRef value() As Variant)
Attribute InternalItems.VB_Description = "Assigns the passed array to this.items"
    This.items = value
End Property


''''''''''''''''''
' Public Methods '
''''''''''''''''''

'@Description("Adds the provided argments to the array at the next available index.")
Public Function Push(ParamArray args() As Variant) As Variant
Attribute Push.VB_Description = "Adds the provided argments to the array at the next available index."
    Dim Element As Variant

    If This.ArrayType = ArrayTypes.BA_UNALLOCATED Or _
       This.ArrayType = ArrayTypes.BA_UNDEFINED Then
        This.ArrayType = ArrayTypes.BA_ONEDIMENSION
    End If

    For Each Element In args
        If This.Length = This.Capacity Then
            EnsureCapacity This.Length + 1
        End If
        If IsArray(Element) Then
            Dim ArrayElement() As Variant
            Dim ArrayElementType As ArrayTypes
            ArrayElement = Element
            If This.ArrayType = BA_ONEDIMENSION Then
                ArrayElementType = GetArrayType(ArrayElement)
                If ArrayElementType = BA_MULTIDIMENSION Then
                    This.ArrayType = BA_MULTIDIMENSION
                Else
                    This.ArrayType = BA_JAGGED
                End If
            End If
            If LBound(ArrayElement) <> This.LowerBound Then
                ArrayElement = Rebase(ArrayElement, ArrayElementType)
            End If
            LetOrSetElement This.items(This.Length + This.LowerBound), ArrayElement
        Else
            LetOrSetElement This.items(This.Length + This.LowerBound), Element
        End If
        inc This.Length
    Next
    Push = This.Length
End Function

'@Description("Removes the last element from the array and returns that element.")
Public Function Pop() As Variant
Attribute Pop.VB_Description = "Removes the last element from the array and returns that element."
    Dim result As Variant
    Dim NewItems() As Variant
    If This.Length > 0 Then
        result = This.items(Me.UpperBound)
        NewItems = Me.Slice(This.LowerBound, Me.UpperBound)
        Me.items = NewItems
    End If
    Pop = result
End Function

'@Description("Removes the first element from the array and returns that element.")
Public Function Shift() As Variant
Attribute Shift.VB_Description = "Removes the first element from the array and returns that element."
    Dim NewItems() As Variant
    Dim result As Variant
    If This.Length > 0 Then
        result = This.items(This.LowerBound)
        NewItems = Me.Slice(This.LowerBound + 1)
        Me.items = NewItems
    End If
    Shift = result
End Function

'@Description("Adds one or more elements to the beginning of an array. Returns the new length of the array.")
Public Function Unshift(ParamArray args() As Variant) As Long
Attribute Unshift.VB_Description = "Adds one or more elements to the beginning of an array. Returns the new length of the array."
    Dim NewItems() As Variant
    Dim OldItems() As Variant
    Dim OldType As ArrayTypes
    
    NewItems = args
    OldType = This.ArrayType
    OldItems = InternalItems
    Me.items = NewItems
    If OldType <> BA_UNALLOCATED And OldType <> BA_UNDEFINED Then
        '@Ignore FunctionReturnValueDiscarded
        Me.Concat OldItems
        This.ArrayType = OldType
    End If
    Unshift = This.Length
End Function

'@Description("Returns a string representing the array structure and its elements.")
Public Function ToString( _
        Optional ByVal PrettyPrint As Boolean, _
        Optional ByVal Separator As String = CHR_COMMA, _
        Optional ByVal OpeningDelimiter As String = "{", _
        Optional ByVal ClosingDelimiter As String = "}", _
        Optional ByVal QuoteStrings As Boolean _
    ) As String
Attribute ToString.VB_Description = "Returns a string representing the array structure and its elements."
    Dim LocalArrayType As ArrayTypes
    Dim result As String
    Dim LocalItems() As Variant
    Dim Sep As String
    
    Sep = IIf(PrettyPrint, Separator & Space(1), Separator)
    
    LocalItems = InternalItems
    LocalArrayType = GetArrayType(LocalItems)
    If Not LocalArrayType = ArrayTypes.BA_UNDEFINED And _
       Not LocalArrayType = ArrayTypes.BA_UNALLOCATED Then
        If LocalArrayType = ArrayTypes.BA_MULTIDIMENSION Then
            LocalItems = MultiToJagged(LocalItems)
        End If
        RecursiveToString _
            SourceArray:=LocalItems, _
            PrettyPrint:=PrettyPrint, _
            Separator:=Sep, _
            OpeningDelimiter:=OpeningDelimiter, _
            ClosingDelimiter:=ClosingDelimiter, _
            QuoteStrings:=QuoteStrings
        result = StringBuilder(Final:=True)
    End If
    ToString = result
End Function

'@Description("Determines whether the array includes a certain value among its entries.")
Public Function Includes( _
        ByVal SearchElement As Variant, _
        Optional ByVal fromIndex As Long = MISSING_LONG, _
        Optional ByVal Recurse As Boolean _
    ) As Boolean
Attribute Includes.VB_Description = "Determines whether the array includes a certain value among its entries."
    Dim LocalLength As Long
    Dim CurrentIndex As Long
    Dim SearchArray() As Variant
    
    SearchArray = InternalItems
    LocalLength = This.Length
    
    If LocalLength = 0 Then
        Includes = False
        Exit Function
    End If
        
    If fromIndex > This.LowerBound Then
        CurrentIndex = fromIndex
    ElseIf fromIndex = MISSING_LONG Then
        CurrentIndex = This.LowerBound
    Else
        CurrentIndex = LocalLength + fromIndex
    End If
    
    Includes = RecursiveIncludes(SearchElement, SearchArray, CurrentIndex, Recurse:=Recurse)
End Function

'@Description("Determines whether the type of any of the elements in the array matches the typename.")
Public Function IncludesType( _
        ByVal SearchTypeName As String, _
        Optional ByVal fromIndex As Long = MISSING_LONG, _
        Optional ByVal Recurse As Boolean _
    ) As Boolean
Attribute IncludesType.VB_Description = "Determines whether the type of any of the elements in the array matches the typename."
    Dim LocalLength As Long
    Dim CurrentIndex As Long
    Dim SearchArray() As Variant
    
    SearchArray = InternalItems
    LocalLength = This.Length
    
    If LocalLength = 0 Then
        IncludesType = False
        Exit Function
    End If
        
    If fromIndex > This.LowerBound Then
        CurrentIndex = fromIndex
    ElseIf fromIndex = MISSING_LONG Then
        CurrentIndex = This.LowerBound
    Else
        CurrentIndex = LocalLength + fromIndex
    End If
    
    IncludesType = RecursiveIncludes(SearchTypeName, SearchArray, CurrentIndex, True, Recurse)
End Function

'@Description("Determines whether all elements in the array match a value.")
Public Function Every( _
        ByVal SearchElement As Variant, _
        Optional ByVal fromIndex As Long = MISSING_LONG _
    ) As Boolean
Attribute Every.VB_Description = "Determines whether all elements in the array match a value."
    
    Dim LocalLength As Long
    Dim CurrentIndex As Long
    Dim SearchArray() As Variant
    
    SearchArray = InternalItems
    LocalLength = This.Length
    
    If LocalLength = 0 Then
        Every = False
        Exit Function
    End If
        
    If fromIndex > This.LowerBound Then
        CurrentIndex = fromIndex
    ElseIf fromIndex = MISSING_LONG Then
        CurrentIndex = This.LowerBound
    Else
        CurrentIndex = LocalLength + fromIndex
    End If
    
    Every = RecursiveEvery(SearchElement, SearchArray, CurrentIndex)
End Function

'@Description("Determines whether the type of all elements match the provided typename.")
Public Function EveryType( _
        ByVal SearchTypeName As String, _
        Optional ByVal fromIndex As Long = MISSING_LONG _
    ) As Boolean
Attribute EveryType.VB_Description = "Determines whether the type of all elements match the provided typename."
    
    Dim LocalLength As Long
    Dim CurrentIndex As Long
    Dim SearchArray() As Variant
    
    SearchArray = InternalItems
    LocalLength = This.Length
    
    If LocalLength = 0 Then
        EveryType = False
        Exit Function
    End If
        
    If fromIndex > This.LowerBound Then
        CurrentIndex = fromIndex
    ElseIf fromIndex = MISSING_LONG Then
        CurrentIndex = This.LowerBound
    Else
        CurrentIndex = LocalLength + fromIndex
    End If
    
    EveryType = RecursiveEvery(SearchTypeName, SearchArray, CurrentIndex, True)
    
End Function

'@Description("Returns a new array that contains the keys for each index in the array.")
Public Function keys() As Variant()
Attribute keys.VB_Description = "Returns a new array that contains the keys for each index in the array."
    Dim i As Long
    Dim LocalLowerBound As Long
    Dim result() As Variant
    
    If This.ArrayType = BA_UNDEFINED Then
        RaiseError EC_UNDEFINED_ARRAY, "Keys"
    ElseIf This.ArrayType = BA_UNALLOCATED Then
        RaiseError EC_UNALLOCATED_ARRAY, "Keys"
    Else
        LocalLowerBound = This.LowerBound
        ReDim result(0 To This.Length - 1)
        For i = LBound(result) To UBound(result)
            result(i) = i + LocalLowerBound
        Next
    End If
    keys = result
End Function

'@Description("Returns the largest value in a list of values. If arguments are provided it will return the largest argument. If no arguments are provided it will return the largest element in the stored array")
Public Function max(ParamArray args() As Variant) As Variant
Attribute max.VB_Description = "Returns the largest value in a list of values. If arguments are provided it will return the largest argument. If no arguments are provided it will return the largest element in the stored array"
    Dim LocalItems() As Variant
    
    If UBound(args) < LBound(args) Then          'Test if no args provided
        LocalItems = InternalItems
    Else
        LocalItems = args
    End If
    
    max = RecursiveMax(LocalItems)
End Function

'@Description("Returns the smallest value in a list of values. If arguments are provided it will return the smallest argument. If no arguments are provided it will return the smallest element in the stored array")
Public Function Min(ParamArray args() As Variant) As Variant
Attribute Min.VB_Description = "Returns the smallest value in a list of values. If arguments are provided it will return the smallest argument. If no arguments are provided it will return the smallest element in the stored array"
    Dim LocalItems() As Variant
    
    If UBound(args) < LBound(args) Then          'Test if no args provided
        LocalItems = InternalItems
    Else
        LocalItems = args
    End If
    
    Min = RecursiveMin(LocalItems)
End Function

'@Description("Returns a new array containing the portion of the stored array between the StartIndex and EndIndex")
Public Function Slice( _
        ByVal startIndex As Long, _
        Optional ByVal endIndex As Long = MISSING_LONG _
    ) As Variant()
Attribute Slice.VB_Description = "Returns a new array containing the portion of the stored array between the StartIndex and EndIndex"
    Dim LocalLength As Long
    Dim RelativeStart As Long
    Dim RelativeEnd As Long
    Dim OldIndex As Long
    Dim Final As Long
    Dim count As Long
    Dim NewIndex As Long
    Dim LocalItems() As Variant
    Dim result() As Variant
    
    LocalItems = InternalItems
    LocalLength = This.Length
    RelativeStart = startIndex
    If RelativeStart < LBound(LocalItems) Then
        If RelativeStart < 0 Then
            OldIndex = max((LocalLength + RelativeStart), LBound(LocalItems))
        Else
            OldIndex = max((LocalLength - RelativeStart), LBound(LocalItems))
        End If
    Else
        OldIndex = Min(RelativeStart, LocalLength)
    End If
    If endIndex = MISSING_LONG Then
        RelativeEnd = LocalLength + LBound(LocalItems)
    Else
        RelativeEnd = endIndex
    End If
    If RelativeEnd < LBound(LocalItems) Then
        Final = max((LocalLength + RelativeEnd), LBound(LocalItems))
    Else
        Final = Min(RelativeEnd, LocalLength + LBound(LocalItems))
    End If
    
    NewIndex = LBound(LocalItems)
    count = max(Final - OldIndex, 0) + LBound(LocalItems)
    If count > NewIndex Then
        ReDim result(NewIndex To count - 1)
        Do While OldIndex < Final
            If OldIndex >= LBound(LocalItems) And OldIndex <= UBound(LocalItems) Then
                LetOrSetElement result(NewIndex), LocalItems(OldIndex)
                inc NewIndex
                inc OldIndex
            End If
        Loop
        If This.ArrayType = BA_MULTIDIMENSION Then
            Slice = JaggedToMulti(result)
        Else
            Slice = result
        End If
    End If
    
End Function

'@Description("Stores the values contained in an Excel range in the internal array")
Public Function FromExcelRange( _
        ByRef FromRange As Object, _
        Optional ByVal DetectLastRow As Boolean, _
        Optional ByVal DetectLastColumn As Boolean _
    ) As BetterArray
Attribute FromExcelRange.VB_Description = "Stores the values contained in an Excel range in the internal array"

    If TypeName(FromRange) = "Range" Then
        Dim StartColumn As Long
        Dim EndColumn As Long
        Dim StartRow As Long
        Dim EndRow As Long
        
        With FromRange
            StartColumn = .column
            StartRow = .Row
            EndColumn = .column + .Columns.count - 1
            EndRow = .Row + .Rows.count - 1
        End With
        
        With FromRange.parent
            If DetectLastColumn Then
                EndColumn = .Cells.item(StartRow, .Columns.count).End(xlToLeft).column
            End If
            
            If DetectLastRow Then
                EndRow = .Cells.item(.Rows.count, StartColumn).End(xlUp).Row
            End If
            
            Me.items = .Range(.Cells(StartRow, StartColumn), .Cells(EndRow, EndColumn)).value
        End With
        
        ' convert arrays of single column or row to 1d
        If StartColumn = EndColumn And StartRow <> EndRow Then
            Me.items = Me.ExtractSegment(, StartColumn)
        ElseIf StartColumn <> EndColumn And StartRow = EndRow Then
            Me.items = Me.ExtractSegment(StartRow)
        End If
        
    Else
        RaiseError ErrorCodes.EC_EXPECTED_RANGE_OBJECT, "FromExcelRange()", "FromRange"
    End If
    
    Set FromExcelRange = Me
End Function

'@Description("Extracts the segment of the array located at the specified indices.")
Public Function ExtractSegment( _
        Optional ByVal RowIndex As Long = MISSING_LONG, _
        Optional ByVal ColumnIndex As Long = MISSING_LONG _
    ) As Variant()
Attribute ExtractSegment.VB_Description = "Extracts the segment of the array located at the specified indices."
    Dim i As Long
    Dim LocalRowIndex As Long
    Dim LocalColumnIndex As Long
    Dim NestedBounds() As Long
    Dim LocalItems() As Variant
    Dim result() As Variant
    
    LocalItems = InternalItems

    If RowIndex = MISSING_LONG Then
        If ColumnIndex = MISSING_LONG Then
            ' if no coords provided return full array
            ' calls to Me.Items instead of returning localItems to covert md from jagged
            result = Me.items
        Else
            ' Row number missing, column number present
            Select Case This.ArrayType
            Case BA_ONEDIMENSION
                If ColumnIndex >= LBound(LocalItems) And ColumnIndex <= UBound(LocalItems) Then
                    LocalColumnIndex = ColumnIndex
                Else
                    LocalColumnIndex = LBound(LocalItems)
                End If
                result = Array(LocalItems(LocalColumnIndex))
            Case BA_JAGGED, BA_MULTIDIMENSION
                ' return full column
                NestedBounds = GetMaxBoundsAtDimension(LocalItems, 2)
                If ColumnIndex >= NestedBounds(0) And ColumnIndex <= NestedBounds(1) Then
                    LocalColumnIndex = ColumnIndex
                Else
                    LocalColumnIndex = This.LowerBound
                End If
                ReDim result(LBound(LocalItems) To UBound(LocalItems))
                For i = LBound(LocalItems) To UBound(LocalItems)
                    result(i) = LocalItems(i)(LocalColumnIndex)
                Next
            Case BA_UNALLOCATED
                result = LocalItems
            Case Else
            End Select
        End If
    Else
        ' Row Number not missing
        If RowIndex >= LBound(LocalItems) And RowIndex <= UBound(LocalItems) Then
            LocalRowIndex = RowIndex
        Else
            LocalRowIndex = LBound(LocalItems)
        End If
        If ColumnIndex = MISSING_LONG Then
            ' Row number present, column number missing
            Select Case This.ArrayType
            Case BA_ONEDIMENSION
                result = Array(LocalItems(LocalRowIndex))
            Case BA_JAGGED, BA_MULTIDIMENSION
                result = LocalItems(LocalRowIndex)
            Case BA_UNALLOCATED
                result = LocalItems
            Case Else
            End Select
        Else
            ' Row number present, column number present
            Select Case This.ArrayType
            Case BA_ONEDIMENSION
                ' ignore column number arg as array is 1d
                result = Array(LocalItems(LocalRowIndex))
            Case BA_JAGGED, BA_MULTIDIMENSION
                ' return intersection of rowIndex and column number
                NestedBounds = GetMaxBoundsAtDimension(LocalItems, 2)
                If ColumnIndex >= NestedBounds(0) And ColumnIndex <= NestedBounds(1) Then
                    LocalColumnIndex = ColumnIndex
                Else
                    LocalColumnIndex = This.LowerBound
                End If
                If IsArray(LocalItems(LocalRowIndex)(LocalColumnIndex)) Then
                    result = LocalItems(LocalRowIndex)(LocalColumnIndex)
                Else
                    result = Array(LocalItems(LocalRowIndex)(LocalColumnIndex))
                End If
            Case BA_UNALLOCATED
                result = LocalItems
            Case Else
            End Select
        End If
    End If
    
    ExtractSegment = result
End Function

'@Description("Writes the values stored in the array to an excel worksheet starting at the specified range.")
Public Function ToExcelRange( _
        ByRef Destination As Object, _
        Optional ByVal TransposeValues As Boolean _
    ) As Object
Attribute ToExcelRange.VB_Description = "Writes the values stored in the array to an excel worksheet starting at the specified range."
    Const TARGET_APPLICATION As String = "Microsoft Excel"
    Const TARGET_OBJECT As String = "Range"
    Dim LocalRange As Object
    Dim LocalItems() As Variant
    Dim Depth As Long
    Dim LengthRows As Long
    Dim LengthColumns As Long
    Dim AvailableRows As Long
    Dim AvailableColumns As Long
    Dim DestType As String
    Dim DestApplication As String
       
    On Error Resume Next
        DestType = TypeName(Destination)
        DestApplication = Destination.Application.name
    On Error GoTo 0
    
    If DestType = TARGET_OBJECT And DestApplication = TARGET_APPLICATION Then
        AvailableRows = Destination.parent.Rows.count - Destination.Row + 1
        AvailableColumns = Destination.parent.Columns.count - Destination.column + 1
        LocalItems = InternalItems
        Depth = GetJaggedArrayDepth(LocalItems)
        If Depth > 0 Then 'check array is allocated
            If Depth = 1 Then
                '1d array
                LocalItems = ConvertOneDimensionArrayToJagged(LocalItems)
            End If
            'jagged array - multidim arrays should also be stored as jagged internally
            Const OutputDepth As Long = 2 'need 2 dimensions for tabular representation output
            LocalItems = JaggedToMulti(LocalItems, OutputDepth, EnsureScalar:=True)
            If TransposeValues Then
                LocalItems = Transpose2DArray(LocalItems)
            End If
            
            LocalItems = TrimColumnsMultidimensionArray(LocalItems, AvailableColumns)
            LocalItems = TrimRowsMultidimensionArray(LocalItems, AvailableRows)
            LengthRows = UBound(LocalItems, 1) - LBound(LocalItems, 1) + 1
            LengthColumns = UBound(LocalItems, 2) - LBound(LocalItems, 2) + 1
                        
            Set LocalRange = Destination.Resize( _
                RowSize:=LengthRows, _
                ColumnSize:=LengthColumns _
            )
            LocalRange.value = LocalItems
        '@Ignore EmptyElseBlock
        Else
            'array must be unalloc
            'do nothing(?)
        End If
    Else
        RaiseError ErrorCodes.EC_EXPECTED_RANGE_OBJECT, "ToExcelRange()", "Destination"
    End If
    
    Set ToExcelRange = LocalRange
End Function

'@Description("Tests if the stored array is sorted in ascending order.")
Public Function isSorted(Optional ByVal ColumnIndex As Long = MISSING_LONG) As Boolean
Attribute isSorted.VB_Description = "Tests if the stored array is sorted in ascending order."
    Dim i As Long
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim LocalColumnIndex As Long
    Dim Depth As Long
    Dim result As Boolean
    Dim LocalItems() As Variant
    Dim StoredType As ArrayTypes
    
    result = True
    
    LocalItems = InternalItems
    LocalLowerBound = LBound(LocalItems)
    LocalUpperBound = UBound(LocalItems)
    StoredType = GetArrayType(LocalItems)
    If StoredType <> BA_UNDEFINED And StoredType <> BA_UNALLOCATED Then
        Select Case StoredType
        Case BA_ONEDIMENSION
            For i = LocalLowerBound To LocalUpperBound - 1
                If LocalItems(i) > LocalItems(i + 1) Then
                    result = False
                    Exit For
                End If
            Next
        Case BA_MULTIDIMENSION
            ' should never be true as all md arrays stored as jagged internally
            If ColumnIndex = MISSING_LONG Then
                LocalColumnIndex = LBound(LocalItems, 2)
            Else
                LocalColumnIndex = CLng(ColumnIndex)
            End If
            For i = LocalLowerBound To LocalUpperBound - 1
                If LocalItems(i, LocalColumnIndex) > LocalItems(i + 1, LocalColumnIndex) Then
                    result = False
                    Exit For
                End If
            Next
        Case BA_JAGGED
            Depth = GetJaggedArrayDepth(LocalItems)
            If Depth > 2 Then
                isSorted = False
                RaiseError EC_EXCEEDS_MAX_SORT_DEPTH, "IsSorted"
            Else
                For i = LocalLowerBound To LocalUpperBound - 1
                    If i = LocalLowerBound Then
                        If ColumnIndex = MISSING_LONG Then
                            LocalColumnIndex = LBound(LocalItems(i))
                        Else
                            LocalColumnIndex = CLng(ColumnIndex)
                        End If
                    End If
                    If LocalItems(i)(LocalColumnIndex) > LocalItems(i + 1)(LocalColumnIndex) Then
                        result = False
                        Exit For
                    End If
                Next
            End If
        End Select
    End If
    isSorted = result
End Function



'@Description("returns the first index at which a given element can be found in the array")
Public Function IndexOf( _
        ByVal SearchElement As Variant, _
        Optional ByVal fromIndex As Long = MISSING_LONG, _
        Optional ByVal CompType As ComparisonType _
    ) As Long
Attribute IndexOf.VB_Description = "returns the first index at which a given element can be found in the array"
    Dim LocalItems() As Variant
    Dim RelativeStart As Long
    Dim CurrentIndex As Long
    
    If CompType = CT_LIKENESS Then
        If TypeName(SearchElement) <> "String" Then
            RaiseError EC_STRING_TYPE_EXPECTED, "IndexOf", "searchElement"
        End If
    End If
    
    If This.Length = 0 Then
        IndexOf = MISSING_LONG
        Exit Function
    End If

    LocalItems = InternalItems

    If fromIndex = MISSING_LONG Then
        RelativeStart = LBound(LocalItems)
    Else
        RelativeStart = fromIndex
    End If
    
    If RelativeStart >= LBound(LocalItems) Then
        CurrentIndex = RelativeStart
    Else
        If RelativeStart > 0 Then
            CurrentIndex = LBound(LocalItems)
        Else
            CurrentIndex = UBound(LocalItems) + RelativeStart
        End If
        
        If CurrentIndex < LBound(LocalItems) Then
            CurrentIndex = LBound(LocalItems)
        End If
    End If
    
    Dim IsMatch As Boolean
    Do While CurrentIndex <= UBound(LocalItems)
        Select Case CompType
        Case ComparisonType.CT_LIKENESS
            IsMatch = CStr(LocalItems(CurrentIndex)) Like CStr(SearchElement)
        Case Else
            IsMatch = ElementsAreEqual(SearchElement, LocalItems(CurrentIndex))
        End Select
        If IsMatch Then
            IndexOf = CurrentIndex
            Exit Function
        End If
        inc CurrentIndex
    Loop
    
    IndexOf = MISSING_LONG
End Function

'@Description("returns the first index at which a given element can be found in the array")
Public Function LastIndexOf( _
        ByVal SearchElement As Variant, _
        Optional ByVal fromIndex As Long = MISSING_LONG, _
        Optional ByVal CompType As ComparisonType _
    ) As Long
Attribute LastIndexOf.VB_Description = "returns the first index at which a given element can be found in the array"
    Dim LocalItems() As Variant
    Dim CurrentIndex As Long
    
    If CompType = CT_LIKENESS Then
        If TypeName(SearchElement) <> "String" Then
            RaiseError EC_STRING_TYPE_EXPECTED, "IndexOf", "searchElement"
        End If
    End If
    
    If This.Length = 0 Then
        LastIndexOf = MISSING_LONG
        Exit Function
    End If

    LocalItems = InternalItems
    
    If fromIndex = MISSING_LONG Then
        CurrentIndex = UBound(LocalItems)
    Else
        If fromIndex >= LBound(LocalItems) Then
            CurrentIndex = Min(fromIndex, UBound(LocalItems))
        ElseIf fromIndex < 0 Then
            CurrentIndex = UBound(LocalItems) + fromIndex
        End If
        
    End If
    
    Dim IsMatch As Boolean
    Do While CurrentIndex >= LBound(LocalItems)
        Select Case CompType
        Case ComparisonType.CT_LIKENESS
            IsMatch = CStr(LocalItems(CurrentIndex)) Like CStr(SearchElement)
        Case Else
            IsMatch = ElementsAreEqual(SearchElement, LocalItems(CurrentIndex))
        End Select
        If IsMatch Then
            LastIndexOf = CurrentIndex
            Exit Function
        End If
        dec CurrentIndex
    Loop
    
    LastIndexOf = MISSING_LONG
End Function

'@Description("Removes the item at the specified index. Returns the new length of the array")
Public Function Remove(ByVal index As Long) As Long
Attribute Remove.VB_Description = "Removes the item at the specified index. Returns the new length of the array"
    Dim RelativeIndex As Long
    Dim LocalType As ArrayTypes
    
    RelativeIndex = MISSING_LONG
    LocalType = This.ArrayType
    
    If index >= This.LowerBound Then
        If index <= Me.UpperBound Then
            RelativeIndex = index
        End If
    Else
        If index < 0 Then
            RelativeIndex = Me.UpperBound + index
            If RelativeIndex < This.LowerBound Then RelativeIndex = MISSING_LONG
        End If
    End If
    
    If RelativeIndex <> MISSING_LONG Then
        Dim BeforeSlice() As Variant
        Dim AfterSlice() As Variant
        Dim BeforeExists As Boolean
        Dim AfterExists As Boolean
        
        If RelativeIndex > This.LowerBound Then
            BeforeSlice = Me.Slice(This.LowerBound, RelativeIndex)
            BeforeExists = True
        End If
        If RelativeIndex < Me.UpperBound Then
            AfterSlice = Me.Slice(RelativeIndex + 1)
            AfterExists = True
        End If
        If BeforeExists Then
            If AfterExists Then
                Me.items = InternalConcat(BeforeSlice, AfterSlice)
            Else
                Me.items = BeforeSlice
            End If
        Else
            If AfterExists Then
                Me.items = AfterSlice
            End If
        End If
        If BeforeExists Or AfterExists Then
            If This.ArrayType = BA_JAGGED And LocalType = BA_MULTIDIMENSION Then
                This.ArrayType = LocalType
            End If
        Else
            Me.Clear
        End If
    End If
    
    Remove = This.Length
End Function

'@Description("changes the contents of the array by removing or replacing existing elements and/or adding new elements in place.")
Public Function Splice( _
        ByVal startIndex As Long, _
        ParamArray args() As Variant _
    ) As Variant()
Attribute Splice.VB_Description = "changes the contents of the array by removing or replacing existing elements and/or adding new elements in place."
    Dim LocalItems() As Variant
    Dim ActualStart As Long
    Dim ActualDeleteCount As Long
    Dim argsCount As Long
    Dim LocalLength As Long
    Dim LocalArgs() As Variant
    Dim result() As Variant
    Dim i As Long
    Dim tempArray() As Variant
    Dim ItemCount As Long
    Dim TempItems() As Variant
    
    LocalArgs = args
    LocalItems = InternalItems
    LocalLength = UBound(LocalItems) + 1
    If startIndex < LBound(LocalItems) Then
        ActualStart = max(LocalLength + startIndex, LBound(LocalItems))
    Else
        ActualStart = Min(startIndex, LocalLength)
    End If
    argsCount = GetArrayLength(LocalArgs)
    If argsCount = 0 Then
        ' only a start index has been supplied
        ' we're gonna delete everything
        ItemCount = 0
        ActualDeleteCount = LocalLength - ActualStart
    Else
        ' first arg in args is delete count, second onwards are insertion items
        ItemCount = argsCount - 1
        'deleteCount is first element in args.
        ActualDeleteCount = Min(max(args(LBound(args)), 0), LocalLength - ActualStart)
    End If
    
    If LocalLength + ItemCount - ActualDeleteCount > MAX_ARRAY_LENGTH Then
        RaiseError EC_EXCEEDS_MAX_ARRAY_LENGTH, "Splice"
        Exit Function
    End If

    ' get items to be deleted for return val
    If ActualDeleteCount > 0 Then
        ReDim result(0 To ActualDeleteCount - 1)
        i = 0
        Do While i < ActualDeleteCount
            LetOrSetElement result(i), LocalItems(ActualStart + i)
            inc i
        Loop
    Else
        ' return empty array if nothing to be deleted
        ReDim result(0)
    End If
    
    If ItemCount > 0 Then
        ReDim TempItems(0 To ItemCount - 1)
        For i = LBound(TempItems) To UBound(TempItems)
            TempItems(i) = args(i + 1)
        Next
    End If
            
    If ItemCount < ActualDeleteCount Then
        i = ActualStart
        Do While i < (LocalLength - ActualDeleteCount)
            LetOrSetElement LocalItems(i + ItemCount), LocalItems(i + ActualDeleteCount)
            inc i
        Loop
        
        
        i = LocalLength
        Dim TempBetterArray As BetterArray
        Set TempBetterArray = New BetterArray
        TempBetterArray.items = LocalItems
        Do While i > (LocalLength - ActualDeleteCount + ItemCount)
            TempBetterArray.Remove (i - 1)
            ' Perform ? DeletePropertyOrThrow(O, ! ToString(k - 1)).
            dec i
        Loop
        LocalItems = TempBetterArray.items
        Set TempBetterArray = Nothing
    ElseIf ItemCount > ActualDeleteCount Then
        i = (LocalLength - ActualDeleteCount)
        tempArray = LocalItems
        ReDim Preserve tempArray(This.LowerBound To i + ItemCount - 1) As Variant
        Do While i > ActualStart
            LetOrSetElement tempArray(i + ItemCount - 1), LocalItems(i + ActualDeleteCount - 1)
            dec i
        Loop
        LocalItems = tempArray
    End If
    
    If ItemCount > 0 Then
        i = ActualStart
        Dim j As Long
        For j = LBound(TempItems) To UBound(TempItems)
            LocalItems(i) = TempItems(j)
            inc i
        Next
    End If
    
    InternalItems = LocalItems
    
    Splice = result
End Function


''''''''''''''''''''''''''''
' Chainable Public Methods '
''''''''''''''''''''''''''''

'@Description("Recursvely fills (modifies) all the elements of an array from a start index (default zero) to an end index (default array length) with a passed value.")
Public Function Fill( _
        ByVal value As Variant, _
        Optional ByVal startIndex As Long = MISSING_LONG, _
        Optional ByVal endIndex As Long = MISSING_LONG _
    ) As BetterArray
Attribute Fill.VB_Description = "Recursvely fills (modifies) all the elements of an array from a start index (default zero) to an end index (default array length) with a passed value."
    Dim LocalItems() As Variant
    Dim RelativeStart As Long
    Dim RelativeEnd As Long
    
    LocalItems = InternalItems
    If startIndex = MISSING_LONG Then
        RelativeStart = LBound(LocalItems)
    Else
        If startIndex < 0 Then
            RelativeStart = max(Me.UpperBound + startIndex, LBound(LocalItems))
        ElseIf startIndex < LBound(LocalItems) Then
            RelativeStart = LBound(LocalItems)
        Else
            RelativeStart = Min(startIndex, Me.UpperBound)
        End If
    End If
    
    If endIndex = MISSING_LONG Then
        RelativeEnd = Me.UpperBound
    Else
        If endIndex < 0 And endIndex < LBound(LocalItems) Then
            RelativeEnd = max(Me.UpperBound + endIndex, LBound(LocalItems))
        ElseIf endIndex < LBound(LocalItems) Then
            RelativeEnd = Me.UpperBound
        Else
            RelativeEnd = Min(endIndex, Me.UpperBound)
        End If
    End If
    InternalItems = RecursiveFill(LocalItems, value, RelativeStart, RelativeEnd)
    Set Fill = Me
End Function

'@Description("Parses a string representing an array and stores the outcome.")
Public Function ParseFromString( _
        ByVal SourceString As String, _
        Optional ByVal ValueSeparator As String = CHR_COMMA, _
        Optional ByVal ArrayOpenDelimiter As String, _
        Optional ByVal ArrayClosingDelimiter As String _
    ) As BetterArray
Attribute ParseFromString.VB_Description = "Parses a string representing an array and stores the outcome."
    Dim Opener As String
    Dim Closer As String
    Dim ArraysAreDelimited As Boolean
    Dim result() As Variant
    
    If Len(SourceString) > 0 Then
        If ArrayOpenDelimiter = vbNullString And ArrayClosingDelimiter = vbNullString Then
            Dim FirstChar As String
            Dim LastChar As String
            FirstChar = Left$(SourceString, 1)
            LastChar = Right$(SourceString, 1)
            ' no array delimiters supplied
            ' see if the diff in ascii values are for first and last chars are less than 3
            ' this covers () [] {} <>
            ' escape numbers incase array is "1,2,3"
            ' escape ucase letters incase array is "A,B,C"
            ' escape lcase letters incase array is "a,b,c"
            If (Asc(LastChar) - Asc(FirstChar) < 3) And _
                    (Asc(FirstChar) < 65 Or Asc(FirstChar) > 90) And _
                    (Asc(FirstChar) < 97 Or Asc(FirstChar) > 122) And _
                    Not IsNumeric(FirstChar) Then
                Opener = FirstChar
                Closer = LastChar
                ArraysAreDelimited = True
            End If
        Else
            ' array delimiters supplied
            Opener = ArrayOpenDelimiter
            Closer = ArrayClosingDelimiter
            ArraysAreDelimited = True
        End If
        
        If ArraysAreDelimited Then
            result = ParseDelimitedArrayString(SourceString, ValueSeparator, Opener, Closer)
        Else
            result = ParseArraySegmentFromString(SourceString, 1, 0, ValueSeparator)
        End If
    Else
        RaiseError EC_NULL_STRING, "ParseFromString", "SourceString"
    End If
    
    Me.items = result
    Set ParseFromString = Me
End Function

'@Description("Swaps rows for columns & vice versa.")
Public Function Transpose() As BetterArray
Attribute Transpose.VB_Description = "Swaps rows for columns & vice versa."
    Dim result() As Variant
    Dim LocalItems() As Variant
    Dim LocalType As ArrayTypes
    Dim StoredType As ArrayTypes
    StoredType = This.ArrayType
    LocalItems = InternalItems
    LocalType = GetArrayType(LocalItems)
    Select Case LocalType
    Case ArrayTypes.BA_ONEDIMENSION
        result = Transpose1DArray(LocalItems)
    Case ArrayTypes.BA_MULTIDIMENSION
        ' Sould Never match as MD arrs stored as jagged internally
        result = Transpose2DArray(LocalItems)
    Case ArrayTypes.BA_JAGGED
        result = TransposeArrayOfArrays(LocalItems)
    Case Else
        result = LocalItems
    End Select
    Me.items = result
    If StoredType = BA_MULTIDIMENSION Then
        LocalType = GetArrayType(result)
        If LocalType = BA_JAGGED Then
            This.ArrayType = BA_MULTIDIMENSION
        End If
    End If
    Set Transpose = Me
End Function

'@Description("Returns a new BetterArray instance containing the same values as the current instance.")
Public Function Clone() As BetterArray
Attribute Clone.VB_Description = "Returns a new BetterArray instance containing the same values as the current instance."
    Dim result As BetterArray
    Set result = New BetterArray
    result.LowerBound = This.LowerBound
    result.items = Me.items
    Set Clone = result
End Function

'@Description("Clears the current array and resets capacity to default value")
Public Function ResetToDefault() As BetterArray
Attribute ResetToDefault.VB_Description = "Clears the current array and resets capacity to default value"
    This.LowerBound = 0
    ReDim This.items(This.LowerBound To DEFAULT_CAPACITY + This.LowerBound)
    This.Length = 0
    Me.Capacity = DEFAULT_CAPACITY
    This.ArrayType = BA_UNALLOCATED
    Set ResetToDefault = Me
End Function

'@Description("Clears all entries in the current array but retains existing capacity.")
Public Function Clear() As BetterArray
Attribute Clear.VB_Description = "Clears all entries in the current array but retains existing capacity."
    Dim OldCapacity As Long
    OldCapacity = This.Capacity
    ReDim This.items(This.LowerBound To OldCapacity + This.LowerBound)
    This.ArrayType = BA_UNALLOCATED
    This.Length = 0
    Me.Capacity = OldCapacity
    Set Clear = Me
End Function

'@Description("Joins one or more arrays onto the end of the current array.")
Public Function Concat(ParamArray args() As Variant) As BetterArray
Attribute Concat.VB_Description = "Joins one or more arrays onto the end of the current array."
    Dim i As Long
    Dim result() As Variant
    Dim Stored() As Variant
    Dim StoredType As ArrayTypes
    Dim StoredCapacity As Long
    StoredType = This.ArrayType
    StoredCapacity = This.Capacity
    If StoredType <> BA_JAGGED And StoredType <> BA_MULTIDIMENSION Then
        For i = LBound(args) To UBound(args)
            If IsArray(args(i)) Then
                If IsMultidimensionalArray(args(i)) Then
                    StoredType = BA_MULTIDIMENSION
                    Exit For
                End If
            End If
        Next
    End If
    If This.ArrayType <> BA_UNALLOCATED Then
        Stored = InternalItems
    End If
    result = ConcatDelegate(Stored, args)
    Me.items = result
    If StoredType = BA_MULTIDIMENSION Then
        This.ArrayType = StoredType
    End If
    If This.Capacity < StoredCapacity Then
        Me.Capacity = StoredCapacity
    End If
    Set Concat = Me
End Function

'@Description("Writes all of the elements in a Collection objet to the internal array.")
Public Function CopyFromCollection(ByVal SourceCollection As Collection) As BetterArray
Attribute CopyFromCollection.VB_Description = "Writes all of the elements in a Collection objet to the internal array."
    If SourceCollection Is Nothing Then
        RaiseError EC_EXPECTED_COLLECTION_OBJECT, "CopyFromCollection", "SourceCollection"
    End If
    Dim i As Long
    Dim NewItems() As Variant
    This.Length = SourceCollection.count
    If This.Length = 0 Then
        NewItems = GetEmptyArray
    Else
        ReDim NewItems(This.LowerBound To (This.Length - This.LowerBound - 1))
        For i = 1 To This.Length
            NewItems(i + This.LowerBound - 1) = SourceCollection.item(i)
        Next
    End If
    Me.items = NewItems
    Set CopyFromCollection = Me
End Function

'@Description("Sorts the stored array in ascending order.")
Public Function Sort(Optional ByVal SortColumn As Long = MISSING_LONG) As BetterArray
Attribute Sort.VB_Description = "Sorts the stored array in ascending order."
    Dim LocalItems() As Variant
    Dim sortedItems() As Variant
    Dim LocalArrayType As ArrayTypes
    Dim LocalSortColumn As Long
    Dim FirstChildLowerBound As Long
    
    Set Sort = Me
    
    If Me.IncludesType("Object") Then
        RaiseError EC_CANNOT_SORT_OBJECTS, "Sort"
        Exit Function
    End If
    
    LocalItems = InternalItems
    LocalArrayType = GetArrayType(LocalItems)
    If LocalArrayType = ArrayTypes.BA_UNALLOCATED Or _
       LocalArrayType = ArrayTypes.BA_UNDEFINED Then
        sortedItems = GetEmptyArray
    Else
        If This.Length > 0 Then
            If LocalArrayType <> BA_ONEDIMENSION Then
                If LocalArrayType = ArrayTypes.BA_MULTIDIMENSION Then
                    LocalItems = MultiToJagged(LocalItems)
                End If
                Dim Depth As Long
                Depth = GetJaggedArrayDepth(LocalItems)
                If Depth > 2 Then
                    RaiseError EC_EXCEEDS_MAX_SORT_DEPTH, "Sort"
                End If
                FirstChildLowerBound = LBound(LocalItems(LBound(LocalItems)))
                If SortColumn = MISSING_LONG Then
                    LocalSortColumn = FirstChildLowerBound
                Else
                    LocalSortColumn = SortColumn - 1 + FirstChildLowerBound
                End If
            End If
            ApplySortMethod _
                LocalItems, _
                LocalArrayType, _
                LocalSortColumn
        End If
        sortedItems = LocalItems
    End If
    Me.items = sortedItems
End Function

'@Description("Copies a section of the stored array to another location in the same array.")
Public Function CopyWithin( _
        ByVal Target As Long, _
        Optional ByVal startIndex As Long = MISSING_LONG, _
        Optional ByVal endIndex As Long = MISSING_LONG _
    ) As BetterArray
Attribute CopyWithin.VB_Description = "Copies a section of the stored array to another location in the same array."
    Dim LocalLength As Long
    Dim RelativeTarget As Long
    Dim RelativeStart As Long
    Dim RelativeEnd As Long
    Dim toIndex As Long
    Dim fromIndex As Long
    Dim Final As Long
    Dim count As Long
    Dim Direction As Long
    Dim LocalItems() As Variant
    
    Select Case This.ArrayType
    Case ArrayTypes.BA_UNDEFINED
        RaiseError EC_EXPECTED_ARRAY, "CopyWithin"
    Case ArrayTypes.BA_UNALLOCATED
        RaiseError EC_UNALLOCATED_ARRAY, "CopyWithin"
    Case Else
        LocalItems = InternalItems
        LocalLength = This.Length
        RelativeTarget = Target
        
        If RelativeTarget < 0 Then
            toIndex = max((LocalLength + RelativeTarget), 0)
        Else
            toIndex = Min(RelativeTarget, LocalLength)
        End If
        
        RelativeStart = IIf(startIndex = MISSING_LONG, LBound(LocalItems), startIndex)
        
        If RelativeStart < 0 Then
            fromIndex = max((LocalLength + RelativeStart), 0)
        Else
            fromIndex = Min(RelativeStart, LocalLength)
        End If
        
        If endIndex = MISSING_LONG Then
            RelativeEnd = LocalLength
        Else
            RelativeEnd = endIndex
        End If
        
        If RelativeEnd < 0 Then
            Final = max((LocalLength + RelativeEnd), 0)
        Else
            Final = Min(RelativeEnd, LocalLength)
        End If
        
        count = Min(Final - fromIndex, LocalLength - toIndex)
        
        If fromIndex < toIndex And toIndex < fromIndex + count Then
            Direction = -1
            inc fromIndex, count - 1
            inc toIndex, count - 1
        Else
            Direction = 1
        End If
        
        Do While count > 0
            If fromIndex >= LBound(LocalItems) And fromIndex <= UBound(LocalItems) Then
                LocalItems(toIndex) = LocalItems(fromIndex)
            End If
            inc fromIndex, Direction
            inc toIndex, Direction
            dec count
        Loop
        
        Me.items = LocalItems
    End Select
    
    Set CopyWithin = Me

End Function

'@Ignore ShadowedDeclaration
'@Description("Filters the stored array to include only values that meet the specified criteria.")
Public Function Filter( _
        ByVal Match As Variant, _
        Optional ByVal Include As Boolean, _
        Optional ByVal Recurse As Boolean, _
        Optional ByVal ColumnIndex As Long = MISSING_LONG _
        ) As BetterArray
    Dim result() As Variant
    Dim LocalItems() As Variant
    Dim LocalType As ArrayTypes
    LocalType = This.ArrayType
    Select Case LocalType
    Case ArrayTypes.BA_UNDEFINED
        RaiseError EC_EXPECTED_ARRAY, "Filter"
    Case ArrayTypes.BA_UNALLOCATED
        result = GetEmptyArray
    Case ArrayTypes.BA_ONEDIMENSION
        LocalItems = InternalItems
        result = FilterOneDimension( _
            LocalItems, _
            Match, _
            Include _
        )
    Case Else
        LocalItems = InternalItems
        If ColumnIndex <> MISSING_LONG Then
            result = FilterByColumn( _
                LocalItems, _
                Match, _
                Include, _
                ColumnIndex _
            )
        Else
            result = RecursiveFilter( _
                LocalItems, _
                Match, _
                Include, _
                Recurse _
            )

        End If
    End Select
    Me.items = result
    If LocalType = BA_MULTIDIMENSION Then This.ArrayType = LocalType
    Set Filter = Me
End Function


'@Description("Filters the stored array to include only values of the specified type.")
Public Function FilterType( _
        ByVal SearchTypeName As String, _
        Optional ByVal Include As Boolean, _
        Optional ByVal Recurse As Boolean, _
        Optional ByVal ColumnIndex As Long = MISSING_LONG _
        ) As BetterArray
Attribute FilterType.VB_Description = "Filters the stored array to include only values of the specified type."
    Dim result() As Variant
    Dim LocalItems() As Variant
    Dim LocalType As ArrayTypes
    LocalType = This.ArrayType
    
    Select Case LocalType
    Case ArrayTypes.BA_UNDEFINED
        RaiseError EC_EXPECTED_ARRAY, "FilterType"
    Case ArrayTypes.BA_UNALLOCATED
        result = GetEmptyArray
    Case ArrayTypes.BA_ONEDIMENSION
        LocalItems = InternalItems
        result = FilterOneDimension( _
            LocalItems, _
            SearchTypeName, _
            Include, _
            True _
        )
    Case Else
        LocalItems = InternalItems
        If ColumnIndex <> MISSING_LONG Then
            result = FilterByColumn( _
                LocalItems, _
                SearchTypeName, _
                Include, _
                ColumnIndex, _
                True _
            )
        Else
            result = RecursiveFilter( _
                LocalItems, _
                SearchTypeName, _
                Include, _
                Recurse, _
                True _
            )

        End If
    End Select
      
    Me.items = result
    If LocalType = BA_MULTIDIMENSION Then This.ArrayType = LocalType
    Set FilterType = Me
End Function

'@Description("Reverses the order of elements in the array. The first array element becomes the last, and the last array element becomes the first.")
Public Function Reverse(Optional ByVal Recurse As Boolean) As BetterArray
Attribute Reverse.VB_Description = "Reverses the order of elements in the array. The first array element becomes the last, and the last array element becomes the first."
    Dim LocalItems() As Variant
    Dim result() As Variant
    Dim CurrentType As ArrayTypes
    CurrentType = This.ArrayType
    Select Case CurrentType
    Case BA_UNDEFINED
        RaiseError EC_UNDEFINED_ARRAY, "Reverse"
    Case BA_ONEDIMENSION, BA_MULTIDIMENSION, BA_JAGGED
        LocalItems = InternalItems
        result = RecursiveReverse(LocalItems, Recurse)
        Me.items = result
        This.ArrayType = CurrentType
    End Select
    
    Set Reverse = Me
End Function

'@Description("Flattens a Multi-Dimensional or Jagged array into One Dimension arrays.")
Public Function Flatten() As BetterArray
Attribute Flatten.VB_Description = "Flattens a Multi-Dimensional or Jagged array into One Dimension arrays."
    Dim LocalItems() As Variant
    Dim result As BetterArray
        
    ' noop on 1d or unalloc
    ' md & jagged both stored as jagged
    Select Case This.ArrayType
    Case BA_UNDEFINED
        RaiseError EC_UNDEFINED_ARRAY, "Flatten"
    Case BA_MULTIDIMENSION, BA_JAGGED
        LocalItems = InternalItems
        Set result = New BetterArray
        result.LowerBound = This.LowerBound
        RecursiveFlatten LocalItems, result
        Me.items = result.items
    End Select

    Set Flatten = Me
End Function

'@Description("shuffles the order of the stored array.")
Public Function Shuffle(Optional ByVal Recurse As Boolean) As BetterArray
Attribute Shuffle.VB_Description = "shuffles the order of the stored array."
    Dim LocalItems() As Variant
    Dim result() As Variant
    Dim CurrentType As ArrayTypes
    CurrentType = This.ArrayType
    Select Case CurrentType
    Case BA_UNDEFINED
        RaiseError EC_UNDEFINED_ARRAY, "Shuffle"
    Case BA_ONEDIMENSION, BA_MULTIDIMENSION, BA_JAGGED
        LocalItems = InternalItems
        result = RecursiveShuffle(LocalItems, Recurse)
        Me.items = result
        This.ArrayType = CurrentType
    End Select
    
    Set Shuffle = Me
End Function

'@Description("filters the array to contain only unique elements")
Public Function Unique(Optional ByVal ColumnIndex As Long = MISSING_LONG) As BetterArray
Attribute Unique.VB_Description = "filters the array to contain only unique elements"
    Dim LocalItems() As Variant
    Dim LocalType As ArrayTypes
    Dim result As BetterArray
    Dim i As Long
    
    LocalItems = InternalItems
    LocalType = This.ArrayType
    Set result = New BetterArray
    result.LowerBound = Me.LowerBound
    If (LocalType = BA_JAGGED Or LocalType = BA_MULTIDIMENSION) And ColumnIndex <> MISSING_LONG Then
        Dim LocalDepth As Long
        LocalDepth = GetJaggedArrayDepth(LocalItems)
        If LocalDepth = 2 Then
            ' if array is 2d we can compare unique based on specified column
            Dim LocalColumn As Long
            Dim ComparisonValues() As Variant
            Dim Comparator As BetterArray
            Dim Bounds() As Long
            
            Set Comparator = New BetterArray
            Bounds = GetMaxBoundsAtDimension(LocalItems, 2)
            Comparator.LowerBound = Bounds(0)
            LocalColumn = ColumnIndex - 1 + Bounds(0)
            ' determine column to compare - columnIndex if valid, lbound if not
            If LocalColumn < Bounds(0) Or LocalColumn > Bounds(1) Then
                LocalColumn = Bounds(0)
            End If
            
            ' extract the column
            ComparisonValues = Me.ExtractSegment(ColumnIndex:=LocalColumn)
            ' store the extracted column in a new BetterArray instance
            Comparator.items = ComparisonValues
            
            For i = LBound(LocalItems) To UBound(LocalItems)
                If Comparator.IndexOf(LocalItems(i)(LocalColumn)) = i Then
                    result.Push LocalItems(i)
                End If
            Next
            
        Else
            ' array has more then two dimensions
            ' just return unique in outermost array
            For i = LBound(LocalItems) To UBound(LocalItems)
                If Me.IndexOf(LocalItems(i)) = i Then
                    result.Push LocalItems(i)
                End If
            Next
        
        End If
    Else
        ' no columnIndex arg or array is one-dimension
        ' unique in outermost array
        For i = LBound(LocalItems) To UBound(LocalItems)
            If Me.IndexOf(LocalItems(i)) = i Then
                result.Push LocalItems(i)
            End If
        Next
    End If
    

    Me.items = result.items
    This.ArrayType = LocalType
    Set Unique = Me
End Function

'@Description("Accepts a path argument pointing to a comma-separated values (CSV) file and stores the delimited values contained within to the internal array.")
Public Function FromCSVFile( _
    ByVal path As String, _
    Optional ByVal columnDelimiter As String = CHR_COMMA, _
    Optional ByVal RowDelimiter As String, _
    Optional ByVal Quote As String = CHR_QUOTE, _
    Optional ByVal IgnoreFirstRow As Boolean, _
    Optional ByVal DuckType As Boolean _
) As BetterArray
Attribute FromCSVFile.VB_Description = "Accepts a path argument pointing to a comma-separated values (CSV) file and stores the delimited values contained within to the internal array."
    Dim RawData As String
    Dim LineEnding As String
    
    RawData = ReadStringFromFile(path)
    If RowDelimiter = vbNullString Then
        LineEnding = DetectLineEndings(RawData)
    Else
        LineEnding = RowDelimiter
    End If
    
    Set FromCSVFile = FromCSVString( _
        CSVString:=RawData, _
        columnDelimiter:=columnDelimiter, _
        RowDelimiter:=LineEnding, _
        Quote:=Quote, _
        IgnoreFirstRow:=IgnoreFirstRow, _
        DuckType:=DuckType _
    )
End Function

'@Description("Accepts a string argument representing the contents of a comma-separated values (CSV) file and stores the delimited values contained within to the internal array.")
Public Function FromCSVString( _
    ByVal CSVString As String, _
    Optional ByVal columnDelimiter As String = CHR_COMMA, _
    Optional ByVal RowDelimiter As String, _
    Optional ByVal Quote As String = CHR_QUOTE, _
    Optional ByVal IgnoreFirstRow As Boolean, _
    Optional ByVal DuckType As Boolean _
) As BetterArray
Attribute FromCSVString.VB_Description = "Accepts a string argument representing the contents of a comma-separated values (CSV) file and stores the delimited values contained within to the internal array."
    Dim LocalType As ArrayTypes
    Dim Parsed() As Variant
    Dim LineEnding As String
    
    If RowDelimiter = vbNullString Then
        LineEnding = DetectLineEndings(CSVString)
    Else
        LineEnding = RowDelimiter
    End If
    
    If This.ArrayType = BA_MULTIDIMENSION Then
        LocalType = BA_MULTIDIMENSION
    Else
        LocalType = BA_JAGGED
    End If
     
    Parsed = ParseCSV( _
        expression:=CSVString, _
        columnDelimiter:=columnDelimiter, _
        RowDelimiter:=LineEnding, _
        Quote:=Quote, _
        IgnoreFirstRow:=IgnoreFirstRow, _
        ReturnJagged:=IIf(LocalType = BA_JAGGED, True, False), _
        Base:=This.LowerBound, _
        DuckType:=DuckType _
    )

    This.Length = GetArrayLength(Parsed)
    This.Capacity = UBound(Parsed)
    This.ArrayType = LocalType
    This.items = Parsed

    Set FromCSVString = Me
End Function

'@Description("Returns a string representation of the stored array to use as output for a CSV file")
Public Function ToCSVString( _
    Optional ByRef Headers As Variant, _
    Optional ByVal columnDelimiter As String = CHR_COMMA, _
    Optional ByVal RowDelimiter As String = vbCrLf, _
    Optional ByVal Quote As String = CHR_QUOTE, _
    Optional ByVal EncloseAllInQuotes As Boolean, _
    Optional ByVal DateFormat As String, _
    Optional ByVal NumberFormat As String _
) As String
Attribute ToCSVString.VB_Description = "Returns a string representation of the stored array to use as output for a CSV file"
    Dim LocalItems() As Variant
    Dim Depth As Long
    Dim EncodedRecords() As String
    Dim result As String
    Dim HeadersArray() As Variant
    
    LocalItems = InternalItems
    Depth = GetJaggedArrayDepth(LocalItems)
    If Depth > 0 Then 'check array is allocated
        If Depth = 1 Then
            '1d array
            LocalItems = ConvertOneDimensionArrayToJagged(LocalItems)
        End If
        If Not IsMissing(Headers) Then
            If IsArray(Headers) Then
                HeadersArray = Array(Headers)
                LocalItems = InternalConcat(HeadersArray, LocalItems)
            End If
        End If
        EncodedRecords = EncodeCSVRecords(LocalItems, columnDelimiter, RowDelimiter, EncloseAllInQuotes, DateFormat, NumberFormat)
        result = BuildCSVString(EncodedRecords, columnDelimiter, RowDelimiter)
    End If
    ToCSVString = result
End Function

'@Description("Writes stored array to CSV file")
Public Function ToCSVFile( _
    ByVal path As String, _
    Optional ByRef Headers As Variant, _
    Optional ByVal columnDelimiter As String = CHR_COMMA, _
    Optional ByVal RowDelimiter As String = vbCrLf, _
    Optional ByVal Quote As String = CHR_QUOTE, _
    Optional ByVal EncloseAllInQuotes As Boolean, _
    Optional ByVal DateFormat As String, _
    Optional ByVal NumberFormat As String _
) As String
Attribute ToCSVFile.VB_Description = "Writes stored array to CSV file"
    Dim Content As String
    ' test if path valid - if not handle err
    
    ' get ouput string
    If IsMissing(Headers) Then
        Content = Me.ToCSVString( _
            columnDelimiter:=columnDelimiter, _
            RowDelimiter:=RowDelimiter, _
            Quote:=Quote, _
            EncloseAllInQuotes:=EncloseAllInQuotes, _
            DateFormat:=DateFormat, _
            NumberFormat:=NumberFormat _
        )
    Else
        Content = Me.ToCSVString( _
            Headers, _
            columnDelimiter, _
            RowDelimiter, _
            Quote, _
            EncloseAllInQuotes, _
            DateFormat, _
            NumberFormat _
        )
    End If
    
    ' write output
    PrintStringToFile path, Content
    
    ToCSVFile = Content
End Function


'''''''''''''''''''
' Private Methods '

'''''''''''''''''''
'@Description("Tries to guess line endings used in a string")
Private Function DetectLineEndings(ByVal Source As String) As String
Attribute DetectLineEndings.VB_Description = "Tries to guess line endings used in a string"
    If VBA.Strings.InStr(Source, vbCrLf) Then
        DetectLineEndings = vbCrLf
        Exit Function
    End If
    If VBA.Strings.InStr(Source, vbLf) Then
        DetectLineEndings = vbLf
        Exit Function
    End If
    If VBA.Strings.InStr(Source, vbCr) Then
        DetectLineEndings = vbCr
        Exit Function
    End If
    ' default to CRLF
    DetectLineEndings = vbCrLf
End Function

'@Description("Reads a text file to a string from a specified path")
Private Function ReadStringFromFile(ByVal path As String) As String
Attribute ReadStringFromFile.VB_Description = "Reads a text file to a string from a specified path"
    Dim File As Long
    Dim byteCount As Long
    Dim result As String
    If Dir(path) <> vbNullString Then
        File = FreeFile
        Open path For Input Access Read Lock Read As File
            byteCount = LOF(File)
            result = Input$(byteCount, File)
        Close File
    End If
    ReadStringFromFile = result
End Function

'@Description("Prints as string to a text file at the specified path")
Private Sub PrintStringToFile(ByVal path As String, ByVal Content As String)
Attribute PrintStringToFile.VB_Description = "Prints as string to a text file at the specified path"
    Dim File As Long
    File = FreeFile
    Open path For Output Access Write As File
        Print #File, Content
    Close File
End Sub

'@Description("Surrounds the provided string with quote characters")
Private Function WrapQuote(Optional ByVal Source As String = vbNullString) As String
Attribute WrapQuote.VB_Description = "Surrounds the provided string with quote characters"
    WrapQuote = CHR_QUOTE & Source & CHR_QUOTE
End Function

'@Description("Concatenates all records into a single string to be output as a csv")
Private Function BuildCSVString( _
    ByRef Records() As String, _
    ByVal columnDelimiter As String, _
    ByVal RowDelimiter As String _
) As String
Attribute BuildCSVString.VB_Description = "Concatenates all records into a single string to be output as a csv"
    Dim i As Long
    Dim j As Long
    Dim LastRow As Long
    Dim LastCol As Long
    Dim ValidRow As Boolean
    
    LastRow = UBound(Records)
    LastCol = UBound(Records, 2)
    StringBuilder NewString:=True
    For i = LBound(Records) To UBound(Records)
        'skip entirely blank rows
        ValidRow = False
        For j = LBound(Records, 2) To UBound(Records, 2)
            If Records(i, j) <> vbNullString Then
                ValidRow = True
                Exit For
            End If
        Next
        If ValidRow Then
            For j = LBound(Records, 2) To UBound(Records, 2)
                StringBuilder Records(i, j)
                If j = LastCol Then
                    If i <> LastRow Then
                        StringBuilder RowDelimiter
                    End If
                Else
                    StringBuilder columnDelimiter
                End If
            Next
        End If
    Next
    BuildCSVString = StringBuilder(Final:=True)
End Function

'@Description("Applies RFC 4180 business rules to all fields")
Private Function EncodeCSVRecords( _
    ByRef Records() As Variant, _
    ByVal columnDelimiter As String, _
    ByVal RowDelimiter As String, _
    ByVal EncloseAllInQuotes As Boolean, _
    ByVal DateFormat As String, _
    ByVal NumberFormat As String _
) As String()
Attribute EncodeCSVRecords.VB_Description = "Applies RFC 4180 business rules to all fields"
    Dim FirstDimBounds() As Long
    Dim SecondDimBounds() As Long
    Dim i As Long
    Dim j As Long
    Dim result() As String
    Dim CurrentField As Variant
    
    FirstDimBounds = GetArrayBounds(Records)
    SecondDimBounds = GetMaxBoundsAtDimension(Records, 2)
    ReDim result( _
        FirstDimBounds(0) To FirstDimBounds(1), _
        SecondDimBounds(0) To SecondDimBounds(1) _
    )
    For i = FirstDimBounds(0) To FirstDimBounds(1)
        If IsArray(Records(i)) Then
            For j = LBound(Records(i)) To UBound(Records(i))
                CurrentField = GetScalarRepresentation(Records(i)(j))
                result(i, j) = EncodeCSVField( _
                    CurrentField, _
                    columnDelimiter, _
                    RowDelimiter, _
                    EncloseAllInQuotes, _
                    DateFormat, _
                    NumberFormat _
                )
            Next
        End If
    Next
    EncodeCSVRecords = result
End Function

'@Description("Applies RFC 4180 business rules to the encoding of a CSV field")
Private Function EncodeCSVField( _
    ByVal Field As Variant, _
    ByVal columnDelimiter As String, _
    ByVal RowDelimiter As String, _
    ByVal EncloseAllInQuotes As Boolean, _
    ByVal DateFormat As String, _
    ByVal NumberFormat As String _
) As String
Attribute EncodeCSVField.VB_Description = "Applies RFC 4180 business rules to the encoding of a CSV field"
    ' https://tools.ietf.org/html/rfc4180#section-2
    Dim LocalField As String
    
    If IsDate(Field) And DateFormat <> vbNullString Then
        On Error Resume Next
        LocalField = format$(Field, DateFormat)
        On Error GoTo 0
    ElseIf IsNumeric(Field) And NumberFormat <> vbNullString Then
        On Error Resume Next
        LocalField = format$(Field, NumberFormat)
        On Error GoTo 0
    End If
    If LocalField = vbNullString Then
        LocalField = CStr(Field)
    End If
    LocalField = EscapeCharInString(LocalField, CHR_QUOTE, CHR_QUOTE)
    If EncloseAllInQuotes Or _
        VBA.Strings.InStr(LocalField, RowDelimiter) Or _
        VBA.Strings.InStr(LocalField, columnDelimiter) Then
        LocalField = WrapQuote(LocalField)
    End If
    
    EncodeCSVField = LocalField
    Exit Function
End Function

'@Description("Escapes a substring in a string using a provided Escape substring")
Private Function EscapeCharInString( _
    ByVal Destination As String, _
    ByVal Target As String, _
    ByVal Escape As String, _
    Optional ByVal BothSides As Boolean _
) As String
Attribute EscapeCharInString.VB_Description = "Escapes a substring in a string using a provided Escape substring"
    Dim index As Long
    Dim TargetLength As Long
    Dim EscapeLength As String
    Dim result As String
    
    TargetLength = Len(Target)
    EscapeLength = Len(Escape)
    result = Destination
    index = 1
    Do
        index = VBA.Strings.InStr(index, Destination, Target)
        If index Then
            result = InsertIntoStringAtIndex(result, Escape, index)
            inc index, EscapeLength + TargetLength
            If BothSides Then
                result = InsertIntoStringAtIndex(result, Escape, index)
            End If
        End If
    Loop While index > 0
    
    EscapeCharInString = result
End Function

'@Description("Inserts a substring into a destination string starting at the provided index")
Private Function InsertIntoStringAtIndex( _
    ByVal Destination As String, _
    ByVal Source As String, _
    ByVal index As Long _
) As String
Attribute InsertIntoStringAtIndex.VB_Description = "Inserts a substring into a destination string starting at the provided index"
    Dim Front As String
    Dim Back As String
    If index > 1 Then
        Front = Left$(Destination, index)
    End If
    
    If index < Len(Destination) Then
        Back = Right$(Destination, Len(Destination) - index)
    End If
    InsertIntoStringAtIndex = Front & Source & Back
End Function

'@Description("Utility function which delegates to ConcatDelegate")
Private Function InternalConcat(ParamArray args() As Variant) As Variant()
Attribute InternalConcat.VB_Description = "Utility function which delegates to ConcatDelegate"
    Dim First() As Variant
    First = Array()
    InternalConcat = ConcatDelegate(First, args)
End Function

'@Description("Concatenates Passed Arrays. Do not use as utility function")
Private Function ConcatDelegate(ByRef First() As Variant, ParamArray args() As Variant) As Variant()
Attribute ConcatDelegate.VB_Description = "Concatenates Passed Arrays. Do not use as utility function"
    ' NOTE don't call this internally as a utility. use InternalConcat
    Dim i As Long
    Dim Cursor As Long
    Dim NewLength As Long
    Dim Rest() As Variant
    Dim result() As Variant
    Dim Current() As Variant
    
    Rest = args(0)
    Cursor = This.LowerBound
    If IsArrayAllocated(First) Then
        Cursor = LBound(First)
    ElseIf IsArray(Rest(LBound(Rest))) Then
        If IsArrayAllocated(Rest(LBound(Rest))) Then
            Cursor = LBound(Rest(LBound(Rest)))
        End If
    End If
    
    NewLength = GetArrayLength(First) + GetTotalLengthOfNestedArrays(Rest)
    ReDim result(Cursor To max(Cursor + NewLength - 1, Cursor))
    
    InsertArrayAtIndex result, First, Cursor
    For i = LBound(Rest) To UBound(Rest)
        If IsArray(Rest(i)) Then
            Current = Rest(i)
            If IsMultidimensionalArray(Current) Then
                Current = MultiToJagged(Current)
            End If
            InsertArrayAtIndex result, Current, Cursor
        Else
            LetOrSetElement result(Cursor), Rest(i)
            inc Cursor
        End If
    Next
    ConcatDelegate = result
End Function

'@Description("Places the Values in Source within Destination starting at Index")
Private Sub InsertArrayAtIndex(ByRef Destination() As Variant, ByRef Source() As Variant, ByRef index As Long)
Attribute InsertArrayAtIndex.VB_Description = "Places the Values in Source within Destination starting at Index"
    Dim Offset As Long
    Dim SourceLength As Long
    Dim StartingIndex As Long
    If IsArrayAllocated(Source) Then
        StartingIndex = index
        Offset = StartingIndex - LBound(Source)
        SourceLength = GetArrayLength(Source)
        Do While index < StartingIndex + SourceLength
            LetOrSetElement Destination(index), Source(index - Offset)
            inc index
        Loop
    End If
End Sub

'@Description("Returns the combined length of all nested arrays")
Private Function GetTotalLengthOfNestedArrays(ByRef Source() As Variant) As Long
Attribute GetTotalLengthOfNestedArrays.VB_Description = "Returns the combined length of all nested arrays"
    Dim i As Long
    Dim result As Long
    Dim Current() As Variant
    For i = LBound(Source) To UBound(Source)
        If IsArray(Source(i)) Then
            Current = Source(i)
            inc result, GetArrayLength(Current)
        End If
    Next
    GetTotalLengthOfNestedArrays = result
End Function

'@Description("Truncates the number of rows in an array to the given amount")
Private Function TrimColumnsMultidimensionArray(ByRef Original() As Variant, ByVal AvailableColumns As Long) As Variant()
Attribute TrimColumnsMultidimensionArray.VB_Description = "Truncates the number of rows in an array to the given amount"
    Dim i As Long
    Dim j As Long
    Dim result() As Variant
    Dim OuterBounds(0 To 1) As Long
    Dim InnerBounds(0 To 1) As Long
    Dim CountColumns As Long
    
    CountColumns = UBound(Original, 2) - LBound(Original, 2) + 1
    If CountColumns > AvailableColumns Then
        OuterBounds(0) = LBound(Original, 1)
        OuterBounds(1) = UBound(Original, 1)
        InnerBounds(0) = LBound(Original, 2)
        InnerBounds(1) = InnerBounds(0) + AvailableColumns - 1
        ReDim result(OuterBounds(0) To OuterBounds(1), InnerBounds(0) To InnerBounds(1))
        For i = OuterBounds(0) To OuterBounds(1)
            For j = InnerBounds(0) To InnerBounds(1)
                LetOrSetElement result(i, j), Original(i, j)
            Next
        Next
    Else
        result = Original
    End If
    TrimColumnsMultidimensionArray = result
End Function

'@Description("Truncates the number of rows in an array to the given amount")
Private Function TrimRowsMultidimensionArray(ByRef Original() As Variant, ByVal AvailableRows As Long) As Variant()
Attribute TrimRowsMultidimensionArray.VB_Description = "Truncates the number of rows in an array to the given amount"
    Dim i As Long
    Dim j As Long
    Dim result() As Variant
    Dim OuterBounds(0 To 1) As Long
    Dim InnerBounds(0 To 1) As Long
    Dim CountRows As Long
    
    CountRows = UBound(Original, 1) - LBound(Original, 1) + 1
    If CountRows > AvailableRows Then
        OuterBounds(0) = LBound(Original, 1)
        OuterBounds(1) = OuterBounds(0) + AvailableRows - 1
        InnerBounds(0) = LBound(Original, 2)
        InnerBounds(1) = UBound(Original, 2)
        ReDim result(OuterBounds(0) To OuterBounds(1), InnerBounds(0) To InnerBounds(1))
        For i = OuterBounds(0) To OuterBounds(1)
            For j = InnerBounds(0) To InnerBounds(1)
                LetOrSetElement result(i, j), Original(i, j)
            Next
        Next
    Else
        result = Original
    End If
    TrimRowsMultidimensionArray = result
End Function

'@Description("Increments an index")
Private Function inc(ByRef index As Long, Optional ByVal value As Long = 1) As Long
Attribute inc.VB_Description = "Increments an index"
    index = index + value
    inc = index
End Function

'@Description("Decrements an index")
Private Function dec(ByRef index As Long, Optional ByVal value As Long = 1) As Long
Attribute dec.VB_Description = "Decrements an index"
    index = index - value
    dec = index
End Function

'@Description("Populates TString Types")
Private Function StringFactory(ByRef expression As String) As TString
Attribute StringFactory.VB_Description = "Populates TString Types"
    Dim result As TString
    result.Text = expression
    result.Length = Len(expression)
    result.ByteLength = LenB(expression)
    StringFactory = result
End Function

'@Description("Returns the next byte position in the string of the specified delimiter")
Private Function NextDelimBytePos( _
    ByRef expression As TString, _
    ByRef Delimiter As TString, _
    Optional ByRef startIndex As Long = 1 _
) As Long
Attribute NextDelimBytePos.VB_Description = "Returns the next byte position in the string of the specified delimiter"
    Dim result As Long
    result = InStrB(startIndex, expression.Text, Delimiter.Text, vbBinaryCompare)
    Do Until (result And 1) Or (result = 0)
        result = InStrB(result + 1, expression.Text, Delimiter.Text, vbBinaryCompare)
    Loop
    NextDelimBytePos = result
End Function

'@Description("Splits the provided string into rows based on the provided row delimiter")
Private Function GetCSVRows( _
    ByRef Expr As TString, _
    ByRef RowDelim As TString, _
    ByRef LiteralDelim As TString _
) As String()
Attribute GetCSVRows.VB_Description = "Splits the provided string into rows based on the provided row delimiter"
    Dim RowDelimIndex As Long
    Dim LastRowDelim As Long
    Dim RowIndex As Long
    Dim BlankTrailingRows As Long
    Dim MaxRows As Long
    Dim LineEndIndices() As Long
    Dim LiteralDelimIndex As Long
    Dim CountRows As Long
    Dim i As Long
    Dim Cursor As Long
    Dim CSVRows() As String

    ' find first row delimiter
    RowDelimIndex = NextDelimBytePos(Expr, RowDelim)

    ' find last row delimiter
    LastRowDelim = InStrRev(Expr.Text, RowDelim.Text)
    RowIndex = Expr.Length + 1
    Do While RowIndex - LastRowDelim = RowDelim.Length
        RowIndex = LastRowDelim
        LastRowDelim = InStrRev(Expr.Text, RowDelim.Text, LastRowDelim)
        inc BlankTrailingRows
    Loop
    RowIndex = Empty

    ' estimate max number of rows - does not account for line breaks within strings
    MaxRows = (Expr.Length - Len(Replace(Expr.Text, RowDelim.Text, vbNullString))) / RowDelim.Length + 1
    ReDim LineEndIndices(0 To MaxRows)
        
    ' find first quote delimiter
    LiteralDelimIndex = NextDelimBytePos(Expr, LiteralDelim)
    
    Do While RowDelimIndex > 0
        ' see if row delimiter before quote delimiter
        If RowDelimIndex + RowDelim.ByteLength <= LiteralDelimIndex Or LiteralDelimIndex = 0 Then
            ' no preceding literalDelim - iterate row count
            LineEndIndices(CountRows) = RowDelimIndex
            RowDelimIndex = NextDelimBytePos(Expr, RowDelim, RowDelimIndex + RowDelim.ByteLength)
            inc CountRows
        Else
            ' quote delimiter before row delimiter
            ' find next quote delim
            LiteralDelimIndex = NextDelimBytePos(Expr, LiteralDelim, LiteralDelimIndex + 2)
           ' see if closing quote delim found
            If LiteralDelimIndex Then
                ' find next row delim
                RowDelimIndex = NextDelimBytePos(Expr, RowDelim, LiteralDelimIndex + 2)
                ' if next row delim found
                If RowDelimIndex Then
                    ' find next quote delim
                    LiteralDelimIndex = NextDelimBytePos(Expr, LiteralDelim, LiteralDelimIndex + 2)
                End If
            End If
        End If
    Loop
    
    LineEndIndices(CountRows) = Expr.ByteLength + 1
    dec CountRows, BlankTrailingRows - 1
    ReDim CSVRows(0 To CountRows - 1)
    Cursor = 1
    For i = 0 To CountRows - 1
        CSVRows(i) = MidB$(Expr.Text, Cursor, LineEndIndices(i) - Cursor)
        Cursor = LineEndIndices(i) + RowDelim.ByteLength
    Next
    GetCSVRows = CSVRows
End Function

'@Description("Counts the number of delimited columns in a provided string")
Private Function CountCSVColumns(ByRef FirstRow As String, ByRef LiteralDelim As TString, ByVal columnDelimiter As String) As Long
Attribute CountCSVColumns.VB_Description = "Counts the number of delimited columns in a provided string"
    ' Each line should contain the same number of fields
    ' https://www.loc.gov/preservation/digital/formats/fdd/fdd000323.shtml
    ' so we can just count the columnd in the first row
    Dim i As Long
    Dim InQuote As Boolean
    Dim CountColumns As Long
    Dim Char As String
    CountColumns = 1
    For i = 0 To Len(FirstRow)
        Char = Mid$(FirstRow, i + 1, 1)
        Select Case Char
            Case LiteralDelim.Text
                InQuote = Not InQuote
            Case columnDelimiter
                If Not InQuote Then
                    inc CountColumns
                End If
        End Select
    Next
    CountCSVColumns = CountColumns
End Function

'@Description("Parses CSV string into a 2d or jagged array")
Private Function ParseCSV( _
    ByRef expression As String, _
    ByVal columnDelimiter As String, _
    ByVal RowDelimiter As String, _
    ByRef Quote As String, _
    ByVal IgnoreFirstRow As Boolean, _
    ByVal ReturnJagged As Boolean, _
    ByVal Base As Long, _
    ByVal DuckType As Boolean _
) As Variant()
Attribute ParseCSV.VB_Description = "Parses CSV string into a 2d or jagged array"
    ' CSV is a simple format for representing a rectangular array (matrix) of numeric and textual values.
    ' It an example of a "flat file" format.
    ' It is a delimited data format that has fields/columns separated by the comma character %x2C (Hex 2C)
    ' and records/rows/lines separated by characters indicating a line break.
    ' RFC 4180 stipulates the use of CRLF pairs to denote line breaks, where CR is %x0D (Hex 0D) and LF is %x0A (Hex 0A).
    ' Each line should contain the same number of fields.
    ' Fields that contain a special character (comma, CR, LF, or double quote), must be "escaped" by enclosing them in double quotes (Hex 22).
    ' An optional header line may appear as the first line of the file with the same format as normal record lines.
    ' This header will contain names corresponding to the fields in the file and should contain the same number of fields as the records in the rest of the file.
    ' CSV commonly employs US-ASCII as character set, but other character sets are permitted.
    ' https://www.loc.gov/preservation/digital/formats/fdd/fdd000323.shtml

    Dim i As Long
    Dim j As Long
    Dim Cursor As Long
    Dim ColumnIndex As Long
    Dim CountColumns As Long
    Dim CountRows As Long
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim StartRow As Long
    Dim InQuote As Boolean
    Dim Buffer As String
    Dim Char As String
    Dim Frag As String
    Dim CSVRows() As String
    Dim JaggedRow() As Variant
    Dim Results() As Variant
    Dim Element As Variant
    Dim Expr As TString
    Dim RowDelim As TString
    Dim LiteralDelim As TString
    
    Expr = StringFactory(expression)
    RowDelim = StringFactory(RowDelimiter)
    LiteralDelim = StringFactory(Quote)
    
    If Expr.ByteLength = 0 Or RowDelim.ByteLength = 0 Then
        ' TODO: Raise Err?
        ParseCSV = Array()
        Exit Function
    End If

    CSVRows = GetCSVRows(Expr, RowDelim, LiteralDelim)
    CountColumns = CountCSVColumns(CSVRows(0), LiteralDelim, columnDelimiter)
    CountRows = UBound(CSVRows) - LBound(CSVRows) + 1
    
    If IgnoreFirstRow Then
        StartRow = 1
    Else
        StartRow = 0
    End If
    
    LastRow = Base + CountRows - 1 - StartRow
    LastColumn = Base + CountColumns - 1
    If ReturnJagged Then
        ReDim JaggedRow(Base To LastColumn)
        ReDim Results(Base To LastRow)
    Else
        ReDim Results(Base To LastRow, Base To LastColumn)
    End If
    
    For i = StartRow To LastRow
        ColumnIndex = Base
        Buffer = CSVRows(i)
        Cursor = 1
        For j = 0 To Len(Buffer)
            Char = Mid$(Buffer, j + 1, 1)
            If Char = LiteralDelim.Text Then
                InQuote = Not InQuote
            ElseIf Char = columnDelimiter Or j = Len(Buffer) Then
                If Not InQuote Then
                    Frag = Mid$(Buffer, Cursor, j - Cursor + 1)
                    If DuckType Then
                        Element = DuckTypeElement(Frag)
                    Else
                        Element = UnquoteString(Frag)
                    End If
                    If ReturnJagged Then
                        JaggedRow(ColumnIndex) = Element
                    Else
                        Results(i - StartRow, ColumnIndex) = Element
                    End If
                    Cursor = j + 2
                    inc ColumnIndex
                    Frag = vbNullString
                End If
            End If
        Next
        If ReturnJagged Then
            Results(i - StartRow) = JaggedRow
            ' ReDim jaggedRow(Base To lastColumn)
        End If
    Next
    ParseCSV = Results
End Function

'@Description("Recursvely fills the array (and nested arrays) with the passed value.")
Private Function RecursiveFill( _
        ByRef SourceArray() As Variant, _
        ByVal value As Variant, _
        Optional ByVal startIndex As Long = MISSING_LONG, _
        Optional ByVal endIndex As Long = MISSING_LONG _
    ) As Variant()
Attribute RecursiveFill.VB_Description = "Recursvely fills the array (and nested arrays) with the passed value."
    Dim RelativeStart As Long
    Dim RelativeEnd As Long
    Dim i As Long
    
    If startIndex = MISSING_LONG Then
        RelativeStart = LBound(SourceArray)
    Else
        RelativeStart = startIndex
    End If
    If endIndex = MISSING_LONG Then
        RelativeEnd = UBound(SourceArray)
    Else
        RelativeEnd = endIndex
    End If
    For i = RelativeStart To RelativeEnd
        If IsArray(SourceArray(i)) Then
            Dim PassThru() As Variant
            PassThru = RecursiveFill(PassThru, value)
        Else
            LetOrSetElement SourceArray(i), value
        End If
    Next
    RecursiveFill = SourceArray
End Function

'@Description("Recursively checks if the passed array only includes the search element.")
Private Function RecursiveEvery( _
        ByVal SearchElement As Variant, _
        ByRef SearchArray() As Variant, _
        Optional ByVal fromIndex As Long, _
        Optional ByVal CompareTypeNames As Boolean _
    ) As Boolean
Attribute RecursiveEvery.VB_Description = "Recursively checks if the passed array only includes the search element."
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim i As Long
    
    LocalLowerBound = LBound(SearchArray)
    If fromIndex > LocalLowerBound Then LocalLowerBound = fromIndex
    LocalUpperBound = UBound(SearchArray)

    For i = LocalLowerBound To LocalUpperBound
        If IsArray(SearchArray(i)) Then
            Dim PassThru() As Variant
            PassThru = SearchArray(i)
            If Not RecursiveEvery(SearchElement, PassThru, CompareTypeNames:=CompareTypeNames) Then
                Exit Function
            End If
        ElseIf CompareTypeNames Then
            If Not InStr( _
                    UCase$(TypeName(SearchArray(i))), _
                    UCase$(CStr(SearchElement)) _
                ) > 0 Then
                Exit Function
            End If
        Else
            If Not ElementsAreEqual(SearchElement, SearchArray(i)) Then
                Exit Function
            End If
        End If
    Next
    
    RecursiveEvery = True
End Function

'@Description("shuffles the order of the passed array.")
Private Function RecursiveShuffle( _
        ByRef SourceArray() As Variant, _
        Optional ByVal Recurse As Boolean _
    ) As Variant()
Attribute RecursiveShuffle.VB_Description = "shuffles the order of the passed array."
    ' https://en.wikipedia.org/wiki/Fisher%E2%80%93Yates_shuffle#The_modern_algorithm
    Dim i As Long
    Dim j As Long
    Dim Lower As Long
    Dim Upper As Long
    Dim Nested() As Variant
    Lower = LBound(SourceArray)
    Upper = UBound(SourceArray)
    Randomize
    For i = Upper To Lower + 1 Step -1
        If IsArray(SourceArray(i)) And Recurse Then
            Nested = SourceArray(i)
            SourceArray(i) = RecursiveShuffle(Nested, Recurse)
        End If
        j = Int(Rnd * (i - Lower) + 1)
        If IsArray(SourceArray(j)) And Recurse Then
            Nested = SourceArray(j)
            SourceArray(j) = RecursiveShuffle(Nested, Recurse)
        End If
        Swap SourceArray, i, j
    Next
    RecursiveShuffle = SourceArray
End Function

'@Description("reverses the order of the passed array.")
Private Function RecursiveReverse( _
        ByRef SourceArray() As Variant, _
        Optional ByVal Recurse As Boolean _
    ) As Variant()
Attribute RecursiveReverse.VB_Description = "reverses the order of the passed array."
    
    Dim LocalLength As Long
    Dim Middle As Long
    Dim Lower As Long
    Dim Upper As Long
    Dim PassThruArray() As Variant
    
    LocalLength = GetArrayLength(SourceArray)
    Lower = LBound(SourceArray)
    Middle = Int(LocalLength / 2) + Lower
    
    Do While Lower <> Middle
        ' TODO: check this works with any lb
        Upper = UBound(SourceArray) + LBound(SourceArray) - Lower
        If IsArray(SourceArray(Lower)) And Recurse Then
            PassThruArray = SourceArray(Lower)
            SourceArray(Lower) = RecursiveReverse(PassThruArray, Recurse)
        End If
        If IsArray(SourceArray(Upper)) And Recurse Then
            PassThruArray = SourceArray(Upper)
            SourceArray(Upper) = RecursiveReverse(PassThruArray, Recurse)
        End If
        Swap SourceArray, Lower, Upper
        inc Lower
    Loop
    RecursiveReverse = SourceArray
End Function

'@Description("returns the length of the passed array.")
Private Function GetArrayLength(ByRef SourceArray() As Variant) As Long
Attribute GetArrayLength.VB_Description = "returns the length of the passed array."
    Dim result As Long
    result = 0
    If IsArrayAllocated(SourceArray) Then
        If Not IsArrayEmpty(SourceArray) Then
            result = UBound(SourceArray) - LBound(SourceArray) + 1
        End If
    End If
    GetArrayLength = result
End Function

'@Description("converts all the values in a jagged array to a 1d array.")
Private Sub RecursiveFlatten( _
        ByRef SourceArray() As Variant, _
        ByRef Results As BetterArray _
    )
Attribute RecursiveFlatten.VB_Description = "converts all the values in a jagged array to a 1d array."
    Dim Element As Variant
    For Each Element In SourceArray
        If IsArray(Element) Then
            Dim ArrayElement() As Variant
            ArrayElement = Element
            RecursiveFlatten ArrayElement, Results
        Else
            Results.Push Element
        End If
    Next
End Sub

'@Description("Recursively finds the largest value in an array.")
Private Function RecursiveMax(ByRef SourceArray() As Variant) As Variant
Attribute RecursiveMax.VB_Description = "Recursively finds the largest value in an array."
    Dim i As Long
    Dim result As Variant
    
    If IsMultidimensionalArray(SourceArray) Then
        RaiseError EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION, "Max", "Args"
    Else
        For i = LBound(SourceArray) To UBound(SourceArray)
            If IsArray(SourceArray(i)) Then
                Dim NestedArray() As Variant
                Dim NestedResult As Variant
                NestedArray = SourceArray(i)
                NestedResult = RecursiveMax(NestedArray)
                If IsEmpty(result) Then
                    result = NestedResult
                Else
                    If NestedResult > result Then result = NestedResult
                End If
            Else
                Dim CurrentValue As Variant
                CurrentValue = GetScalarRepresentation(SourceArray(i))
                If CurrentValue <> OBJECT_REPR Then
                    If IsEmpty(result) Then
                        result = CurrentValue
                    Else
                        If CurrentValue > result Then result = CurrentValue
                    End If
                End If
            End If
        Next
    End If
    RecursiveMax = result
End Function

'@Description("Recursively finds the smallest value in an array.")
Private Function RecursiveMin(ByRef SourceArray() As Variant) As Variant
Attribute RecursiveMin.VB_Description = "Recursively finds the smallest value in an array."
    Dim i As Long
    Dim result As Variant
    
    If IsMultidimensionalArray(SourceArray) Then
        RaiseError EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION, "Max", "Args"
    Else
        For i = LBound(SourceArray) To UBound(SourceArray)
            If IsArray(SourceArray(i)) Then
                Dim NestedArray() As Variant
                Dim NestedResult As Variant
                NestedArray = SourceArray(i)
                NestedResult = RecursiveMin(NestedArray)
                If IsEmpty(result) Then
                    result = NestedResult
                Else
                    If NestedResult < result Then result = NestedResult
                End If
            Else
                Dim CurrentValue As Variant
                CurrentValue = GetScalarRepresentation(SourceArray(i))
                If CurrentValue <> OBJECT_REPR Then
                    If IsEmpty(result) Then
                        result = CurrentValue
                    Else
                        If CurrentValue < result Then result = CurrentValue
                    End If
                End If
            End If
        Next
    End If
    RecursiveMin = result
End Function

'@Description("Recursively checks if the passed array includes the search element.")
Private Function RecursiveIncludes( _
        ByVal SearchElement As Variant, _
        ByRef SearchArray() As Variant, _
        Optional ByVal fromIndex As Long, _
        Optional ByVal CompareTypeNames As Boolean, _
        Optional ByVal Recurse As Boolean _
    ) As Boolean
Attribute RecursiveIncludes.VB_Description = "Recursively checks if the passed array includes the search element."
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim i As Long
    LocalLowerBound = LBound(SearchArray)
    If fromIndex > LocalLowerBound Then LocalLowerBound = fromIndex
    LocalUpperBound = UBound(SearchArray)
    For i = LocalLowerBound To LocalUpperBound
        If CompareTypeNames Then
            If InStr( _
                    UCase$(TypeName(SearchArray(i))), _
                    UCase$(CStr(SearchElement)) _
                ) > 0 _
                Or _
                UCase$(CStr(SearchElement)) = "OBJECT" And _
                IsObject(SearchArray(i)) _
                Then
                RecursiveIncludes = True
                Exit Function
            End If
        End If
        If IsArray(SearchArray(i)) And Recurse Then
            Dim PassThru() As Variant
            PassThru = SearchArray(i)
            If RecursiveIncludes( _
                    SearchElement, _
                    PassThru, _
                    CompareTypeNames:=CompareTypeNames, _
                    Recurse:=Recurse _
                ) Then
                RecursiveIncludes = True
                Exit Function
            End If
        ElseIf Not CompareTypeNames Then
            If ElementsAreEqual(SearchElement, SearchArray(i)) Then
                RecursiveIncludes = True
                Exit Function
            End If
        End If
    Next
    RecursiveIncludes = False
End Function

'@Description("Filters a one dimension array.")
Private Function FilterOneDimension( _
        ByRef SourceArray() As Variant, _
        ByVal Match As Variant, _
        ByVal Include As Boolean, _
        Optional ByVal CompareTypeNames As Boolean _
    ) As Variant()
    Dim result As BetterArray
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim i As Long
    Dim IsMatch As Boolean
    
    LocalLowerBound = LBound(SourceArray)
    LocalUpperBound = UBound(SourceArray)
    Set result = New BetterArray
    result.LowerBound = This.LowerBound
    For i = LocalLowerBound To LocalUpperBound
        If CompareTypeNames Then
            IsMatch = (InStr( _
                UCase$(TypeName(SourceArray(i))), _
                UCase$(CStr(Match)) _
            ) > 0)
        Else
            IsMatch = ElementsAreEqual(Match, SourceArray(i))
        End If
        If (Include And IsMatch) Or (Not Include And Not IsMatch) Then
            result.Push SourceArray(i)
        End If
    Next
    FilterOneDimension = result.items
End Function

'@Description("Filters a jaggged array by column.")
Private Function FilterByColumn( _
        ByRef SourceArray() As Variant, _
        ByVal Match As Variant, _
        ByVal Include As Boolean, _
        ByVal ColumnIndex As Long, _
        Optional ByVal CompareTypeNames As Boolean _
    ) As Variant()
    Dim result As BetterArray
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim i As Long
    Dim IsMatch As Boolean
    Dim ComparisonElement As Variant
    
    LocalLowerBound = LBound(SourceArray)
    LocalUpperBound = UBound(SourceArray)
    Set result = New BetterArray
    result.LowerBound = This.LowerBound

    For i = LocalLowerBound To LocalUpperBound
        LetOrSetElement ComparisonElement, SourceArray(i)(ColumnIndex)
        If CompareTypeNames Then
            IsMatch = (InStr( _
                UCase$(TypeName(ComparisonElement)), _
                UCase$(CStr(Match)) _
            ) > 0)
        Else
            IsMatch = ElementsAreEqual(Match, ComparisonElement)
        End If
        If (Include And IsMatch) Or (Not Include And Not IsMatch) Then
            result.Push SourceArray(i)
        End If
    Next
    FilterByColumn = result.items
End Function

'@Description("Recursively filters a passed array.")
Private Function RecursiveFilter( _
        ByRef SourceArray() As Variant, _
        ByVal Match As Variant, _
        ByVal Include As Boolean, _
        ByVal Recurse As Boolean, _
        Optional ByVal CompareTypeNames As Boolean _
    ) As Variant()
Attribute RecursiveFilter.VB_Description = "Recursively filters a passed array."
    Dim result As BetterArray
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim i As Long
    Dim IsMatch As Boolean
    
    LocalLowerBound = LBound(SourceArray)
    LocalUpperBound = UBound(SourceArray)
    Set result = New BetterArray
    result.LowerBound = This.LowerBound
    For i = LocalLowerBound To LocalUpperBound
        If IsArray(SourceArray(i)) And Recurse Then
            If IsArrayAllocated(SourceArray(i)) Then
                Dim LocalItems() As Variant
                LocalItems = SourceArray(i)
                LocalItems = RecursiveFilter( _
                    LocalItems, _
                    Match, _
                    Include, _
                    Recurse, _
                    CompareTypeNames _
                )
                ' don't push a returned empty array
                If Not RecursiveEvery(Empty, LocalItems) Then
                    result.Push LocalItems
                End If
            Else
                ' push an empty array if current is an unalocated rray
                If CompareTypeNames Then
                    IsMatch = (InStr( _
                        UCase$(TypeName(SourceArray(i))), _
                        UCase$(CStr(Match)) _
                    ) > 0)
                    ' push an empty array if only asked to exclude specific types or type
                    ' matches the requested type (even though it's empty...)
                    If (Include And IsMatch) Or (Not Include And Not IsMatch) Then
                        result.Push GetEmptyArray
                    End If
                Else
                    If Not Include Then
                        ' push an empty array if only asked to exclude specific values
                        result.Push GetEmptyArray
                    End If
                End If
            End If
        Else
            If CompareTypeNames Then
                IsMatch = (InStr( _
                    UCase$(TypeName(SourceArray(i))), _
                    UCase$(CStr(Match)) _
                ) > 0)
            Else
                IsMatch = ElementsAreEqual(Match, SourceArray(i))
            End If
            If (Include And IsMatch) Or (Not Include And Not IsMatch) Then
                result.Push SourceArray(i)
            End If
        End If
    Next
    RecursiveFilter = result.items
End Function

'@Description("Compares two values for equality. Doesn't support multidimensional arrays.")
'@Ignore ShadowedDeclaration
Private Function ElementsAreEqual( _
        ByVal Expected As Variant, _
        ByVal Actual As Variant _
    ) As Boolean
Attribute ElementsAreEqual.VB_Description = "Compares two values for equality. Doesn't support multidimensional arrays."
    ' Using 13dp of precision for EPSILON rather than IEEE 754 standard of 2^-52
    ' some roundings in type conversions cause greater diffs than machine epsilon
    Const epsilon As Double = 0.0000000000001
    Dim result As Boolean
    Dim i As Long
    
    On Error GoTo ErrHandler
    If IsArray(Expected) Or IsArray(Actual) Then
        If IsArray(Expected) And IsArray(Actual) Then
            If LBound(Expected) = LBound(Actual) And _
                    UBound(Expected) = UBound(Actual) Then
                Dim CurrentlyEqual As Boolean
                CurrentlyEqual = True
                For i = LBound(Expected) To UBound(Actual)
                    If Not ElementsAreEqual(Expected(i), Actual(i)) Then
                        CurrentlyEqual = False
                        Exit For
                    End If
                Next
                result = CurrentlyEqual
            End If
        End If
    ElseIf IsEmpty(Expected) Or IsEmpty(Actual) Then
        If IsEmpty(Expected) And IsEmpty(Actual) Then result = True
    ElseIf IsObject(Expected) Or IsObject(Actual) Then
        If IsObject(Expected) And IsObject(Actual) Then
            If Expected Is Actual Then result = True
        End If
    ElseIf IsNumeric(Expected) Or IsNumeric(Actual) Then
        If IsNumeric(Expected) And IsNumeric(Actual) Then
            Dim Diff As Double
            Diff = Abs(Expected - Actual)
            If Diff <= (IIf( _
                    Abs(Expected) < Abs(Actual), _
                    Abs(Actual), _
                    Abs(Expected) _
                ) * epsilon) Then
                result = True
            End If
        End If
    ElseIf Expected = Actual Then
        result = True
    End If
    ElementsAreEqual = result
    Exit Function
ErrHandler:
    ElementsAreEqual = False
End Function

' TODO: Refactor to remove recursion
'@Description("Recursively parses nested arrays from a string.")
Private Function ParseDelimitedArrayString( _
        ByVal SourceString As String, _
        ByVal ValueSeparator As String, _
        ByVal Opener As String, _
        ByVal Closer As String, _
        Optional ByRef Cursor As Long = 2 _
    ) As Variant()
Attribute ParseDelimitedArrayString.VB_Description = "Recursively parses nested arrays from a string."
    Dim CurrentChar As String
    Dim LocalResult() As Variant
    Dim i As Long
    Dim BreakLoop As Boolean
    Dim NextOpener As Long
    Dim NextCloser As Long
    
    CurrentChar = Mid$(SourceString, Cursor, 1)
    NextOpener = InStr(Cursor, SourceString, Opener)
    NextCloser = InStr(Cursor, SourceString, Closer)
    
    If CurrentChar <> Opener And (NextCloser < NextOpener Or NextOpener = 0) Then
        ' we're inside an array to be parsed
        ParseDelimitedArrayString = ParseArraySegmentFromString(SourceString, Cursor, NextCloser, ValueSeparator)
        Exit Function
    End If
        
    ' traverse nested arrays
    i = This.LowerBound
    Do
        If CurrentChar = Opener Then
            If i > This.LowerBound Then
                ReDim Preserve LocalResult(i)
            Else
                LocalResult = GetEmptyArray
            End If
            inc Cursor
            LocalResult(i) = ParseDelimitedArrayString( _
                SourceString, _
                ValueSeparator, _
                Opener, _
                Closer, _
                Cursor _
            )
            inc i
        End If
        CurrentChar = Mid$(SourceString, Cursor, 1)
        Do Until CurrentChar = Opener
            inc Cursor
            CurrentChar = Mid$(SourceString, Cursor, 1)
            If CurrentChar = Closer Or Cursor >= Len(SourceString) Then
                BreakLoop = True
                Exit Do
            End If
        Loop
    Loop Until BreakLoop
    
    ParseDelimitedArrayString = LocalResult
End Function

'@Description("Returns an array from a delimited string")
Private Function ParseArraySegmentFromString( _
        ByVal SourceString As String, _
        ByRef Cursor As Long, _
        ByVal NextCloser As Long, _
        ByVal ValueSeparator As String _
    ) As Variant()
Attribute ParseArraySegmentFromString.VB_Description = "Returns an array from a delimited string"
    Dim SegmentLength As Long
    Dim Segment() As String
    SegmentLength = IIf(NextCloser = 0, Len(SourceString), NextCloser - Cursor)
    Segment = Split(Mid$(SourceString, Cursor, SegmentLength), ValueSeparator)
    Cursor = NextCloser
    ParseArraySegmentFromString = DuckTypeStringArray(Segment)
End Function

'@Description("Trims enclosing quote chars from string if present")
Private Function UnquoteString(ByRef Element As String) As String
Attribute UnquoteString.VB_Description = "Trims enclosing quote chars from string if present"
    Dim LocalElement As String
    LocalElement = Trim$(Element)
    If LocalElement = vbNullString Then
        UnquoteString = vbNullString
    Else
        If Asc(Left$(LocalElement, 1)) = 34 And Asc(Right$(LocalElement, 1)) = 34 Then
            'trim string delimiters
            LocalElement = Mid$(LocalElement, 2, Len(LocalElement) - 2)
        End If
        UnquoteString = Trim$(LocalElement)
    End If
End Function

'@Description("Takes a string. Returns a Variant with the element converted to an appropriate type")
Private Function DuckTypeElement(ByRef Element As String) As Variant
Attribute DuckTypeElement.VB_Description = "Takes a string. Returns a Variant with the element converted to an appropriate type"
    Dim result As Variant
    Dim LocalElement As String
    LocalElement = Element
    If UCase$(LocalElement) = "TRUE" Or UCase$(Element) = "FALSE" Then
        result = CBool(LocalElement)
    ElseIf IsNumeric(LocalElement) Then
        Const epsilon As Double = 2 ^ -52
        Dim Diff As Double
        ' check if integer
        Diff = Abs(Fix(LocalElement) - LocalElement)
        If Diff > epsilon Then
            result = CDbl(LocalElement)
        Else
            result = CLng(LocalElement)
        End If
    Else
        result = UnquoteString(LocalElement)
    End If
    DuckTypeElement = result
End Function

'@Description("Takes a string array. Returns a Variant array with the elements converted to an appropriate type")
Private Function DuckTypeStringArray(ByRef SourceArray() As String) As Variant()
Attribute DuckTypeStringArray.VB_Description = "Takes a string array. Returns a Variant array with the elements converted to an appropriate type"
    Dim i As Long
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim result() As Variant
    
    LocalLowerBound = LBound(SourceArray)
    LocalUpperBound = UBound(SourceArray)
    ReDim result(LocalLowerBound To LocalUpperBound)
    For i = LocalLowerBound To LocalUpperBound
        result(i) = DuckTypeElement(SourceArray(i))
    Next
    DuckTypeStringArray = result
End Function

'@Description("Populates the ErrorDefinitions array.")
Private Sub PopulateErrorDefinitions()
Attribute PopulateErrorDefinitions.VB_Description = "Populates the ErrorDefinitions array."
    Dim Source As String
    Source = TypeName(Me)
    This.ErrorDefinitions(ErrorCodes.EC_EXPECTED_RANGE_OBJECT) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_EXPECTED_RANGE_OBJECT, _
        Source:=Source, _
        Description:="Range Object Expected" _
    )
    This.ErrorDefinitions(ErrorCodes.EC_MAX_DIMENSIONS_LIMIT) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_MAX_DIMENSIONS_LIMIT, _
        Source:=Source, _
        Description:="Cannot convert structure of arrays with more than 20 dimensions." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_EXPECTED_COLLECTION_OBJECT) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_EXPECTED_COLLECTION_OBJECT, _
        Source:=Source, _
        Description:="Valid Collection Object Expected" _
    )
    This.ErrorDefinitions(ErrorCodes.EC_EXCEEDS_MAX_SORT_DEPTH) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_EXCEEDS_MAX_SORT_DEPTH, _
        Source:=Source, _
        Description:="Cannot sort on arrays with more than 2 dimensions" _
    )
    This.ErrorDefinitions(ErrorCodes.EC_EXPECTED_JAGGED_ARRAY) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_EXPECTED_JAGGED_ARRAY, _
        Source:=Source, _
        Description:="Expected jagged array." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_EXPECTED_MULTIDIMENSION_ARRAY) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_EXPECTED_MULTIDIMENSION_ARRAY, _
        Source:=Source, _
        Description:="Expected multidimension array." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_EXPECTED_ARRAY) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_EXPECTED_ARRAY, _
        Source:=Source, _
        Description:="Expected array." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_NULL_STRING) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_NULL_STRING, _
        Source:=Source, _
        Description:="Cannot parse from a null string. Expected string with length greater than 0." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_UNALLOCATED_ARRAY) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_UNALLOCATED_ARRAY, _
        Source:=Source, _
        Description:="Cannot operate on unallocated array." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_UNDEFINED_ARRAY) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_UNDEFINED_ARRAY, _
        Source:=Source, _
        Description:="Array is undefined." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_INVALID_MULTIDIMENSIONAL_ARRAY_OPERATION, _
        Source:=Source, _
        Description:="Unable to perform the requested operation on a multidimensional array." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_EXPECTED_VARIANT_ARRAY) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_EXPECTED_VARIANT_ARRAY, _
        Source:=Source, _
        Description:="Unable to perform the requested operation on a typed array." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_EXCEEDS_MAX_ARRAY_LENGTH) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_EXCEEDS_MAX_ARRAY_LENGTH, _
        Source:=Source, _
        Description:="The requested operation would result in an array which exceeds the maximum possible length." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_STRING_TYPE_EXPECTED) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_STRING_TYPE_EXPECTED, _
        Source:=Source, _
        Description:="Expected a String or String-coercible type." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_CANNOT_CONVERT_TO_REQUESTED_STRUCTURE, _
        Source:=Source, _
        Description:="The stored array cannot be converted to the requested structure." _
    )
    This.ErrorDefinitions(ErrorCodes.EC_CANNOT_SORT_OBJECTS) = ErrorDefinitionFactory( _
        Number:=ErrorCodes.EC_CANNOT_SORT_OBJECTS, _
        Source:=Source, _
        Description:="The array contains Objects which cannot be sorted." _
    )
    
End Sub

'@Description("Returns an ErrorDefinition Type populated with the provided arguments.")
Private Function ErrorDefinitionFactory( _
        ByVal Number As Long, _
        ByVal Source As String, _
        ByVal Description As String _
    ) As ErrorDefinition
Attribute ErrorDefinitionFactory.VB_Description = "Returns an ErrorDefinition Type populated with the provided arguments."
    Dim result As ErrorDefinition
    result.Number = Number
    result.Source = Source
    result.Description = Description
    ErrorDefinitionFactory = result
End Function

'@Description("Raises an application error based on the passed ErrorCode and the definitions defined in ErrorDefinitions.")
Private Sub RaiseError( _
        ByVal ErrorCode As ErrorCodes, _
        ByVal Caller As String, _
        Optional ByVal ArgName As String _
    )
Attribute RaiseError.VB_Description = "Raises an application error based on the passed ErrorCode and the definitions defined in ErrorDefinitions."
    Dim CurrentError As ErrorDefinition
    Dim LocalArgName As String
    If ArgName <> vbNullString Then LocalArgName = "Argument " & ArgName & ": "
    CurrentError = This.ErrorDefinitions(ErrorCode)
    err.Raise CurrentError.Number, _
              CurrentError.Source & "." & Caller, _
              LocalArgName & CurrentError.Description
End Sub

'Used by ToExcelRange
'@Description("Ensures each element of passed array can be represented as a scalar value.")
Private Function EnsureScalar1DArray(ByRef SourceArray() As Variant) As Variant()
Attribute EnsureScalar1DArray.VB_Description = "Ensures each element of passed array can be represented as a scalar value."
    Dim i As Long
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim result() As Variant
    
    LocalLowerBound = LBound(SourceArray)
    LocalUpperBound = UBound(SourceArray)
    ReDim result(LocalLowerBound To LocalUpperBound)
    For i = LocalLowerBound To LocalUpperBound
        result(i) = GetScalarRepresentation(SourceArray(i))
    Next
    EnsureScalar1DArray = result
End Function

'@Description("Converts a 1d array to a 1 column wide jagged array.")
Private Function ConvertOneDimensionArrayToJagged(ByRef SourceArray() As Variant) As Variant()
Attribute ConvertOneDimensionArrayToJagged.VB_Description = "Converts a 1d array to a 1 column wide jagged array."
    Dim i As Long
    Dim LocalUpperBound As Long
    Dim LocalLowerBound As Long
    Dim result() As Variant
    LocalUpperBound = UBound(SourceArray)
    LocalLowerBound = LBound(SourceArray)
    ReDim result(LocalLowerBound To LocalUpperBound)
    For i = LocalLowerBound To LocalUpperBound
        result(i) = Array(SourceArray(i))
    Next
    ConvertOneDimensionArrayToJagged = result
End Function

'@Description("Converts a 1d array to a 1 column wide 2d array.")
Private Function Transpose1DArray(ByRef SourceArray() As Variant) As Variant()
Attribute Transpose1DArray.VB_Description = "Converts a 1d array to a 1 column wide 2d array."
    Dim i As Long
    Dim j As Long
    Dim LocalUpperBound As Long
    Dim LocalLowerBound As Long
    Dim result() As Variant
    LocalUpperBound = UBound(SourceArray)
    LocalLowerBound = LBound(SourceArray)
    ReDim result(LocalLowerBound To LocalUpperBound, _
                 LocalLowerBound To LocalLowerBound)
    j = LocalLowerBound
    For i = LocalLowerBound To LocalUpperBound
        result(j, LocalLowerBound) = SourceArray(i)
        inc j
    Next
    Transpose1DArray = result
End Function

'@Description("Transposes the values in a 2d array. Rows become columns, columns become rows.")
Private Function Transpose2DArray(ByRef SourceArray() As Variant) As Variant()
Attribute Transpose2DArray.VB_Description = "Transposes the values in a 2d array. Rows become columns, columns become rows."
    Dim CurrentRow As Long
    Dim LowerBoundRow As Long
    Dim UpperBoundRow As Long
    Dim CurrentColumn As Long
    Dim LowerBoundCol As Long
    Dim UpperBoundCol As Long
    Dim result() As Variant
    
    LowerBoundCol = LBound(SourceArray, 1)
    UpperBoundCol = UBound(SourceArray, 1)
    LowerBoundRow = LBound(SourceArray, 2)
    UpperBoundRow = UBound(SourceArray, 2)
    ReDim result(LowerBoundRow To UpperBoundRow, LowerBoundCol To UpperBoundCol)
    For CurrentRow = LowerBoundRow To UpperBoundRow
        For CurrentColumn = LowerBoundCol To UpperBoundCol
            result(CurrentRow, CurrentColumn) = SourceArray(CurrentColumn, CurrentRow)
        Next
    Next
    Transpose2DArray = result
End Function

'@Description("Transposes the values in a jagged array with a depth of 2. Rows become columns, columns become rows.")
Private Function TransposeArrayOfArrays(ByRef SourceArray() As Variant) As Variant()
Attribute TransposeArrayOfArrays.VB_Description = "Transposes the values in a jagged array with a depth of 2. Rows become columns, columns become rows."
    Dim CurrentRow As Long
    Dim LowerBoundRow As Long
    Dim UpperBoundRow As Long
    Dim CurrentColumn As Long
    Dim LowerBoundCol As Long
    Dim UpperBoundCol As Long
    Dim result() As Variant
    Dim NestedBounds() As Long
    Dim Nested() As Variant
    
    NestedBounds = GetMaxBoundsAtDimension(SourceArray, 2)
    LowerBoundCol = LBound(SourceArray)
    UpperBoundCol = UBound(SourceArray)
    LowerBoundRow = NestedBounds(0)
    UpperBoundRow = NestedBounds(1)
    'check if array can fit into 1 dimension
    If LowerBoundRow = UpperBoundRow Then
        ReDim result(LowerBoundCol To UpperBoundCol)
        CurrentRow = LowerBoundRow
        For CurrentColumn = LowerBoundCol To UpperBoundCol
            result(CurrentColumn) = SourceArray(CurrentColumn)(CurrentRow)
        Next
    ElseIf LowerBoundCol = UpperBoundCol Then
        ReDim result(LowerBoundRow To UpperBoundRow)
        CurrentColumn = LowerBoundCol
        For CurrentRow = UpperBoundRow To UpperBoundRow
            result(CurrentRow) = SourceArray(CurrentColumn)(CurrentRow)
        Next
    Else
        ReDim result(LowerBoundRow To UpperBoundRow)
        For CurrentRow = LowerBoundRow To UpperBoundRow
            ReDim Nested(LowerBoundCol To UpperBoundCol)
            For CurrentColumn = LowerBoundCol To UpperBoundCol
                Nested(CurrentColumn) = SourceArray(CurrentColumn)(CurrentRow)
            Next
            result(CurrentRow) = Nested
        Next
    End If
    TransposeArrayOfArrays = result
End Function

'@Description("Returns an empty 1d array with 1 slot.")
Private Function GetEmptyArray() As Variant()
Attribute GetEmptyArray.VB_Description = "Returns an empty 1d array with 1 slot."
    Dim result() As Variant
    ReDim result(This.LowerBound To This.LowerBound)
    GetEmptyArray = result
End Function

'@Description("Used to apply the correct specified sort algorithm.")
Private Sub ApplySortMethod( _
        ByRef SourceArray() As Variant, _
        ByVal LocalArrayType As ArrayTypes, _
        Optional ByVal col As Long _
    )
Attribute ApplySortMethod.VB_Description = "Used to apply the correct specified sort algorithm."
    Select Case This.SortMethod
        Case SortMethods.SM_TIMSORT
            TimSort SourceArray, LocalArrayType, col
        Case SortMethods.SM_QUICKSORT_RECURSIVE
            QuickSortRecursive _
                SourceArray, _
                LBound(SourceArray), _
                UBound(SourceArray), _
                LocalArrayType, _
                col
        Case SortMethods.SM_QUICKSORT_ITERATIVE
            QuickSortIterative _
                SourceArray, _
                LBound(SourceArray), _
                UBound(SourceArray), _
                LocalArrayType, _
                col
    End Select
End Sub

'@Description("Returns an field of a one dimension or two dimension jagged array")
Private Function GetComparisonItem( _
    ByRef Source() As Variant, _
    ByVal index As Long, _
    ByVal col As Long, _
    ByVal LocalArrayType As ArrayTypes _
) As Variant
Attribute GetComparisonItem.VB_Description = "Returns an field of a one dimension or two dimension jagged array"
    If IsObject(Source(index)) Then
        GetComparisonItem = ObjPtr(Source(index))
    Else
        If LocalArrayType = BA_JAGGED Then
            GetComparisonItem = Source(index)(col)
        Else
            GetComparisonItem = Source(index)
        End If
    End If
End Function

'@Description("Sorts the passed array in place using InsertionSort.")
Private Sub InsertionSort( _
    ByRef Source() As Variant, _
    ByVal LeftIndex As Variant, _
    ByVal RightIndex As Variant, _
    ByVal LocalArrayType As ArrayTypes, _
    Optional ByVal col As Long _
)
Attribute InsertionSort.VB_Description = "Sorts the passed array in place using InsertionSort."
    Dim i As Long
    Dim j As Long
    Dim Pivot As Variant
    Dim PivotCompVal As Variant
    Dim Current As Variant
    Dim CurrentCompval As Variant
        
    ' Loop from the element indicated by
    ' `left` until the element indicated by `right`
            
    For i = LeftIndex + 1 To RightIndex
        ' This is the element we want to position in its
        ' correct place
        LetOrSetElement Pivot, Source(i)
        PivotCompVal = GetComparisonItem(Source, i, col, LocalArrayType)
        ' Initialize the variable that will be used to
        ' find the correct position of the element referenced
        ' by `key_item`
        j = i - 1
        ' Run through the list of items (the left
        ' portion of the array) and find the correct position
        ' of the element referenced by `key_item`. Do this only
        ' if the `key_item` is smaller than its adjacent values.
        Do While j >= LeftIndex
            LetOrSetElement Current, Source(j)
            CurrentCompval = GetComparisonItem(Source, j, col, LocalArrayType)
            If CurrentCompval > PivotCompVal Then
                ' Shift the value one position to the left
                ' and reposition `j` to point to the next element
                ' (from right to left)
                LetOrSetElement Source(j + 1), Current
                dec j
            Else
                Exit Do
            End If
        Loop
    
        ' When you finish shifting the elements, position
        ' the `key_item` in its correct location
        '@Ignore UnassignedVariableUsage
        LetOrSetElement Source(j + 1), Pivot
    Next
End Sub

'@Description("Sorts the passed array in place using MergeSort.")
Private Sub MergeSort( _
    ByRef Source() As Variant, _
    ByVal startIndex As Long, _
    ByVal MidIndex As Long, _
    ByVal endIndex As Long, _
    ByVal LocalArrayType As ArrayTypes, _
    Optional ByVal col As Long _
)
Attribute MergeSort.VB_Description = "Sorts the passed array in place using MergeSort."
    ' Original array is broken in two parts
    ' left and right array
    Dim LowLength As Long
    Dim HighLength As Long
    Dim i As Long
    Dim LowIndex As Long
    Dim HighIndex As Long
    Dim LowPartition() As Variant
    Dim HighPartition() As Variant
    '@Ignore VariableNotAssigned
    Dim LowItem As Variant
    '@Ignore VariableNotAssigned
    Dim HighItem As Variant
    
    LowLength = MidIndex - startIndex + 1
    HighLength = endIndex - MidIndex
        
    ReDim LowPartition(0 To max(LowLength - 1, 0))
    ReDim HighPartition(0 To max(HighLength - 1, 0))
    
    i = 0
    Do While i < LowLength
        LetOrSetElement LowPartition(i), Source(startIndex + i)
        inc i
    Loop
    
    i = 0
    Do While i < HighLength
        LetOrSetElement HighPartition(i), Source(MidIndex + 1 + i)
        inc i
    Loop
    
    ' After comparing, we
    ' merge those two array
    ' in larger sub array
    i = startIndex
    Do While LowIndex < LowLength And HighIndex < HighLength
        LowItem = GetComparisonItem(LowPartition, LowIndex, col, LocalArrayType)
        HighItem = GetComparisonItem(HighPartition, HighIndex, col, LocalArrayType)
        If LowItem <= HighItem Then
            LetOrSetElement Source(i), LowPartition(LowIndex)
            inc LowIndex
        Else
            LetOrSetElement Source(i), HighPartition(HighIndex)
            inc HighIndex
        End If
        inc i
    Loop
    
    ' Copy remaining elements of left, if any
    Do While LowIndex < LowLength
        LetOrSetElement Source(i), LowPartition(LowIndex)
        inc LowIndex
        inc i
    Loop
    Do While HighIndex < HighLength
        LetOrSetElement Source(i), HighPartition(HighIndex)
        inc HighIndex
        inc i
    Loop
    
End Sub

'@Description("Iterative Timsort")
Private Sub TimSort( _
    ByRef Source() As Variant, _
    ByVal LocalArrayType As ArrayTypes, _
    Optional ByVal col As Long _
)
Attribute TimSort.VB_Description = "Iterative Timsort"
    Const MIN_RUN As Long = 32
    Dim i As Long
    Dim SourceLength As Long
    SourceLength = GetArrayLength(Source)
    ' Sort individual subarrays of size RUN
    For i = LBound(Source) To UBound(Source) Step MIN_RUN
        InsertionSort Source, i, Min((i + MIN_RUN - 1), UBound(Source)), LocalArrayType, col
    Next
    
    ' Now you can start merging the sorted slices.
    ' Start from `min_run`, doubling the size on
    ' each iteration until you surpass the length of
    ' the array.
    Dim Size As Long
    Dim start As Long
    Dim Midpoint As Long
    Dim Endpoint As Long
    Dim LastIndex As Long
    
    LastIndex = UBound(Source)
    Size = MIN_RUN
    Do While Size < SourceLength
        ' Determine the arrays that will
        ' be merged together
        start = LBound(Source)
        Do While start < UBound(Source)
            ' Compute the `midpoint` (where the first array ends
            ' and the second starts) and the `endpoint` (where
            ' the second array ends)
            Midpoint = Min(start + Size - 1, LastIndex)
            Endpoint = Min(start + Size * 2 - 1, LastIndex)
            ' Merge the two subarrays.
            ' The `left` array should go from `start` to
            ' `midpoint + 1`, while the `right` array should
            ' go from `midpoint + 1` to `end + 1`.
            MergeSort Source, start, Midpoint, Endpoint, LocalArrayType, col
            inc start, Size * 2
        Loop
        ' Each iteration should double the size of your arrays
        inc Size, Size
    Loop

End Sub

'@Description("Sorts in place a passed array using a Recursive implementation of the QuickSort algorithm.")
Private Sub QuickSortRecursive( _
        ByRef SourceArray() As Variant, _
        ByVal Low As Long, _
        ByVal High As Long, _
        ByVal LocalArrayType As ArrayTypes, _
        Optional ByVal col As Long _
    )
Attribute QuickSortRecursive.VB_Description = "Sorts in place a passed array using a Recursive implementation of the QuickSort algorithm."
    Dim PartitionIndex As Long
    If Low < High Then
        PartitionIndex = QsPartition(SourceArray, Low, High, LocalArrayType, col)
        QuickSortRecursive SourceArray, Low, PartitionIndex - 1, LocalArrayType, col
        QuickSortRecursive SourceArray, PartitionIndex + 1, High, LocalArrayType, col
    End If
End Sub

'@Description("Sorts in place a passed array using an Iterative implementation of the QuickSort algorithm")
Private Sub QuickSortIterative( _
        ByRef SourceArray() As Variant, _
        ByVal Low As Long, _
        ByVal High As Long, _
        ByVal LocalArrayType As ArrayTypes, _
        Optional ByVal col As Long _
    )
Attribute QuickSortIterative.VB_Description = "Sorts in place a passed array using an Iterative implementation of the QuickSort algorithm"
    Dim PartitionIndex As Long
    Dim Stack() As Long
    Dim Top As Long
    Dim LocalLow As Long
    Dim LocalHigh As Long
    
    LocalLow = Low
    LocalHigh = High
    
    ' Create an auxiliary stack
    ReDim Stack(0 To LocalHigh - LocalLow + 1)
    
    ' initialize top of stack
    Top = -1
    ' push initial values of l and h to stack
    Stack(inc(Top)) = LocalLow
    Stack(inc(Top)) = LocalHigh
    
    ' Keep popping from stack while is not empty
    Do While Top >= 0
        ' Pop high and low
        LocalHigh = Stack(Top)
        LocalLow = Stack(dec(Top))
        dec Top
        ' Set pivot element at its correct position
        ' in sorted array
        PartitionIndex = QsPartition(SourceArray, LocalLow, LocalHigh, LocalArrayType, col)
        
        ' If there are elements on left side of pivot,
        ' then push left side to stack
        If PartitionIndex - LocalLow > 1 Then
            Stack(inc(Top)) = LocalLow
            Stack(inc(Top)) = PartitionIndex - 1
        End If
       
        ' If there are elements on right side of pivot,
        ' then push right side to stack
        If PartitionIndex + 1 < LocalHigh Then
            Stack(inc(Top)) = PartitionIndex + 1
            Stack(inc(Top)) = LocalHigh
        End If
        
    Loop
    
End Sub


'@Description("Child function of quickSort. Returns the next partition index.")
Private Function QsPartition( _
        ByRef SourceArray() As Variant, _
        ByVal Low As Long, _
        ByVal High As Long, _
        ByVal LocalArrayType As ArrayTypes, _
        Optional ByVal col As Long _
    ) As Long
Attribute QsPartition.VB_Description = "Child function of quickSort. Returns the next partition index."
    Dim i As Long
    Dim j As Long
    Dim Pivot As Variant
    Dim Current As Variant
    
    Pivot = GetComparisonItem(SourceArray, High, col, LocalArrayType)
    
    i = Low - 1
    For j = Low To High - 1
        Current = GetComparisonItem(SourceArray, j, col, LocalArrayType)
        If Current <= Pivot Then
            inc i
            Swap SourceArray, i, j
        End If
    Next
    Swap SourceArray, i + 1, High
    QsPartition = i + 1
End Function

'@Description("Swaps the location of two elements in an array.")
Private Sub Swap(ByRef SourceArray() As Variant, ByVal i As Long, ByVal j As Long)
Attribute Swap.VB_Description = "Swaps the location of two elements in an array."
    Dim Element As Variant
    LetOrSetElement Element, SourceArray(i)
    LetOrSetElement SourceArray(i), SourceArray(j)
    '@Ignore UnassignedVariableUsage
    LetOrSetElement SourceArray(j), Element
End Sub

'@Description("Sets or lets the value of a variable by reference depending on the type of the element to be assigned")
Private Sub LetOrSetElement(ByRef Destination As Variant, ByRef Source As Variant)
Attribute LetOrSetElement.VB_Description = "Sets or lets the value of a variable by reference depending on the type of the element to be assigned"
    If IsObject(Source) Then
        Set Destination = Source
    Else
        Destination = Source
    End If
End Sub

'@Description("Efficiently concatenates string fragments to a byte array buffer")
Private Function StringBuilder( _
    Optional ByVal Fragment As String = vbNullString, _
    Optional ByVal Final As Boolean, _
    Optional ByVal NewString As Boolean _
) As String
Attribute StringBuilder.VB_Description = "Efficiently concatenates string fragments to a byte array buffer"
    Static Buffer() As Byte
    Static bufferLength As Long
    Static BufferCapacity As Long
    Static Cursor As Long
    Static FirstIndex As Long
    
    Dim FragLength As Long
    Dim LastIndex As Long
    Dim i As Long

    FragLength = LenB(Fragment)
    If Cursor < 1 Or NewString Then
        Buffer = Fragment
        bufferLength = FragLength
        BufferCapacity = FragLength
        FirstIndex = LBound(Buffer)
        Cursor = UBound(Buffer)
    ElseIf FragLength > 0 Then
        inc bufferLength, FragLength
        LastIndex = FirstIndex + bufferLength - 1
        If bufferLength > BufferCapacity Then
            Do
                BufferCapacity = BufferCapacity * 2
            Loop While bufferLength > BufferCapacity
            ReDim Preserve Buffer(FirstIndex To FirstIndex + BufferCapacity)
        End If
        For i = Cursor + 1 To LastIndex
            Buffer(i) = AscB(MidB$(Fragment, i - Cursor, 1))
        Next
        Cursor = LastIndex
    End If
    
    If Final Then
        ReDim Preserve Buffer(FirstIndex To Cursor)
        StringBuilder = CStr(Buffer)
        Erase Buffer
        Cursor = Empty
    End If
    
End Function

'@Description("Recursively parses an array to a string representation.")
Private Sub RecursiveToString( _
        ByRef SourceArray() As Variant, _
        ByVal PrettyPrint As Boolean, _
        ByVal Separator As String, _
        ByVal OpeningDelimiter As String, _
        ByVal ClosingDelimiter As String, _
        ByVal QuoteStrings As Boolean, _
        Optional ByVal Tabs As Long = 1 _
    )
Attribute RecursiveToString.VB_Description = "Recursively parses an array to a string representation."
    Const TabWidth As Long = 2
    Dim i As Long
    Dim Sep As String
    
    StringBuilder OpeningDelimiter
    For i = LBound(SourceArray) To UBound(SourceArray)
        Sep = IIf(i = UBound(SourceArray), ClosingDelimiter, Separator)
        If IsArray(SourceArray(i)) Then
            Dim Nested() As Variant
            Nested = SourceArray(i)
            If PrettyPrint Then
                StringBuilder vbCrLf & Space(TabWidth * Tabs)
            End If
            RecursiveToString Nested, PrettyPrint, Separator, OpeningDelimiter, ClosingDelimiter, QuoteStrings, Tabs + 1
            If i = UBound(SourceArray) And PrettyPrint Then
                StringBuilder vbCrLf
                StringBuilder Space(TabWidth * (Tabs - 1))
            End If
            StringBuilder Sep
        Else
            ' ensure no arrays passed to GetScalar as it clears string builder
            StringBuilder Replace( _
                    CStr(GetScalarRepresentation(SourceArray(i), QuoteStrings)), _
                    Separator, _
                    vbNullString _
                ) & Sep
        End If
    Next
End Sub

'@Description("Returns scalar values as is or appropriate representation if not scalar.")
Private Function GetScalarRepresentation( _
        ByRef Element As Variant, _
        Optional ByVal QuoteStrings As Boolean _
    ) As Variant
Attribute GetScalarRepresentation.VB_Description = "Returns scalar values as is or appropriate representation if not scalar."
    Dim result As Variant
    If IsObject(Element) Then
        On Error Resume Next
        'Try to retrieve default member
        result = Element
        On Error GoTo 0
        ' if there is no default member or it returns an object result = Empty
        If IsEmpty(result) Then
            result = OBJECT_REPR
        End If
    Else
        result = Element
    End If
    ' if the element is array or default member returns an array then convert tostring
    If IsArray(result) And Not IsEmpty(result) Then
        Dim PassThruArray() As Variant
        PassThruArray = result
        StringBuilder NewString:=True
        RecursiveToString PassThruArray, False, CHR_COMMA, "{", "}", QuoteStrings
        result = StringBuilder(Final:=True)
    End If
    If Not IsNumeric(result) And QuoteStrings Then
        GetScalarRepresentation = Chr$(34) & result & Chr$(34)
    Else
        GetScalarRepresentation = result
    End If
    
End Function

'@Description("Returns the type of a passed array.")
Private Function GetArrayType(ByVal SourceArray As Variant) As ArrayTypes
Attribute GetArrayType.VB_Description = "Returns the type of a passed array."
    Dim result As ArrayTypes
    If IsArray(SourceArray) Then
        If Not IsArrayAllocated(SourceArray) Then
            result = BA_UNALLOCATED
        Else
            If IsMultidimensionalArray(SourceArray) Then
                result = BA_MULTIDIMENSION
            ElseIf IsJaggedArray(SourceArray) Then
                result = BA_JAGGED
            Else
                result = BA_ONEDIMENSION
            End If
        End If
    Else
        If IsEmpty(SourceArray) Then
            result = BA_UNDEFINED
        Else
            RaiseError EC_EXPECTED_ARRAY, "getArrayType", "sourceArray"
        End If
    End If
    GetArrayType = result
End Function

'@Description("Tests if an array has a single empty slot.")
Private Function IsArrayEmpty(ByVal SourceArray As Variant) As Boolean
Attribute IsArrayEmpty.VB_Description = "Tests if an array has a single empty slot."
    Dim result As Boolean
    If LBound(SourceArray) = UBound(SourceArray) Then
        On Error Resume Next ' Error handler necessary in case of multidimension array
        result = (SourceArray(LBound(SourceArray)) = Empty)
        On Error GoTo 0
    End If
    IsArrayEmpty = result
End Function

'@Description("Tests if an array has been allocated.")
Private Function IsArrayAllocated(ByVal SourceArray As Variant) As Boolean
Attribute IsArrayAllocated.VB_Description = "Tests if an array has been allocated."
    On Error Resume Next
    IsArrayAllocated = ( _
        IsArray(SourceArray) And _
        Not IsError(LBound(SourceArray, 1)) And _
        LBound(SourceArray, 1) <= UBound(SourceArray, 1) _
    )
    On Error GoTo 0
End Function

'@Description("Tests if an array is jagged.")
Private Function IsJaggedArray(ByVal SourceArray As Variant) As Boolean
Attribute IsJaggedArray.VB_Description = "Tests if an array is jagged."
    If IsArray(SourceArray) Then
        On Error GoTo ErrHandler
        Dim Element As Variant
        For Each Element In SourceArray
            If IsArray(Element) Then
                IsJaggedArray = True
                Exit Function
            End If
        Next
        On Error GoTo 0
    End If
    Exit Function
ErrHandler:
    err.Clear
End Function

'@Description("Tests if an array is multidimensioned.")
Private Function IsMultidimensionalArray(ByVal SourceArray As Variant) As Boolean
Attribute IsMultidimensionalArray.VB_Description = "Tests if an array is multidimensioned."
    If IsArray(SourceArray) Then
        On Error GoTo ErrHandler
        '@Ignore VariableNotUsed
        Dim LocalUpperBound As Long
        '@Ignore AssignmentNotUsed
        LocalUpperBound = UBound(SourceArray, 2)
        IsMultidimensionalArray = True
        On Error GoTo 0
    End If
    Exit Function
ErrHandler:
    err.Clear
End Function

'@Description("re-indexes array to conform to new bounds. Does not change length.")
Private Function Rebase(Optional ByRef SourceArray As Variant, Optional ByVal CurrentArrayType As ArrayTypes) As Variant()
Attribute Rebase.VB_Description = "re-indexes array to conform to new bounds. Does not change length."
    Dim LocalItems() As Variant
    Dim LocalType As ArrayTypes

    If IsMissing(SourceArray) Or Not IsArray(SourceArray) Then
        LocalItems = InternalItems
        LocalType = This.ArrayType
    Else
        LocalItems = SourceArray
        If CurrentArrayType = 0 Then
            LocalType = GetArrayType(LocalItems)
        Else
            LocalType = CurrentArrayType
        End If
    End If
    If LocalType = BA_MULTIDIMENSION Then
        ' ensure all nested arrays have the correct base
        ' if they don't and additional arrays are added via concat the resulting
        ' md array from jaggedToMulti will have unexpected bounds
        Rebase = RecursiveRebase(LocalItems, True)
    Else
        Rebase = RecursiveRebase(LocalItems, False)
    End If
End Function

'@Description("re-indexes nested arrays to conform to new bounds. Child function of Rebase)
Private Function RecursiveRebase( _
        ByRef SourceArray() As Variant, _
        ByVal Recurse As Boolean _
    ) As Variant()
    Dim i As Long
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim Offset As Long
    Dim NewItems() As Variant
    
    LocalLowerBound = LBound(SourceArray)
    LocalUpperBound = UBound(SourceArray)
    Offset = This.LowerBound - LocalLowerBound
    ReDim NewItems(This.LowerBound To LocalUpperBound + Offset)
    For i = LocalLowerBound To LocalUpperBound
        If IsArray(SourceArray(i)) And Recurse And Offset <> 0 Then
            Dim Nested() As Variant
            Nested = SourceArray(i)
            NewItems(i + Offset) = RecursiveRebase(Nested, Recurse)
        Else
            NewItems(i + Offset) = SourceArray(i)
        End If
    Next
    
    RecursiveRebase = NewItems
End Function

'@Description("Increases capacity of internal array if required.")
Private Sub EnsureCapacity(ByVal MinimumCapacity As Long)
Attribute EnsureCapacity.VB_Description = "Increases capacity of internal array if required."
    If This.Capacity < MinimumCapacity Then
        Dim NewCapacity As Long
        NewCapacity = IIf(This.Capacity = 0, DEFAULT_CAPACITY, This.Capacity * 2)
        If NewCapacity > MAX_ARRAY_LENGTH Then
            NewCapacity = MAX_ARRAY_LENGTH
        End If
        If NewCapacity < MinimumCapacity Then
            NewCapacity = MinimumCapacity
        End If
        Me.Capacity = NewCapacity
    End If
End Sub

'@Description("Converts a typed array to a variant() array")
Private Function ConvertArrayForStorage( _
    ByRef values As Variant, _
    ByVal CurrentArrayType As ArrayTypes, _
    Optional ByVal ConvertMultiToJagged As Boolean, _
    Optional ByVal ConvertJaggedNestedArrays As Boolean _
) As Variant()
Attribute ConvertArrayForStorage.VB_Description = "Converts a typed array to a variant() array"
    Dim result() As Variant
    Select Case CurrentArrayType
    Case ArrayTypes.BA_MULTIDIMENSION
        If ConvertMultiToJagged Then
            result = MultiToJagged(values)
        ElseIf TypeName(values) <> "Variant()" Then
            result = TypedMultiToVariantMulti(values)
        Else
            result = values
        End If
    Case ArrayTypes.BA_ONEDIMENSION, ArrayTypes.BA_JAGGED
        If TypeName(values) <> "Variant()" _
            Or (CurrentArrayType = BA_JAGGED And ConvertJaggedNestedArrays) _
            Or LBound(values) <> This.LowerBound Then
            result = TypedJaggedToVariantJagged( _
                values, _
                ConvertJaggedNestedArrays, _
                ConvertMultiToJagged _
            )
        Else
            result = values
        End If
    End Select
    ConvertArrayForStorage = result
End Function

'@Description("Converts a typed jagged array to a variant() jagged array")
Private Function TypedJaggedToVariantJagged( _
    ByRef SourceArray As Variant, _
    Optional ByVal ConvertJaggedNestedArrays As Boolean, _
    Optional ByVal ConvertMultiToJagged As Boolean _
) As Variant()
Attribute TypedJaggedToVariantJagged.VB_Description = "Converts a typed jagged array to a variant() jagged array"
    Dim result() As Variant
    Dim i As Long
    Dim Bounds(0 To 1) As Long
    Dim Offset As Long
    
    Bounds(0) = LBound(SourceArray)
    Bounds(1) = UBound(SourceArray)
    Offset = This.LowerBound - Bounds(0)
    ReDim result(Bounds(0) + Offset To Bounds(1) + Offset)
    For i = Bounds(0) To Bounds(1)
        If IsArray(SourceArray(i)) Then
            result(i + Offset) = ConvertArrayForStorage( _
                SourceArray(i), _
                GetArrayType(SourceArray(i)), _
                ConvertMultiToJagged, _
                ConvertJaggedNestedArrays _
            )
        Else
            LetOrSetElement result(i + Offset), SourceArray(i)
        End If
    Next
    TypedJaggedToVariantJagged = result
End Function

'@Description("Converts a typed md array to a variant() md array")
Private Function TypedMultiToVariantMulti( _
    ByRef SourceArray As Variant _
) As Variant()
Attribute TypedMultiToVariantMulti.VB_Description = "Converts a typed md array to a variant() md array"
    Dim LocalDepth As Long
    Dim result() As Variant
    
    LocalDepth = GetMultidimensionalArrayDepth(SourceArray)
    
    If LocalDepth = 2 Then
        Dim FirstDimBounds(0 To 1) As Long
        Dim SecondDimBounds(0 To 1) As Long
        Dim FirstDimOffset As Long
        Dim SecondDimOffset As Long
        Dim i As Long
        Dim j As Long
        
        FirstDimBounds(0) = LBound(SourceArray, 1)
        FirstDimBounds(1) = UBound(SourceArray, 1)
        SecondDimBounds(0) = LBound(SourceArray, 2)
        SecondDimBounds(1) = UBound(SourceArray, 2)
        FirstDimOffset = This.LowerBound - FirstDimBounds(0)
        SecondDimOffset = This.LowerBound - SecondDimBounds(0)
        ReDim result( _
            FirstDimBounds(0) + FirstDimOffset To FirstDimBounds(1) + FirstDimOffset, _
            SecondDimBounds(0) + SecondDimOffset To SecondDimBounds(1) + SecondDimOffset _
        )
        For i = FirstDimBounds(0) To FirstDimBounds(1)
            For j = SecondDimBounds(0) To SecondDimBounds(1)
                LetOrSetElement result(i + FirstDimOffset, j + SecondDimOffset), SourceArray(i, j)
            Next
        Next
    ElseIf LocalDepth > 2 And LocalDepth <= 20 Then
        result = RecursiveTypedMultiToVariantMulti(SourceArray, LocalDepth)
    Else
        RaiseError EC_MAX_DIMENSIONS_LIMIT, "MultiToJagged()", "sourceArray()"
    End If
    
    TypedMultiToVariantMulti = result
        
End Function

'@Description("Returns an array representing the structure of a Multidimension array")
Private Function MapMultidimensionArray( _
        ByRef SourceArray As Variant, _
        Optional ByVal KnownDepth As Long _
    ) As Variant()
Attribute MapMultidimensionArray.VB_Description = "Returns an array representing the structure of a Multidimension array"
    Dim i As Long
    Dim LocalMap() As Variant
    Dim LocalDepth As Long
    Dim CurrentBounds(0 To 1) As Long

    If Not IsArray(SourceArray) Then Exit Function

    If KnownDepth > 0 Then
        LocalDepth = KnownDepth
    Else
        LocalDepth = GetMultidimensionalArrayDepth(SourceArray)
    End If

    ' crumbs are base 0
    ReDim LocalMap(0 To LocalDepth - 1)

    For i = LBound(LocalMap) To UBound(LocalMap)
        CurrentBounds(0) = LBound(SourceArray, i + 1)
        CurrentBounds(1) = UBound(SourceArray, i + 1)
        LocalMap(i) = CurrentBounds
    Next

    MapMultidimensionArray = LocalMap
End Function

'@Description("Recursively Converts typred multidimension arrays to variant md arrays arrays. For md arrays with more than 2d.")
Private Function RecursiveTypedMultiToVariantMulti( _
        ByRef SourceArray As Variant, _
        Optional ByVal Depth As Long, _
        Optional ByVal CurrentDepth As Long, _
        Optional ByRef Crumbs As Variant, _
        Optional ByRef result As Variant, _
        Optional ByVal EnsureScalar As Boolean _
    ) As Variant()
Attribute RecursiveTypedMultiToVariantMulti.VB_Description = "Recursively Converts typred multidimension arrays to variant md arrays arrays. For md arrays with more than 2d."
    Dim i As Long
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim LocalDepth As Long
    Dim LocalCurrentDepth As Long
    Dim LocalCrumbs() As Variant
    Dim LocalResult() As Variant
    Dim CurrentElement As Variant
    
    LocalDepth = Depth
    LocalCurrentDepth = CurrentDepth

    If IsMissing(Crumbs) Then
        ReDim LocalCrumbs(LocalDepth - 1)
    Else
        LocalCrumbs = Crumbs
    End If

    If IsMissing(result) Then
        LocalResult = CreateMultidimensionalArray(MapMultidimensionArray(SourceArray, LocalDepth))
    Else
        LocalResult = result
    End If

    inc LocalCurrentDepth
    LocalLowerBound = LBound(SourceArray, LocalCurrentDepth)
    LocalUpperBound = UBound(SourceArray, LocalCurrentDepth)
    For i = LocalLowerBound To LocalUpperBound
        LocalCrumbs(LocalCurrentDepth - 1) = i
        If LocalCurrentDepth = LocalDepth Then
            LetOrSetElement CurrentElement, GetElementByBreadcrumb(SourceArray, LocalCrumbs)
            If EnsureScalar Then
                LocalResult = LetElementByBreadcrumb(LocalResult, LocalCrumbs, GetScalarRepresentation(CurrentElement))
            Else
                LocalResult = LetElementByBreadcrumb(LocalResult, LocalCrumbs, CurrentElement)
            End If
        Else
            LocalResult = RecursiveTypedMultiToVariantMulti( _
                SourceArray:=SourceArray, _
                Depth:=LocalDepth, _
                CurrentDepth:=LocalCurrentDepth, _
                Crumbs:=LocalCrumbs, _
                result:=LocalResult, _
                EnsureScalar:=EnsureScalar _
            )
        End If
    Next

    RecursiveTypedMultiToVariantMulti = LocalResult
End Function

'@Description("Converts jagged arrays to multidimension arrays")
Private Function JaggedToMulti( _
        ByRef SourceArray() As Variant, _
        Optional ByVal Depth As Long, _
        Optional ByVal EnsureScalar As Boolean _
    ) As Variant()
Attribute JaggedToMulti.VB_Description = "Converts jagged arrays to multidimension arrays"
    Dim LocalDepth As Long
    Dim result() As Variant
    LocalDepth = Depth
    If LocalDepth = 0 Then
        If Not IsJaggedArray(SourceArray) Then
            RaiseError EC_EXPECTED_JAGGED_ARRAY, "JaggedToMulti", "sourceArray()"
            Exit Function
        End If
        LocalDepth = GetJaggedArrayDepth(SourceArray)
    End If
    
    If LocalDepth <= 1 Then
        result = SourceArray
    ElseIf LocalDepth = 2 Then
        Dim FirstDimBounds() As Long
        Dim SecondDimBounds() As Long
        Dim i As Long
        Dim j As Long
        FirstDimBounds = GetArrayBounds(SourceArray)
        SecondDimBounds = GetMaxBoundsAtDimension(SourceArray, 2)
        ReDim result(FirstDimBounds(0) To FirstDimBounds(1), _
            SecondDimBounds(0) To SecondDimBounds(1))
        For i = FirstDimBounds(0) To FirstDimBounds(1)
            If IsArray(SourceArray(i)) Then
                For j = LBound(SourceArray(i)) To UBound(SourceArray(i))
                    If EnsureScalar Then
                        LetOrSetElement result(i, j), _
                            GetScalarRepresentation(SourceArray(i)(j))
                    Else
                        LetOrSetElement result(i, j), _
                            SourceArray(i)(j)
                    End If
                Next
            Else
                If EnsureScalar Then
                    LetOrSetElement result(i, LBound(result, 2)), _
                        GetScalarRepresentation(SourceArray(i))
                Else
                    LetOrSetElement result(i, LBound(result, 2)), _
                        SourceArray(i)
                End If
            End If
        Next
    ElseIf LocalDepth > 2 And LocalDepth <= 20 Then
        result = RecursiveJaggedToMulti(SourceArray, LocalDepth, EnsureScalar:=EnsureScalar)
    Else
        RaiseError EC_MAX_DIMENSIONS_LIMIT, "JaggedToMulti()", "sourceArray()"
        result = SourceArray
    End If
    JaggedToMulti = result
End Function

Private Function MultiToJagged( _
    ByRef SourceArray As Variant, _
    Optional ByVal Depth As Long _
) As Variant()
    Dim LocalDepth As Long
    Dim result() As Variant
    
    LocalDepth = Depth
    
    If LocalDepth = 0 Then
        If Not IsMultidimensionalArray(SourceArray) Then
            RaiseError EC_EXPECTED_MULTIDIMENSION_ARRAY, "MultiToJagged", "sourceArray()"
            result = SourceArray
            Exit Function
        End If
        LocalDepth = GetMultidimensionalArrayDepth(SourceArray)
    End If
    
    If LocalDepth <= 1 Then
        result = SourceArray
    ElseIf LocalDepth = 2 Then
        Dim FirstDimBounds(0 To 1) As Long
        Dim SecondDimBounds(0 To 1) As Long
        Dim FirstDimOffset As Long
        Dim SecondDimOffset As Long
        Dim Nested() As Variant
        Dim i As Long
        Dim j As Long
        
        FirstDimBounds(0) = LBound(SourceArray, 1)
        FirstDimBounds(1) = UBound(SourceArray, 1)
        SecondDimBounds(0) = LBound(SourceArray, 2)
        SecondDimBounds(1) = UBound(SourceArray, 2)
        FirstDimOffset = This.LowerBound - FirstDimBounds(0)
        SecondDimOffset = This.LowerBound - SecondDimBounds(0)
        ReDim result(FirstDimBounds(0) + FirstDimOffset To FirstDimBounds(1) + FirstDimOffset)
        For i = FirstDimBounds(0) To FirstDimBounds(1)
            ReDim Nested(SecondDimBounds(0) + SecondDimOffset To SecondDimBounds(1) + SecondDimOffset)
            For j = SecondDimBounds(0) To SecondDimBounds(1)
                LetOrSetElement Nested(j + SecondDimOffset), SourceArray(i, j)
            Next
            result(i + FirstDimOffset) = Nested
        Next
    ElseIf LocalDepth > 2 And LocalDepth <= 20 Then
        result = RecursiveMultiToJagged(SourceArray, LocalDepth)
    Else
        RaiseError EC_MAX_DIMENSIONS_LIMIT, "MultiToJagged()", "sourceArray()"
        result = SourceArray
    End If
    
    MultiToJagged = result
End Function

'@Description("Recursively Converts multidimension arrays to jagged arrays. For md arrays with more than 2d.")
Private Function RecursiveMultiToJagged( _
        ByRef SourceArray As Variant, _
        Optional ByVal Depth As Long, _
        Optional ByVal CurrentDepth As Long, _
        Optional ByRef Crumbs As Variant _
    ) As Variant()
Attribute RecursiveMultiToJagged.VB_Description = "Recursively Converts multidimension arrays to jagged arrays. For md arrays with more than 2d."
    Dim i As Long
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim LocalDepth As Long
    Dim LocalCurrentDepth As Long
    Dim LocalCrumbs() As Variant
    Dim result() As Variant
    Dim BoundOffset As Long
    
    LocalDepth = Depth
    LocalCurrentDepth = CurrentDepth
    
    If LocalDepth > 1 Then
        If IsMissing(Crumbs) Then
            ReDim LocalCrumbs(LocalDepth - 1)
        Else
            LocalCrumbs = Crumbs
        End If
        inc LocalCurrentDepth
        LocalLowerBound = LBound(SourceArray, LocalCurrentDepth)
        LocalUpperBound = UBound(SourceArray, LocalCurrentDepth)
        BoundOffset = This.LowerBound - LocalLowerBound
        ReDim result(LocalLowerBound + BoundOffset To LocalUpperBound + BoundOffset)
        For i = LocalLowerBound To LocalUpperBound
            LocalCrumbs(LocalCurrentDepth - 1) = i
            If LocalCurrentDepth = LocalDepth Then
                LetOrSetElement result(i + BoundOffset), GetElementByBreadcrumb(SourceArray, LocalCrumbs)
            Else
                result(i + BoundOffset) = RecursiveMultiToJagged(SourceArray, LocalDepth, LocalCurrentDepth, LocalCrumbs)
            End If
        Next
    Else
        result = SourceArray
    End If
    RecursiveMultiToJagged = result
End Function


'@Description("Recursively Converts jagged arrays to multidimension arrays. For jagged arrays with more than 2d.")
Private Function RecursiveJaggedToMulti( _
        ByRef SourceArray() As Variant, _
        Optional ByVal Depth As Long, _
        Optional ByVal CurrentDepth As Long, _
        Optional ByRef Crumbs As Variant, _
        Optional ByRef result As Variant, _
        Optional ByVal EnsureScalar As Boolean _
    ) As Variant()
Attribute RecursiveJaggedToMulti.VB_Description = "Recursively Converts jagged arrays to multidimension arrays. For jagged arrays with more than 2d."
    Dim i As Long
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim LocalDepth As Long
    Dim LocalCurrentDepth As Long
    Dim LocalCrumbs() As Variant
    Dim LocalResult() As Variant


    LocalDepth = Depth
    LocalCurrentDepth = CurrentDepth

    If IsMissing(Crumbs) Then
        ReDim LocalCrumbs(LocalDepth - 1)
    Else
        LocalCrumbs = Crumbs
    End If

    If IsMissing(result) Then
        LocalResult = CreateMultidimensionalArray(MapJaggedArray(SourceArray, KnownDepth:=LocalDepth))
    Else
        LocalResult = result
    End If

    inc LocalCurrentDepth
    LocalLowerBound = LBound(SourceArray)
    LocalUpperBound = UBound(SourceArray)
    For i = LocalLowerBound To LocalUpperBound
        LocalCrumbs(LocalCurrentDepth - 1) = i
        If LocalCurrentDepth = LocalDepth Then
            If EnsureScalar Then
                LocalResult = LetElementByBreadcrumb(LocalResult, LocalCrumbs, GetScalarRepresentation(SourceArray(i)))
            Else
                LocalResult = LetElementByBreadcrumb(LocalResult, LocalCrumbs, SourceArray(i))
            End If
        Else
            Dim Nested() As Variant
            If Not IsArray(SourceArray(i)) Then
                Nested = Array(SourceArray(i))
            Else
                Nested = SourceArray(i)
            End If
            LocalResult = RecursiveJaggedToMulti( _
                SourceArray:=Nested, _
                Depth:=LocalDepth, _
                CurrentDepth:=LocalCurrentDepth, _
                Crumbs:=LocalCrumbs, _
                result:=LocalResult, _
                EnsureScalar:=EnsureScalar _
            )
        End If
    Next
    RecursiveJaggedToMulti = LocalResult
End Function

'@Description("Returns a number representing how many levels of nested arrays are contained in array.")
Private Function GetJaggedArrayDepth(ByRef SourceArray() As Variant) As Long
Attribute GetJaggedArrayDepth.VB_Description = "Returns a number representing how many levels of nested arrays are contained in array."
    Dim i As Long
    Dim LocalLowerBound As Long
    Dim LocalUpperBound As Long
    Dim CurrentDepth As Long
    Dim MaxDepth As Long
    Dim Depth As Long
    
    If IsArray(SourceArray) Then
        inc Depth
        LocalLowerBound = LBound(SourceArray)
        LocalUpperBound = UBound(SourceArray)
        For i = LocalLowerBound To LocalUpperBound
            If IsArray(SourceArray(i)) Then
                Dim Nested() As Variant
                Nested = SourceArray(i)
                CurrentDepth = GetJaggedArrayDepth(Nested)
                If CurrentDepth > MaxDepth Then MaxDepth = CurrentDepth
            End If
        Next
        inc Depth, MaxDepth
    End If
    GetJaggedArrayDepth = Depth
End Function

'@Description("Gets the maximium bound values for a specified dimension of a jagged array.")
Private Function GetMaxBoundsAtDimension( _
        ByRef SourceArray() As Variant, _
        Optional ByVal Dimension As Long = 1, _
        Optional ByVal CurrentDimension As Long = 1 _
    ) As Long()
Attribute GetMaxBoundsAtDimension.VB_Description = "Gets the maximium bound values for a specified dimension of a jagged array."
    Dim MaxBounds() As Long
    
    If CurrentDimension < Dimension Then
        Dim i As Long
        Dim CurrentBounds() As Long
        Dim LocalLowerBound As Long
        Dim LocalUpperBound As Long
        
        LocalLowerBound = LBound(SourceArray)
        LocalUpperBound = UBound(SourceArray)
        For i = LocalLowerBound To LocalUpperBound
            If IsArray(SourceArray(i)) Then
                Dim Nested() As Variant
                Nested = SourceArray(i)
                CurrentBounds = GetMaxBoundsAtDimension(Nested, Dimension, CurrentDimension + 1)
                If Not IsArrayAllocated(MaxBounds) Then
                    MaxBounds = CurrentBounds
                Else
                    If CurrentBounds(0) < MaxBounds(0) Then MaxBounds(0) = CurrentBounds(0)
                    If CurrentBounds(1) > MaxBounds(1) Then MaxBounds(1) = CurrentBounds(1)
                End If
            End If
        Next
    Else
        MaxBounds = GetArrayBounds(SourceArray)
    End If
    
    GetMaxBoundsAtDimension = MaxBounds
End Function

'@Description("Returns an array representing the structure of a jagged array")
Private Function MapJaggedArray( _
        ByRef SourceArray() As Variant, _
        Optional ByVal KnownDepth As Long _
    ) As Variant()
Attribute MapJaggedArray.VB_Description = "Returns an array representing the structure of a jagged array"
    Dim i As Long
    Dim LocalMap() As Variant
    Dim LocalDepth As Long

    If Not IsArray(SourceArray) Then Exit Function

    If KnownDepth > 0 Then
        LocalDepth = KnownDepth
    Else
        LocalDepth = GetJaggedArrayDepth(SourceArray)
    End If

    ' crumbs are base 0
    ReDim LocalMap(0 To LocalDepth - 1)

    For i = LBound(LocalMap) To UBound(LocalMap)
        LocalMap(i) = GetMaxBoundsAtDimension(SourceArray, i + 1)
    Next

    MapJaggedArray = LocalMap
End Function


'@Description("Returns a 0-based array with two slots. First slot is the lowerbound of the passed array. second slot (index 0) is the upper bound.")
Private Function GetArrayBounds(ByVal SourceArray As Variant) As Long()
Attribute GetArrayBounds.VB_Description = "Returns a 0-based array with two slots. First slot is the lowerbound of the passed array. second slot (index 0) is the upper bound."
    Dim result(0 To 1) As Long
    If IsArray(SourceArray) Then
        result(0) = LBound(SourceArray)
        result(1) = UBound(SourceArray)
    End If
    GetArrayBounds = result
End Function

'@Description("Returns a number representing the total number of dimensions of a multidimension array.")
Private Function GetMultidimensionalArrayDepth(ByVal SourceArray As Variant) As Long
Attribute GetMultidimensionalArrayDepth.VB_Description = "Returns a number representing the total number of dimensions of a multidimension array."
    Dim i As Long
    '@Ignore VariableNotUsed
    Dim Void As Long
    
    On Error Resume Next
    Do
        inc i
        Void = UBound(SourceArray, i)
    Loop Until err.Number <> 0
    err.Clear
    On Error GoTo 0
    GetMultidimensionalArrayDepth = i - 1
End Function

'@Description("Returns the element stored in an array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot.")
Private Function GetElementByBreadcrumb( _
        ByVal SourceArray As Variant, _
        ByRef Crumb() As Variant _
    ) As Variant
Attribute GetElementByBreadcrumb.VB_Description = "Returns the element stored in an array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot."
    Dim result As Variant
    Select Case UBound(Crumb)
    Case 0
        LetOrSetElement result, SourceArray(Crumb(0))
    Case 1
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1))
    Case 2
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2))
    Case 3
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3))
    Case 4
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4))
    Case 5
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5))
    Case 6
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6))
    Case 7
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7))
    Case 8
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8))
    Case 9
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9))
    Case 10
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10))
    Case 11
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11))
    Case 12
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12))
    Case 13
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13))
    Case 14
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14))
    Case 15
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15))
    Case 16
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16))
    Case 17
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17))
    Case 18
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18))
    Case 19
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18), Crumb(19))
    Case 20
        LetOrSetElement result, SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), _
            Crumb(9), Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18), Crumb(19), Crumb(20))
    End Select
    '@Ignore UnassignedVariableUsage
    GetElementByBreadcrumb = result
End Function

'@Description("Assigns an element to a location in a multidimensonal array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot.")
Private Function LetElementByBreadcrumb( _
        ByRef SourceArray() As Variant, _
        ByRef Crumb() As Variant, _
        ByVal Element As Variant _
    ) As Variant
Attribute LetElementByBreadcrumb.VB_Description = "Assigns an element to a location in a multidimensonal array based on a crumb argument. A crumb is an array representing the location of the multidimension array slot."
    Select Case UBound(Crumb)
    Case 0
        LetOrSetElement SourceArray(Crumb(0)), Element
    Case 1
        LetOrSetElement SourceArray(Crumb(0), Crumb(1)), Element
    Case 2
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2)), Element
    Case 3
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3)), Element
    Case 4
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4)), Element
    Case 5
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5)), Element
    Case 6
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6)), Element
    Case 7
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7)), Element
    Case 8
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8)), Element
    Case 9
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9)), _
            Element
    Case 10
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10)), Element
    Case 11
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11)), Element
    Case 12
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12)), Element
    Case 13
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12), Crumb(13)), Element
    Case 14
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14)), Element
    Case 15
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15)), Element
    Case 16
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
        Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16)), Element
    Case 17
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17)), Element
    Case 18
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
            Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18)), Element
    Case 19
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
        Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18), Crumb(19)), Element
    Case 20
        LetOrSetElement SourceArray(Crumb(0), Crumb(1), Crumb(2), Crumb(3), Crumb(4), Crumb(5), Crumb(6), Crumb(7), Crumb(8), Crumb(9), _
        Crumb(10), Crumb(11), Crumb(12), Crumb(13), Crumb(14), Crumb(15), Crumb(16), Crumb(17), Crumb(18), Crumb(19), Crumb(20)), _
        Element
    End Select

    LetElementByBreadcrumb = SourceArray
End Function

'@Description("Returns an empty multidimensional array with the dimensions specified in the crumb argument. A crumb is an array representing the structure of the multidimension array.")
Private Function CreateMultidimensionalArray(ByRef Crumb() As Variant) As Variant()
Attribute CreateMultidimensionalArray.VB_Description = "Returns an empty multidimensional array with the dimensions specified in the crumb argument. A crumb is an array representing the structure of the multidimension array."
    Dim result() As Variant
    Select Case UBound(Crumb)
    Case 0
        ReDim result(Crumb(0)(0) To Crumb(0)(1))
    Case 1
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1))
    Case 2
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1))
    Case 3
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1))
    Case 4
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1))
    Case 5
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1))
    Case 6
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1))
    Case 7
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1))
    Case 8
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1))
    Case 9
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1))
    Case 10
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1))
    Case 11
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1))
    Case 12
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1))
    Case 13
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1))
    Case 14
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1))
    Case 15
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1))
    Case 16
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1), _
                     Crumb(16)(0) To Crumb(16)(1))
    Case 17
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1), _
                     Crumb(16)(0) To Crumb(16)(1), Crumb(17)(0) To Crumb(17)(1))
    Case 18
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1), _
                     Crumb(16)(0) To Crumb(16)(1), Crumb(17)(0) To Crumb(17)(1), Crumb(18)(0) To Crumb(18)(1))
    Case 19
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1), _
                     Crumb(16)(0) To Crumb(16)(1), Crumb(17)(0) To Crumb(17)(1), Crumb(18)(0) To Crumb(18)(1), Crumb(19)(0) To Crumb(19)(1))
    Case 20
        ReDim result(Crumb(0)(0) To Crumb(0)(1), Crumb(1)(0) To Crumb(1)(1), Crumb(2)(0) To Crumb(2)(1), Crumb(3)(0) To Crumb(3)(1), _
                     Crumb(4)(0) To Crumb(4)(1), Crumb(5)(0) To Crumb(5)(1), Crumb(6)(0) To Crumb(6)(1), Crumb(7)(0) To Crumb(7)(1), _
                     Crumb(8)(0) To Crumb(8)(1), Crumb(9)(0) To Crumb(9)(1), Crumb(10)(0) To Crumb(10)(1), Crumb(11)(0) To Crumb(11)(1), _
                     Crumb(12)(0) To Crumb(12)(1), Crumb(13)(0) To Crumb(13)(1), Crumb(14)(0) To Crumb(14)(1), Crumb(15)(0) To Crumb(15)(1), _
                     Crumb(16)(0) To Crumb(16)(1), Crumb(17)(0) To Crumb(17)(1), Crumb(18)(0) To Crumb(18)(1), Crumb(19)(0) To Crumb(19)(1), _
                     Crumb(20)(0) To Crumb(20)(1))
    End Select
    CreateMultidimensionalArray = result
End Function
