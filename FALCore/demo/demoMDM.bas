Attribute VB_Name = "demoMDM"

Option Explicit

Public Sub RunAdvancedMdmDemo()
    '@brief Demonstrates the advanced creation, manipulation, and export capabilities of the refactored clsMDM class.

    ' --- 1. Create and Populate an MDM Object with Complex Numbers ---
    Dim mdm As New clsMDM
    Debug.Print "--- Step 1: Creating and populating a new clsMDM object with complex numbers ---"

    ' Add a structured ICCAP Input Parameter using new properties and enums
    Dim freqInput As New clsMdmInputParameter
    freqInput.Name = "Freq"
    freqInput.Mode = mdmModeF ' Frequency mode
    freqInput.SweepType = mdmSweepLOG
    freqInput.SweepOrder = 1
    freqInput.Start = 1000
    freqInput.Stop = 1000000
    freqInput.NumPoints = 3
    mdm.AddIccapInput freqInput

    ' Add a structured ICCAP Output Parameter for a complex number
    Dim zOutput As New clsMdmOutputParameter
    zOutput.Name = "Z"
    zOutput.Mode = mdmTypeZ ' Z-parameters are complex
    zOutput.Type = "B" ' Both measured and simulated
    zOutput.Unit = "Ohm"
    mdm.AddIccapOutput zOutput

    ' Add a data block
    Dim db1 As clsMdmDataBlock
    Set db1 = mdm.AddDataBlock()

    ' Add data, including complex number columns R:Z and I:Z
    db1.AddValue "Freq", 0, "1000"
    db1.AddValue "R:Z", 0, "100"
    db1.AddValue "I:Z", 0, "-50"

    db1.AddValue "Freq", 1, "10000"
    db1.AddValue "R:Z", 1, "95"
    db1.AddValue "I:Z", 1, "-150"

    db1.AddValue "Freq", 2, "1000000"
    db1.AddValue "R:Z", 2, "20"
    db1.AddValue "I:Z", 2, "-300"

    Debug.Print "Population complete."
    Debug.Print ""

    ' --- 2. Generate and Parse MDM String ---
    Debug.Print "--- Step 2: Generating and parsing MDM string ---"
    Dim mdmString As String
    mdmString = mdm.ToMdmString()
    Debug.Print "Generated MDM String:"
    Debug.Print mdmString

    Dim parsedMdm As New clsMDM
    parsedMdm.ParseMdmString mdmString
    Debug.Print "Parsing successful. Note that complex numbers are now represented by a clsComplexNumber object."
    Debug.Print ""

    ' --- 3. Convert Data Types ---
    Debug.Print "--- Step 3: Converting string data to numeric types ---"
    ' Before conversion, the data is stored as strings (as it is in the file)
    Debug.Print "Type of Freq value before conversion: " & TypeName(parsedMdm.DataBlocks(0).Data("Freq")(0))

    parsedMdm.ConvertDataTypes

    ' After conversion, strings representing numbers are converted to actual numeric types
    Debug.Print "Type of Freq value after conversion: " & TypeName(parsedMdm.DataBlocks(0).Data("Freq")(0))
    Debug.Print "Complex number Z at row 0: " & parsedMdm.DataBlocks(0).Data("Z")(0).ToString()
    Debug.Print ""

    ' --- 4. Export Directly to a New Worksheet ---
    Debug.Print "--- Step 4: Exporting data directly to a new Excel worksheet ---"

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("MDM_Advanced_Demo").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wks.Name = "MDM_Advanced_Demo"

    ' Use the new, simplified export method
    parsedMdm.ExportToWorksheet wks, "A1"

    wks.Columns.AutoFit
    Debug.Print "Data successfully exported to worksheet '" & wks.Name & "'."
    wks.Activate
    Debug.Print ""

    ' --- 5. Export to JSON ---
    Debug.Print "--- Step 5: Exporting data to a JSON string ---"
    Dim jsonString As String
    jsonString = parsedMdm.ToJSON(PrettyPrint:=True)
    Debug.Print "Generated JSON String (note the complex number objects):"
    Debug.Print jsonString
    Debug.Print ""

    ' --- 6. Validate Data Integrity ---
    Debug.Print "--- Step 6: Demonstrating the Validation Framework ---"

    ' a) Validate the existing, valid MDM object
    Debug.Print "Validating the parsed MDM object..."
    If parsedMdm.Validate() Then
        Debug.Print "Result: VALID. No errors found."
    Else
        Debug.Print "Result: INVALID. This was not expected."
    End If
    Debug.Print ""

    ' b) Create an intentionally invalid MDM object to demonstrate error detection
    Debug.Print "Creating an invalid MDM object to test validation..."
    Dim invalidMdm As New clsMDM

    ' Add a sweep definition that won't match the data, using new properties and enums
    Dim vIn As New clsMdmInputParameter
    vIn.Name = "Vd"
    vIn.Mode = mdmModeV
    vIn.SweepType = mdmSweepLIN
    vIn.NumPoints = 5 ' Mismatch: We will only add 2 points
    vIn.Start = 0
    vIn.Stop = 1
    invalidMdm.AddIccapInput vIn

    ' Add a valid output parameter
    Dim iOut As New clsMdmOutputParameter
    iOut.Name = "Id"
    iOut.Mode = mdmTypeI
    iOut.Type = "M"
    invalidMdm.AddIccapOutput iOut

    ' Add a data block with errors
    Dim invalidDb As clsMdmDataBlock
    Set invalidDb = invalidMdm.AddDataBlock()

    ' Error 1: Row count mismatch. 'Vd' will have 2 rows, 'Id' will have 1.
    invalidDb.AddValue "Vd", 0, 0
    invalidDb.AddValue "Vd", 1, 1

    invalidDb.AddValue "Id", 0, 0.1

    ' Error 2: Undefined header. 'I_leak' is not in ICCAP_OUTPUTS.
    invalidDb.AddValue "I_leak", 0, 0.001

    Debug.Print "Validating the invalid MDM object..."
    If Not invalidMdm.Validate() Then
        Debug.Print "Result: INVALID. Errors were correctly detected:"
        ' Iterate through the new collection of structured error objects
        Dim err As clsMdmValidationError
        For Each err In invalidMdm.ValidationErrors
            Debug.Print "  - " & err.ToString()
        Next err
    Else
        Debug.Print "Result: VALID. Validation failed to detect errors."
    End If
    Debug.Print ""

    Debug.Print "--- Demo Complete ---"

End Sub

Public Sub RunFluentInterfaceDemo()
    '@brief Demonstrates the new fluent interface and simplified data accessors.

    ' --- 1. Create and Populate an MDM Object using Method Chaining ---
    Debug.Print "--- Step 1: Creating a new clsMDM object using the fluent interface ---"

    ' Create parameter definitions using new strongly-typed properties and enums
    Dim vIn As New clsMdmInputParameter
    vIn.Name = "Vd"
    vIn.Mode = mdmModeV
    vIn.SweepType = mdmSweepLIN
    vIn.SweepOrder = 1
    vIn.NumPoints = 3
    vIn.Start = 0
    vIn.Stop = 1

    Dim iOut As New clsMdmOutputParameter
    iOut.Name = "Id"
    iOut.Mode = mdmTypeI
    iOut.Type = "M"

    ' Use method chaining to build the object and its data block in a more concise way
    Dim mdm As New clsMDM
    mdm.AddUserInput "Device", "MOSFET_A1" _
       .AddIccapInput vIn _
       .AddIccapOutput iOut

    mdm.AddDataBlock() _
       .AddInputValue "Temp", "25" _
       .AddValue "Vd", 0, 0 _
       .AddValue "Id", 0, 0.01 _
       .AddValue "Vd", 1, 0.5 _
       .AddValue "Id", 1, 0.52 _
       .AddValue "Vd", 2, 1 _
       .AddValue "Id", 2, 1.05

    Debug.Print "MDM object created successfully using method chaining."
    Debug.Print "Generated MDM String:"
    Debug.Print mdm.ToMdmString()
    Debug.Print ""

    ' --- 2. Demonstrate Simplified Data Accessors ---
    Debug.Print "--- Step 2: Using simplified accessors to retrieve data ---"

    ' a) Get a single value
    Dim singleValue As Variant
    singleValue = mdm.GetValue(blockIndex:=0, header:="Id", rowIndex:=1)
    Debug.Print "Retrieved single value for Id at rowIndex 1: " & singleValue
    Debug.Print ""

    ' b) Get an entire column
    Dim columnData As Variant
    columnData = mdm.GetColumn(blockIndex:=0, header:="Vd")

    Debug.Print "Retrieved the entire 'Vd' column:"
    If IsArray(columnData) Then
        Dim i As Long
        For i = LBound(columnData) To UBound(columnData)
            Debug.Print "  Vd(" & i & ") = " & columnData(i)
        Next i
    Else
        Debug.Print "  Failed to retrieve column data as an array."
    End If
    Debug.Print ""

    ' c) Demonstrate error handling for accessors
    Dim badValue As Variant
    badValue = mdm.GetValue(blockIndex:=0, header:="Vd", rowIndex:=99) ' Invalid row
    Debug.Print "Attempting to get value at invalid row index 99..."
    If IsError(badValue) Then
        Debug.Print "  Correctly returned an error: " & CStr(badValue)
    Else
        Debug.Print "  Did not return an error as expected."
    End If
    Debug.Print ""

    Debug.Print "--- Fluent Interface Demo Complete ---"

End Sub