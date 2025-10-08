Attribute VB_Name = "demoMDM"

Option Explicit

Public Sub RunAdvancedMdmDemo()
    '@brief Demonstrates the advanced creation, manipulation, and export capabilities of the refactored clsMDM class.

    ' --- 1. Create and Populate an MDM Object with Complex Numbers ---
    Dim mdm As New clsMDM
    Debug.Print "--- Step 1: Creating and populating a new clsMDM object with complex numbers ---"

    ' Add a structured ICCAP Input Parameter
    Dim freqInput As New clsMdmInputParameter
    freqInput.Name = "Freq"
    freqInput.SweepType = "LOG"
    freqInput.SweepOptions("Start") = "1e3"
    freqInput.SweepOptions("Stop") = "1e6"
    freqInput.SweepOptions("NumPoints") = 3
    mdm.AddIccapInput freqInput

    ' Add a structured ICCAP Output Parameter for a complex number
    Dim zOutput As New clsMdmOutputParameter
    zOutput.Name = "Z"
    zOutput.Type = "C" ' C for Complex
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

    ' Add a sweep definition that won't match the data
    Dim vIn As New clsMdmInputParameter
    vIn.Name = "Vd"
    vIn.SweepType = "LIN"
    vIn.SweepOptions("NumPoints") = 5 ' Mismatch: We will only add 2 points
    vIn.SweepOptions("Start") = 0
    vIn.SweepOptions("Stop") = 1
    invalidMdm.AddIccapInput vIn

    ' Add a valid output parameter
    Dim iOut As New clsMdmOutputParameter
    iOut.Name = "Id"
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
        Dim errKey As Variant
        For Each errKey In invalidMdm.ValidationErrors.Keys
            Debug.Print "  - ERROR: " & invalidMdm.ValidationErrors(errKey)
        Next errKey
    Else
        Debug.Print "Result: VALID. Validation failed to detect errors."
    End If
    Debug.Print ""

    Debug.Print "--- Demo Complete ---"

End Sub