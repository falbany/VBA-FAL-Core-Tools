Attribute VB_Name = "demoMDM"

Option Explicit

Public Sub RunAdvancedMdmDemo()
    '@brief Demonstrates the advanced creation, manipulation, and export capabilities of the refactored clsMDM class.

    ' --- 1. Create and Populate an MDM Object with Structured Parameters ---
    Dim mdm As New clsMDM
    Debug.Print "--- Step 1: Creating and populating a new clsMDM object with structured parameters ---"

    ' Add User Inputs and Values (these are simple key-value pairs)
    mdm.AddUserInput "Temp", "25"
    mdm.AddIccapValue "L", "1e-5"

    ' Create and add a structured ICCAP Input Parameter
    Dim vdInput As New clsMdmInputParameter
    vdInput.Name = "Vd"
    vdInput.Mode = "V"
    vdInput.ModeOptions("PlusNode") = "D"
    vdInput.ModeOptions("MinusNode") = "S"
    vdInput.ModeOptions("Unit") = "V"
    vdInput.ModeOptions("Compliance") = "0.1"
    vdInput.SweepType = "LIN"
    vdInput.SweepOptions("SweepOrder") = 1
    vdInput.SweepOptions("Start") = 0
    vdInput.SweepOptions("Stop") = 1
    vdInput.SweepOptions("NumPoints") = 3
    vdInput.SweepOptions("StepSize") = 0.5
    mdm.AddIccapInput vdInput

    ' Create and add a structured ICCAP Output Parameter
    Dim idOutput As New clsMdmOutputParameter
    idOutput.Name = "Id"
    idOutput.Mode = "I"
    idOutput.ModeOptions("ToNode") = "D"
    idOutput.ModeOptions("FromNode") = "S"
    idOutput.Unit = "A"
    idOutput.Compliance = "DEFAULT"
    idOutput.Type = "M"
    mdm.AddIccapOutput idOutput

    ' Add a data block using the new structured method
    Dim db1 As clsMdmDataBlock
    Set db1 = mdm.AddDataBlock()
    db1.AddInputValue "Vg", "1.0"

    ' Add headers
    db1.AddHeader "Vd"
    db1.AddHeader "Id"

    ' Add data rows using the new, cell-by-cell AddValue method
    db1.AddValue "Vd", 0, 0
    db1.AddValue "Id", 0, 0.001

    db1.AddValue "Vd", 1, 0.5
    db1.AddValue "Id", 1, 0.05

    db1.AddValue "Vd", 2, 1
    db1.AddValue "Id", 2, 0.1

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
    Debug.Print "Parsing successful."
    Debug.Print ""

    ' --- 3. Verify Structured Data ---
    Debug.Print "--- Step 3: Verifying structured data from parsed object ---"
    Dim parsedInput As clsMdmInputParameter
    Set parsedInput = parsedMdm.IccapInputs("Vd")
    Debug.Print "Parsed Input 'Vd' Sweep Type: " & parsedInput.SweepType
    Debug.Print "Parsed Input 'Vd' Start Value: " & parsedInput.SweepOptions("Start")

    Dim parsedOutput As clsMdmOutputParameter
    Set parsedOutput = parsedMdm.IccapOutputs("Id")
    Debug.Print "Parsed Output 'Id' ToNode: " & parsedOutput.ModeOptions("ToNode")
    Debug.Print ""

    ' --- 4. Export to 2D Array and Paste to Excel ---
    Debug.Print "--- Step 4: Exporting data to a 2D Array ---"
    Dim dataArray As Variant
    dataArray = parsedMdm.To2DArray()

    If Not IsEmpty(dataArray) Then
        Dim wks As Worksheet
        Set wks = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wks.Name = "MDM_Advanced_Demo"
        wks.Range("A1").Resize(UBound(dataArray, 1), UBound(dataArray, 2)).value = dataArray
        wks.Columns.AutoFit
        Debug.Print "Data successfully exported to worksheet '" & wks.Name & "'."
        wks.Activate
    Else
        Debug.Print "Failed to export data to 2D array."
    End If
    Debug.Print ""

    ' --- 5. Export to JSON ---
    Debug.Print "--- Step 5: Exporting data to a JSON string ---"
    Dim jsonString As String
    jsonString = parsedMdm.ToJSON(PrettyPrint:=True)
    Debug.Print "Generated JSON String:"
    Debug.Print jsonString

    Debug.Print "--- Demo Complete ---"

End Sub