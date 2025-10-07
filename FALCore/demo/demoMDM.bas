Attribute VB_Name = "demoMDM"

Option Explicit

Public Sub RunMdmDemo()
    '@brief Demonstrates the creation, manipulation, and export capabilities of the clsMDM class.

    ' --- 1. Create and Populate an MDM Object ---
    Dim mdm As New clsMDM
    Debug.Print "--- Step 1: Creating and populating a new clsMDM object ---"

    ' Add header information
    mdm.AddUserInput "Temp", "LIN 1 25 25 1"
    mdm.AddIccapInput "Vd", "V COMMON 1 0 V 0"
    mdm.AddIccapInput "Vg", "V COMMON 2 0 V 0"
    mdm.AddIccapOutput "Id", "I 1 0 A M"
    mdm.AddIccapValue "L", "1e-5"
    mdm.AddIccapValue "W", "2e-5"

    ' Add first data block
    Dim db1 As clsMdmDataBlock
    Set db1 = mdm.AddDataBlock()
    db1.AddInputValue "Temp", 25
    db1.AddInputValue "Vg", 1

    Dim data1(1 To 4, 1 To 2) As Variant
    data1(1, 1) = "Vd":   data1(1, 2) = "Id"
    data1(2, 1) = 0:      data1(2, 2) = 0
    data1(3, 1) = 0.5:    data1(3, 2) = 0.05
    data1(4, 1) = 1:      data1(4, 2) = 0.1
    db1.Data = data1

    ' Add second data block
    Dim db2 As clsMdmDataBlock
    Set db2 = mdm.AddDataBlock()
    db2.AddInputValue "Temp", 25
    db2.AddInputValue "Vg", 2

    Dim data2(1 To 4, 1 To 2) As Variant
    data2(1, 1) = "Vd":   data2(1, 2) = "Id"
    data2(2, 1) = 0:      data2(2, 2) = 0
    data2(3, 1) = 0.5:    data2(3, 2) = 0.1
    data2(4, 1) = 1:      data2(4, 2) = 0.2
    db2.Data = data2

    Debug.Print "Population complete. MDM object has " & mdm.DataBlocks.Count & " data blocks."
    Debug.Print ""

    ' --- 2. Generate MDM String ---
    Debug.Print "--- Step 2: Generating MDM string output ---"
    Dim mdmString As String
    mdmString = mdm.ToMdmString()
    Debug.Print "Generated MDM String:"
    Debug.Print mdmString
    Debug.Print ""

    ' --- 3. Parse the Generated String ---
    Debug.Print "--- Step 3: Parsing the generated MDM string into a new object ---"
    Dim parsedMdm As New clsMDM
    If parsedMdm.ParseMdmString(mdmString) Then
        Debug.Print "Parsing successful."
        Debug.Print "Parsed MDM object has " & parsedMdm.IccapInputs.Count & " ICCAP inputs."
        Debug.Print "Parsed MDM object has " & parsedMdm.DataBlocks.Count & " data blocks."
    Else
        Debug.Print "Parsing failed."
        Exit Sub
    End If
    Debug.Print ""

    ' --- 4. Export to 2D Array and Paste to Excel ---
    Debug.Print "--- Step 4: Exporting data to a 2D Array and pasting to a new worksheet ---"
    Dim dataArray As Variant
    dataArray = parsedMdm.To2DArray()

    If Not IsEmpty(dataArray) Then
        On Error Resume Next
        Dim wks As Worksheet
        Set wks = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Err.Number <> 0 Then
            Set wks = ThisWorkbook.Worksheets.Add
        End If
        On Error GoTo 0
        wks.Name = "MDM_Demo_Output"

        ' Paste the data
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