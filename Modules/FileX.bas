Attribute VB_Name = "FileX"
' **************************************************************************************
' Module    : FileX
' Author    : Forent ALBANY
' Website   :
' Purpose   : Manipulation of Files
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2023-07-00              Initial Release
'---------------------------------------------------------------------------------------
' Revision Propositions:
' ~~~~~~~~~~~~~~~~
' Dependencies:
' ~~~~~~~~~~~~~~~~
'   GLOBAL_VAR_MOD
'   LANG_MOD
'   EXCEL_MOD
' **************************************************************************************


'Option Explicit

Private Sub HandleError(ByVal ErrMsg As String)
    ' @brief Affiche un message d'erreur standardis� � l'utilisateur.
    ' @param ErrMsg Le message d'erreur sp�cifique � afficher.
    MsgBox "Une erreur est survenue :" & vbCrLf & ErrMsg, vbOKOnly + vbCritical, "Op�ration impossible"
End Sub


Public Function Logger(sFile As String, Optional sType As String = "", Optional sSource As String = "", Optional sDetails As String = "", Optional maxFileSize As Long = 20000) As Boolean
    ' * @brief Logs events to a specified file with optional archiving based on file size.
    ' * @param sFile The file path to log events to.
    ' * @param sType (Optional) The type of the event (e.g., INFO, WARNING, ERROR).
    ' * @param sSource (Optional) The source of the event (e.g., module, function).
    ' * @param sDetails (Optional) Additional details about the event.
    ' * @param maxFileSize (Optional) The maximum file size before archiving (default is 20,000 bytes).
    ' *        If set to 0 or not provided, no archiving based on file size will occur.
    ' * @return Boolean indicating if the logging operation was successful.
    ' *
    ' * @details
    ' * Usage Example:
    ' * @code{vba}
    ' * Logger "C:\Path\To\Log.txt", "INFO", "Module1", "Event details", 50000
    ' * @endcode
    ' *
    ' * This will log an INFO event with the specified details to the file "C:\Path\To\Log.txt".
    ' * If the file size exceeds 50,000 bytes, it will be archived.
    ' */
    Dim stext       As String
    Dim sExtension  As String
    
    sExtension = "." & FileX.Get_FileExtension(sFile)

    ' Archive file at certain size
    If maxFileSize <> 0 Then
        If FileX.FileExist(sFile) Then
            If FileLen(sFile) > maxFileSize Then
                FileCopy sFile, Replace(sFile, sExtension, format(Now, "_ddmmyyyy-hhmmss") & sExtension)
                Kill sFile
            End If
        End If
    End If
    
    stext = "[" & format(Now, "dd-mm-yyyy hh:mm:ss") & "]" & vbTab & _
            IIf(sType <> "", sType & vbTab, "") & _
            IIf(sSource <> "", sSource & vbTab, "") & _
            IIf(sDetails <> "", sDetails & vbTab, "") & _
            "(user: " & Application.UserName & vbTab & _
            "version: " & GLOBAL_VAR_MOD.RELEASE & ")"

    Logger = FileX.AppendTxt(sFile, stext, True)

End Function

Function Get_SaveFilePath_WithDialog(Optional InitialFileName As String = "MyFile", Optional InitialFileExtension As String = ".txt", Optional FileFilter As String = "Text Files (*.txt), *.xlsx") As String
    Dim filePath As Variant

    ' Prompt user to choose a file path and name
    filePath = Application.GetSaveAsFilename( _
                InitialFileName:=InitialFileName & InitialFileExtension, _
                FileFilter:=FileFilter)

    ' Check if a file path was selected
    If filePath <> False Then Get_SaveFilePath_WithDialog = filePath Else Get_SaveFilePath_WithDialog = ""

End Function

Public Function Select_Files(Optional FileType As String = "", Optional AllowMultiSelect As Boolean = True) As Variant
    '* @brief Displays a file selection dialog for the user to choose one or more files.
    '* @param FileType The type of files to filter (e.g., "xls", "xlsx", "mdm", "csv"). Leave blank for all files.
    '* @param AllowMultiSelect True if multiple file selection is allowed, False otherwise.
    '* @return An array of selected file paths.
    On Error GoTo ifError
    Dim i               As Integer
    Dim count           As Integer
    Dim filePaths()     As String
    
    ' Selection de fichiers
    Set File = Application.FileDialog(msoFileDialogFilePicker)
    With File
        .filters.Clear
        Select Case UCase(FileType)
            Case "":
            Case "XLS", "XLSX"
                .filters.Add "Excel files", "*.xlsx; *.xls; *.xlsm", 1
                .filters.Add "Excel Converted", "*.mdm.xlsx; *.mdm.xlsx", 2
            Case "MDM": .filters.Add "ICCAP", "*.mdm", 1
            Case "CSV": .filters.Add "CSV", "*.csv", 1
            Case Else: .filters.Add "Specified", Replace(FileType, ".", "*."), 1
        End Select
        .filters.Add "All files", "*.*"
        .title = "File Selection"
        .AllowMultiSelect = AllowMultiSelect
        .InitialView = msoFileDialogViewDetails
        .ButtonName = "Select"
        .Show
        count = File.SelectedItems.count
        ReDim filePaths(1 To count)
        For i = 1 To count
            filePaths(i) = File.SelectedItems(i)
        Next
    End With
    
    Select_Files = filePaths
    Exit Function
ifError:
    Select_Files = CVErr(2001)
End Function

Public Function Select_Folder(Optional AllowMultiSelect As Boolean = False) As String
    On Error GoTo ifError
    ' Folder selection.
    Set File = Application.FileDialog(msoFileDialogFolderPicker)
    With File
        .title = "Folder selection"
        .AllowMultiSelect = AllowMultiSelect
        .InitialView = msoFileDialogViewDetails
        .ButtonName = "Select"
        '.Show
        If File.Show Then
            Select_Folder = File.SelectedItems(1)
        Else
            Select_Folder = ""  ' if aborted.
        End If
    End With
    Exit Function
ifError:
    Select_Folder = ""
End Function

Public Function Select_AllFiles(Repertoire As String, FileType As String, FileTypeNot As String, Filter1 As String, Filter2 As String, Filter3 As String, Filter4 As String, Filter5 As String, Filter6 As String, ByRef FileList() As String, ByRef FolderList() As String, Optional ByRef nb_file As Long) As Boolean
    'On Error GoTo IfError
    ' R�pertorie tous les fichiers d'un r�pertoire avec possibilit� de filtrage.
    ' N�cessite d'activer la r�f�rence "Microsoft Scripting RunTime"
        'Dans l'�diteur de macros (Alt+F11):
        'Menu Outils
        'R�f�rences
        'Cochez la ligne "Microsoft Scripting RunTime".
        'Cliquez sur le bouton OK pour valider.
        
    ' Filter1 & Filter2 & Filter3 -> .And
    ' Filter4 & Filter5 & Filter6 -> .NAnd
    
    Dim Fso As Scripting.FileSystemObject
    Dim SourceFolder As Scripting.Folder
    Dim SubFolder As Scripting.Folder
    Dim FileItem As Scripting.File
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = Fso.GetFolder(Repertoire)
    
    ' Filtres : Mise en forme
    And1 = "*" & Filter1 & "*"
    And2 = "*" & Filter2 & "*"
    And3 = "*" & Filter3 & "*"
    ExtensionFile = "*" & FileType
    AndNot1 = "*" & Filter4 & "*"
    AndNot2 = "*" & Filter5 & "*"
    AndNot3 = "*" & Filter6 & "*"
    ExtensionFileNot = "*" & FileTypeNot & "*"
    
    ' Gestion de l'absence de filtre
    If Filter1 = "" Then And1 = "*.*"
    If Filter2 = "" Then And2 = "*.*"
    If Filter3 = "" Then And3 = "*.*"
    If FileType = "" Then ExtensionFile = "*.*"
    If Filter4 = "" Then AndNot1 = "*FAlbanY*"
    If Filter5 = "" Then AndNot2 = "*FAlbanY*"
    If Filter6 = "" Then AndNot3 = "*FAlbanY*"
    If FileTypeNot = "" Then ExtensionFileNot = "*FAlbanY*"
 
    'Boucle sur tous les fichiers du r�pertoire
    For Each FileItem In SourceFolder.Files
        File = FileItem.name
        ' Extension.
        If Not (File Like ExtensionFile) Then GoTo NextFile
        If (File Like ExtensionFileNot) Then GoTo NextFile
        ' Filtres And.
        If Not (File Like And1) Then GoTo NextFile
        If Not (File Like And2) Then GoTo NextFile
        If Not (File Like And3) Then GoTo NextFile
        ' Filtres NAnd.
        If (File Like AndNot1) Then GoTo NextFile
        If (File Like AndNot2) Then GoTo NextFile
        If (File Like AndNot3) Then GoTo NextFile
        ' Si Match.
        ReDim Preserve FileList(nb_file)
        ReDim Preserve FolderList(nb_file)
        FileList(nb_file) = FileItem.name                   ' Inscrit le nom du fichier dans le tableau
        FolderList(nb_file) = FileItem.ParentFolder         ' Inscrit le path du fichier dans le tableau
        ' Incr�mentation.
        nb_file = nb_file + 1
        ' Progression.
        Application.StatusBar = "Recherche des fichiers : " & nb_file & " fichiers trouv�s.     " & FolderList(nb_file - 1) & "\" & FileList(nb_file - 1)
NextFile:
    Next FileItem
    
    If IsMyProgressBarShow Then Call MyProgressBar.actualiser("Dossier : ..." & Right(SourceFolder.ParentFolder, 65), 0, nb_file, "Analyse des dossiers")

    '--- Appel r�cursif pour lister les fichier dans les sous-r�pertoires ---.
    For Each SubFolder In SourceFolder.SubFolders
       Call Select_AllFiles(SubFolder.path, FileType, FileTypeNot, Filter1, Filter2, Filter3, Filter4, Filter5, Filter6, FileList, FolderList, nb_file)
    Next SubFolder
    
'    Select_AllFiles = True
'    Exit Function
'IfError:
'    Select_AllFiles = False
End Function

Public Function Sort_Files(ByVal FilePaths As Variant, ByVal SortKey As String, Optional ByVal SortOrder As XlSortOrder = xlAscending, Optional ByVal CustomPattern As String = "") As Variant
    ' @brief Sorts an array of file paths based on specified criteria.
    ' @param FilePaths An array of full file path strings.
    ' @param SortKey The criteria to sort by. Valid options: "Date", "FileName", "FileSize", "CustomNumber".
    ' @param SortOrder (Optional) The sort direction, xlAscending (default) or xlDescending.
    ' @param CustomPattern (Optional) A VBScript RegExp pattern used to extract a number from the filename when SortKey is "CustomNumber".
    ' @return A variant array of strings containing the sorted file paths. Returns Empty on failure.
    ' @dependencies Microsoft Scripting Runtime, Microsoft VBScript Regular Expressions 5.5
    ' @details
    ' For "CustomNumber" sorting, the function uses a regular expression to find a number in the filename.
    ' It is best practice to use a capturing group `()` in your `CustomPattern` to isolate the exact numeric part.
    ' If a capturing group is present, its value is used for sorting. Otherwise, the entire matched pattern is used.
    ' @example
    ' ' To sort files like "DataFile_10.csv", "DataFile_2.csv" numerically:
    ' Dim files As Variant
    ' files = Array("C:\Data\DataFile_10.csv", "C:\Data\DataFile_2.csv")
    '
    ' ' This pattern finds a number (\d+) surrounded by an underscore and a dot.
    ' ' The parentheses `()` capture just the number for sorting.
    ' Dim pattern As String
    ' pattern = "_(\d+)\."
    '
    ' Dim sortedFiles As Variant
    ' sortedFiles = Sort_Files(files, "CustomNumber", xlAscending, pattern)

    On Error GoTo ErrorHandler
    Sort_Files = Empty ' Default return value

    ' 1. Input Validation
    If Not IsArray(FilePaths) Then
        Err.Raise vbObjectError, , "Invalid input for Sort_Files. An array of file paths is expected."
        Exit Function
    End If
    If LBound(FilePaths) > UBound(FilePaths) Then Exit Function ' Empty array, return Empty

    Dim lb As Long: lb = LBound(FilePaths)
    Dim ub As Long: ub = UBound(FilePaths)
    Dim i As Long, j As Long

    ' 2. Prepare data structure for sorting
    ' Using a 2D array to hold [SortValue, OriginalFilePath]
    Dim sortData() As Variant
    ReDim sortData(lb To ub, 1 To 2)

    Dim fso As Object ' Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim regEx As Object ' VBScript_RegExp_55.RegExp
    Dim regExMatches As Object

    If LCase(SortKey) = "customnumber" Then
        If CustomPattern = "" Then
            Err.Raise vbObjectError, , "A 'CustomPattern' must be provided when SortKey is 'CustomNumber'."
            Exit Function
        End If
        Set regEx = CreateObject("VBScript.RegExp")
        With regEx
            .Global = False
            .IgnoreCase = True
            .Pattern = CustomPattern
        End With
    End If

    ' 3. Populate the sort data array
    For i = lb To ub
        Dim file As Object
        Set file = fso.GetFile(FilePaths(i))

        Select Case LCase(SortKey)
            Case "date"
                sortData(i, 1) = file.DateLastModified
            Case "filename"
                sortData(i, 1) = file.Name
            Case "filesize"
                sortData(i, 1) = file.Size
            Case "customnumber"
                Set regExMatches = regEx.Execute(file.Name)
                If regExMatches.Count > 0 Then
                    If regExMatches(0).SubMatches.Count > 0 Then
                        ' A capturing group was used, extract the number from the first submatch.
                        sortData(i, 1) = Val(regExMatches(0).SubMatches(0))
                    Else
                        ' No capturing group, use the entire match.
                        sortData(i, 1) = Val(regExMatches(0).Value)
                    End If
                Else
                    sortData(i, 1) = -1 ' Default value for non-matching files
                End If
            Case Else
                Err.Raise vbObjectError, , "Invalid 'SortKey' provided: '" & SortKey & "'. Valid options are 'Date', 'FileName', 'FileSize', 'CustomNumber'."
                Exit Function
        End Select
        sortData(i, 2) = FilePaths(i)
    Next i

    ' 4. Sort the array
    For i = lb To ub - 1
        For j = i + 1 To ub
            If (SortOrder = xlAscending And sortData(i, 1) > sortData(j, 1)) Or (SortOrder = xlDescending And sortData(i, 1) < sortData(j, 1)) Then
                Dim tempSortValue As Variant
                Dim tempFilePath As String

                tempSortValue = sortData(i, 1)
                tempFilePath = sortData(i, 2)

                sortData(i, 1) = sortData(j, 1)
                sortData(i, 2) = sortData(j, 2)

                sortData(j, 1) = tempSortValue
                sortData(j, 2) = tempFilePath
            End If
        Next j
    Next i

    ' 5. Prepare the result array
    Dim result() As String
    ReDim result(lb To ub)

    For i = lb To ub
        result(i) = sortData(i, 2)
    Next i

    Sort_Files = result
    Exit Function

ErrorHandler:
    Sort_Files = Empty
End Function


Public Function SortBy(ByRef FileList() As String, ByRef FolderList() As String, ByRef nb_file As Long, SortFile As String, Marker As Variant, ByRef SortFn() As Variant, Optional SortSens As String = "+") As Boolean
    On Error GoTo ifError
    'OnMacroNameError
    ' Note d'utilisation du Marqueur de tri (ex "_SXX++"):
    '   "_S" repr�sente le pointeur. Le choix du pointeur est libre. (ex "_f", "_test", "test")
    '   "XX" repr�sente le nombre de digit num�riques � utiliser pour le tri. Seul le nombre de "X" est libre. (ex "X", "XX", "XXX")
    '   "++" et "--" indiquent le sens du tri. "+" par d�faut.
    '   option : Tri par date avec le marqueur "Date" (ex "date++").


    ' Affichage de la progression dans la barre de statut
    Application.StatusBar = "Tri des fichiers"
    
    ' Variables de la fonction.
    ByDate = False
    ByMark = True
    ReDim SortFn(nb_file)
    NoMarkInFile = ""
    AllFileMark = True
    SortformNb = 0
    SortMarkerNb = 0
    SortSheet_Name = "Fichiers"
    
    ' Mise en forme du fichier.
        ' Creation de la feuille.
        If IsSheetHere(SortFile, SortSheet_Name, False) = 0 Then Workbooks(SortFile).Sheets.Add After:=Workbooks(SortFile).Sheets(1): Workbooks(SortFile).Sheets(Workbooks(SortFile).Sheets.count).name = SortSheet_Name
        ' Titres de colonnes.
        Workbooks(SortFile).Sheets(SortSheet_Name).Range("A1").value = "SortKey"
        Workbooks(SortFile).Sheets(SortSheet_Name).Range("B1").value = "Id"
        Workbooks(SortFile).Sheets(SortSheet_Name).Range("C1").value = "Date"
        Workbooks(SortFile).Sheets(SortSheet_Name).Range("D1").value = "Path"
        Workbooks(SortFile).Sheets(SortSheet_Name).Range("E1").value = "File"
    
    ' Pr�sence du marqueur de tri : "Date" ou Marqueur.
    If Marker = "" Then
        ByMark = False
        ByDate = False
    End If
             
    If InStr(1, Marker, "Date", vbTextCompare) > 0 Then
        ByDate = True
        ByMark = False
    End If
    
    ' Conditionnement de l'op�rateur de tri.
    ' Sens de tri. (defaut : "+")
    If Marker Like "*++*" Then
        Marker = Replace(Marker, "+", "")
        SortSens = "+"
    End If
    If Marker Like "*--*" Then
        Marker = Replace(Marker, "-", "")
        SortSens = "-"
    End If
    TMarker = Marker
    ' Marqueur de tri.
    If ByMark Then
        For f = 1 To Len(Marker)
            ' D�compte du nombre d'Id de tri.
            If TMarker Like "*X*" Then
                TMarker = Replace(TMarker, "X", "", 1, 1, vbBinaryCompare)
                SortformNb = SortformNb + 1     ' Nombre de chiffres pour le tri.
            Else
                Exit For
            End If
            Workbooks(SortFile).Sheets(SortSheet_Name).Range("A2:E" & nb_file + 1).Sort key1:=Range("A2"), Order1:=xlDescending, header:=xlNo
    End If
   
    ' Remplissage des Colones et detection du marqueur.
    For f = 1 To nb_file
        SortFn(f) = f
        Workbooks(SortFile).Sheets(SortSheet_Name).Cells(f + 1, 1).value = SortFn(f)
        Workbooks(SortFile).Sheets(SortSheet_Name).Cells(f + 1, 2).value = f
        Workbooks(SortFile).Sheets(SortSheet_Name).Cells(f + 1, 3).value = FileDateTime(FolderList(f) & "/" & FileList(f))
        Workbooks(SortFile).Sheets(SortSheet_Name).Cells(f + 1, 4).value = FolderList(f)
        Workbooks(SortFile).Sheets(SortSheet_Name).Cells(f + 1, 5).value = FileList(f)
        ' Detection de l'absence du marqueur. Tous les fichiers doivent contenir le marqueur.
        If (FileList(f) Like ("*" & TMarker & "*")) = False And ByMark Then
            NoMarkInFile = NoMarkInFile & Chr(10) & FileList(f)
            AllFileMark = False
        End If
    Next
    
    ' R�alisation du tri par marqueur de titre.
    If ByMark And AllFileMark Then
        ' Nombre de caract�res du marqueur.
        SortMarkerNb = Len(Marker) - SortformNb
        ' Indice du marqueur de tri. Lecture de la valeur des "XX".
        For f = 1 To nb_file
            pos = InStr(1, FileList(f), TMarker)
            SortKey = Mid(FileList(f), pos + SortMarkerNb, SortformNb)
            Workbooks(SortFile).Sheets(SortSheet_Name).Cells(f + 1, 1).value = val(SortKey)
        Next
        End If
        ' Enregistrement dans SortFn.
        For f = 1 To nb_file
            SortFn(f) = Workbooks(SortFile).Sheets(SortSheet_Name).Cells(f + 1, 2).value
        Next
    Else
        ' Message d'erreur.
        If Not AllFileMark Then
            MsgBox "Marqueur de tri " & TMarker & " absent dans les fichiers : " _
                & Chr(10) & NoMarkInFile & Chr(10) & " . Poursuite sans tri."
        End If
    End If
    
    If ByDate Then
        ' Tri selon le sens.
        If SortSens = "+" Or SortSens = "pls" Or SortSens = "xlAscending" Then
            Workbooks(SortFile).Sheets(SortSheet_Name).Range("A2:E" & nb_file + 1).Sort key1:=Range("C2"), Order1:=xlAscending, header:=xlNo
        Else
            Workbooks(SortFile).Sheets(SortSheet_Name).Range("A2:E" & nb_file + 1).Sort key1:=Range("C2"), Order1:=xlDescending, header:=xlNo
        End If
        ' Enregistrement dans SortFn.
        For f = 1 To nb_file
            SortFn(f) = Workbooks(SortFile).Sheets(SortSheet_Name).Cells(f + 1, 2).value
        Next
    End If

    SortBy = True
    Exit Function

ifError:
    SortBy = False

End Function

Public Function SaveAs(path, SaveName, Optional FileToSave As Variant, Optional format As String = ".xlsx", Optional AutoClose As Boolean = True) As Boolean
    ' Sauvegarde du fichier actif sauf si l'option "WorkbookName" est sp�cifi�.
    ' Ex. Path = "\\serveur.ims.bordeaux\171_equipe35\These_Florent_ALBANY\"
    ' SaveName = "nom du fichier" &
    ' Format = "Extension". Valeur si omis : ".xlsx"
    ' FileToSave permet de sp�cifier le fichier ouvert � sauvegarder. Pr�sence de l'extension requise.
    Dim myFile As String
    
    ' Remplacement des caract�res interdits.
    SaveName = CaracteresInterdits(CStr(SaveName), "\/:*?""<>|;", " ")
    'testvar = InStrRev(Path, "\")
    If Not InStrRev(path, "\") = Len(path) Then path = path & "\"   ' /!\ experimental
    myFile = CStr(FileToSave)
    SaveName = Replace(SaveName, format, "")
    ' Sauvegarde.
    Select Case True
        Case IsMissing(FileToSave)
            With ActiveWorkbook
                .title = SaveName
                .Author = "Florent ALBANY"
                .SaveAs fileName:=path & SaveName & format _
                    , FileFormat:=xlWorkbookDefault, Password:="", WriteResPassword:="", _
                    ReadOnlyRecommended:=False, CreateBackUp:=False, ConflictResolution:=xlUserResolution
            End With
        Case Else
            With Workbooks(myFile)
                .title = SaveName
                .Author = "Florent ALBANY"
                .SaveAs fileName:=path & SaveName & format _
                    , FileFormat:=xlWorkbookDefault, Password:="", WriteResPassword:="", _
                    ReadOnlyRecommended:=False, CreateBackUp:=False, ConflictResolution:=xlUserResolution
            End With
    End Select
    If AutoClose Then Workbooks(SaveName & format).Close   ' ActiveWorkbook.Close
    SaveAs = True
End Function

Public Function check_dir(directory As String) As Boolean
      On Error GoTo ErrNotExist
      ChDir (directory)
      check_dir = True
      Exit Function
ErrNotExist:
      check_dir = False
End Function



Public Function check_file(fFile As String) As Boolean
      Dim pReadFile As Integer
      On Error GoTo ErrNotExist
      pReadFile = FreeFile()
      Open fFile For Input As #pReadFile
      Close #pReadFile
      check_file = True
      
      Exit Function
ErrNotExist:
      check_file = False
End Function

Public Function create_folder(ByVal DirectoryPath As String) As Boolean
    Dim i                   As Integer
    Dim strArray()          As String
    Dim strSubFolder        As String
    
    DirectoryPath = Clean_FilePath(DirectoryPath)
    DirectoryPath = Get_FileDirectory(DirectoryPath)
    
    strArray = Split(DirectoryPath, "\")
    strSubFolder = Trim(strArray(0))
    For i = 1 To UBound(strArray)
        strSubFolder = strSubFolder & "\" & Trim(strArray(i))
        If Dir(strSubFolder, vbDirectory) = "" Then MkDir strSubFolder
    Next
    create_folder = check_dir(DirectoryPath)
End Function

Public Function read_file2str(fullFilename As String) As String
    Dim pReadFile As Integer
    Dim DataLine As String
    Dim strout As String
    strout = ""
    pReadFile = FreeFile()
    On Error GoTo ErrNotExist
    Open fullFilename For Input As #pReadFile
    While Not EOF(pReadFile)
        Line Input #pReadFile, DataLine
        strout = strout & DataLine & vbCrLf
    Wend
    Close #pReadFile
    read_file2str = strout
    Exit Function
    
ErrNotExist:
      read_file2str = strout
End Function

Public Function write_str2file(stext As String, sFile As String) As Integer
    If OverwriteTxt(sFile, stext) Then write_str2file = 0 Else write_str2file = 1
'    Dim pWriteFile As Integer
'    pWriteFile = FreeFile()
'    Open filepath For Output As #pWriteFile
'    Print #pWriteFile, StrText
'    Close #pWriteFile
'    write_str2file = 0
End Function

Public Function Save_Str2File(stext As String, sFile As String) As Variant
    ' Ecrire une variable string dans un fichier texte et qui cr�e le r�pertoire inclus dans le fichier texte s�il n�existe pas.
    ' cr�er le r�pertoire inclus dans le fichier texte s'il n'existe pas
    If Not create_folder(sFile) Then Save_Str2File = CVErr(2001): Exit Function
    If Not OverwriteTxt(sFile, stext) Then Save_Str2File = CVErr(2001): Exit Function
    Save_Str2File = FileExist(sFile)
    
'    Save_Str2File = False
'    If Dir(Left(sFile, InStrRev(sFile, "\"))) = "" Then
'        If Not create_folder(Left(sFile, InStrRev(sFile, "\"))) Then GoTo ifError
'    End If
'    Dim pWriteFile As Integer
'    pWriteFile = FreeFile()
'    Open sFile For Output As #pWriteFile
'    Print #pWriteFile, sText
'    Close #pWriteFile
'    Save_Str2File = True
'    Exit Function
'ifError:
'    Save_Str2File = CVErr(2001)
End Function

Public Function append_str2file(StrText As String, filePath As String) As Integer
    Dim pWriteFile As Integer
    pWriteFile = FreeFile()
    Open filePath For Append As #pWriteFile
    Print #pWriteFile, StrText
    Close #pWriteFile
    append_str2file = 0
End Function

Public Function rm_file_if_exist(FileToDelete As String) As Boolean
    If check_file(FileToDelete) Then 'See above
        ' First remove readonly attribute, if set
        SetAttr FileToDelete, vbNormal
        ' Then delete the file
        Kill FileToDelete
        rm_file_if_exist = True
    Else
        rm_file_if_exist = False
    End If
End Function

Function store_data_to_mdm(filePath As String) As Boolean
    Dim StrTmp          As String
    Dim sLabels         As String
    Dim aStrTmp()       As String
    Dim str_line        As String
    Dim i               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim m               As Integer
    Dim n               As Integer
    Dim idx             As Integer
    Dim nbSync          As Integer
    Dim SyncNo          As Integer
    Dim iSpaceNb        As Integer
    Dim colsize         As Integer
    Dim nbswlin         As Integer
    Dim nbInput         As Integer
    Dim valcst          As Double
    Dim Ratio           As Double
    Dim Offset          As Double
    Dim aInputs()       As Double
    Dim full_size_input As Long
    Dim nb_2nd_sweep_order As Long
    
    ' Initialize
    colsize = 20
    nbInput = 0
    nbswlin = 0
    full_size_input = 1
    
    For i = 1 To NBSMU
        If VAR_CB_ONOFF_SMU(i) Then
            nbInput = nbInput + 1
            If VAR_SWEEP_TYPE(i) = "LIN" Then nbswlin = nbswlin + 1
        End If
    Next i
    
    For i = 1 To nbswlin: full_size_input = full_size_input * VAR_SIZE_INPUT(VAR_SORT_INDEX_INPUT(i)): Next i
    nb_2nd_sweep_order = full_size_input / VAR_SIZE_INPUT(VAR_SORT_INDEX_INPUT(1))
        
    ' Build the Label line.
    sLabels = " #" & VAR_INPUT_NAME(VAR_SORT_INDEX_INPUT(1)) & "::"
    For i = 1 To nbInput: sLabels = IIf(VAR_SWEEP_TYPE(VAR_SORT_INDEX_INPUT(i)) = "SYNC", sLabels & VAR_INPUT_NAME(VAR_SORT_INDEX_INPUT(i)) & "::", sLabels): Next i
    For i = 1 To nbInput: sLabels = IIf(VAR_EXPORT(VAR_SORT_INDEX_INPUT(i)), sLabels & VAR_OUTPUT_NAME(VAR_SORT_INDEX_INPUT(i)) & "::", sLabels): Next i
    aStrTmp = Split(sLabels, "::")
    sLabels = ""
    For i = LBound(aStrTmp) To UBound(aStrTmp)
        If Len(aStrTmp(i)) < colsize Then iSpaceNb = colsize - Len(aStrTmp(i)) Else iSpaceNb = 1
        sLabels = sLabels & aStrTmp(i) & Space(iSpaceNb)
    Next
    
    ' Build The INPUTs Array (SWEEP1 + SYNCs)
    For i = 1 To nbInput: nbSync = IIf(VAR_SWEEP_TYPE(VAR_SORT_INDEX_INPUT(i)) = "SYNC", nbSync + 1, nbSync): Next i
    ReDim aInputs(0 To VAR_SIZE_INPUT(VAR_SORT_INDEX_INPUT(1)) - 1, 0 To nbSync)
    For l = 0 To UBound(aInputs, 1)
        aInputs(l, 0) = (CStrToCDbl(VAR_START(VAR_SORT_INDEX_INPUT(1)))) + l * (CStrToCDbl(VAR_STEP(VAR_SORT_INDEX_INPUT(1))))
        aInputs(l, 0) = IIf(Abs(aInputs(l, 0)) < VALMIN, 0, aInputs(l, 0))
        SyncNo = 1
        For i = 1 To nbInput
            idx = VAR_SORT_INDEX_INPUT(i)
            If VAR_SWEEP_TYPE(idx) = "SYNC" Then
                Ratio = (CStrToCDbl(VAR_STOP(idx)) - CStrToCDbl(VAR_START(idx))) / (CStrToCDbl(VAR_STOP(VAR_SORT_INDEX_INPUT(1))) - CStrToCDbl(VAR_START(VAR_SORT_INDEX_INPUT(1))))
                Offset = (CStrToCDbl(VAR_START(idx))) - (CStrToCDbl(VAR_START(VAR_SORT_INDEX_INPUT(1)))) * Ratio
                valcst = (CStrToCDbl(VAR_START(VAR_SORT_INDEX_INPUT(1)))) + l * (CStrToCDbl(VAR_STEP(VAR_SORT_INDEX_INPUT(1)))) * Ratio + Offset
                valcst = IIf(Abs(valcst) < VALMIN, 0, valcst)
                aInputs(l, SyncNo) = valcst
                SyncNo = SyncNo + 1
            End If
        Next i
    Next
    
    StrTmp = "! VERSION = 6.00" & vbCrLf
'    ' FALBANY Add in 2023.2.3 : MEASUREMENT_PARAMETER.
    StrTmp = StrTmp & "! BEGIN_MEASX" & vbCrLf
    StrTmp = StrTmp & "!   VERSION" & vbTab & RELEASE & vbCrLf
    StrTmp = StrTmp & "!   SETUP" & vbTab & VAR_MsState.setup & vbTab & VAR_MsState.Mdm & vbCrLf
    StrTmp = StrTmp & "!   DEVICE" & vbTab & VAR_MsState.device & vbTab & VAR_MsState.Structure & vbCrLf
    StrTmp = StrTmp & "!   TEMP" & vbTab & VAR_MsState.Temperature & vbCrLf
    StrTmp = StrTmp & "!   TIME" & vbTab & format(VAR_MsState.Start_Time, "yyyy/mm/dd_h:mm:ss") & vbTab & format(VAR_MsState.Stop_Time, "yyyy/mm/dd_h:mm:ss") & vbTab & VAR_MsState.Duration & vbCrLf
    StrTmp = StrTmp & "!   NAME" & vbTab & VAR_MsState.Mdm & "_" & VAR_MsState.device & "_" & VAR_MsState.Structure & "_" & VAR_MsState.Temperature & "_" & format(VAR_MsState.Start_Time, "yyyy-mm-dd_h-mm-ss") & vbCrLf
    StrTmp = StrTmp & "! END_MEASX" & vbCrLf & vbCrLf
'    ' End FALBANY Add in 2023.2.3

    ' <<<<<<<< BEGIN_HEADER >>>>>>>>>
    StrTmp = StrTmp & "BEGIN_HEADER" & vbCrLf
        ' Build ICCAP_INPUTS.
        StrTmp = StrTmp & " ICCAP_INPUTS" & vbCrLf
        For i = 1 To nbInput
            idx = VAR_SORT_INDEX_INPUT(i)
            StrTmp = StrTmp & Space(2) & VAR_INPUT_NAME(idx) & Space(8) & VAR_MODE(idx) & Space(2) & VAR_NODE(idx) & " GROUND SMU" & VAR_SMU_UNIT(idx) & " " & CStrToStr(VAR_COMPLIANCE(idx)) & " " & VAR_SWEEP_TYPE(idx)
            Select Case VAR_SWEEP_TYPE(idx)
                Case "LIN": StrTmp = StrTmp & Space(8) & VAR_SWEEP_ORDER(idx) & Space(4) & CStrToStr(VAR_START(idx)) & Space(6) & CStrToStr(VAR_STOP(idx)) & "       " & VAR_SIZE_INPUT(idx) & "   " & CStrToStr(VAR_STEP(idx)) & vbCrLf
                Case "CON": StrTmp = StrTmp & Space(8) & CStrToStr(VAR_START(idx)) & vbCrLf
                Case "SYNC"
                    Ratio = (CStrToCDbl(VAR_STOP(idx)) - CStrToCDbl(VAR_START(idx))) / (CStrToCDbl(VAR_STOP(VAR_SORT_INDEX_INPUT(1))) - CStrToCDbl(VAR_START(VAR_SORT_INDEX_INPUT(1))))
                    Offset = CStrToCDbl(VAR_START(idx)) - CStrToCDbl(VAR_START(VAR_SORT_INDEX_INPUT(1))) * Ratio
                    StrTmp = StrTmp & Space(8) & CCDblToStr(Ratio) & " " & CCDblToStr(Offset) & " " & VAR_INPUT_NAME(VAR_SORT_INDEX_INPUT(1)) & vbCrLf
            End Select
        Next i
        ' Build ICCAP_OUTPUTS.
        StrTmp = StrTmp & " ICCAP_OUTPUTS" & vbCrLf
        For i = 1 To nbInput
            idx = VAR_SORT_INDEX_INPUT(i)
            If VAR_EXPORT(idx) = True Then
                StrTmp = StrTmp & Space(2) & VAR_OUTPUT_NAME(idx) & Space(8) & IIf(VAR_MODE(idx) = "V", "I", "V") & Space(2) & VAR_NODE(idx) & " GROUND SMU" & VAR_SMU_UNIT(idx) & " B" & vbCrLf
            End If
        Next i
    StrTmp = StrTmp & "END_HEADER" & vbCrLf & vbCrLf
    ' <<<<<<<< END_HEADER >>>>>>>>>

    If nbswlin = 0 Then m = 0 Else m = 1
    For k = 0 To nb_2nd_sweep_order - 1
        StrTmp = StrTmp & "BEGIN_DB" & vbCrLf
        ' <<<<<<<< ICCAP_VAR >>>>>>>>>
        If nbInput > 1 Then
            For i = (m + 1) To nbInput
                idx = VAR_SORT_INDEX_INPUT(i)
                If VAR_SWEEP_TYPE(idx) <> "SYNC" Then
                    valcst = 1
                    l = 2
                    Do While l < VAR_SWEEP_ORDER(idx)
                        valcst = valcst * VAR_SIZE_INPUT(VAR_SORT_INDEX_INPUT(l))
                        l = l + 1
                    Loop
                    valcst = (CStrToCDbl(VAR_START(idx))) + (CStrToCDbl(VAR_STEP(idx))) * k
                    StrTmp = StrTmp & " ICCAP_VAR " & VAR_INPUT_NAME(idx) & Space(8) & CStrToStr(format(valcst, "##0.0##########")) & vbCrLf
                End If
            Next i
        End If
        StrTmp = StrTmp & vbCrLf
        ' <<<<<<<< ICCAP_VAR >>>>>>>>>

        ' <<<<<<<< LABELS >>>>>>>>>
        StrTmp = StrTmp & sLabels & vbCrLf
        ' <<<<<<<< LABELS >>>>>>>>>

        ' <<<<<<<< INPUTS & OUTPUTS VALUES >>>>>>>>>
        For l = 0 To VAR_SIZE_INPUT(VAR_SORT_INDEX_INPUT(1)) - 1
            str_line = Space(2)
            ' INPUTs.
            For i = LBound(aInputs, 2) To UBound(aInputs, 2): str_line = str_line & CStrToStr(format(aInputs(l, i), "Scientific")) & "::": Next i
            ' OUTPUTs.
            For i = 1 To nbInput
                idx = VAR_SORT_INDEX_INPUT(i)
                If VAR_EXPORT(idx) Then
                    valcst = VAR_ARRAY_2D_DATA(idx, l + k * VAR_SIZE_INPUT(VAR_SORT_INDEX_INPUT(1)))
                    If Abs(valcst) < VALMIN Then valcst = 0
                    Select Case VAR_SWEEP_TYPE(idx)
                        Case "SYNC": str_line = str_line & CStrToStr(format(valcst, "Scientific")) & "::"
                        Case Else: str_line = str_line & CStrToStr(format(valcst, "##0.0##########")) & "::"
                    End Select
                End If
            Next i
            ' FORMAT the line in columns and add to StrTmp.
            aStrTmp = Split(str_line, "::")
            str_line = ""
            For i = LBound(aStrTmp) To UBound(aStrTmp)
                If Len(aStrTmp(i)) < colsize Then iSpaceNb = colsize - Len(aStrTmp(i)) Else iSpaceNb = 1
                str_line = str_line & aStrTmp(i) & Space(iSpaceNb)
            Next
            StrTmp = StrTmp & str_line & vbCrLf
        Next l
        ' <<<<<<<< INPUTS & OUTPUTS VALUES >>>>>>>>>
        StrTmp = StrTmp & "END_DB" & vbCrLf & vbCrLf
    Next k

    ' Write the file
    store_data_to_mdm = OverwriteTxt(filePath, StrTmp)
End Function


Function store_data_to_excel_file(filePath As String) As Integer
    Dim fl_Update As Boolean
    Dim fl_disp As Boolean
    Dim sh As Worksheet
    Dim wb As Workbook

    Dim eapp As Excel.Application
    Dim wbook_src As Workbook
    Dim wbook_dest As Workbook
    Dim wsheet As Worksheet
    Dim pathExcel_Output As String
    
    fl_Update = Application.ScreenUpdating
    fl_disp = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbook_src = ActiveWorkbook
    Set wbook_dest = Workbooks.Add
   
    ' Delete Default sheet excepted the first one
    Do While wbook_dest.Sheets.count > 1
        wbook_dest.Sheets(wbook_dest.Sheets.count).Delete
    Loop
    
    'Copy Sheets to new workbook
    For Each wsheet In wbook_src.Sheets
        If wsheet.name <> "MeasX" Then
            wsheet.Copy After:=wbook_dest.Sheets(wbook_dest.Sheets.count)
        End If
    Next wsheet
    
    ' Delete the first sheet
    ' wbook_dest.Sheets(1).Delete

    wbook_dest.SaveAs filePath
    wbook_dest.Close SaveChanges:=True
    store_data_to_excel_file = 0
    
    If fl_disp = True Then Application.DisplayAlerts = True
    If fl_Update = True Then Application.ScreenUpdating = True
    
End Function

'-----------------------------------------'
'--------------- FILE MOD ----------------'
'-----------------------------------------'
'
'Public bSectionExists         As Boolean
'Public bKeyExists             As Boolean

'---------------------------------------------------------------------------------------
' Procedure : Ini_ReadKeyVal
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Read an Ini file's Key
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
' Req'd Refs: Uses Late Binding, so none required
'             No APIs either! 100% VBA
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sIniFile  : Full path and filename of the ini file to read
' sSection  : Ini Section to search for the Key to read the Key from
' sKey      : Name of the Key to read the value of
'
' Usage:
' ~~~~~~
' ? Ini_Read(Application.CurrentProject.Path & "\MyIniFile.ini", "LINKED TABLES", "Path")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-08-09              Initial Release
'---------------------------------------------------------------------------------------
Public Function Ini_ReadKeyVal(ByVal sIniFIle As String, _
                        ByVal sSection As String, _
                        ByVal sKey As String) As String
    On Error GoTo Error_Handler
    Dim sIniFileContent         As String
    Dim aIniLines()             As String
    Dim sLine                   As String
    Dim i                       As Long
    Dim bSectionExists          As Boolean
    Dim bKeyExists              As Boolean
    
    sIniFileContent = ""
    bSectionExists = False
    bKeyExists = False

    'Validate that the file actually exists
    If FileExist(sIniFIle) = False Then
        MsgBox "The specified ini file: " & vbCrLf & vbCrLf & _
               sIniFIle & vbCrLf & vbCrLf & _
               "could not be found.", vbCritical + vbOKOnly, "File not found"
        GoTo Error_Handler_Exit
    End If

    sIniFileContent = ReadFile(sIniFIle)    'Read the file into memory
    aIniLines = Split(sIniFileContent, vbCrLf)
    For i = 0 To UBound(aIniLines)
        sLine = Trim(aIniLines(i))
        If bSectionExists = True And Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
            Exit For    'Start of a new section
        End If
        If sLine = "[" & sSection & "]" Then
            bSectionExists = True
        End If
        If bSectionExists = True Then
            If Len(sLine) > Len(sKey) Then
                If Left(sLine, Len(sKey) + 1) = sKey & "=" Then
                    bKeyExists = True
                    Ini_ReadKeyVal = Mid(sLine, InStr(sLine, "=") + 1)
                End If
            End If
        End If
    Next i

Error_Handler_Exit:
    On Error Resume Next
    Exit Function

Error_Handler:
    'Err.Number = 75 'File does not exist, Permission issues to write is denied,
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: Ini_ReadKeyVal" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : Ini_WriteKeyVal
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Writes a Key value to the specified Ini file's Section
'               If the file does not exist, it will be created
'               If the Section does not exist, it will be appended to the existing content
'               If the Key does not exist, it will be appended to the existing Section content
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
' Req'd Refs: Uses Late Binding, so none required
'             No APIs either! 100% VBA
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sIniFile  : Full path and filename of the ini file to edit
' sSection  : Ini Section to search for the Key to edit
' sKey      : Name of the Key to edit
' sValue    : Value to associate to the Key
'
' Usage:
' ~~~~~~
' Call Ini_WriteKeyVal(Application.CurrentProject.Path & "\MyIniFile.ini", "LINKED TABLES", "Paths", "D:\")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-08-09              Initial Release
' 2         2020-01-27              Fix to address issue flagged by users
'---------------------------------------------------------------------------------------
Public Function Ini_WriteKeyVal(ByVal sIniFIle As String, _
                         ByVal sSection As String, _
                         ByVal sKey As String, _
                         ByVal sValue As String) As Boolean
    On Error GoTo Error_Handler
    Dim sIniFileContent         As String
    Dim aIniLines()             As String
    Dim sLine                   As String
    Dim sNewLine                As String
    Dim i                       As Long
    Dim bFileExist              As Boolean
    Dim bInSection              As Boolean
    Dim bKeyAdded               As Boolean
    Dim bSectionExists          As Boolean
    Dim bKeyExists              As Boolean


    sIniFileContent = ""
    bSectionExists = False
    bKeyExists = False

    'Validate that the file actually exists
    If FileExist(sIniFIle) = False Then GoTo SectionDoesNotExist Else bFileExist = True

    sIniFileContent = ReadFile(sIniFIle)    'Read the file into memory
    aIniLines = Split(sIniFileContent, vbCrLf)    'Break the content into individual lines
    sIniFileContent = ""    'Reset it
    For i = 0 To UBound(aIniLines)    'Loop through each line
        sNewLine = ""
        sLine = Trim(aIniLines(i))
        If sLine = "[" & sSection & "]" Then
            bSectionExists = True
            bInSection = True
        End If
        If bInSection = True Then
            If sLine <> "[" & sSection & "]" Then
                If Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then    ' Original wrong code : If Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
                    'Our section exists, but the key wasn't found, so append it
                    If bKeyAdded = False Then
                        sNewLine = sKey & "=" & sValue
                        i = i - 1
                        'bInSection = False    ' we're switching section
                        bKeyAdded = True
                    End If
                    bInSection = False    ' we're switching section
                End If
            End If
            If Len(sLine) > Len(sKey) Then
                If Split(sLine, "=")(0) = sKey Then 'Original code : If Left(sLine, Len(sKey) + 1) = sKey & "=" Then
                    sNewLine = sKey & "=" & sValue
                    bKeyExists = True
                    bKeyAdded = True
                End If
            End If
        End If
        If Len(sIniFileContent) > 0 Then sIniFileContent = sIniFileContent & vbCrLf
        If sNewLine = "" Then sIniFileContent = sIniFileContent & sLine Else sIniFileContent = sIniFileContent & sNewLine
    Next i

SectionDoesNotExist:
    'if not found, add it to the end
    If bSectionExists = False Then
        If Len(sIniFileContent) > 0 Then sIniFileContent = sIniFileContent & vbCrLf
        sIniFileContent = sIniFileContent & "[" & sSection & "]"
    End If
    If bKeyAdded = False Then
        sIniFileContent = sIniFileContent & vbCrLf & sKey & "=" & sValue
    End If

    'Write to the ini file the new content
    Call OverwriteTxt(sIniFIle, sIniFileContent)
    Ini_WriteKeyVal = True

Error_Handler_Exit:
    On Error Resume Next
    Exit Function

Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: Ini_WriteKeyVal" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : FileExist
' DateTime  : 2023.10
' Author    : Florent ALBANY
' Website   :
' Purpose   : Test for the existance of a directory; Returns True/False
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' Path - name of the Path to be tested for including full path
'---------------------------------------------------------------------------------------
Public Function DirExist(path As String) As Boolean
    On Error Resume Next
    DirExist = IIf(Dir(path, vbDirectory) <> "", True, False)
End Function

'---------------------------------------------------------------------------------------
' Procedure : FileExist
' DateTime  : 2007-Mar-06 13:51
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Test for the existance of a file; Returns True/False
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' strFile - name of the file to be tested for including full path
'---------------------------------------------------------------------------------------
Public Function FileExist(strFile As String) As Boolean
    On Error GoTo err_handler
    
    strFile = Clean_FilePath(strFile)
    If Len(Dir(strFile)) > 0 Then FileExist = True Else FileExist = False
    
Exit_Err_Handler:
    Exit Function

err_handler:
    MsgBox "The following error has occurred." & vbCrLf & vbCrLf & _
            "Error Number: " & err.Number & vbCrLf & _
            "Error Source: FileExist" & vbCrLf & _
            "Error Description: " & err.Description, _
            vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

'---------------------------------------------------------------------------------------
' Procedure : OverwriteTxt
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Output Data to an external file (*.txt or other format)
'             ***Do not forget about access' DoCmd.OutputTo Method for
'             exporting objects (queries, report,...)***
'             Will overwirte any data if the file already exists
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - name of the file that the text is to be output to including the full path
' sText - text to be output to the file
'
' Usage:
' ~~~~~~
' Call OverwriteTxt("C:\Users\Vance\Documents\EmailExp2.txt", "Text2Export")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-Jul-06                 Initial Release
'---------------------------------------------------------------------------------------
Public Function OverwriteTxt(sFile As String, stext As String) As Boolean
On Error GoTo err_handler
    Dim fileNumber As Integer
    
    sFile = Clean_FilePath(sFile)
    
    fileNumber = FreeFile()                     ' Get unused file number
    Open sFile For Output As #fileNumber        ' Connect to the file
    Print #fileNumber, stext;                   ' Append our string
    Close #fileNumber                           ' Close the file
    OverwriteTxt = True
    Exit Function
    
Exit_Err_Handler:
    Exit Function
 
err_handler:
    MsgBox "The following error has occurred." & vbCrLf & vbCrLf & _
            "Error Number: " & err.Number & vbCrLf & _
            "Error Source: OverwriteTxt" & vbCrLf & _
            "Error Description: " & err.Description, _
            vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

'---------------------------------------------------------------------------------------
' Procedure : AppendTxt
' Author    : Florent ALBANY.
' Website   :
' Purpose   : Append Data to an external file (*.txt or other format)
'             ***Do not forget about access' DoCmd.OutputTo Method for
'             exporting objects (queries, report,...)***
'             Will overwirte any data if the file already exists
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - name of the file that the text is to be output to including the full path
' sText - text to be output to the file
'
' Usage:
' ~~~~~~
' Call AppendTxt("C:\Users\Vance\Documents\EmailExp2.txt", "Text2Export")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2023-Oct-12                 Initial Release
'---------------------------------------------------------------------------------------
Public Function AppendTxt(sFile As String, stext As String, Optional appendWithLineBreak As Boolean = False) As Boolean
    On Error GoTo err_handler
    Dim fileNumber As Integer
    
    sFile = Clean_FilePath(sFile)
    
    fileNumber = FreeFile()                     ' Get unused file number
    Open sFile For Append As #fileNumber        ' Connect to the file in append mode
    
    If appendWithLineBreak Then
        If FileLen(sFile) > 0 Then Print #fileNumber, vbCrLf & stext;      ' Insert a new line before appending the string
    Else
        Print #fileNumber, stext;               ' Append our string without a new line
    End If
    
    Close #fileNumber                           ' Close the file
    AppendTxt = True
    Exit Function
    
Exit_Err_Handler:
    Exit Function
 
err_handler:
    MsgBox "The following error has occurred." & vbCrLf & vbCrLf & _
            "Error Number: " & err.Number & vbCrLf & _
            "Error Source: AppendTxt" & vbCrLf & _
            "Error Description: " & err.Description, _
            vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadFile
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Faster way to read text file all in RAM rather than line by line
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
' Input Variables:
' ~~~~~~~~~~~~~~~~
' strFile - name of the file that is to be read
'
' Usage Example:
' ~~~~~~~~~~~~~~~~
' MyTxt = ReadText("c:\tmp\test.txt")
' MyTxt = ReadText("c:\tmp\test.sql")
' MyTxt = ReadText("c:\tmp\test.csv")
'---------------------------------------------------------------------------------------
Public Function ReadFile(ByVal strFile As String) As String
On Error GoTo Error_Handler
    Dim fileNumber  As Integer
    Dim sFile       As String 'Variable contain file content
 
    fileNumber = FreeFile
    Open strFile For Binary Access Read As fileNumber
    sFile = Space(LOF(fileNumber))
    Get #fileNumber, , sFile
    Close fileNumber
 
    ReadFile = sFile
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred." & vbCrLf & vbCrLf & _
            "Error Number: " & err.Number & vbCrLf & _
            "Error Source: ReadFile" & vbCrLf & _
            "Error Description: " & err.Description, _
            vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Function Ini_ReadToDictionary(sIniFIle As String, Optional KeyCAseSensitive As Boolean = True, Optional sDelimiter As String = "=") As Dictionary
'---------------------------------------------------------------------------------------
' Procedure : Ini_ReadToDictionary
' Author    : Florent ALBANY
' Purpose   : Read an ini file and save it into a Dictionary structure.
'             /!\ sSection is set to empty if an empty line is encoutered.
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sIniFile - name of the file that is to be read
' KeyCAseSensitive - set the dictionary CompareMode
' sDelimiter - Separator between sSection and sKey
'
' Usage Example:
' ~~~~~~~~~~~~~~~~
' MyDict = Ini_ReadToDictionary("c:\tmp\test.txt", True)
' MyDict = Ini_ReadToDictionary("c:\tmp\test.sql")
'---------------------------------------------------------------------------------------
    Dim dict As New Dictionary
    Dim sSection As String
    Dim sKey As String
    Dim sValue As String
    Dim sIniFileContent As String
    Dim aIniLines() As String
    Dim i As Long
    
    If FileExist(sIniFIle) = False Then
        MsgBox "The specified ini file: " & vbCrLf & vbCrLf & _
               sIniFIle & vbCrLf & vbCrLf & _
               "could not be found.", vbCritical + vbOKOnly, "File not found"
        GoTo Error_Handler_Exit
    End If
    
    If KeyCAseSensitive Then dict.CompareMode = BinaryCompare Else dict.CompareMode = TextCompare
    dict("delimiter") = sDelimiter                      '
    sIniFileContent = ReadFile(sIniFIle)                ' Read the file into memory
    aIniLines = Split(sIniFileContent, vbCrLf)          ' Split file by lines
    For i = LBound(aIniLines) To UBound(aIniLines)
        If Left(aIniLines(i), 1) = "[" And Right(aIniLines(i), 1) = "]" Then sSection = "[" & Mid(aIniLines(i), 2, Len(aIniLines(i)) - 2) & "]": GoTo nextLine
        If aIniLines(i) = "" Then sSection = "": GoTo nextLine
        If Not aIniLines(i) Like "*" & sDelimiter & "*" Then GoTo nextLine
        sKey = Split(aIniLines(i), "=")(0)
        sValue = Split(aIniLines(i), "=")(1)
        dict(sSection & sKey) = sValue
nextLine:
    Next

    Set Ini_ReadToDictionary = dict
    Exit Function
Error_Handler_Exit:
End Function

Public Function Ini_ReadToCollection(sIniFIle As String, Optional sDelimiter As String = "=") As Collection
'---------------------------------------------------------------------------------------
' Procedure : Ini_ReadToCollection
' Author    : Florent ALBANY
' Purpose   : Read an ini file and save it into a collection structure.
'             /!\ sSection is set to empty if an empty line is encoutered.
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sIniFile - name of the file that is to be read
' KeyCAseSensitive - set the dictionary CompareMode
' sDelimiter - Separator between sSection and sKey
'
' Usage Example:
' ~~~~~~~~~~~~~~~~
' MyDict = Ini_ReadToDictionary("c:\tmp\test.txt", True)
' MyDict = Ini_ReadToDictionary("c:\tmp\test.sql")
'---------------------------------------------------------------------------------------
    Dim i               As Long
    Dim coll            As New Collection
    Dim sSection        As String
    Dim sKey            As String
    Dim sValue          As String
    Dim sIniFileContent As String
    Dim aIniLines()     As String
    
    
    If FileExist(sIniFIle) = False Then
        MsgBox "The specified ini file: " & vbCrLf & vbCrLf & _
               sIniFIle & vbCrLf & vbCrLf & _
               "could not be found.", vbCritical + vbOKOnly, "File not found"
        GoTo Error_Handler_Exit
    End If
    
    coll.Add Key:="delimiter", item:=sDelimiter                      '
    sIniFileContent = ReadFile(sIniFIle)                ' Read the file into memory
    aIniLines = Split(sIniFileContent, vbCrLf)          ' Split file by lines
    For i = LBound(aIniLines) To UBound(aIniLines)
        If Left(aIniLines(i), 1) = "[" And Right(aIniLines(i), 1) = "]" Then sSection = "[" & Mid(aIniLines(i), 2, Len(aIniLines(i)) - 2) & "]": GoTo nextLine
        If aIniLines(i) = "" Then sSection = "": GoTo nextLine
        If Not aIniLines(i) Like "*" & sDelimiter & "*" Then GoTo nextLine
        sKey = Split(aIniLines(i), "=")(0)
        sValue = Split(aIniLines(i), "=")(1)
        coll.Add Key:=sSection & sKey, item:=sValue
        
nextLine:
    Next

    Set Ini_ReadToCollection = coll
    Exit Function
Error_Handler_Exit:
End Function
'With the above code, you can now easily read any INI file key value and write/update key values by doing something like

Public Function CheckCreate_Sheet(Worksheet_Name As String, Optional Workbook_Name As Variant = "ActiveWorkbook") As Worksheet      ' FALBANY Add in 2023.4.1
    ' Check worksheet_name presence in workbook_name and return the Worksheet.
    ' If worksheet_name don't exist, create and return the Worksheet.
    Dim mySheet As Worksheet
    
    If Workbook_Name = "ActiveWorkbook" Then Workbook_Name = ActiveWorkbook.name
    
    On Error Resume Next
    Set mySheet = Workbooks(Workbook_Name).Worksheets(Worksheet_Name)
    On Error GoTo 0
    
    If mySheet Is Nothing Then
        Set mySheet = Workbooks(Workbook_Name).Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        mySheet.name = Worksheet_Name
    End If
    Set CheckCreate_Sheet = Feuille
End Function

Public Function Check_Sheet(Worksheet_Name As String, Optional Workbook_Name As Variant = "ActiveWorkbook") As Boolean              ' FALBANY Add in 2023.4.1
    '* @brief Check_Sheet function checks if a worksheet exists in a workbook.
    '* @param Worksheet_Name The name of the worksheet to check.
    '* @param Workbook_Name Optional parameter specifying the name of the workbook. Default value is "ActiveWorkbook".
    '* @return Boolean value indicating whether the worksheet exists in the specified workbook or not.
    '* @note If the workbook name is not provided, the function will check for the worksheet in the active workbook.
    Dim mySheet As Worksheet
    On Error GoTo ifError
    
    If Workbook_Name = "ActiveWorkbook" Then Workbook_Name = ActiveWorkbook.name
    Set mySheet = Workbooks(Workbook_Name).Worksheets(Worksheet_Name)
    Check_Sheet = True
    Exit Function
ifError:
    Check_Sheet = False
End Function

Sub ClearTextFile(filePath As String)
    '* @brief ClearTextFile subroutine clears the contents of a text file.
    '* @param filePath The path of the text file to be cleared.
    '* @details This subroutine checks if the specified text file exists using the FileX.FileExist function.
    '* If the file exists, it opens the file for output and immediately closes it, effectively clearing its contents.
    If FileX.FileExist(filePath) Then Open filePath For Output As #1: Close #1
End Sub

Sub Open_File(filePath As String)
    '* @brief Open_File subroutine opens a file using its default application.
    '* @param filePath The full path of the file to be opened.
    '* @details This subroutine uses the Shell.Application object to open a file using its default application.
    '* @note The file path must include the full path to the file.
    Dim objShell As Object
    If FileX.FileExist(filePath) Then
        Set objShell = CreateObject("Shell.Application")
        objShell.Open (filePath)
    End If
End Sub

Sub Open_Directory(path As String)
    '* @brief Open_Directory subroutine opens a directory in the file explorer.
    '* @param directoryPath The full path of the directory to be opened.
    '* @details This subroutine uses the Shell.Application object to open a directory in the file explorer.
    '* @note The directory path must include the full path to the directory.

    Dim objShell As Object
    
    path = FileX.Clean_FilePath(path)
    If FileX.DirExist(path) Then
        Set objShell = CreateObject("Shell.Application")
        objShell.Open (path)
    Else
        MsgBox "The directory '" & path & "' does not exist.", vbExclamation, "Directory Not Found"
    End If

End Sub

Public Function Get_FileDirectory(filePath As String) As String
    '* @brief Get_FileDirectory function returns the directory of a file.
    '* @param filePath The full path of the file.
    '* @return The directory of the file. If no directory is specified in the file path, an empty string is returned.
    Dim lastBackslashIndex  As Long
    Dim lastPointIndex      As Long
    lastBackslashIndex = InStrRev(filePath, "\")
    lastPointIndex = InStrRev(filePath, ".")
    If lastBackslashIndex > 0 Then
        If lastBackslashIndex < lastPointIndex Then Get_FileDirectory = Left(filePath, lastBackslashIndex - 1) Else Get_FileDirectory = filePath
    Else
        Get_FileDirectory = ""
    End If
        
End Function

Public Function Get_FileName(filePath As String) As String
    '* @brief Get_FileName function extracts the name of a file from the given file path.
    '* @param filePath The full path of the file.
    '* @return The name of the file. If no directory is specified in the file path, filePath is returned.
    Dim lastBackslashIndex  As Long
    Dim lastPointIndex      As Long
    Dim lenFileName         As Long
    lastBackslashIndex = InStrRev(filePath, "\")
    lastPointIndex = InStrRev(filePath, ".")
    lenFileName = IIf(lastBackslashIndex < lastPointIndex, lastPointIndex - lastBackslashIndex - 1, Len(filePath) - lastBackslashIndex)
    If lastBackslashIndex > 0 Then Get_FileName = Mid(filePath, lastBackslashIndex + 1, lenFileName) Else Get_FileName = filePath
End Function

Public Function Get_FileExtension(filePath As String) As String
    '* @brief   This function retrieves the file extension from a given file path.
    '* @param   filePath   The full path of the file.
    '* @return  The extension of the file. If the file path is invalid or doesn't contain a file name, it returns an empty string.
    '* @details This function extracts the file extension from the file name obtained from the given file path.
    '*          If the file name is an empty string, it means that the path doesn't contain a valid file name, so the function returns an empty string.
    '*          If the file name contains a dot, it extracts the extension using the method of finding the last occurrence of the dot.
    '*          Otherwise, it returns an empty string.
    Dim lastPointIndex As Long
    lastPointIndex = InStrRev(filePath, ".")
    If lastPointIndex > 0 Then Get_FileExtension = Mid(filePath, lastPointIndex + 1) Else Get_FileExtension = ""
End Function

Public Function Clean_FilePath(filePath As String) As String
    '* @brief Cleans a file path string by replacing forward slashes with backslashes.
    '* @param filePath The file path string to be cleaned.
    '* @return The cleaned file path string.
    Clean_FilePath = Replace(Replace(filePath, "/", "\"), "\\", "\")
End Function


Public Sub Import_CsvToSelection()
    '* @brief Imports CSV data into a selected range in Excel.
    '* @details This procedure prompts the user to select a destination cell, then prompts for a CSV file to import.
    '*  It then reads the file and writes the data into the selected range.
    On Error GoTo ifError
    Dim filePath        As Variant
    Dim rng             As Range
    Dim sCsv            As String
    
    ' Check if a range is selected, if not, prompt the user.
    Select Case TypeName(Selection)
        Case "Range": Set rng = Selection
        Case Else: MsgBox "You need to first select the destination cell", vbExclamation: GoTo ifError
    End Select
    
    ' Prompt the user to select a CSV file.
    filePath = FileX.Select_Files("csv", False)
    
    ' Check if an error occurred during file selection.
    If IsError(filePath) Then MsgBox "Aucun fichier s�lectionn�.", vbExclamation: GoTo ifError
    
    ' Read the selected CSV file.
    sCsv = ReadFile(filePath(LBound(filePath)))
    
    ' Check if the CSV file is empty.
    If Trim(sCsv) = "" Then MsgBox FileX.Get_FileName(CStr(filePath(LBound(filePath)))) & "." & FileX.Get_FileExtension(CStr(filePath(LBound(filePath)))) & " file is empty.", vbExclamation: GoTo ifError
    
    ' Convert and write the CSV data to the selected range.
    If Not ArrayX.Csv_To_Range(sCsv, rng, ",", Chr(34)) Then MsgBox "Write operation failed.", vbExclamation: GoTo ifError
    
    Exit Sub

ifError:
    ' Error handling goes here if needed.
End Sub

Public Sub Import_CsvToNewSpreadsheet()
    '* @brief Imports CSV data into a new Spreadsheet in Excel.
    '* @details This procedure prompts for a CSV file to import.
    '*  It then reads the file and writes the data into a new Spreadsheet in Excel.
'    On Error GoTo ifError
    Dim i               As Integer
    Dim filePath        As Variant
    Dim sCsv            As String
    Dim fileName        As String
    Dim SheetName       As String
    Dim wks             As Worksheet
    Dim wbk             As Workbook
    Dim Delimiter       As String
    Dim QuoteChar       As String
       
    Delimiter = ","
    QuoteChar = ""  ' chr(34)
    
    ' Prompt the user to select a CSV file.
    filePath = FileX.Select_Files("csv", True)
    
    ' Check if an error occurred during file selection.
    If IsError(filePath) Then MsgBox "Aucun fichier s�lectionn�.", vbExclamation: GoTo ifError
    
    For i = LBound(filePath) To UBound(filePath)
    
        ' Read the selected CSV file.
        fileName = FileX.Get_FileName(CStr(filePath(i)))
        sCsv = ReadFile(filePath(i))
        
        ' Check if the CSV file is empty.
        If Trim(sCsv) = "" Then MsgBox fileName & "." & FileX.Get_FileExtension(CStr(filePath(i))) & " file is empty.", vbExclamation: GoTo ifError
        
        ' Create a new Sheet for the import.
        Set wbk = ActiveWorkbook
        SheetName = LANG_MOD.Clear_CharInString(fileName, "_-:\/?*[]; ")
        SheetName = Resize_String(SheetName, 31)
        If EXCEL_MOD.wbk_SheetExist(wbk, SheetName) Then
            Set wks = wbk.Sheets(SheetName)
            wks.Cells.Clear
        Else
            Set wks = wbk.Sheets.Add
            wks.name = SheetName
            wks.Move After:=wbk.Sheets(wbk.Sheets.count)
        End If
            
        ' Convert and write the CSV data to the selected range.
        If Not ArrayX.Csv_To_Range(sCsv, wks.Range("A1"), Delimiter, QuoteChar) Then MsgBox "Write operation failed.", vbExclamation: GoTo ifError
    Next i
    
    Exit Sub

ifError:
    ' Error handling goes here if needed.
End Sub

Sub Export_SelectionToCsv()
    Dim Arr2D           As Variant
    Dim sCsv            As String
    Dim path            As String
    Dim filePath        As String
    Dim fileName        As String
    Dim Delimiter       As String
    Dim QuoteChar       As String
       
    Delimiter = ","
    QuoteChar = Chr(34)
    
    ' R�cup�rer la s�lection en tant que tableau 2D
    Arr2D = ArrayX.Selection_To_a2D
    If IsError(Arr2D) Then
        MsgBox "Aucune s�lection valide trouv�e.", vbExclamation
        Exit Sub
    End If
    
    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "csv Files (*.csv),*csv")
    If filePath = "" Then Exit Sub
    
    ' Convertir le tableau 2D en cha�ne CSV
    sCsv = ArrayX.a2D_ToCsv(Arr2D, Delimiter, QuoteChar)
    
    ' Enregistrer le fichier CSV
    FileX.OverwriteTxt filePath, sCsv
    
    MsgBox "Fichier CSV cr�� avec succ�s : " & filePath, vbInformation
End Sub

Sub Export_SheetToCsv()
    Dim Arr2D           As Variant
    Dim sCsv            As String
    Dim path            As String
    Dim filePath        As String
    Dim fileName        As String
    Dim Delimiter       As String
    Dim QuoteChar       As String
       
    Delimiter = ","
    QuoteChar = Chr(34)
    
    ' R�cup�rer la s�lection en tant que tableau 2D
    Arr2D = ArrayX.Selection_To_a2D
    If IsError(Arr2D) Then
        MsgBox "Aucune feuille valide trouv�e.", vbExclamation
        Exit Sub
    End If

    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "csv Files (*.csv),*csv")
    If filePath = "" Then Exit Sub
    
    ' Convertir le tableau 2D en cha�ne CSV
    sCsv = ArrayX.a2D_ToCsv(Arr2D, Delimiter, QuoteChar)
    
    ' Enregistrer le fichier CSV
    FileX.OverwriteTxt filePath, sCsv
    
    MsgBox "Fichier CSV cr�� avec succ�s : " & filePath, vbInformation
End Sub

' @Description("Reads a text file to a string from a specified path")
' @Param(path, String, "The full path to the text file.")
' @Returns(String, "The content of the file as a string, or an empty string if an error occurs or file doesn't exist.")
Public Function ReadStringFromFile(ByVal path As String) As String
    On Error GoTo ErrHandler ' G�re les erreurs d'ex�cution
    
    Dim fileNumber As Long
    Dim byteCount As Long
    Dim resultString As String
    
    ' 1. V�rifier l'existence du fichier.
    ' Dir retourne le nom du fichier s'il existe, sinon une cha�ne vide.
    ' Il ne v�rifie pas les droits d'acc�s ici, seulement l'existence physique.
    If Dir(path, vbNormal + vbHidden + vbSystem + vbArchive + vbReadOnly) <> vbNullString Then
        fileNumber = FreeFile ' Obtient le prochain num�ro de fichier disponible
        
        ' 2. Ouvrir le fichier en mode binaire pour lire tous les octets (y compris les caract�res nuls si pr�sents)
        ' Utilisation de Binary pour un contr�le plus fin et pour g�rer potentiellement l'encodage par la suite si n�cessaire.
        ' Cependant, Input$ reste orient� caract�res. Pour un encodage strict, voir l'alternative ci-dessous.
        Open path For Binary Access Read Shared As #fileNumber
            ' V�rifie si le fichier est vide
            If LOF(fileNumber) > 0 Then
                byteCount = LOF(fileNumber) ' Obtient la taille du fichier en octets
                resultString = Space(byteCount) ' Alloue la m�moire pour la cha�ne
                Get #fileNumber, , resultString ' Lit tous les octets du fichier dans la cha�ne
            Else
                resultString = vbNullString ' Fichier vide
            End If
        Close #fileNumber ' Toujours fermer le fichier
    Else
        ' Fichier non trouv� ou path invalide
        Call HandleError("Le fichier sp�cifi� n'existe pas ou le chemin est invalide : " & path)
        resultString = vbNullString
        GoTo CleanExit
    End If
    
    ReadStringFromFile = resultString
    
CleanExit:
    Exit Function

ErrHandler:
    ' Gestionnaire d'erreurs
    If fileNumber <> 0 Then Close #fileNumber ' S'assurer que le fichier est ferm� en cas d'erreur
    Call HandleError("Erreur lors de la lecture du fichier '" & path & "' : " & err.Description & " (Code: " & err.Number & ")")
    ReadStringFromFile = vbNullString ' Retourne une cha�ne vide en cas d'erreur
    Resume CleanExit ' Reprend � CleanExit pour une sortie propre
End Function

' @Description("D�termine l'encodage d'un fichier texte en v�rifiant les Byte Order Marks (BOMs).")
' @Param(filePath, String, "Le chemin complet du fichier � analyser.")
' @Returns(String, "L'encodage d�tect� (ex: 'UTF-8', 'UTF-16 LE', 'UTF-16 BE', 'ANSI'), ou 'Inconnu' si l'encodage ne peut �tre d�termin�.")
Public Function GetFileEncoding(ByVal filePath As String) As String
    On Error GoTo ErrHandler
    
    Dim fileNumber As Long
    Dim bytes() As Byte ' D�clar� sans taille initiale pour �viter "Tableau d�j� dimensionn�"
    Dim actualBytesRead As Long
    Dim result As String
    Dim maxBytesToRead As Long
    
    Debug.Print "--- D�but GetFileEncoding pour : " & filePath & " ---" ' Datalogging
    
    ' V�rifier si le fichier existe
    If Dir(filePath) = vbNullString Then
        Call HandleError("Le fichier sp�cifi� n'existe pas : " & filePath)
        GetFileEncoding = "Fichier non trouv�"
        Debug.Print "Fichier non trouv� : " & filePath ' Datalogging
        Exit Function
    End If
    
    fileNumber = FreeFile
    Open filePath For Binary Access Read As #fileNumber
    Debug.Print "Fichier ouvert en mode binaire avec num�ro " & fileNumber & "." ' Datalogging
    
    ' Lire les premiers octets (jusqu'� 4)
    If LOF(fileNumber) > 0 Then
        maxBytesToRead = 4 ' Nous voulons lire jusqu'� 4 octets pour les BOMs les plus longs
        
        ' Dimensionner le tableau en fonction de la taille r�elle du fichier, sans d�passer maxBytesToRead
        If LOF(fileNumber) < maxBytesToRead Then
            ReDim bytes(0 To LOF(fileNumber) - 1)
            Debug.Print "Taille du fichier (" & LOF(fileNumber) & " octets) inf�rieure � 4. Lecture de " & LOF(fileNumber) & " octets." ' Datalogging
        Else
            ReDim bytes(0 To maxBytesToRead - 1)
            Debug.Print "Lecture des " & maxBytesToRead & " premiers octets du fichier." ' Datalogging
        End If
        
        Get #fileNumber, , bytes
        actualBytesRead = UBound(bytes) - LBound(bytes) + 1
        
        ' Datalogging des octets lus en hexad�cimal
        Dim byteStr As String
        Dim j As Long
        For j = LBound(bytes) To UBound(bytes)
            byteStr = byteStr & Right("00" & Hex(bytes(j)), 2) & " "
        Next j
        Debug.Print "Octets lus (hex) : " & Trim(byteStr) ' Datalogging
        
    Else
        ' Fichier vide
        GetFileEncoding = "Vide"
        Close #fileNumber
        Debug.Print "Fichier vide : " & filePath ' Datalogging
        Exit Function
    End If
    
    Close #fileNumber ' Fermer le fichier imm�diatement apr�s avoir lu le BOM
    Debug.Print "Fichier ferm�." ' Datalogging
    
    ' D�tection de l'encodage bas� sur le BOM (Byte Order Mark)
    Select Case True
        ' UTF-8 BOM: EF BB BF
        Case actualBytesRead >= 3 And _
             bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF
            result = "UTF-8"
            
        ' UTF-16 Little Endian BOM: FF FE
        Case actualBytesRead >= 2 And _
             bytes(0) = &HFF And bytes(1) = &HFE
            result = "UTF-16 LE" ' Little Endian

        ' UTF-16 Big Endian BOM: FE FF
        Case actualBytesRead >= 2 And _
             bytes(0) = &HFE And bytes(1) = &HFF
            result = "UTF-16 BE" ' Big Endian
            
        ' UTF-32 Little Endian BOM: FF FE 00 00
        Case actualBytesRead >= 4 And _
             bytes(0) = &HFF And bytes(1) = &HFE And bytes(2) = &H0 And bytes(3) = &H0
            result = "UTF-32 LE" ' Little Endian

        ' UTF-32 Big Endian BOM: 00 00 FE FF
        Case actualBytesRead >= 4 And _
             bytes(0) = &H0 And bytes(1) = &H0 And bytes(2) = &HFE And bytes(3) = &HFF
            result = "UTF-32 BE" ' Big Endian
            
        Case Else
            result = "ANSI"
    End Select
    
    GetFileEncoding = result
    Debug.Print "Encodage d�tect� : " & result ' Datalogging
    
CleanExit:
    Debug.Print "--- Fin GetFileEncoding ---" ' Datalogging
    Exit Function

ErrHandler:
    If fileNumber <> 0 Then Close #fileNumber
    Call HandleError("Erreur lors de la d�tection de l'encodage du fichier '" & filePath & "' : " & err.Description & " (Code: " & err.Number & ")")
    GetFileEncoding = "Erreur"
    Debug.Print "ERREUR dans GetFileEncoding. Fichier : " & filePath & ", Code : " & err.Number & ", Description : " & err.Description ' Datalogging
    Resume CleanExit
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReadFile
' Author    : Adapt� par Florent ALBANY (original CARDA Consultants Inc.)
' Purpose   : Reads a text file into a string, automatically detecting its encoding
'             using the GetFileEncoding function and ADODB.Stream for robust reading.
' Input Variables:
' ~~~~~~~~~~~~~~~~
' strFile - The full path to the file that is to be read.
'
' Returns:
' ~~~~~~~~
' String - The content of the file as a VBA String (internal Unicode/UTF-16 LE),
'          or an empty string if an error occurs, the file is not found, or is empty.
'
' Dependencies:
' ~~~~~~~~~~~~~~
' - Microsoft ActiveX Data Objects x.x Library (Required reference for ADODB.Stream)
' - GetFileEncoding (Function in this module to detect file encoding)
' - HandleError (external function for centralized error reporting)
'
' Usage Example:
' ~~~~~~~~~~~~~~~~
' Dim fileContent As String
' fileContent = ReadFile("C:\path\to\your_diverse_file.txt")
' If fileContent <> "" Then
'     MsgBox "Contenu du fichier : " & vbCrLf & fileContent
' Else
'     MsgBox "Impossible de lire le fichier ou fichier vide.", vbExclamation
' End If
'---------------------------------------------------------------------------------------
Public Function ReadFile_WithADO(ByVal strFile As String) As String
    On Error GoTo Error_Handler
    
    Dim adoStream As ADODB.Stream ' Requires reference to Microsoft ActiveX Data Objects Library
    Dim detectedEncoding As String
    Dim tempPath As String ' Variable temporaire pour Dir$
    
    ReadFile_WithADO = "" ' Initialize the result to an empty string
    
    ' 1. Verify file existence and validity.
    tempPath = Dir$(strFile, vbNormal + vbHidden + vbSystem + vbArchive + vbReadOnly)
    If tempPath = vbNullString Then
        Call HandleError("The specified file does not exist or the path is invalid: " & strFile)
        GoTo Error_Handler_Exit ' Exit cleanly
    End If
    
    ' 2. Detect the file's encoding using our dedicated function.
    detectedEncoding = GetFileEncoding(strFile)
    
    ' Handle cases where GetFileEncoding might return an error or specific status
    Select Case detectedEncoding
        Case "Fichier non trouv�"
            ' Error already handled by GetFileEncoding, just exit
            GoTo Error_Handler_Exit
        Case "Vide"
            ' File is empty, already handled. ReadFile is already ""
            GoTo Error_Handler_Exit
        Case "Erreur"
            ' An error occurred during encoding detection, already handled.
            GoTo Error_Handler_Exit
    End Select
    
    ' 3. Read the file content using ADODB.Stream with the detected encoding.
    Set adoStream = New ADODB.Stream
    
    With adoStream
        .Type = 2 ' adTypeText
        .Open
        
        ' Set the charset based on the detected encoding.
        ' Map detected encodings to ADO.Stream charsets.
        Select Case detectedEncoding
            Case "UTF-8"
                .Charset = "utf-8"
            Case "UTF-16 LE"
                .Charset = "utf-16" ' ADO treats this as LE by default when loading from file with BOM
            Case "UTF-16 BE"
                .Charset = "utf-16BE" ' ADO has a specific charset for Big Endian
            Case "UTF-32 LE" ' ADO does not natively support UTF-32 for direct char reading,
                ' This would require loading as binary and manual conversion.
                ' For simplicity and common use cases, we'll revert to a common text encoding.
                ' If strict UTF-32 is needed, this function requires a deeper change (e.g., direct byte manipulation).
                Call HandleError("UTF-32 encoding detected for '" & strFile & "'. ADODB.Stream does not natively support direct text reading of UTF-32. Attempting with UTF-8 fallback.")
                .Charset = "utf-8" ' Fallback for less common encodings
            Case "UTF-32 BE" ' Same for UTF-32 BE
                 Call HandleError("UTF-32BE encoding detected for '" & strFile & "'. ADODB.Stream does not natively support direct text reading of UTF-32. Attempting with UTF-8 fallback.")
                .Charset = "utf-8" ' Fallback
            Case "ANSI"
                .Charset = "Windows-1252" ' Common ANSI encoding for Western European languages
            Case Else ' Fallback for any unknown or unhandled detection
                Call HandleError("Unknown or unhandled encoding detected ('" & detectedEncoding & "') for '" & strFile & "'. Attempting with UTF-8 fallback.")
                .Charset = "utf-8"
        End Select
        
        On Error Resume Next ' Temporarily disable error handling for LoadFromFile to allow fallback
        .LoadFromFile strFile
        If err.Number <> 0 Then
            ' If the initial load with detected encoding fails (e.g., bad BOM detection, corrupted file, or specific ADO limitation)
            ' Try a common fallback: ANSI.
            err.Clear
            If .State = 1 Then .Close ' Close if already open but errored
            .Open ' Re-open for a fresh attempt
            .Charset = "Windows-1252" ' Try ANSI as a common fallback
            .LoadFromFile strFile
            If err.Number <> 0 Then
                ' If even ANSI fails, then it's a critical error.
                Call HandleError("Irrecoverable error reading file '" & strFile & "' with any supported encoding. Error: " & err.Description)
                .Close
                GoTo Error_Handler_Exit
            End If
        End If
        On Error GoTo Error_Handler ' Re-enable normal error handling
        
        ' Read the file content into the String variable.
        ReadFile_WithADO = .ReadText
        .Close
    End With
    
Error_Handler_Exit:
    On Error Resume Next ' Ensure no errors during cleanup
    If Not adoStream Is Nothing Then
        If adoStream.State = 1 Then adoStream.Close ' adStateOpen = 1
        Set adoStream = Nothing
    End If
    Exit Function
    
Error_Handler:
    ' Centralized error message box as per original function's spirit
    MsgBox "An error occurred while reading the file." & vbCrLf & vbCrLf & _
            "Error Number: " & err.Number & vbCrLf & _
            "Error Source: ReadFile_WithADO" & vbCrLf & _
            "Error Description: " & err.Description & vbCrLf & _
            "File: " & strFile, _
            vbCritical, "File Reading Error"
    Resume Error_Handler_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReadFile_WithoutADO
' Author    : Adapt� par Florent ALBANY
' Purpose   : Reads a text file into a string, attempting to detect its encoding
'             based on BOM (Byte Order Mark) and falling back to ANSI.
'             Note: Less robust for non-BOM UTF-8 or complex encodings than ADODB.Stream.
' Input Variables:
' ~~~~~~~~~~~~~~~~
' strFile - The full path to the file that is to be read.
'
' Returns:
' ~~~~~~~~
' String - The content of the file as a VBA String (internal Unicode/UTF-16 LE),
'          or an empty string if an error occurs, the file is not found, or is empty.
'
' Dependencies:
' ~~~~~~~~~~~~~~
' - GetFileEncoding (Function in this module to detect file encoding)
' - HandleError (external function for centralized error reporting)
'
' Usage Example:
' ~~~~~~~~~~~~~~~~
' Dim fileContent As String
' fileContent = ReadFile_WithoutADO("C:\path\to\your_file.txt")
' If fileContent <> "" Then
'     Debug.Print "Contenu du fichier : " & vbCrLf & fileContent
' Else
'     Debug.Print "Impossible de lire le fichier ou fichier vide."
' End If
'---------------------------------------------------------------------------------------
Public Function ReadFile_WithoutADO(ByVal strFile As String) As String
    On Error GoTo Error_Handler
    
    Dim fileNumber As Long
    Dim detectedEncoding As String
    Dim fileContent As String
    Dim byteCount As Long
    Dim initialBytes(0 To 3) As Byte ' To read potential BOM
    'Dim actualBytesRead As Long ' Not directly used here, GetFileEncoding handles it
    'Dim i As Long ' Not used
    Dim tempPath As String
    
    ReadFile_WithoutADO = "" ' Initialize the result to an empty string
    
    ' 1. Verify file existence and validity.
    tempPath = Dir$(strFile, vbNormal + vbHidden + vbSystem + vbArchive + vbReadOnly)
    If tempPath = vbNullString Then
        Call HandleError("Le fichier sp�cifi� n'existe pas ou le chemin est invalide : " & strFile)
        GoTo Error_Handler_Exit
    End If
    
    ' 2. Detect the file's encoding using GetFileEncoding.
    detectedEncoding = GetFileEncoding(strFile)
    
    ' Handle cases where GetFileEncoding might return an error or specific status
    Select Case detectedEncoding
        Case "Fichier non trouv�", "Erreur"
            GoTo Error_Handler_Exit
        Case "Vide"
            Exit Function ' File is empty, ReadFile_WithoutADO is already ""
    End Select
    
    fileNumber = FreeFile
    
    ' --- L'ERREUR �TAIT ICI : TOUTE LA LOGIQUE DE LECTURE DOIT �TRE DANS CE SEUL SELECT CASE ---
    Select Case detectedEncoding
        Case "UTF-8"
            ' For UTF-8, we must read in Binary mode and convert.
            ' If there's a BOM, we skip it.
            Open strFile For Binary Access Read As #fileNumber
            
            ' Skip BOM if present
            ' GetFileEncoding already handled the BOM detection and file positioning,
            ' but if we need to explicitly skip bytes here, we must do it carefully.
            ' Let's re-open from the beginning and skip.
            If LOF(fileNumber) >= 3 Then
                Get #fileNumber, , initialBytes(0)
                Get #fileNumber, , initialBytes(1)
                Get #fileNumber, , initialBytes(2)
                If Not (initialBytes(0) = &HEF And initialBytes(1) = &HBB And initialBytes(2) = &HBF) Then
                    ' No BOM, rewind to start if it was detected as UTF-8 without BOM by GetFileEncoding heuristically
                    Seek #fileNumber, 1
                End If
            End If
            
            ' Read remaining bytes (potentially less if BOM was skipped)
            byteCount = LOF(fileNumber) - (Seek(fileNumber) - 1) ' Remaining bytes
            If byteCount > 0 Then
                Dim utf8Bytes() As Byte
                ReDim utf8Bytes(0 To byteCount - 1)
                Get #fileNumber, , utf8Bytes
                
                ' Attempt to convert UTF-8 bytes to String
                fileContent = StrConv(utf8Bytes, vbUnicode) ' Converts to UTF-16 LE (VBA's internal string format)
            End If
            Close #fileNumber
            
        Case "UTF-16 LE", "UTF-16 BE"
            ' For UTF-16, Open For Input is generally better than Binary + StrConv for direct text.
            Close #fileNumber ' Close the file opened for BOM detection in GetFileEncoding
            Open strFile For Input Access Read As #fileNumber
            
            ' Skip BOM if present by reading 2 bytes
            If LOF(fileNumber) >= 2 Then
                Dim dummyChar As String * 1 ' Read 2 characters which correspond to the BOM (FF FE or FE FF)
                Input #fileNumber, dummyChar, dummyChar ' Skip the two bytes of the BOM by reading into a string (interprets as 2 chars)
                                                        ' This is a common VBA trick for BOMs in Input mode.
            End If
            
            fileContent = Input$(LOF(fileNumber), #fileNumber) ' Read remaining as text
            Close #fileNumber
            
        Case "ANSI" ' Explicitly handle ANSI
            Close #fileNumber ' Close the file used for BOM detection
            Open strFile For Input Access Read As #fileNumber
            fileContent = Input$(LOF(fileNumber), #fileNumber)
            Close #fileNumber
            
        Case Else ' Fallback for any other detected encoding or if detection failed and we default to ANSI.
            ' This handles cases like "UTF-32" or "Inconnu" from GetFileEncoding, falling back to ANSI.
            Close #fileNumber ' Close the file used for BOM detection
            Open strFile For Input Access Read As #fileNumber
            fileContent = Input$(LOF(fileNumber), #fileNumber)
            Close #fileNumber
            
    End Select ' --- CE END SELECT EST CRUCIAL POUR FERMER LE SELECT CASE PRINCIPAL ---
    
    ReadFile_WithoutADO = fileContent
    
Error_Handler_Exit:
    On Error Resume Next ' Ensure no errors during cleanup
    If fileNumber <> 0 Then Close #fileNumber ' Ensure the file is closed if it was opened
    Exit Function
    
Error_Handler:
    If fileNumber <> 0 Then Close #fileNumber
    MsgBox "An error occurred while reading the file." & vbCrLf & vbCrLf & _
            "Error Number: " & err.Number & vbCrLf & _
            "Error Source: ReadFile_WithoutADO" & vbCrLf & _
            "Error Description: " & err.Description & vbCrLf & _
            "File: " & strFile, _
            vbCritical, "File Reading Error"
    Resume Error_Handler_Exit
End Function


Public Function Copy_File(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Overwrite As Boolean = True) As Boolean
    ' @brief Copies a file from a source path to a destination path.
    ' @param SourcePath The full path of the file to copy.
    ' @param DestinationPath The full path for the new file. Can be a folder or a full new file path.
    ' @param Overwrite (Optional) If True, an existing file at the destination will be overwritten. Defaults to True.
    ' @return True if the copy was successful, False otherwise.
    ' @dependencies Microsoft Scripting Runtime, FileX.create_folder, HandleError
    On Error GoTo ifError
    
    Dim fso As Object 'Scripting.FileSystemObject
    Dim destFolder As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if source file exists
    If Not fso.FileExists(SourcePath) Then
        Call HandleError("Source file not found: " & SourcePath)
        Exit Function ' Returns False
    End If
    
    ' Ensure destination directory exists
    destFolder = fso.GetParentFolderName(DestinationPath)
    If Not fso.FolderExists(destFolder) Then
        Call create_folder(destFolder)
    End If
    
    ' Perform the copy
    fso.CopyFile Source:=SourcePath, Destination:=DestinationPath, OverwriteFiles:=Overwrite
    Copy_File = True
    
    Exit Function
ifError:
    Call HandleError("Failed to copy file from '" & SourcePath & "' to '" & DestinationPath & "'. " & vbCrLf & "Error: " & Err.Description)
    Copy_File = False
End Function


Public Function Move_File(ByVal SourcePath As String, ByVal DestinationPath As String) As Boolean
    ' @brief Moves a file from a source path to a destination path.
    ' @param SourcePath The full path of the file to move.
    ' @param DestinationPath The full path for the new file location. Can be a folder or a full new file path.
    ' @return True if the move was successful, False otherwise.
    ' @dependencies Microsoft Scripting Runtime, FileX.create_folder, HandleError
    On Error GoTo ifError
    
    Dim fso As Object 'Scripting.FileSystemObject
    Dim destFolder As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if source file exists
    If Not fso.FileExists(SourcePath) Then
        Call HandleError("Source file not found: " & SourcePath)
        Exit Function ' Returns False
    End If
    
    ' Ensure destination directory exists
    destFolder = fso.GetParentFolderName(DestinationPath)
    If Not fso.FolderExists(destFolder) Then
        Call create_folder(destFolder)
    End If
    
    ' Perform the move
    fso.MoveFile Source:=SourcePath, Destination:=DestinationPath
    Move_File = True
    
    Exit Function
ifError:
    Call HandleError("Failed to move file from '" & SourcePath & "' to '" & DestinationPath & "'. " & vbCrLf & "Error: " & Err.Description)
    Move_File = False
End Function


Public Function Get_FileSize(ByVal FilePath As String) As Double
    ' @brief Gets the size of a file in bytes.
    ' @param FilePath The full path of the file.
    ' @return The size of the file in bytes. Returns -1 if the file does not exist or an error occurs.
    ' @dependencies FileX.FileExist, HandleError
    On Error GoTo ifError
    
    ' Use existing robust file check
    If Not FileX.FileExist(FilePath) Then
        Get_FileSize = -1
        Exit Function
    End If
    
    Get_FileSize = FileLen(FilePath)
    
    Exit Function
ifError:
    Call HandleError("Could not get file size for '" & FilePath & "'. " & vbCrLf & "Error: " & Err.Description)
    Get_FileSize = -1
End Function


Public Function Combine_Paths(ByVal BasePath As String, ByVal RelativePath As String) As String
    ' @brief Safely combines two path segments into a single path, ensuring correct separator usage.
    ' @param BasePath The first part of the path (e.g., "C:\Users\Test").
    ' @param RelativePath The second part of the path to append (e.g., "MyFolder\file.txt").
    ' @return A correctly formatted full path string.
    ' @dependencies FileX.Clean_FilePath
    
    Dim cleanBasePath As String
    Dim cleanRelativePath As String
    
    ' Clean both paths to use a consistent separator
    cleanBasePath = FileX.Clean_FilePath(BasePath)
    cleanRelativePath = FileX.Clean_FilePath(RelativePath)
    
    ' Remove trailing slash from base path if it exists
    If Right(cleanBasePath, 1) = "\" Then
        cleanBasePath = Left(cleanBasePath, Len(cleanBasePath) - 1)
    End If
    
    ' Remove leading slash from relative path if it exists
    If Left(cleanRelativePath, 1) = "\" Then
        cleanRelativePath = Mid(cleanRelativePath, 2)
    End If
    
    ' Combine them with a single backslash
    Combine_Paths = cleanBasePath & "\" & cleanRelativePath
End Function


Public Sub Zip_Files(ByVal FilePaths As Variant, ByVal ZipFilePath As String)
    ' @brief Creates a zip archive containing specified files.
    ' @param FilePaths An array of full file paths (or a single path as a string) to include in the zip.
    ' @param ZipFilePath The full path for the new zip file to be created.
    ' @details This procedure will create an empty zip file and then add each specified file to it.
    '          It will overwrite an existing zip file at the destination.
    ' @dependencies Shell.Application object, FileX.FileExist, HandleError
    
    On Error GoTo ifError
    
    Dim shellApp As Object
    Dim fileItem As Variant
    Dim i As Long
    
    ' 1. Create an empty zip file
    ' This is a standard trick: create a file with a .zip extension and write the PK header.
    If FileX.FileExist(ZipFilePath) Then Kill ZipFilePath
    Open ZipFilePath For Output As #1
    Print #1, "PK" & Chr(5) & Chr(6) & String(18, 0)
    Close #1
    
    ' 2. Use Shell.Application to work with the zip file
    Set shellApp = CreateObject("Shell.Application")
    
    ' 3. Copy files into the zip folder
    If IsArray(FilePaths) Then
        For i = LBound(FilePaths) To UBound(FilePaths)
            If FileX.FileExist(CStr(FilePaths(i))) Then
                ' The Namespace method treats a .zip file as a folder
                shellApp.Namespace(ZipFilePath).CopyHere CStr(FilePaths(i))
                ' Wait for the shell to finish copying. A loop is more robust than a fixed wait.
                Do While shellApp.Namespace(ZipFilePath).Items.Count <= i
                    Application.Wait Now + TimeValue("00:00:01")
                Loop
            Else
                Debug.Print "Zip_Files: Skipping non-existent file: " & FilePaths(i)
            End If
        Next i
    ElseIf VarType(FilePaths) = vbString Then
        ' Handle a single file path passed as a string
        If FileX.FileExist(CStr(FilePaths)) Then
            shellApp.Namespace(ZipFilePath).CopyHere CStr(FilePaths)
            Application.Wait Now + TimeValue("00:00:01")
        End If
    End If
    
    ' Clean up
    Set shellApp = Nothing
    Exit Sub
    
ifError:
    Call HandleError("Failed to create zip file '" & ZipFilePath & "'. " & vbCrLf & "Error: " & Err.Description)
    If Not shellApp Is Nothing Then Set shellApp = Nothing
End Sub
