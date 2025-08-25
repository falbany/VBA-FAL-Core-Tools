Attribute VB_Name = "FalCSV"
' --------------------------------------------------------------------------------------------------
' Module: FalCSV
' Author: Florent ALBANY
' Date: 2025-07-15
' Version: 1.0
' Description:
'   Ce module fournit un ensemble de fonctions utilitaires pour la manipulation de fichiers CSV
'   (importation et exportation) au sein d'applications Excel VBA. Il s'appuie sur la classe
'   'VBABetterArray' pour une gestion efficace et performante des donn�es en m�moire,
'   facilitant ainsi les transferts entre les feuilles de calcul Excel et les formats CSV.
'
'   Les fonctionnalit�s incluent :
'   - Importation de fichiers CSV uniques ou multiples vers de nouvelles feuilles de calcul.
'   - Importation d'un fichier CSV vers une plage de cellules s�lectionn�e.
'   - Exportation d'une plage de cellules s�lectionn�e ou d'une feuille enti�re vers un fichier CSV.
'   - Options avanc�es pour la personnalisation de l'exportation CSV (d�limiteurs, encodage).
'   - Capacit� � fusionner plusieurs fichiers CSV en une seule feuille.
'   - Fonctionnalit� d'aper�u pour valider le formatage avant l'importation compl�te.
'
'   Con�u avec une approche robuste de gestion des erreurs et une interface utilisateur intuitive.
'
' Dependencies:
'   - Classe VBABetterArray (https://github.com/Senipah/VBA-Better-Array)
'   - Module FileX (pour les dialogues de s�lection/sauvegarde de fichiers)
'
' License: MIT License
'
' Change Log:
' 1.0 (2025-07-15) - Initial release with core CSV import/export functionalities.
'                   Added merge and custom export features.
'
' --------------------------------------------------------------------------------------------------
Option Explicit
' Option Base 0 ' Ou Option Base 1, selon la pr�f�rence par d�faut de tes tableaux BetterArray


Private Sub HandleError(ByVal ErrMsg As String)
    ' @brief Affiche un message d'erreur standardis� � l'utilisateur.
    ' @param ErrMsg Le message d'erreur sp�cifique � afficher.
    MsgBox "Une erreur est survenue :" & vbCrLf & ErrMsg, vbOKOnly + vbCritical, "Op�ration impossible"
End Sub


' --- Function: CSV_FromFilesToWorksheets ---
' Description:
'   Importe le contenu d'un ou plusieurs fichiers CSV s�lectionn�s par l'utilisateur
'   dans de nouvelles feuilles de calcul au sein du classeur Excel actif.
'   Chaque fichier CSV est import� dans sa propre feuille. La fonction g�re
'   l'importation des donn�es en utilisant la classe BetterArray pour l'efficacit�.
'
' Parameters: None
'
' Returns: None
'   Cette proc�dure ne retourne aucune valeur, mais elle modifie le classeur Excel
'   en ajoutant de nouvelles feuilles contenant les donn�es CSV.
'
' Usage:
'   Appelle cette proc�dure depuis n'importe quel autre module VBA ou assigne-la
'   � un bouton ou une forme dans Excel pour permettre aux utilisateurs de lancer
'   l'importation de fichiers CSV.
'   Exemple: Call FalCSV.FromCSVFilesToWorksheets
'
' Error Handling:
'   - Affiche un message si aucun fichier n'est s�lectionn� ou si l'op�ration est annul�e.
'   - G�re les erreurs individuelles lors de la lecture de chaque fichier CSV,
'     permettant de continuer l'importation des autres fichiers m�me en cas de probl�me.
'   - Fournit des messages d'erreur g�n�riques pour les erreurs inattendues.
'
' Dependencies:
'   - BetterArray: Utilis� pour parser et manipuler les donn�es CSV en m�moire.
'   - FileX.Select_Files: Pour le dialogue de s�lection de fichiers.
'
' Change Log:
' 1.0 (2025-07-15) - Impl�mentation initiale.
' --------------------------------------------------------------------------------------------------
Public Sub CSV_FromFilesToWorksheets()
    On Error GoTo ErrHandler ' Meilleure gestion d'erreurs
    Dim MyArray As BetterArray
    Dim filePaths As Variant ' Renomm� pour plus de clart�
    Dim i As Long ' Utiliser Long pour les index, plus robuste
    Dim outputSheet As Worksheet
    Dim fileNameOnly As String ' Variable pour stocker le nom du fichier sans l'extension
    Dim inputDelimiter As String
    Dim finalDelimiter As String
    Dim inputQuote As String

    Set MyArray = New BetterArray

    ' 1. Demander � l'utilisateur de s�lectionner des fichiers CSV.
    filePaths = FileX.Select_Files(FileType:="csv", AllowMultiSelect:=True)

    ' 2. V�rifier si l'utilisateur a annul� ou s'il y a eu une erreur.
    If IsEmpty(filePaths) Or (IsArray(filePaths) And UBound(filePaths) < LBound(filePaths)) Then
        MsgBox "Aucun fichier CSV s�lectionn� ou l'op�ration a �t� annul�e.", vbInformation
        GoTo CleanUp ' Sortie propre
    End If

    ' 2.1 Demander le d�limiteur initial � l'utilisateur
    inputDelimiter = InputBox("Veuillez saisir le d�limiteur � utiliserpar ex. , ou ; ou tabulation via 'TAB'):", "D�limiteur d'aper�u CSV", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 2.2 Demander le Quote initial � l'utilisateur
    inputQuote = InputBox("Veuillez saisir le caract�re d'ouverture et fermeture de cellule � utiliser (par ex. """"):", "Quote d'aper�u CSV", """")
    
    ' 3. Traiter chaque fichier s�lectionn�.
    For i = LBound(filePaths) To UBound(filePaths)
        ' G�rer les erreurs sp�cifiques � chaque fichier
        On Error Resume Next ' G�re les erreurs pour un fichier individuel (ex: fichier corrompu)
        

        MyArray.FromCSVFile path:=CStr(filePath(LBound(filePath))), columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False ' DuckType False pour voir les valeurs brutes
        If err.Number <> 0 Then
            Call HandleError("Erreur lors de la lecture du fichier CSV : " & filePaths(i) & ". " & err.Description)
            err.Clear ' Efface l'erreur pour continuer avec le fichier suivant
            GoTo NextFile ' Passe au fichier suivant
        End If
        On Error GoTo ErrHandler ' R�tablit le gestionnaire global

        Set outputSheet = ActiveWorkbook.Sheets.Add
        ' Donner un nom � la nouvelle feuille (bas� sur le nom du fichier CSV)
        On Error Resume Next ' G�rer les erreurs si le nom de la feuille est trop long ou d�j� pris
        
        ' Extraire le nom du fichier sans le chemin
        fileNameOnly = Mid(CStr(filePaths(i)), InStrRev(CStr(filePaths(i)), "\") + 1)
        ' Extraire le nom de base (sans l'extension)
        If InStr(fileNameOnly, ".") > 0 Then
            fileNameOnly = Left(fileNameOnly, InStr(fileNameOnly, ".") - 1)
        End If
        
        ' Appliquer le nom � la feuille, tronqu� � 31 caract�res
        outputSheet.name = Left(fileNameOnly, 31)
        
        If err.Number <> 0 Then
            ' Si le nom est en conflit ou trop long, Excel ajoutera un num�ro, pas grave
            err.Clear
        End If
        On Error GoTo ErrHandler ' R�tablit le gestionnaire global

        MyArray.ToExcelRange outputSheet.Range("A1")
        MyArray.Clear ' Lib�re l'array pour le prochain fichier, si la classe a une m�thode Clear
NextFile:
    Next i

    MsgBox "Les donn�es des fichiers CSV ont �t� import�es dans de nouvelles feuilles.", vbInformation

CleanUp:
    Set MyArray = Nothing ' Lib�rer l'objet
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans FromCSVFilesToWorksheets : " & err.Description)
    Resume CleanUp ' Va au nettoyage et sort
End Sub


' --- Function: CSV_FromFileToSelection ---
' Description:
'   Importe le contenu d'un fichier CSV s�lectionn� par l'utilisateur directement
'   dans une plage de cellules sp�cifi�e dans la feuille de calcul active.
'   Cette proc�dure est utile pour ins�rer des donn�es CSV � un emplacement pr�cis
'   sans cr�er une nouvelle feuille. Elle utilise la classe BetterArray pour une
'   lecture et une �criture efficaces.
'
' Parameters: None
'
' Returns: None
'   Cette proc�dure ne retourne aucune valeur. Elle modifie la feuille Excel active
'   en y �crivant les donn�es du fichier CSV.
'
' Usage:
'   1. S�lectionnez la cellule de destination (par exemple, "A1") dans votre feuille Excel.
'   2. Ex�cutez cette proc�dure. Un dialogue de s�lection de fichier s'ouvrira.
'   3. Choisissez le fichier CSV � importer.
'   Exemple d'appel depuis un autre module: Call FalCSV.FromCSVFileToSelection
'
' Error Handling:
'   - V�rifie si une plage de cellules est s�lectionn�e avant de proc�der.
'   - G�re l'annulation de la s�lection de fichier par l'utilisateur.
'   - Fournit un message d'erreur clair en cas d'erreur inattendue lors de l'importation.
'
' Dependencies:
'   - BetterArray: N�cessaire pour l'analyse du fichier CSV et l'�criture des donn�es dans Excel.
'   - FileX.Select_Files: Utilis� pour afficher le dialogue de s�lection de fichier CSV.
'
' Change Log:
' 1.0 (2025-07-15) - Impl�mentation initiale.
' --------------------------------------------------------------------------------------------------
Public Sub CSV_FromFileToSelection()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Dim rng As Range ' Correctement typ�
    Set MyArray = New BetterArray
    Dim inputDelimiter As String
    Dim finalDelimiter As String
    Dim inputQuote As String

    ' 1. V�rifier la s�lection de la plage de destination.
    If TypeName(Selection) <> "Range" Then
        MsgBox "Veuillez s�lectionner la cellule de destination o� les donn�es CSV doivent �tre �crites (par ex. A1).", vbExclamation
        GoTo CleanUp
    End If
    Set rng = Selection.Cells(1, 1) ' Prendre juste la premi�re cellule de la s�lection

    ' 2. Demander � l'utilisateur de s�lectionner un fichier CSV.
    filePath = FileX.Select_Files(FileType:="csv", AllowMultiSelect:=False)

    ' 3. V�rifier si l'utilisateur a annul� ou s'il y a eu une erreur.
    If IsEmpty(filePath) Or (IsArray(filePath) And UBound(filePath) < LBound(filePath)) Then
        MsgBox "Aucun fichier CSV s�lectionn� ou l'op�ration a �t� annul�e.", vbInformation
        GoTo CleanUp
    End If

    ' 3.1 Demander le d�limiteur initial � l'utilisateur
    inputDelimiter = InputBox("Veuillez saisir le d�limiteur � utiliserpar ex. , ou ; ou tabulation via 'TAB'):", "D�limiteur d'aper�u CSV", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 3.2 Demander le Quote initial � l'utilisateur
    inputQuote = InputBox("Veuillez saisir le caract�re d'ouverture et fermeture de cellule � utiliser (par ex. """"):", "Quote d'aper�u CSV", """")
    
    ' 4. Importer et �crire le fichier CSV.
    MyArray.FromCSVFile path:=CStr(filePath(LBound(filePath))), columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False ' DuckType False pour voir les valeurs brutes
    MyArray.ToExcelRange rng

    MsgBox "Le fichier CSV a �t� import� dans la feuille � partir de la cellule " & rng.Address(False, False) & ".", vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans FromCSVFileToSelection : " & err.Description)
    Resume CleanUp
End Sub


' --- Function: CSV_FromSelectionToFile ---
' Description:
'   Exporte les donn�es de la plage de cellules Excel actuellement s�lectionn�e
'   par l'utilisateur vers un nouveau fichier CSV. Cette fonction est id�ale
'   pour sauvegarder des sous-ensembles sp�cifiques de donn�es d'une feuille Excel
'   dans un format CSV, facilitant ainsi l'�change ou l'int�gration avec d'autres
'   applications. Elle utilise la classe BetterArray pour une extraction et une
'   �criture efficaces des donn�es.
'
' Parameters: None
'
' Returns: None
'   Cette proc�dure ne retourne aucune valeur. Elle interagit avec l'utilisateur
'   pour le chemin de sauvegarde et cr�e un fichier CSV.
'
' Usage:
'   1. S�lectionnez la plage de cellules dans Excel que vous souhaitez exporter.
'      Si seule une cellule est s�lectionn�e, la fonction exportera automatiquement
'      toute la r�gion contigu� de donn�es (`CurrentRegion`) autour de cette cellule.
'   2. Ex�cutez cette proc�dure (par exemple, via un bouton ou une macro).
'   3. Un dialogue "Enregistrer sous" s'ouvrira, vous permettant de choisir
'      le nom et l'emplacement du fichier CSV de sortie.
'   Exemple d'appel depuis un autre module: Call FalCSV.CSV_FromSelectionToFile
'
' Error Handling:
'   - Affiche un message d'erreur si aucune plage de cellules n'est s�lectionn�e.
'   - G�re l'annulation du dialogue de sauvegarde de fichier par l'utilisateur.
'   - Fournit un message d'erreur g�n�rique si une erreur inattendue survient
'     pendant le processus d'exportation.
'
' Dependencies:
'   - BetterArray: Indispensable pour la lecture des donn�es depuis la plage Excel
'     et leur formatage en sortie CSV.
'   - FileX.Get_SaveFilePath_WithDialog: Utilis� pour afficher le dialogue de sauvegarde.
'
' Change Log:
' 1.0 (2025-07-15) - Impl�mentation initiale.
' --------------------------------------------------------------------------------------------------
Public Sub CSV_FromSelectionToFile()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Dim rng As Range ' Correctement typ�
    Set MyArray = New BetterArray

    ' 1. V�rifier que l'utilisateur a s�lectionn� une plage.
    If TypeName(Selection) <> "Range" Then
        MsgBox "Veuillez s�lectionner la plage de cellules � exporter vers le fichier CSV.", vbExclamation
        GoTo CleanUp
    End If
    Set rng = Selection

    ' 2. Demander � l'utilisateur le chemin de sauvegarde.
    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "Fichiers CSV (*.csv),*.csv")
    If CStr(filePath) = "" Then ' L'utilisateur a annul�
        MsgBox "Op�ration de sauvegarde annul�e.", vbInformation
        GoTo CleanUp
    End If

    ' 3. Lire les donn�es de la plage Excel.
    ' Utilise DetectLastRow/Column=True pour s'assurer que la plage est bien d�tect�e si l'utilisateur ne s�lectionne qu'une cellule.
    ' Ou mieux, si l'objectif est la r�gion contig�e, utiliser CurrentRegion.
    MyArray.FromExcelRange FromRange:=rng.CurrentRegion, DetectLastRow:=True, DetectLastColumn:=True

    ' 4. �crire les donn�es dans le fichier CSV.
    ' Consid�re l'ajout d'options comme Headers, ColumnDelimiter si tu veux plus de contr�le.
    ' Par d�faut, BetterArray utilise la virgule et les guillemets si besoin.
    MyArray.ToCSVFile CStr(filePath)

    MsgBox "Le fichier CSV a �t� cr�� avec succ�s : " & CStr(filePath), vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans FromSelectionToCSVFile : " & err.Description)
    Resume CleanUp
End Sub

' --- Function: CSV_FromWorksheetToFile ---
' Description:
'   Exporte l'int�gralit� des donn�es utilis�es dans la feuille de calcul active
'   vers un nouveau fichier CSV. Cette fonction permet de sauvegarder rapidement
'   tout le contenu pertinent d'une feuille Excel dans un format CSV standard,
'   id�al pour le partage de donn�es ou l'int�gration avec d'autres syst�mes.
'   Elle utilise la classe BetterArray pour extraire les donn�es d'Excel et les
'   formater en CSV.
'
' Parameters: None
'
' Returns: None
'   Cette proc�dure ne retourne aucune valeur. Elle cr�e un fichier CSV � l'emplacement
'   sp�cifi� par l'utilisateur.
'
' Usage:
'   1. Assurez-vous que la feuille de calcul active contient les donn�es que vous souhaitez exporter.
'   2. Ex�cutez cette proc�dure. Un dialogue de sauvegarde de fichier s'ouvrira.
'   3. Sp�cifiez le nom et l'emplacement du fichier CSV de sortie.
'   Exemple d'appel depuis un autre module: Call FalCSV.FromWorksheetToCSVFile
'
' Error Handling:
'   - G�re l'annulation de la bo�te de dialogue de sauvegarde par l'utilisateur.
'   - Fournit un message d'erreur clair en cas d'erreur inattendue lors de l'exportation.
'
' Dependencies:
'   - BetterArray: Indispensable pour la lecture des donn�es depuis Excel et leur �criture en CSV.
'   - FileX.Get_SaveFilePath_WithDialog: Utilis� pour afficher le dialogue de s�lection du chemin de sauvegarde.
'
' Change Log:
' 1.0 (2025-07-15) - Impl�mentation initiale.
' --------------------------------------------------------------------------------------------------
Public Sub CSV_FromWorksheetToFile()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Set MyArray = New BetterArray

    ' 1. Demander � l'utilisateur le chemin de sauvegarde.
    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "Fichiers CSV (*.csv),*.csv")
    If CStr(filePath) = "" Then ' L'utilisateur a annul�
        MsgBox "Op�ration de sauvegarde annul�e.", vbInformation
        GoTo CleanUp
    End If

    ' 2. Lire les donn�es de la plage utilis�e de la feuille active.
    MyArray.FromExcelRange FromRange:=ActiveSheet.UsedRange, DetectLastRow:=True, DetectLastColumn:=True

    ' 3. �crire les donn�es dans le fichier CSV.
    MyArray.ToCSVFile CStr(filePath)

    MsgBox "Le fichier CSV a �t� cr�� avec succ�s : " & CStr(filePath), vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans FromWorksheetToCSVFile : " & err.Description)
    Resume CleanUp
End Sub


' --- Function: CSV_DumpFilesToSheets ---
' Description:
'   Importe le contenu d'un ou plusieurs fichiers CSV s�lectionn�s par l'utilisateur
'   dans des feuilles de calcul s�par�es au sein du classeur Excel actif.
'   Chaque fichier CSV est trait� individuellement et ses donn�es sont �crites
'   dans une nouvelle feuille de travail d�di�e, nomm�e d'apr�s le fichier CSV.
'   Cette fonction est une alternative � l'utilisation d'une interface CSV externe,
'   en centralisant le traitement des donn�es via la classe 'BetterArray'.
'
' Parameters: None
'
' Returns: None
'   Cette proc�dure ne retourne aucune valeur. Elle ajoute de nouvelles feuilles
'   au classeur Excel actif, chacune contenant les donn�es d'un fichier CSV import�.
'
' Usage:
'   Appelle cette proc�dure pour permettre � l'utilisateur de s�lectionner des fichiers CSV.
'   Pour chaque fichier s�lectionn�, une nouvelle feuille sera cr��e et remplie avec les donn�es.
'   Exemple d'appel depuis un autre module: Call FalCSV.DumpToSheet_UsingBetterArray
'
' Error Handling:
'   - Affiche un message si l'utilisateur annule la s�lection de fichiers ou si aucun fichier n'est choisi.
'   - G�re les erreurs individuelles lors de la lecture de chaque fichier CSV,
'     permettant de passer au fichier suivant en cas de corruption ou d'inaccessibilit�.
'   - G�re les erreurs potentielles lors du renommage des feuilles de calcul.
'   - Fournit un message d'erreur g�n�rique pour toute erreur inattendue.
'
' Dependencies:
'   - BetterArray: Essentiel pour la lecture des donn�es CSV et leur �criture dans les plages Excel.
'   - FileX.Select_Files: Pour l'interface de s�lection de fichiers.
'
' Change Log:
' 1.0 (2025-07-15) - Impl�mentation initiale, bas�e sur l'approche BetterArray.
' --------------------------------------------------------------------------------------------------
Public Sub CSV_DumpFilesToNewWorksheets() ' Renomm� pour �viter le conflit
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePaths As Variant
    Dim i As Long
    Dim outputSheet As Worksheet
    Dim fileNameOnly As String ' Variable pour stocker le nom du fichier sans l'extension
    Dim inputDelimiter As String
    Dim finalDelimiter As String
    Dim inputQuote As String


    Set MyArray = New BetterArray

    ' 1. Demander � l'utilisateur de s�lectionner des fichiers CSV.
    filePaths = FileX.Select_Files(FileType:="csv", AllowMultiSelect:=True)

    ' 2. V�rifier si l'utilisateur a annul� ou s'il y a eu une erreur.
    If IsEmpty(filePaths) Or (IsArray(filePaths) And UBound(filePaths) < LBound(filePaths)) Then
        MsgBox "Aucun fichier CSV s�lectionn� ou l'op�ration a �t� annul�e.", vbInformation
        GoTo CleanUp
    End If

    ' 2.1 Demander le d�limiteur initial � l'utilisateur
    inputDelimiter = InputBox("Veuillez saisir le d�limiteur � utiliserpar ex. , ou ; ou tabulation via 'TAB'):", "D�limiteur d'aper�u CSV", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 2.2 Demander le Quote initial � l'utilisateur
    inputQuote = InputBox("Veuillez saisir le caract�re d'ouverture et fermeture de cellule � utiliser (par ex. """"):", "Quote d'aper�u CSV", """")
    
    ' 3. Traiter chaque fichier s�lectionn�.
    For i = LBound(filePaths) To UBound(filePaths)
        On Error Resume Next
        MyArray.FromCSVFile path:=CStr(filePaths(i)), columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False ' DuckType False pour voir les valeurs brutes
        If err.Number <> 0 Then
            Call HandleError("Erreur lors de la lecture du fichier CSV : " & filePaths(i) & ". " & err.Description)
            err.Clear
            GoTo NextFile_BA
        End If
        On Error GoTo ErrHandler

        Set outputSheet = ActiveWorkbook.Sheets.Add
        On Error Resume Next
        
        ' Extraire le nom du fichier sans le chemin
        fileNameOnly = Mid(CStr(filePaths(i)), InStrRev(CStr(filePaths(i)), "\") + 1)
        ' Extraire le nom de base (sans l'extension)
        If InStr(fileNameOnly, ".") > 0 Then
            fileNameOnly = Left(fileNameOnly, InStr(fileNameOnly, ".") - 1)
        End If
        
        ' Appliquer le nom � la feuille, tronqu� � 31 caract�res
        outputSheet.name = Left(fileNameOnly, 31)
        
        If err.Number <> 0 Then err.Clear
        On Error GoTo ErrHandler

        MyArray.ToExcelRange outputSheet.Range("A1")
        MyArray.Clear ' Si BetterArray a une m�thode Clear pour vider son contenu interne
NextFile_BA:
    Next i

    MsgBox "Les donn�es des fichiers CSV ont �t� import�es dans de nouvelles feuilles (via BetterArray).", vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans DumpToSheet_UsingBetterArray : " & err.Description)
    Resume CleanUp
End Sub


Public Sub CSV_MergeFilesToNewWorksheet()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePaths As Variant
    Dim i As Long
    Dim outputSheet As Worksheet
    Dim firstFileProcessed As Boolean
    Set MyArray = New BetterArray

    ' 1. Demander � l'utilisateur de s�lectionner les fichiers CSV � fusionner.
    filePaths = FileX.Select_Files(FileType:="csv", AllowMultiSelect:=True)

    ' 2. V�rifier si des fichiers ont �t� s�lectionn�s.
    If IsEmpty(filePaths) Or (IsArray(filePaths) And UBound(filePaths) < LBound(filePaths)) Then
        MsgBox "Aucun fichier CSV s�lectionn� ou l'op�ration a �t� annul�e.", vbInformation
        GoTo CleanUp
    End If

    ' 3. Cr�er une nouvelle feuille pour la fusion.
    Set outputSheet = ActiveWorkbook.Sheets.Add
    On Error Resume Next ' G�rer les erreurs si le nom est trop long
    outputSheet.name = "Merged_CSV_Data"
    If err.Number <> 0 Then err.Clear
    On Error GoTo ErrHandler

    firstFileProcessed = False

    ' 4. Traiter chaque fichier.
    For i = LBound(filePaths) To UBound(filePaths)
        On Error Resume Next
        If Not firstFileProcessed Then
            ' Premier fichier : Importer normalement (inclut les en-t�tes)
            MyArray.FromCSVFile (CStr(filePaths(i)))
            MyArray.ToExcelRange outputSheet.Range("A1")
            firstFileProcessed = True
        Else
            ' Fichiers suivants : Ignorer la premi�re ligne (en-t�tes) et ajouter � la suite
            Dim tempArray As BetterArray
            Set tempArray = New BetterArray
            tempArray.FromCSVFile path:=CStr(filePaths(i)), IgnoreFirstRow:=True
            ' Trouver la prochaine ligne vide dans la feuille de sortie
            Dim nextRow As Long
            nextRow = outputSheet.Cells(outputSheet.Rows.count, 1).End(xlUp).Row + 1
            tempArray.ToExcelRange outputSheet.Cells(nextRow, 1)
            Set tempArray = Nothing ' Lib�rer l'objet temporaire
        End If

        If err.Number <> 0 Then
            Call HandleError("Erreur lors de la lecture ou de l'�criture du fichier : " & filePaths(i) & ". " & err.Description)
            err.Clear
        End If
        On Error GoTo ErrHandler
        MyArray.Clear ' Lib�rer MyArray pour ne pas accumuler les donn�es en m�moire inutilement
    Next i

    MsgBox "Les donn�es des fichiers CSV s�lectionn�s ont �t� fusionn�es dans la feuille : " & outputSheet.name, vbInformation

CleanUp:
    Set MyArray = Nothing
    Set outputSheet = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans MergeCSVFilesToNewWorksheet : " & err.Description)
    Resume CleanUp
End Sub


Public Sub CSV_ExportSelectedRangeAsCustom()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Dim rng As Range
    Dim columnDelimiter As String
    Dim encloseAll As VbMsgBoxResult ' Pour la bo�te de dialogue Yes/No
    Set MyArray = New BetterArray

    ' 1. V�rifier la s�lection.
    If TypeName(Selection) <> "Range" Then
        MsgBox "Veuillez s�lectionner la plage de cellules � exporter.", vbExclamation
        GoTo CleanUp
    End If
    Set rng = Selection.CurrentRegion ' Exporte la r�gion contigu� de la s�lection

    ' 2. Demander le chemin de sauvegarde.
    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "Fichiers CSV (*.csv),*.csv")
    If CStr(filePath) = "" Then
        MsgBox "Op�ration de sauvegarde annul�e.", vbInformation
        GoTo CleanUp
    End If

    ' 3. Demander le d�limiteur de colonne.
    columnDelimiter = InputBox("Veuillez saisir le d�limiteur de colonne (par ex. , ou ; ou tabulation via 'TAB'):", "D�limiteur CSV", ";")
    If columnDelimiter = "" Then
        MsgBox "D�limiteur non valide. Op�ration annul�e.", vbExclamation
        GoTo CleanUp
    ElseIf UCase(columnDelimiter) = "TAB" Then
        columnDelimiter = vbTab
    End If

    ' 4. Demander si tous les champs doivent �tre entre guillemets.
    encloseAll = MsgBox("Voulez-vous que tous les champs soient entour�s de guillemets ?", vbYesNo + vbQuestion, "Envelopper les champs")

    ' 5. Lire les donn�es.
    MyArray.FromExcelRange FromRange:=rng, DetectLastRow:=True, DetectLastColumn:=True

    ' 6. �crire les donn�es dans le fichier CSV avec les options personnalis�es.
    MyArray.ToCSVFile path:=CStr(filePath), _
                      columnDelimiter:=columnDelimiter, _
                      EncloseAllInQuotes:=(encloseAll = vbYes)

    MsgBox "Fichier CSV personnalis� cr�� avec succ�s : " & CStr(filePath), vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans ExportSelectedRangeAsCustomCSV : " & err.Description)
    Resume CleanUp
End Sub


Public Sub CSV_PreviewFile()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Dim previewSheet As Worksheet
    Dim numLinesToPreview As Long ' Nombre de lignes � afficher en aper�u
    Dim confirmed As VbMsgBoxResult
    Dim inputDelimiter As String
    Dim inputQuote As String
    Dim finalDelimiter As String
    Dim CSVString As String
    Set MyArray = New BetterArray

    numLinesToPreview = 20 ' Nombre de lignes pour l'aper�u

    ' 1. Demander � l'utilisateur de s�lectionner un fichier CSV.
    filePath = FileX.Select_Files(FileType:="csv", AllowMultiSelect:=False)
    If IsEmpty(filePath) Or (IsArray(filePath) And UBound(filePath) < LBound(filePath)) Then
        MsgBox "Aucun fichier CSV s�lectionn� ou l'op�ration a �t� annul�e.", vbInformation
        GoTo CleanUp
    End If

    ' 2. Cr�er une feuille temporaire pour l'aper�u.
    Set previewSheet = ActiveWorkbook.Sheets.Add
    On Error Resume Next
    previewSheet.name = "CSV_Preview_" & format(Now, "HHmmss")
    If err.Number <> 0 Then err.Clear
    On Error GoTo ErrHandler

    ' 3. Demander le d�limiteur initial � l'utilisateur
    inputDelimiter = InputBox("Veuillez saisir le d�limiteur � utiliser pour l'aper�u (par ex. , ou ; ou tabulation via 'TAB'):", "D�limiteur d'aper�u CSV", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 3.5 Demander le Quote initial � l'utilisateur
    inputQuote = InputBox("Veuillez saisir le caract�re d'ouverture et fermeture de cellule � utiliser pour l'aper�u (par ex. """"):", "Quote d'aper�u CSV", """")
    
    ' 4. Importer les premi�res lignes du CSV avec le d�limiteur sugg�r�.
    ' On pourrait lire le fichier ligne par ligne pour ne prendre que les X premi�res lignes,
    ' mais pour l'exemple, BetterArray.FromCSVFile importe tout et on tronque.
    ' Id�alement, BetterArray.FromCSVFile devrait avoir une option pour limiter les lignes lues
    CSVString = FileX.ReadFile_WithADO(CStr(filePath(LBound(filePath))))
    MyArray.FromCSVString CSVString:=CSVString, columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False ' DuckType False pour voir les valeurs brutes
    'MyArray.FromCSVFile path:=CStr(filePath(LBound(filePath))), columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False ' DuckType False pour voir les valeurs brutes

    ' Tronquer l'array si elle est trop grande pour l'aper�u (si BetterArray ne le fait pas nativement)
    If MyArray.UpperBound > numLinesToPreview Then
        ' Assumons une m�thode TrimRows ou une manipulation manuelle de l'array interne si n�cessaire
        ' Pour l'exemple, nous allons juste afficher les X premi�res lignes dans Excel
        Dim trimmedArray As BetterArray
        Set trimmedArray = New BetterArray
        ' Si BetterArray avait un constructeur avec un array, ce serait parfait.
        ' Puisque ce n'est pas le cas, on se basera sur ToExcelRange et on copiera-collera.
        ' C'est une limite actuelle qui pourrait �tre une suggestion d'am�lioration pour BetterArray
    End If

    ' 5. Afficher l'aper�u.
    MyArray.ToExcelRange previewSheet.Range("A1")
    previewSheet.Columns.AutoFit ' Ajuster les colonnes pour une meilleure lisibilit�

    ' 6. Demander confirmation � l'utilisateur.
    confirmed = MsgBox("L'aper�u du fichier CSV a �t� affich� dans la feuille '" & previewSheet.name & "'." & vbNewLine & _
                       "Est-ce que le formatage est correct ? Cliquer sur Non pour ajuster le d�limiteur.", _
                       vbYesNo + vbQuestion, "Confirmer le format CSV")

    If confirmed = vbNo Then
        ' L'utilisateur veut ajuster, on pourrait boucler ici ou appeler une autre fonction
        MsgBox "Veuillez relancer la fonction et essayer un autre d�limiteur.", vbInformation
        Application.DisplayAlerts = False ' Supprime la feuille sans alerte
        previewSheet.Delete
        Application.DisplayAlerts = True
    Else
        MsgBox "Aper�u valid�. Vous pouvez maintenant importer le fichier avec les param�tres choisis.", vbInformation
    End If

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans CSV_PreviewFile : " & err.Description)
    If Not previewSheet Is Nothing Then
        Application.DisplayAlerts = False
        previewSheet.Delete
        Application.DisplayAlerts = True
    End If
    Resume CleanUp
End Sub

