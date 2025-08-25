Attribute VB_Name = "CSVmod"
' --------------------------------------------------------------------------------------------------
' Module: CSVmod
' Author: Florent ALBANY
' Date: 2025-07-15
' Version: 1.0
' Description:
'   Ce module fournit un ensemble de fonctions utilitaires pour la manipulation de fichiers CSV
'   (importation et exportation) au sein d'applications Excel VBA. Il s'appuie sur la classe
'   'VBABetterArray' pour une gestion efficace et performante des données en mémoire,
'   facilitant ainsi les transferts entre les feuilles de calcul Excel et les formats CSV.
'
'   Les fonctionnalités incluent :
'   - Importation de fichiers CSV uniques ou multiples vers de nouvelles feuilles de calcul.
'   - Importation d'un fichier CSV vers une plage de cellules sélectionnée.
'   - Exportation d'une plage de cellules sélectionnée ou d'une feuille entière vers un fichier CSV.
'   - Options avancées pour la personnalisation de l'exportation CSV (délimiteurs, encodage).
'   - Capacité à fusionner plusieurs fichiers CSV en une seule feuille.
'   - Fonctionnalité d'aperçu pour valider le formatage avant l'importation complète.
'
'   Conçu avec une approche robuste de gestion des erreurs et une interface utilisateur intuitive.
'
' Dependencies:
'   - Classe VBABetterArray (https://github.com/Senipah/VBA-Better-Array)
'   - Module FileX (pour les dialogues de sélection/sauvegarde de fichiers)
'
' License: MIT License
'
' Change Log:
' 1.0 (2025-07-15) - Initial release with core CSV import/export functionalities.
'                   Added merge and custom export features.
'
' --------------------------------------------------------------------------------------------------
Option Explicit
' Option Base 0 ' Ou Option Base 1, selon la préférence par défaut de tes tableaux BetterArray


Private Sub HandleError(ByVal ErrMsg As String)
    ' @brief Affiche un message d'erreur standardisé à l'utilisateur.
    ' @param ErrMsg Le message d'erreur spécifique à afficher.
    MsgBox "Une erreur est survenue :" & vbCrLf & ErrMsg, vbOKOnly + vbCritical, "Opération impossible"
End Sub


' --- Function: CSV_FromFilesToWorksheets ---
' Description:
'   Importe le contenu d'un ou plusieurs fichiers CSV sélectionnés par l'utilisateur
'   dans de nouvelles feuilles de calcul au sein du classeur Excel actif.
'   Chaque fichier CSV est importé dans sa propre feuille. La fonction gère
'   l'importation des données en utilisant la classe BetterArray pour l'efficacité.
'
' Parameters: None
'
' Returns: None
'   Cette procédure ne retourne aucune valeur, mais elle modifie le classeur Excel
'   en ajoutant de nouvelles feuilles contenant les données CSV.
'
' Usage:
'   Appelle cette procédure depuis n'importe quel autre module VBA ou assigne-la
'   à un bouton ou une forme dans Excel pour permettre aux utilisateurs de lancer
'   l'importation de fichiers CSV.
'   Exemple: Call CSVmod.FromCSVFilesToWorksheets
'
' Error Handling:
'   - Affiche un message si aucun fichier n'est sélectionné ou si l'opération est annulée.
'   - Gère les erreurs individuelles lors de la lecture de chaque fichier CSV,
'     permettant de continuer l'importation des autres fichiers même en cas de problème.
'   - Fournit des messages d'erreur génériques pour les erreurs inattendues.
'
' Dependencies:
'   - BetterArray: Utilisé pour parser et manipuler les données CSV en mémoire.
'   - FileX.Select_Files: Pour le dialogue de sélection de fichiers.
'
' Change Log:
' 1.0 (2025-07-15) - Implémentation initiale.
' --------------------------------------------------------------------------------------------------
Public Sub CSV_FromFilesToWorksheets()
    On Error GoTo ErrHandler ' Meilleure gestion d'erreurs
    Dim MyArray As BetterArray
    Dim filePaths As Variant ' Renommé pour plus de clarté
    Dim i As Long ' Utiliser Long pour les index, plus robuste
    Dim outputSheet As Worksheet
    Dim fileNameOnly As String ' Variable pour stocker le nom du fichier sans l'extension
    Dim inputDelimiter As String
    Dim finalDelimiter As String
    Dim inputQuote As String

    Set MyArray = New BetterArray

    ' 1. Demander à l'utilisateur de sélectionner des fichiers CSV.
    filePaths = FileX.Select_Files(FileType:="csv", AllowMultiSelect:=True)

    ' 2. Vérifier si l'utilisateur a annulé ou s'il y a eu une erreur.
    If IsEmpty(filePaths) Or (IsArray(filePaths) And UBound(filePaths) < LBound(filePaths)) Then
        MsgBox "Aucun fichier CSV sélectionné ou l'opération a été annulée.", vbInformation
        GoTo CleanUp ' Sortie propre
    End If

    ' 2.1 Demander le délimiteur initial à l'utilisateur
    inputDelimiter = InputBox("Veuillez saisir le délimiteur à utiliserpar ex. , ou ; ou tabulation via 'TAB'):", "Délimiteur d'aperçu CSV", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 2.2 Demander le Quote initial à l'utilisateur
    inputQuote = InputBox("Veuillez saisir le caractère d'ouverture et fermeture de cellule à utiliser (par ex. """"):", "Quote d'aperçu CSV", """")
    
    ' 3. Traiter chaque fichier sélectionné.
    For i = LBound(filePaths) To UBound(filePaths)
        ' Gérer les erreurs spécifiques à chaque fichier
        On Error Resume Next ' Gère les erreurs pour un fichier individuel (ex: fichier corrompu)
        

        MyArray.FromCSVFile path:=CStr(filePath(LBound(filePath))), columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False ' DuckType False pour voir les valeurs brutes
        If err.Number <> 0 Then
            Call HandleError("Erreur lors de la lecture du fichier CSV : " & filePaths(i) & ". " & err.Description)
            err.Clear ' Efface l'erreur pour continuer avec le fichier suivant
            GoTo NextFile ' Passe au fichier suivant
        End If
        On Error GoTo ErrHandler ' Rétablit le gestionnaire global

        Set outputSheet = ActiveWorkbook.Sheets.Add
        ' Donner un nom à la nouvelle feuille (basé sur le nom du fichier CSV)
        On Error Resume Next ' Gérer les erreurs si le nom de la feuille est trop long ou déjà pris
        
        ' Extraire le nom du fichier sans le chemin
        fileNameOnly = Mid(CStr(filePaths(i)), InStrRev(CStr(filePaths(i)), "\") + 1)
        ' Extraire le nom de base (sans l'extension)
        If InStr(fileNameOnly, ".") > 0 Then
            fileNameOnly = Left(fileNameOnly, InStr(fileNameOnly, ".") - 1)
        End If
        
        ' Appliquer le nom à la feuille, tronqué à 31 caractères
        outputSheet.name = Left(fileNameOnly, 31)
        
        If err.Number <> 0 Then
            ' Si le nom est en conflit ou trop long, Excel ajoutera un numéro, pas grave
            err.Clear
        End If
        On Error GoTo ErrHandler ' Rétablit le gestionnaire global

        MyArray.ToExcelRange outputSheet.Range("A1")
        MyArray.Clear ' Libère l'array pour le prochain fichier, si la classe a une méthode Clear
NextFile:
    Next i

    MsgBox "Les données des fichiers CSV ont été importées dans de nouvelles feuilles.", vbInformation

CleanUp:
    Set MyArray = Nothing ' Libérer l'objet
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans FromCSVFilesToWorksheets : " & err.Description)
    Resume CleanUp ' Va au nettoyage et sort
End Sub


' --- Function: CSV_FromFileToSelection ---
' Description:
'   Importe le contenu d'un fichier CSV sélectionné par l'utilisateur directement
'   dans une plage de cellules spécifiée dans la feuille de calcul active.
'   Cette procédure est utile pour insérer des données CSV à un emplacement précis
'   sans créer une nouvelle feuille. Elle utilise la classe BetterArray pour une
'   lecture et une écriture efficaces.
'
' Parameters: None
'
' Returns: None
'   Cette procédure ne retourne aucune valeur. Elle modifie la feuille Excel active
'   en y écrivant les données du fichier CSV.
'
' Usage:
'   1. Sélectionnez la cellule de destination (par exemple, "A1") dans votre feuille Excel.
'   2. Exécutez cette procédure. Un dialogue de sélection de fichier s'ouvrira.
'   3. Choisissez le fichier CSV à importer.
'   Exemple d'appel depuis un autre module: Call CSVmod.FromCSVFileToSelection
'
' Error Handling:
'   - Vérifie si une plage de cellules est sélectionnée avant de procéder.
'   - Gère l'annulation de la sélection de fichier par l'utilisateur.
'   - Fournit un message d'erreur clair en cas d'erreur inattendue lors de l'importation.
'
' Dependencies:
'   - BetterArray: Nécessaire pour l'analyse du fichier CSV et l'écriture des données dans Excel.
'   - FileX.Select_Files: Utilisé pour afficher le dialogue de sélection de fichier CSV.
'
' Change Log:
' 1.0 (2025-07-15) - Implémentation initiale.
' --------------------------------------------------------------------------------------------------
Public Sub CSV_FromFileToSelection()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Dim rng As Range ' Correctement typé
    Set MyArray = New BetterArray
    Dim inputDelimiter As String
    Dim finalDelimiter As String
    Dim inputQuote As String

    ' 1. Vérifier la sélection de la plage de destination.
    If TypeName(Selection) <> "Range" Then
        MsgBox "Veuillez sélectionner la cellule de destination où les données CSV doivent être écrites (par ex. A1).", vbExclamation
        GoTo CleanUp
    End If
    Set rng = Selection.Cells(1, 1) ' Prendre juste la première cellule de la sélection

    ' 2. Demander à l'utilisateur de sélectionner un fichier CSV.
    filePath = FileX.Select_Files(FileType:="csv", AllowMultiSelect:=False)

    ' 3. Vérifier si l'utilisateur a annulé ou s'il y a eu une erreur.
    If IsEmpty(filePath) Or (IsArray(filePath) And UBound(filePath) < LBound(filePath)) Then
        MsgBox "Aucun fichier CSV sélectionné ou l'opération a été annulée.", vbInformation
        GoTo CleanUp
    End If

    ' 3.1 Demander le délimiteur initial à l'utilisateur
    inputDelimiter = InputBox("Veuillez saisir le délimiteur à utiliserpar ex. , ou ; ou tabulation via 'TAB'):", "Délimiteur d'aperçu CSV", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 3.2 Demander le Quote initial à l'utilisateur
    inputQuote = InputBox("Veuillez saisir le caractère d'ouverture et fermeture de cellule à utiliser (par ex. """"):", "Quote d'aperçu CSV", """")
    
    ' 4. Importer et écrire le fichier CSV.
    MyArray.FromCSVFile path:=CStr(filePath(LBound(filePath))), columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False ' DuckType False pour voir les valeurs brutes
    MyArray.ToExcelRange rng

    MsgBox "Le fichier CSV a été importé dans la feuille à partir de la cellule " & rng.Address(False, False) & ".", vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans FromCSVFileToSelection : " & err.Description)
    Resume CleanUp
End Sub


' --- Function: CSV_FromSelectionToFile ---
' Description:
'   Exporte les données de la plage de cellules Excel actuellement sélectionnée
'   par l'utilisateur vers un nouveau fichier CSV. Cette fonction est idéale
'   pour sauvegarder des sous-ensembles spécifiques de données d'une feuille Excel
'   dans un format CSV, facilitant ainsi l'échange ou l'intégration avec d'autres
'   applications. Elle utilise la classe BetterArray pour une extraction et une
'   écriture efficaces des données.
'
' Parameters: None
'
' Returns: None
'   Cette procédure ne retourne aucune valeur. Elle interagit avec l'utilisateur
'   pour le chemin de sauvegarde et crée un fichier CSV.
'
' Usage:
'   1. Sélectionnez la plage de cellules dans Excel que vous souhaitez exporter.
'      Si seule une cellule est sélectionnée, la fonction exportera automatiquement
'      toute la région contiguë de données (`CurrentRegion`) autour de cette cellule.
'   2. Exécutez cette procédure (par exemple, via un bouton ou une macro).
'   3. Un dialogue "Enregistrer sous" s'ouvrira, vous permettant de choisir
'      le nom et l'emplacement du fichier CSV de sortie.
'   Exemple d'appel depuis un autre module: Call CSVmod.CSV_FromSelectionToFile
'
' Error Handling:
'   - Affiche un message d'erreur si aucune plage de cellules n'est sélectionnée.
'   - Gère l'annulation du dialogue de sauvegarde de fichier par l'utilisateur.
'   - Fournit un message d'erreur générique si une erreur inattendue survient
'     pendant le processus d'exportation.
'
' Dependencies:
'   - BetterArray: Indispensable pour la lecture des données depuis la plage Excel
'     et leur formatage en sortie CSV.
'   - FileX.Get_SaveFilePath_WithDialog: Utilisé pour afficher le dialogue de sauvegarde.
'
' Change Log:
' 1.0 (2025-07-15) - Implémentation initiale.
' --------------------------------------------------------------------------------------------------
Public Sub CSV_FromSelectionToFile()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Dim rng As Range ' Correctement typé
    Set MyArray = New BetterArray

    ' 1. Vérifier que l'utilisateur a sélectionné une plage.
    If TypeName(Selection) <> "Range" Then
        MsgBox "Veuillez sélectionner la plage de cellules à exporter vers le fichier CSV.", vbExclamation
        GoTo CleanUp
    End If
    Set rng = Selection

    ' 2. Demander à l'utilisateur le chemin de sauvegarde.
    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "Fichiers CSV (*.csv),*.csv")
    If CStr(filePath) = "" Then ' L'utilisateur a annulé
        MsgBox "Opération de sauvegarde annulée.", vbInformation
        GoTo CleanUp
    End If

    ' 3. Lire les données de la plage Excel.
    ' Utilise DetectLastRow/Column=True pour s'assurer que la plage est bien détectée si l'utilisateur ne sélectionne qu'une cellule.
    ' Ou mieux, si l'objectif est la région contigüe, utiliser CurrentRegion.
    MyArray.FromExcelRange FromRange:=rng.CurrentRegion, DetectLastRow:=True, DetectLastColumn:=True

    ' 4. Écrire les données dans le fichier CSV.
    ' Considère l'ajout d'options comme Headers, ColumnDelimiter si tu veux plus de contrôle.
    ' Par défaut, BetterArray utilise la virgule et les guillemets si besoin.
    MyArray.ToCSVFile CStr(filePath)

    MsgBox "Le fichier CSV a été créé avec succès : " & CStr(filePath), vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans FromSelectionToCSVFile : " & err.Description)
    Resume CleanUp
End Sub

' --- Function: CSV_FromWorksheetToFile ---
' Description:
'   Exporte l'intégralité des données utilisées dans la feuille de calcul active
'   vers un nouveau fichier CSV. Cette fonction permet de sauvegarder rapidement
'   tout le contenu pertinent d'une feuille Excel dans un format CSV standard,
'   idéal pour le partage de données ou l'intégration avec d'autres systèmes.
'   Elle utilise la classe BetterArray pour extraire les données d'Excel et les
'   formater en CSV.
'
' Parameters: None
'
' Returns: None
'   Cette procédure ne retourne aucune valeur. Elle crée un fichier CSV à l'emplacement
'   spécifié par l'utilisateur.
'
' Usage:
'   1. Assurez-vous que la feuille de calcul active contient les données que vous souhaitez exporter.
'   2. Exécutez cette procédure. Un dialogue de sauvegarde de fichier s'ouvrira.
'   3. Spécifiez le nom et l'emplacement du fichier CSV de sortie.
'   Exemple d'appel depuis un autre module: Call CSVmod.FromWorksheetToCSVFile
'
' Error Handling:
'   - Gère l'annulation de la boîte de dialogue de sauvegarde par l'utilisateur.
'   - Fournit un message d'erreur clair en cas d'erreur inattendue lors de l'exportation.
'
' Dependencies:
'   - BetterArray: Indispensable pour la lecture des données depuis Excel et leur écriture en CSV.
'   - FileX.Get_SaveFilePath_WithDialog: Utilisé pour afficher le dialogue de sélection du chemin de sauvegarde.
'
' Change Log:
' 1.0 (2025-07-15) - Implémentation initiale.
' --------------------------------------------------------------------------------------------------
Public Sub CSV_FromWorksheetToFile()
    On Error GoTo ErrHandler
    Dim MyArray As BetterArray
    Dim filePath As Variant
    Set MyArray = New BetterArray

    ' 1. Demander à l'utilisateur le chemin de sauvegarde.
    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "Fichiers CSV (*.csv),*.csv")
    If CStr(filePath) = "" Then ' L'utilisateur a annulé
        MsgBox "Opération de sauvegarde annulée.", vbInformation
        GoTo CleanUp
    End If

    ' 2. Lire les données de la plage utilisée de la feuille active.
    MyArray.FromExcelRange FromRange:=ActiveSheet.UsedRange, DetectLastRow:=True, DetectLastColumn:=True

    ' 3. Écrire les données dans le fichier CSV.
    MyArray.ToCSVFile CStr(filePath)

    MsgBox "Le fichier CSV a été créé avec succès : " & CStr(filePath), vbInformation

CleanUp:
    Set MyArray = Nothing
    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans FromWorksheetToCSVFile : " & err.Description)
    Resume CleanUp
End Sub


' --- Function: CSV_DumpFilesToSheets ---
' Description:
'   Importe le contenu d'un ou plusieurs fichiers CSV sélectionnés par l'utilisateur
'   dans des feuilles de calcul séparées au sein du classeur Excel actif.
'   Chaque fichier CSV est traité individuellement et ses données sont écrites
'   dans une nouvelle feuille de travail dédiée, nommée d'après le fichier CSV.
'   Cette fonction est une alternative à l'utilisation d'une interface CSV externe,
'   en centralisant le traitement des données via la classe 'BetterArray'.
'
' Parameters: None
'
' Returns: None
'   Cette procédure ne retourne aucune valeur. Elle ajoute de nouvelles feuilles
'   au classeur Excel actif, chacune contenant les données d'un fichier CSV importé.
'
' Usage:
'   Appelle cette procédure pour permettre à l'utilisateur de sélectionner des fichiers CSV.
'   Pour chaque fichier sélectionné, une nouvelle feuille sera créée et remplie avec les données.
'   Exemple d'appel depuis un autre module: Call CSVmod.DumpToSheet_UsingBetterArray
'
' Error Handling:
'   - Affiche un message si l'utilisateur annule la sélection de fichiers ou si aucun fichier n'est choisi.
'   - Gère les erreurs individuelles lors de la lecture de chaque fichier CSV,
'     permettant de passer au fichier suivant en cas de corruption ou d'inaccessibilité.
'   - Gère les erreurs potentielles lors du renommage des feuilles de calcul.
'   - Fournit un message d'erreur générique pour toute erreur inattendue.
'
' Dependencies:
'   - BetterArray: Essentiel pour la lecture des données CSV et leur écriture dans les plages Excel.
'   - FileX.Select_Files: Pour l'interface de sélection de fichiers.
'
' Change Log:
' 1.0 (2025-07-15) - Implémentation initiale, basée sur l'approche BetterArray.
' --------------------------------------------------------------------------------------------------
Public Sub CSV_DumpFilesToNewWorksheets() ' Renommé pour éviter le conflit
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

    ' 1. Demander à l'utilisateur de sélectionner des fichiers CSV.
    filePaths = FileX.Select_Files(FileType:="csv", AllowMultiSelect:=True)

    ' 2. Vérifier si l'utilisateur a annulé ou s'il y a eu une erreur.
    If IsEmpty(filePaths) Or (IsArray(filePaths) And UBound(filePaths) < LBound(filePaths)) Then
        MsgBox "Aucun fichier CSV sélectionné ou l'opération a été annulée.", vbInformation
        GoTo CleanUp
    End If

    ' 2.1 Demander le délimiteur initial à l'utilisateur
    inputDelimiter = InputBox("Veuillez saisir le délimiteur à utiliserpar ex. , ou ; ou tabulation via 'TAB'):", "Délimiteur d'aperçu CSV", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 2.2 Demander le Quote initial à l'utilisateur
    inputQuote = InputBox("Veuillez saisir le caractère d'ouverture et fermeture de cellule à utiliser (par ex. """"):", "Quote d'aperçu CSV", """")
    
    ' 3. Traiter chaque fichier sélectionné.
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
        
        ' Appliquer le nom à la feuille, tronqué à 31 caractères
        outputSheet.name = Left(fileNameOnly, 31)
        
        If err.Number <> 0 Then err.Clear
        On Error GoTo ErrHandler

        MyArray.ToExcelRange outputSheet.Range("A1")
        MyArray.Clear ' Si BetterArray a une méthode Clear pour vider son contenu interne
NextFile_BA:
    Next i

    MsgBox "Les données des fichiers CSV ont été importées dans de nouvelles feuilles (via BetterArray).", vbInformation

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

    ' 1. Demander à l'utilisateur de sélectionner les fichiers CSV à fusionner.
    filePaths = FileX.Select_Files(FileType:="csv", AllowMultiSelect:=True)

    ' 2. Vérifier si des fichiers ont été sélectionnés.
    If IsEmpty(filePaths) Or (IsArray(filePaths) And UBound(filePaths) < LBound(filePaths)) Then
        MsgBox "Aucun fichier CSV sélectionné ou l'opération a été annulée.", vbInformation
        GoTo CleanUp
    End If

    ' 3. Créer une nouvelle feuille pour la fusion.
    Set outputSheet = ActiveWorkbook.Sheets.Add
    On Error Resume Next ' Gérer les erreurs si le nom est trop long
    outputSheet.name = "Merged_CSV_Data"
    If err.Number <> 0 Then err.Clear
    On Error GoTo ErrHandler

    firstFileProcessed = False

    ' 4. Traiter chaque fichier.
    For i = LBound(filePaths) To UBound(filePaths)
        On Error Resume Next
        If Not firstFileProcessed Then
            ' Premier fichier : Importer normalement (inclut les en-têtes)
            MyArray.FromCSVFile (CStr(filePaths(i)))
            MyArray.ToExcelRange outputSheet.Range("A1")
            firstFileProcessed = True
        Else
            ' Fichiers suivants : Ignorer la première ligne (en-têtes) et ajouter à la suite
            Dim tempArray As BetterArray
            Set tempArray = New BetterArray
            tempArray.FromCSVFile path:=CStr(filePaths(i)), IgnoreFirstRow:=True
            ' Trouver la prochaine ligne vide dans la feuille de sortie
            Dim nextRow As Long
            nextRow = outputSheet.Cells(outputSheet.Rows.count, 1).End(xlUp).Row + 1
            tempArray.ToExcelRange outputSheet.Cells(nextRow, 1)
            Set tempArray = Nothing ' Libérer l'objet temporaire
        End If

        If err.Number <> 0 Then
            Call HandleError("Erreur lors de la lecture ou de l'écriture du fichier : " & filePaths(i) & ". " & err.Description)
            err.Clear
        End If
        On Error GoTo ErrHandler
        MyArray.Clear ' Libérer MyArray pour ne pas accumuler les données en mémoire inutilement
    Next i

    MsgBox "Les données des fichiers CSV sélectionnés ont été fusionnées dans la feuille : " & outputSheet.name, vbInformation

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
    Dim encloseAll As VbMsgBoxResult ' Pour la boîte de dialogue Yes/No
    Set MyArray = New BetterArray

    ' 1. Vérifier la sélection.
    If TypeName(Selection) <> "Range" Then
        MsgBox "Veuillez sélectionner la plage de cellules à exporter.", vbExclamation
        GoTo CleanUp
    End If
    Set rng = Selection.CurrentRegion ' Exporte la région contiguë de la sélection

    ' 2. Demander le chemin de sauvegarde.
    filePath = FileX.Get_SaveFilePath_WithDialog(ActiveSheet.name, ".csv", "Fichiers CSV (*.csv),*.csv")
    If CStr(filePath) = "" Then
        MsgBox "Opération de sauvegarde annulée.", vbInformation
        GoTo CleanUp
    End If

    ' 3. Demander le délimiteur de colonne.
    columnDelimiter = InputBox("Veuillez saisir le délimiteur de colonne (par ex. , ou ; ou tabulation via 'TAB'):", "Délimiteur CSV", ";")
    If columnDelimiter = "" Then
        MsgBox "Délimiteur non valide. Opération annulée.", vbExclamation
        GoTo CleanUp
    ElseIf UCase(columnDelimiter) = "TAB" Then
        columnDelimiter = vbTab
    End If

    ' 4. Demander si tous les champs doivent être entre guillemets.
    encloseAll = MsgBox("Voulez-vous que tous les champs soient entourés de guillemets ?", vbYesNo + vbQuestion, "Envelopper les champs")

    ' 5. Lire les données.
    MyArray.FromExcelRange FromRange:=rng, DetectLastRow:=True, DetectLastColumn:=True

    ' 6. Écrire les données dans le fichier CSV avec les options personnalisées.
    MyArray.ToCSVFile path:=CStr(filePath), _
                      columnDelimiter:=columnDelimiter, _
                      EncloseAllInQuotes:=(encloseAll = vbYes)

    MsgBox "Fichier CSV personnalisé créé avec succès : " & CStr(filePath), vbInformation

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
    Dim numLinesToPreview As Long ' Nombre de lignes à afficher en aperçu
    Dim confirmed As VbMsgBoxResult
    Dim inputDelimiter As String
    Dim inputQuote As String
    Dim finalDelimiter As String
    Dim CSVString As String
    Set MyArray = New BetterArray

    numLinesToPreview = 20 ' Nombre de lignes pour l'aperçu

    ' 1. Demander à l'utilisateur de sélectionner un fichier CSV.
    filePath = FileX.Select_Files(FileType:="csv", AllowMultiSelect:=False)
    If IsEmpty(filePath) Or (IsArray(filePath) And UBound(filePath) < LBound(filePath)) Then
        MsgBox "Aucun fichier CSV sélectionné ou l'opération a été annulée.", vbInformation
        GoTo CleanUp
    End If

    ' 2. Créer une feuille temporaire pour l'aperçu.
    Set previewSheet = ActiveWorkbook.Sheets.Add
    On Error Resume Next
    previewSheet.name = "CSV_Preview_" & format(Now, "HHmmss")
    If err.Number <> 0 Then err.Clear
    On Error GoTo ErrHandler

    ' 3. Demander le délimiteur initial à l'utilisateur
    inputDelimiter = InputBox("Veuillez saisir le délimiteur à utiliser pour l'aperçu (par ex. , ou ; ou tabulation via 'TAB'):", "Délimiteur d'aperçu CSV", ",")
    If UCase(inputDelimiter) = "TAB" Then
        finalDelimiter = vbTab
    Else
        finalDelimiter = inputDelimiter
    End If

    ' 3.5 Demander le Quote initial à l'utilisateur
    inputQuote = InputBox("Veuillez saisir le caractère d'ouverture et fermeture de cellule à utiliser pour l'aperçu (par ex. """"):", "Quote d'aperçu CSV", """")
    
    ' 4. Importer les premières lignes du CSV avec le délimiteur suggéré.
    ' On pourrait lire le fichier ligne par ligne pour ne prendre que les X premières lignes,
    ' mais pour l'exemple, BetterArray.FromCSVFile importe tout et on tronque.
    ' Idéalement, BetterArray.FromCSVFile devrait avoir une option pour limiter les lignes lues
    CSVString = FileX.ReadFile_WithADO(CStr(filePath(LBound(filePath))))
    MyArray.FromCSVString CSVString:=CSVString, columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False ' DuckType False pour voir les valeurs brutes
    'MyArray.FromCSVFile path:=CStr(filePath(LBound(filePath))), columnDelimiter:=finalDelimiter, Quote:=inputQuote, DuckType:=False ' DuckType False pour voir les valeurs brutes

    ' Tronquer l'array si elle est trop grande pour l'aperçu (si BetterArray ne le fait pas nativement)
    If MyArray.UpperBound > numLinesToPreview Then
        ' Assumons une méthode TrimRows ou une manipulation manuelle de l'array interne si nécessaire
        ' Pour l'exemple, nous allons juste afficher les X premières lignes dans Excel
        Dim trimmedArray As BetterArray
        Set trimmedArray = New BetterArray
        ' Si BetterArray avait un constructeur avec un array, ce serait parfait.
        ' Puisque ce n'est pas le cas, on se basera sur ToExcelRange et on copiera-collera.
        ' C'est une limite actuelle qui pourrait être une suggestion d'amélioration pour BetterArray
    End If

    ' 5. Afficher l'aperçu.
    MyArray.ToExcelRange previewSheet.Range("A1")
    previewSheet.Columns.AutoFit ' Ajuster les colonnes pour une meilleure lisibilité

    ' 6. Demander confirmation à l'utilisateur.
    confirmed = MsgBox("L'aperçu du fichier CSV a été affiché dans la feuille '" & previewSheet.name & "'." & vbNewLine & _
                       "Est-ce que le formatage est correct ? Cliquer sur Non pour ajuster le délimiteur.", _
                       vbYesNo + vbQuestion, "Confirmer le format CSV")

    If confirmed = vbNo Then
        ' L'utilisateur veut ajuster, on pourrait boucler ici ou appeler une autre fonction
        MsgBox "Veuillez relancer la fonction et essayer un autre délimiteur.", vbInformation
        Application.DisplayAlerts = False ' Supprime la feuille sans alerte
        previewSheet.Delete
        Application.DisplayAlerts = True
    Else
        MsgBox "Aperçu validé. Vous pouvez maintenant importer le fichier avec les paramètres choisis.", vbInformation
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

