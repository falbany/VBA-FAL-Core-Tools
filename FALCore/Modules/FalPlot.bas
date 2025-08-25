Attribute VB_Name = "FalPlot"
' Module: FalPlot
' Author: Florent ALBANY
' Date: 2025-07-15
' Version: 2025.7
'
' Description:
' Ce module VBA offre un ensemble complet de fonctions pour la création, la manipulation et le
' formatage avancé de graphiques dans Microsoft Excel. Il est conçu pour simplifier le traçage
' de données à partir de plages ou de feuilles de calcul, en offrant une grande flexibilité
' de personnalisation, notamment pour les titres, les axes, les séries et les styles visuels.
'
' Fonctions principales:
' - Plot_Range: Crée un graphique à partir d'une plage de données spécifiée.
' - DataTrace: Trace automatiquement des données depuis une feuille de calcul, gérant la
'              détection de tableaux et la création de multiples graphiques par série si nécessaire.
' - Trace_Range: Fonction d'enveloppe pour le traçage de plages, potentiellement à partir d'un tableau 2D.
' - ChartFormatting: Applique une multitude d'options de formatage à un objet Chart, y compris
'                    les titres, les échelles d'axes, les légendes, les styles de série et plus encore.
' - ChartText_Format: Formate spécifiquement les indices et exposants dans les titres d'axes pour une meilleure lisibilité.
'
' Dépendances:
' Ce module s'appuie sur des fonctions auxiliaires et d'autres modules (LANG_MOD, EXCEL_MOD, ArrayX)
' pour la manipulation de chaînes, la gestion d'Excel et les opérations sur les tableaux.
'
' Remarques:
' Le module inclut des options de débogage et des gestions d'erreurs pour une robustesse accrue.
' Il vise à automatiser les tâches de plotting répétitives, permettant une visualisation rapide et personnalisée des données.
'
'
'#############################################################################################################################
' Partie 1 : Procédures Publiques (Public Sub)
' ---
' Ces procédures sont les points d'entrée principaux du module,
' accessibles depuis l'extérieur pour exécuter des actions
' spécifiques de traçage et de gestion des graphiques.
'#############################################################################################################################

Public Sub Plot_SelectedRangeWithFormatting()
    '* @brief Crée et formate un graphique à partir de la plage de cellules sélectionnée.
    '* @details Cette procédure demande à l'utilisateur un nom de graphique et des options de formatage,
    '           puis utilise la fonction Plot_RangeWithFormatting pour créer le graphique.
    '* @remarks Affiche des messages d'erreur si aucune plage n'est sélectionnée ou si le traçage échoue.
    On Error GoTo ErrHandler

    Dim selectedRange As Range
    Dim chartTitle As String
    Dim formattingOpts As String
    Dim myChart As Chart

    ' 1. Vérifier si une plage est sélectionnée
    If TypeName(Selection) <> "Range" Then
        Call HandleError("Veuillez sélectionner une plage de cellules à tracer.")
        Exit Sub
    End If
    Set selectedRange = Selection

    ' 2. Demander le nom du graphique et les options de formatage à l'utilisateur
    chartTitle = InputBox("Entrez le titre du graphique (laissez vide pour automatique) :", "Titre du Graphique")
    formattingOpts = InputBox("Entrez les options de formatage (ex: Title=MonTitre;XTitle=Abcisse;YTitle=Unités;ChartType=75) :", "Options de Formatage", "Title=" & chartTitle & ";YTitle=N/A;XTitle=N/A;ChartType=75;PlotBy=-1;AutoLegend=1")

    ' 3. Appeler la fonction Plot_RangeWithFormatting
    Set myChart = Plot_RangeWithFormatting(selectedRange, chartTitle, , , formattingOpts) ' nbSeriesByGroup et ColorStyle sont optionnels ici

    ' 4. Vérifier le succès et informer l'utilisateur
    If myChart Is Nothing Then
        ' Le message d'erreur est déjà géré par Plot_RangeWithFormatting ou HandleError
        Exit Sub
    Else
        MsgBox "Le graphique '" & myChart.name & "' a été créé et formaté avec succès !", vbInformation, "Opération Réussie"
    End If

    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans Plot_SelectedRangeWithFormatting.")
End Sub

Public Sub Create_SmithChart()
    '* @brief Crée un graphique de Smith basé sur le graphique sélectionné.
    '* @details Cette procédure effectue les étapes suivantes:
    '   1. Vérifie si un graphique est sélectionné.
    '   2. Copie le graphique sélectionné.
    '   3. Supprime les séries non pertinentes (non réelles/imaginaires).
    '   4. Applique le formatage spécifique au graphique de Smith.
    '   5. Délie les données du graphique copié pour les rendre statiques.
    '* @remarks Affiche un message d'erreur si une étape échoue.
    On Error GoTo ErrHandler

    Dim srcChart As Chart
    Dim myChart As Chart

    Set srcChart = GetSelectedChart() ' Utilise la fonction d'aide
    If srcChart Is Nothing Then Exit Sub ' Le message d'erreur est géré par GetSelectedChart

    Set myChart = Copy_Chart(srcChart)
    If myChart Is Nothing Then Call HandleError("La copie du graphique a échoué."): Exit Sub

    If Not delete_UnMatchingSeries(myChart, ":", False) Then Call HandleError("Le nettoyage des données Réel/Imaginaire a échoué."): Exit Sub
    If Not SmithChart_Formatting(myChart, True) Then Call HandleError("Le formatage du graphique de Smith a échoué."): Exit Sub
    If Not Delink_ChartData(myChart) Then Call HandleError("La déconnexion des données du graphique a échoué."): Exit Sub

    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans Create_SmithChart.")
End Sub


Public Sub Format_Chart()
    '* @brief Formate un graphique avec les options par défaut.
    '* @details Cette procédure effectue les étapes suivantes:
    '   1. Vérifie si un graphique est sélectionné.
    '   2. Applique le formatage par défaut via ChartFormatting.
    '* @remarks Affiche un message d'erreur si une étape échoue.
    On Error GoTo ErrHandler

    Dim srcChart As Chart
    Dim FormattingOptions As Variant ' Laissée en Variant pour compatibilité, mais pourrait être String si toujours formatée en texte

    Set srcChart = GetSelectedChart() ' Utilise la fonction d'aide
    If srcChart Is Nothing Then Exit Sub

    ' Assurez-vous que FormattingOptions est initialisé correctement si ChartFormatting attend une chaîne vide pour les défauts
    ' Sinon, si vous voulez que les valeurs par défaut internes à ChartFormatting soient utilisées, passez Missing ou Empty.
    ' Pour l'exemple, supposons qu'une chaîne vide déclenche les défauts dans ChartFormatting
    FormattingOptions = ""

    If Not ChartFormatting(srcChart, FormattingOptions) Then Call HandleError("Une erreur est survenue lors du formatage du graphique."): Exit Sub

    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans Format_Chart.")
End Sub


Public Sub Create_YLog()
    '* @brief Crée un graphique avec un axe Y logarithmique basé sur le graphique sélectionné.
    '* @details Cette procédure effectue les étapes suivantes:
    '   1. Vérifie si un graphique est sélectionné.
    '   2. Copie le graphique sélectionné.
    '   3. Applique une échelle logarithmique à l'axe Y du graphique copié.
    '   4. Délie les données du graphique copié.
    '* @remarks Affiche un message d'erreur si une étape échoue.
    On Error GoTo ErrHandler

    Dim srcChart As Chart
    Dim myChart As Chart

    Set srcChart = GetSelectedChart() ' Utilise la fonction d'aide
    If srcChart Is Nothing Then Exit Sub

    Set myChart = Copy_Chart(srcChart)
    If myChart Is Nothing Then Call HandleError("La copie du graphique a échoué."): Exit Sub

    ' S'assure que Chart_YLog retourne un Boolean pour le succès/échec
    If Not Chart_YLog(myChart) Then Call HandleError("La génération de l'échelle Y logarithmique a échoué."): Exit Sub
    If Not Delink_ChartData(myChart) Then Call HandleError("La déconnexion des données du graphique a échoué."): Exit Sub

    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans Create_YLog.")
End Sub


Public Sub Create_Derivative()
    '* @brief Crée un graphique de dérivée basé sur le graphique sélectionné.
    '* @details Cette procédure effectue les étapes suivantes:
    '   1. Vérifie si un graphique est sélectionné.
    '   2. Copie le graphique sélectionné.
    '   3. Génère une dérivée pour le graphique copié.
    '   4. Délie les données du graphique copié.
    '* @remarks Affiche un message d'erreur si une étape échoue.
    On Error GoTo ErrHandler

    Dim srcChart As Chart
    Dim myChart As Chart

    Set srcChart = GetSelectedChart() ' Utilise la fonction d'aide
    If srcChart Is Nothing Then Exit Sub

    Set myChart = Copy_Chart(srcChart)
    If myChart Is Nothing Then Call HandleError("La copie du graphique a échoué."): Exit Sub

    ' S'assure que Chart_Derivate retourne un Boolean pour le succès/échec
    If Not Chart_Derivate(myChart) Then Call HandleError("La génération de la dérivée a échoué."): Exit Sub
    If Not Delink_ChartData(myChart) Then Call HandleError("La déconnexion des données du graphique a échoué."): Exit Sub

    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans Create_Derivative.")
End Sub

Public Sub Export_SelectedChartAsImage()
    '* @brief Exporte le graphique Excel actuellement sélectionné vers un fichier image.
    '* @details Cette procédure est un point d'entrée simple pour l'utilisateur.
    '           Elle appelle la fonction Export_ChartAsImage qui gère la sélection du fichier
    '           via une boîte de dialogue "Enregistrer sous..." et l'exportation réelle.
    '* @remarks Gère les erreurs et informe l'utilisateur du succès ou de l'échec de l'opération.
    On Error GoTo ErrHandler

    Debug.Print "Appel de Export_SelectedChartAsImage."

    ' Appelle la fonction d'exportation.
    ' Si aucun chemin n'est spécifié, une boîte de dialogue "Enregistrer sous" s'ouvrira.
    If Export_ChartAsImage() Then
        MsgBox "Le graphique a été exporté avec succès !", vbInformation, "Exportation Réussie"
    Else
        ' Le message d'erreur est déjà géré par Export_ChartAsImage ou HandleError
        MsgBox "L'exportation du graphique a échoué ou a été annulée.", vbExclamation, "Échec de l'Exportation"
    End If

    Exit Sub

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans Export_SelectedChartAsImage.")
    MsgBox "Une erreur interne est survenue lors de l'exportation.", vbCritical, "Erreur Critique"
End Sub
'#############################################################################################################################
' Partie 2 : Fonctions Publiques et Privées (Public & Private Function)
' ---
' Cette section contient les fonctions. Les fonctions publiques sont accessibles de l'extérieur et retournent une valeur,
' tandis que les fonctions privées sont des outils internes, utilisées par les autres fonctions et procédures du module
' pour des tâches spécifiques et récurrentes.
'#############################################################################################################################

Private Function HandleError(ByVal ErrMsg As String)
    ' @brief Affiche un message d'erreur standardisé à l'utilisateur.
    ' @param ErrMsg Le message d'erreur spécifique à afficher.
    MsgBox "Une erreur est survenue :" & vbCrLf & ErrMsg, vbOKOnly + vbCritical, "Opération impossible"
End Function

 Private Function SanitizeFileName(ByVal fileName As String) As String
     Dim invalidChars As Variant
     Dim i As Long
     invalidChars = Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|") ' Caractères invalides pour les noms de fichiers
     SanitizeFileName = fileName
     For i = LBound(invalidChars) To UBound(invalidChars)
         SanitizeFileName = Replace(SanitizeFileName, invalidChars(i), "_")
     Next i
     ' Remplacer les espaces par des underscores pour des noms de fichiers plus propres
     SanitizeFileName = Replace(SanitizeFileName, " ", "_")
 End Function


' @Description("Exporte le graphique Excel sélectionné sous forme de fichier image.")
' @Param(Optional imagePath As String, "Le chemin complet et le nom du fichier image (ex: 'C:\MonDossier\MonGraphique.png'). Si omis, une boîte de dialogue 'Enregistrer sous' s'affiche.")
' @Param(Optional imageFilterIndex As Long, "L'index du filtre de fichier pour la boîte de dialogue (1: PNG, 2: JPG, 3: GIF, 4: BMP). Par défaut 1 (PNG).")
' @Returns(Boolean, "True si l'exportation a réussi, False sinon.")
Public Function Export_ChartAsImage(Optional ByVal imagePath As String = "", Optional ByVal imageFilterIndex As Long = 1) As Boolean
    On Error GoTo ErrHandler
    
    Dim srcChart As Chart
    Dim defaultFileName As String
    Dim fileExtension As String
    Dim filters As String
    Dim selectedFile As Variant ' Pour la boîte de dialogue SaveAs
    
    Export_ChartAsImage = False ' Initialiser le résultat à False
    
    ' 1. Obtenir le graphique sélectionné
    Set srcChart = GetSelectedChart() ' Assumes GetSelectedChart() est une fonction existante qui retourne le Chart sélectionné
    If srcChart Is Nothing Then
        ' Le message d'erreur est géré par GetSelectedChart()
        Exit Function
    End If
    
    Debug.Print "Début de l'exportation du graphique '" & srcChart.name & "'."
    
    ' 2. Déterminer le chemin et le nom du fichier de sortie
    If imagePath = "" Then
        ' Construire les filtres pour la boîte de dialogue "Enregistrer sous"
        filters = "Fichier PNG (*.png),*.png," & _
                  "Fichier JPEG (*.jpg),*.jpg," & _
                  "Fichier GIF (*.gif),*.gif," & _
                  "Fichier BMP (*.bmp),*.bmp"
                  
        ' Nom de fichier par défaut basé sur le nom du graphique
        defaultFileName = Environ("USERPROFILE") & "\Desktop\" & SanitizeFileName(srcChart.name) & ".png"
        
        ' Afficher la boîte de dialogue "Enregistrer sous"
        With Application.FileDialog(msoFileDialogSaveAs)
            .InitialFileName = defaultFileName
            .title = "Enregistrer le graphique sous..."
            .filters.Clear
            .filters.Add "Images", filters, 1 ' Ajoute tous les filtres en une fois
            .FilterIndex = imageFilterIndex ' Sélectionne le filtre par défaut
            
            If .Show = -1 Then ' L'utilisateur a cliqué sur "Enregistrer"
                selectedFile = .SelectedItems(1)
            Else ' L'utilisateur a cliqué sur "Annuler"
                Debug.Print "Exportation annulée par l'utilisateur."
                Exit Function
            End If
        End With
        imagePath = selectedFile
    End If
    
    ' S'assurer que le chemin a une extension correcte (si l'utilisateur a tapé manuellement ou si imagePath est passé en paramètre)
    If InStr(imagePath, ".") = 0 Then ' Pas d'extension
        Select Case imageFilterIndex
            Case 1: fileExtension = ".png"
            Case 2: fileExtension = ".jpg"
            Case 3: fileExtension = ".gif"
            Case 4: fileExtension = ".bmp"
            Case Else: fileExtension = ".png" ' Fallback
        End Select
        imagePath = imagePath & fileExtension
        Debug.Print "Extension ajoutée: " & imagePath ' Datalogging
    End If
    
    ' 3. Exporter le graphique
    ' Le filtre de fichier est un numéro d'énumération XlChartPictureType
    Dim picType As XlChartPictureType
    Select Case LCase(Right(imagePath, 3)) ' Vérifie les 3 derniers caractères de l'extension
        Case "png": picType = xlPNG
        Case "jpg": picType = xlJPEG
        Case "gif": picType = xlGIF
        Case "bmp": picType = xlBitmap
        Case Else: picType = xlPNG ' PNG par défaut si extension inconnue
    End Select
    
    srcChart.Export fileName:=imagePath, FilterName:=picType
    
    Debug.Print "Graphique exporté avec succès vers : " & imagePath ' Datalogging
    Export_ChartAsImage = True ' Succès
    Exit Function

ErrHandler:
    Call HandleError("Erreur lors de l'exportation du graphique en image: " & err.Description & " (Code: " & err.Number & ")")
    Debug.Print "Erreur Export_ChartAsImage: " & err.Description & " (Code: " & err.Number & ")" ' Datalogging détaillé
End Function

Public Function Plot_RangeWithFormatting(data_src As Range, Optional ChartName As String = "", Optional nbSeriesByGroup As Integer = 0, Optional ColorStyle As String = "DefautStyle", Optional FormattingOptions As String = "") As Chart
    '* @brief Crée et formate un graphique à partir d'une plage de données.
    '* @details Cette fonction est une combinaison de Plot_Range et ChartFormatting.
    '           Elle crée un nouveau graphique à partir des données spécifiées,
    '           puis applique directement les options de formatage fournies.
    '* @param data_src La plage de données source pour le graphique.
    '* @param ChartName (Optionnel) Nom du graphique à créer. Si vide, un nom est généré automatiquement.
    '* @param nbSeriesByGroup (Optionnel) Nombre de séries par groupe pour la coloration. Par défaut à 1.
    '* @param ColorStyle (Optionnel) Style de couleur à appliquer aux séries. Par défaut à "DefautStyle".
    '* @param FormattingOptions (Optionnel) Chaîne de texte contenant les options de formatage (ex: "Title=MonTitre;XTitle=AxeX").
    '* @return L'objet Chart créé et formaté, ou Nothing en cas d'échec.
    On Error GoTo ErrHandler

    Dim myChart As Chart
    Dim defaultFormatting As String ' Pour stocker les options par défaut générées par Plot_Range

    ' Appel à Plot_Range pour créer le graphique initial
    ' Plot_Range va déjà générer certaines options de formatage basiques.
    ' Il est crucial que Plot_Range accepte et construise sur les FormattingOptions passées
    ' ou que nous fusionnions les options ici.
    ' Si Plot_Range construit déjà la chaîne FormattingOptions avec un titre, etc.,
    ' il faut s'assurer que notre FormattingOptions ici le complète ou le surcharge.

    ' Méthode 1: Plot_Range s'occupe de la génération du nom et des options de base,
    ' et nous ajoutons ou surchargeons avec FormattingOptions ensuite.
    Set myChart = Plot_Range(data_src, ChartName, nbSeriesByGroup, ColorStyle)

    If myChart Is Nothing Then
        Call HandleError("Échec de la création du graphique initial à partir de la plage de données.")
        Set Plot_RangeWithFormatting = Nothing
        Exit Function
    End If

    ' Assurez-vous que ChartFormatting est robuste face à des options vides ou manquantes
    If Not ChartFormatting(myChart, FormattingOptions) Then
        Call HandleError("Échec de l'application du formatage au graphique.")
        ' Considérer si le graphique doit être supprimé ici ou si l'utilisateur gère l'échec.
        ' Pour la robustesse, on le laisse mais on signale l'échec.
        Set Plot_RangeWithFormatting = Nothing
        Exit Function
    End If

    Set Plot_RangeWithFormatting = myChart
    Exit Function

ErrHandler:
    Call HandleError("Une erreur inattendue est survenue dans Plot_RangeWithFormatting.")
    Set Plot_RangeWithFormatting = Nothing
End Function


Private Function GetSelectedChart() As Chart
    ' @brief Tente de récupérer le graphique actif.
    ' @return Un objet Chart si un graphique est sélectionné et actif, sinon Nothing.
    On Error GoTo ErrHandle

    If Not ActiveChart Is Nothing Then
        Set GetSelectedChart = ActiveChart
    Else
        Call HandleError("Vous devez sélectionner un graphique pour effectuer cette opération.")
        Set GetSelectedChart = Nothing ' S'assure de retourner Nothing en cas d'absence de sélection
    End If
    Exit Function

ErrHandle:
    Call HandleError("Erreur lors de la récupération du graphique sélectionné.")
    Set GetSelectedChart = Nothing
End Function


Public Function Plot_Range(data_src As Range, Optional ChartName As String = "", Optional nbSeriesByGroup As Integer = -1, Optional ColorStyle As String = "DefautStyle", Optional FormattingOptions As String = "") As Chart
    Dim myChart         As Chart
    
    If ChartName = "" Then ChartName = FalLang.Clear_SubChar(data_src.parent.name, "_-. ")
    FormattingOptions = "Title=" & ChartName & ";XTitle=" & data_src.Cells(1, 1) & ";YTitle=(a.u);ChartType=" & xlXYScatterLinesNoMarkers & ";PlotBy=" & xlColumns & ";AutoLegend=0" & FormattingOptions

    Set myChart = Chart_Create(ChartName, data_src.parent.parent)
    Chart_AddSeries_from_range myChart, data_src
    delete_MatchingSeries myChart, data_src.Cells(1, 1), False
    Chart_ColorSeries myChart, 1, nbSeriesByGroup
    ChartFormatting myChart, FormattingOptions
    
    Set Plot_Range = myChart
End Function

Public Function DataTrace(wks_src As Worksheet, Optional DataTopLeftCell As String = "A1", Optional ColorStyle As String = "DefautStyle", Optional plotSeries As Boolean = True, Optional nbseries As Integer = -1, Optional sPlot As String = " - Plot", Optional FormattingOptions As String = "")
'On Error GoTo ifError
    Dim plot_Src(3)     As String
    Dim plot_Options(3) As String
    Dim aLabel()        As String       ' Noms des colonnes.
    Dim LastCol_Index   As Long         ' Dernière Colonne de données.
    Dim LastLineTbl()   As Long         ' Dernière Ligne de données.
    Dim Chart_Name      As String       ' Nom de la feuille graphique.
    Dim dummy           As Variant
    Dim firstLine       As Long
    Dim TopLeftCell     As Range
    Dim chartTitle      As String
    
    Set TopLeftCell = wks_src.Range(DataTopLeftCell)
    chartTitle = FalLang.Clear_SubString(wks_src.name, ".txt;.xlsx;.mdm;xlsm; ")
    chartTitle = FalLang.Resize_String(chartTitle, 31)
    plot_aColor = ColorPalette(ColorStyle)

    ' Detection du tableau de Données.
        ' Première et dernière ligne.
        LastLine = FalWork.find_LastNonEmptyRowInColumn(wks_src, TopLeftCell.column)
        firstLine = FLine(wks_src.parent.name, wks_src.name, TopLeftCell.column, LastLine)
'        firstLine = FalWork.find_FirstNonEmptyColumnInRowFromCol(wks_src, TopLeftCell.column, LastLine)
        LastCol_Index = FalWork.find_LastNonEmptyColumnInRow(wks_src, firstLine)
        LastCol_Name = FalLang.col(1 * LastCol_Index)
        firstcol_name = FalLang.col(1 * TopLeftCell.column)
        NbLine = LastLine - firstLine
        NbColumn = LastCol_Index - (TopLeftCell.Cells.column - 1)
        ' LastLine par Colones.
        ReDim LastLineTbl(LastCol_Index)
        For ColIndex = TopLeftCell.column To LastCol_Index
            LastLineTbl(ColIndex) = FalWork.find_LastNonEmptyRowInColumnFromLine(wks_src, TopLeftCell.Row, ColIndex)
        Next
        ' Nom des Colones.
        ReDim aLabel(LastCol_Index)
        For ColIndex = TopLeftCell.column To LastCol_Index
            tmpName = wks_src.Cells(firstLine - 1, ColIndex).value
            If tmpName = "" Or tmpName = False Then
                wks_src.Cells(firstLine - 1, ColIndex).value = "Col " & ColIndex    ' Renommage des colonnes sans nom.
            End If
            aLabel(ColIndex) = wks_src.Cells(firstLine - 1, ColIndex).value
        Next
    ' Identification Automatique du nombre de series.
    If nbseries < 1 Then nbseries = FalArray.aXD_count_Occurrence(aLabel, Left(aLabel(2), InStr(aLabel(2), " (") - 1), False, False)
    
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''' GRAPHIQUES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    ' [PLOT ALL SERIES]
    Dim myChart As Chart
    Chart_Name = FalLang.Resize_String(FalLang.Clear_SubChar(wks_src.name, "_-. "), 10) & " - " & sPlot & "(All)"
    FormattingOptions = "Title=" & chartTitle & ";XTitle=" & aLabel(1) & ";YTitle=(a.u);ChartType=" & xlXYScatterLinesNoMarkers & ";PlotBy=" & xlColumns & ";AutoLegend=0" & FormattingOptions

    Set myChart = Chart_Create(Chart_Name, wks_src.parent)
    Chart_AddSeries_from_range myChart, wks_src.Range(firstcol_name & (firstLine - 1) & ":" & LastCol_Name & LastLine)
    delete_MatchingSeries myChart, wks_src.Range(firstcol_name & (firstLine - 1)), False
    Chart_ColorSeries myChart, 1, ((NbColumn / nbseries) - 1)
    ChartFormatting myChart, FormattingOptions
    
    
    

'    Chart_Name = FalLang.Resize_String(FalLang.Clear_SubChar(wks_src.Name, "_-. "), 10) & " - " & sPlot & "(All)"
'    plot_Options(1) = "Title=" & ChartTitle & ";XTitle=" & aLabel(1) & ";YTitle=(a.u);ChartType=" & xlXYScatterLinesNoMarkers & ";PlotBy=" & xlColumns & ";AutoLegend=0"
'    If nbseries = 1 Then
'        plot_Src(1) = firstcol_name & firstLine - 1 & ":" & LastCol_Name & LastLineTbl(TopLeftCell.column)
'        Set myChart = plot(data_src:=wks_src.Range(plot_Src(1)), series:=(LastCol_Index - 1), nbBySeries:=1, plot_Options:=plot_Options(1), plot_aColor:=plot_aColor, chartName:=Chart_Name)
'    Else
'        plot_Src(1) = firstcol_name & firstLine - 1 & ":" & FalLang.col(1 * TopLeftCell.column + 1) & LastLineTbl(TopLeftCell.column)
'        Set myChart = plot(data_src:=wks_src.Range(plot_Src(1)), series:=1, nbBySeries:=1, plot_Options:=plot_Options(1), plot_aColor:=plot_aColor, chartName:=Chart_Name)
'        ' //////// Ajout de données à un graphique en fct du nbre de series \\\\\\\\
'        With wks_src.Parent.Charts(Chart_Name)
'            ' Ajout des autres series de données.
'            stepCol = TopLeftCell.column + (LastCol_Index / nbseries) - 1
'            Clr = 0 ' Boucle de table de couleur.
'            For XCol = 1 To LastCol_Index Step stepCol
'                Xletter = FalLang.col(1 * XCol)
'                For YCol = (XCol + 1) To (XCol + stepCol - 1)
'                    If .FullSeriesCollection.count >= 255 Then
'                        'MsgBox "Attention ! Le graphique contient plus de 256 séries." _
''        '                            & Chr(10) & "(Les séries suivantes ne seront pas tracées.)"
'                        FalWork.DebugPrint "MeasX", "DataTrace", "INFO, desc = max series exceeded"
'                        FalWork.statusBar_Write "Attention ! Le graphique contient plus de 256 séries."
'                        GoTo toomuchseries
'                    End If
'                    If YCol <> 2 Then   ' Première série déjà écrite.
'                        Yletter = FalLang.col(1 * YCol)
'                        ' Série de données titre & X & Y.
'                        XValues = "=" & wks_src.Name & "!$" & Xletter & "$" & firstLine & ":$" & Xletter & "$" & LastLineTbl(YCol)
'                        YValues = "=" & wks_src.Name & "!$" & Yletter & "$" & firstLine & ":$" & Yletter & "$" & LastLineTbl(YCol)
'                        XYTitle = "=" & wks_src.Name & "!$" & Yletter & "$" & firstLine - 1
'                        ' Création de la série.
'                        .SeriesCollection.newSeries     ' Nouvelle serie.
'                        .FullSeriesCollection(.SeriesCollection.count).Name = XYTitle
'                        .FullSeriesCollection(.SeriesCollection.count).XValues = XValues
'                        .FullSeriesCollection(.SeriesCollection.count).Values = YValues
'                        ' couleur des courbes.
'                        .SeriesCollection(.SeriesCollection.count).Border.ColorIndex = plot_aColor(Clr)
'                    End If
'                Next
'                Clr = Clr + 1
'                If Clr >= UBound(plot_aColor) Then Clr = 0
'            Next
'        End With
'toomuchseries:
'        ' \\\\\\\\ ---------------------------|---------------------------- //////// '
'    End If
'    isFormated = ChartFormatting(wks_src.Parent.Sheets(Chart_Name), FormattingOptions)
        
        
    ' [PLOT BY SERIES]
'    If nbSeries > 0 And plotSeries Then
'        Dim serieName       As String
'        Dim Chart_src       As Chart
'        Dim Chart_tmp       As Chart
'        Set Chart_src = wks_src.Parent.Sheets(Chart_Name)
'
'        For Ynum = 1 To stepCol - 1
'            Chart_src.Copy After:=Chart_src
'            Set Chart_tmp = ActiveChart
'            If InStr(aLabel(Ynum + 1), " (") > 0 Then serieName = Left(aLabel(Ynum + 1), InStr(aLabel(Ynum + 1), " (") - 1) Else serieName = aLabel(Ynum + 1)
'            Chart_Name = Replace(Chart_tmp.Name, "(2)", "")
'            Chart_Name = FalLang.Clear_SubChar(Replace(Chart_Name, "All", serieName), ":\/?*[];")
'            Chart_Name = FalLang.Resize_String(Chart_Name, 31)
'            Chart_tmp.Name = Chart_Name
'            If Chart_tmp.Axes(xlValue, xlPrimary).HasTitle Then Chart_tmp.Axes(xlValue, xlPrimary).AxisTitle.Characters.text = serieName
'            FalPlot.delete_UnMatchingSeries Chart_tmp, serieName, False
'        Next
'    End If
        

    ' Tracé par série.
        If nbseries > 0 And plotSeries Then
            ' feuilles de graphique en fonction du nombre de séries.
            stepCol = TopLeftCell.column + (LastCol_Index / nbseries) - 1
            YbyMes = stepCol - 1
            If YbyMes > 1 Then
                For Ynum = 1 To YbyMes
                    ' wks_src.range().value
                    outputname = wks_src.Cells(firstLine - 1, TopLeftCell.column + Ynum).value
                    Chart_Name = FalLang.Resize_String(FalLang.Clear_SubChar(wks_src.name, "_-. "), 10) & " - " & serieName & "(" & Trim(Left(outputname, InStr(outputname, " (") - 1)) & ")"
                    plot_Options(1) = "Title=" & chartTitle & ";XTitle=" & aLabel(1) & ";YTitle=" & Trim(Left(outputname, InStr(outputname, " (") - 1)) & ";ChartType=" & xlXYScatterLinesNoMarkers & ";PlotBy=" & xlColumns & ";AutoLegend=0"
                    plot_Src(1) = firstcol_name & firstLine - 1 & ":" & FalLang.col(1 * TopLeftCell.column + Ynum) & LastLineTbl(TopLeftCell.column)
                    Set myChart = plot(data_src:=wks_src.Range(plot_Src(1)), series:=1, nbBySeries:=1, plot_Options:=plot_Options(1), plot_aColor:=plot_aColor, ChartName:=Chart_Name)
                    
                    'IsPloted = plot(wks_src.Parent.Name, 1, 1, wks_src.Name, plot_Src(1), plot_Options(1), plot_aColor, Chart_Name)
                    ' //////// Ajout de données à un graphique en fct du nbre de series \\\\\\\\
                    With wks_src.parent.Charts(Chart_Name)
                        Clr = 0 ' Boucle de table de couleur.
                        For XCol = 1 To LastCol_Index Step stepCol
                            If .FullSeriesCollection.count >= 256 Then
                                'MsgBox "Attention ! Le graphique contient plus de 256 séries." _
'        '                            & Chr(10) & "(Les séries suivantes ne seront pas tracées.)"
                                FalWork.DebugPrint "MeasX", "DataTrace", "INFO, desc = max series exceeded"
                                FalWork.statusBar_Write "Attention ! Le graphique contient plus de 256 séries."
                                Exit For
                            End If
                            YCol = XCol + Ynum
                            Xletter = FalLang.col(CInt(XCol))
                            Yletter = FalLang.col(CInt(YCol))
                            ' Série de données titre & X & Y.
                            XValues = "=" & wks_src.name & "!$" & Xletter & "$" & firstLine & ":$" & Xletter & "$" & LastLineTbl(YCol)
                            YValues = "=" & wks_src.name & "!$" & Yletter & "$" & firstLine & ":$" & Yletter & "$" & LastLineTbl(YCol)
                            XYTitle = "=" & wks_src.name & "!$" & Yletter & "$" & firstLine - 1
                            ' Création de la série.
                            .SeriesCollection.newSeries
                            .SeriesCollection(.FullSeriesCollection.count).name = XYTitle
                            .SeriesCollection(.FullSeriesCollection.count).XValues = XValues
                            .SeriesCollection(.FullSeriesCollection.count).values = YValues
                            ' couleur des courbes.
                            .SeriesCollection(.FullSeriesCollection.count).Border.ColorIndex = plot_aColor(Clr)

                            Clr = Clr + 1
                            If Clr >= UBound(plot_aColor) Then Clr = 0
                        Next
                    End With
                    ' \\\\\\\\ ---------------------------|---------------------------- //////// '
                    isFormated = ChartFormatting(wks_src.parent.Sheets(Chart_Name), FormattingOptions)
                Next
            End If
        End If
                
EndLine:

    ' Fin de la fonction.
    DataTrace = True
ifError:
    DataTrace = True
    
End Function

Public Function Trace_Range(rng_src As Range, Optional ColorStyle As String = "DefautStyle", Optional plotSeries As Boolean = True, Optional nbseries As Integer = -1, Optional sPlot As String = " - Plot", Optional FormattingOptions As Variant)
    Dim Arr2D       As Variant
    
    Arr2D = rng_src.value
    FalPlot.plot (Arr2D)
End Function

Public Function ChartFormatting(myChart As Chart, Optional FormattingOptions As Variant) As Boolean
    On Error GoTo ErrHandler ' Utilisation d'un gestionnaire d'erreurs sp�cifique

    ' D�clarations de toutes les variables utilis�es (y compris celles initialement comment�es)
    Dim sKey As String
    Dim sValue As String
    Dim Options As Variant
    Dim index As Long ' Utiliser Long pour les index de tableau
    Dim ser As series
    Dim axNo As Axis

    ' Variables de configuration par d�faut
    Dim chartTypeSetting As Long ' Renomm� pour �viter conflit avec VBE ChartType
    Dim chartTitle As String
    Dim Y1Title As String
    Dim Y2Title As String
    Dim X1Title As String
    Dim X2Title As String
    Dim X1ScaleType As Long ' Utiliser Long pour les constantes Excel (xlLinear, xlLogarithmic)
    Dim Y1ScaleType As Long
    Dim X2ScaleType As Long
    Dim Y2ScaleType As Long
    Dim HasLegend As Boolean
    Dim HasTitle As Boolean
    Dim X1HasTitle As Boolean
    Dim Y1HasTitle As Boolean
    Dim X2HasTitle As Boolean
    Dim Y2HasTitle As Boolean
    Dim X1ShowAxis As Boolean
    Dim Y1ShowAxis As Boolean
    Dim X1ShowGridLines As Boolean
    Dim Y1ShowGridLines As Boolean

    Dim FontPolice As String
    Dim TitleFontSize As Integer
    Dim XAxisTitleFontSize As Integer
    Dim YAxisTitleFontSize As Integer
    Dim XAxisTicksFontSize As Integer
    Dim YAxisTicksFontSize As Integer
    Dim LegendFontSize As Integer
    Dim LegendPosition As Long
    Dim LegendInLayout As Boolean
    Dim PlotAreaLineWeight As Double
    Dim SeriesLineWeight As Double
    Dim SeriesLineDashStyle As Long
    Dim SeriesMarkerStyle As Long
    Dim HasDataLabels As String ' Peut �tre "Auto", "True", "False"
    Dim X1LabelNumberFormat As String
    Dim Y1LabelNumberFormat As String
    Dim X2LabelNumberFormat As String
    Dim Y2LabelNumberFormat As String
    Dim Y1Min As Variant ' Utiliser Variant car "Auto" ou Double
    Dim Y1Max As Variant
    Dim X1Min As Variant
    Dim X1Max As Variant
    Dim Y2Min As Variant
    Dim Y2Max As Variant
    Dim X2Min As Variant
    Dim X2Max As Variant
    Dim X1MinimumScaleIsAuto As Boolean
    Dim X1MaximumScaleIsAuto As Boolean
    Dim Y1MinimumScaleIsAuto As Boolean
    Dim Y1MaximumScaleIsAuto As Boolean
    Dim X2MinimumScaleIsAuto As Boolean
    Dim X2MaximumScaleIsAuto As Boolean
    Dim Y2MinimumScaleIsAuto As Boolean
    Dim Y2MaximumScaleIsAuto As Boolean
    Dim SquarePlot As Integer ' 0: non, 1: oui
    Dim CrossesAt As Double
    Dim PlotBy As Long ' xlRows ou xlColumns

    ' Configuration par d�faut (valeurs initiales)
    chartTypeSetting = -1 ' Utiliser -1 pour "Auto" pour les Long/Integer, ou un enum personnalis� si tu en as
    chartTitle = "Auto"
    HasTitle = True
    PlotBy = -1 ' Utiliser -1 pour "Auto" si tu ne veux pas forcer xlRows/xlColumns
    FontPolice = "Auto"
    TitleFontSize = 28
    XAxisTitleFontSize = 24
    YAxisTitleFontSize = 24
    XAxisTicksFontSize = 16
    YAxisTicksFontSize = 16
    HasLegend = True
    X1ShowAxis = True
    Y1ShowAxis = True
    X1ShowGridLines = True
    Y1ShowGridLines = True
    LegendFontSize = 12
    LegendPosition = xlLegendPositionBottom
    LegendInLayout = True
    PlotAreaLineWeight = 1.2
    SeriesLineWeight = 1.75
    SeriesLineDashStyle = msoLineSolid
    SeriesMarkerStyle = xlMarkerStyleNone
    HasDataLabels = "Auto"
    X1LabelNumberFormat = "General"
    Y1LabelNumberFormat = "General"
    X2LabelNumberFormat = "General"
    Y2LabelNumberFormat = "General"
    X1ScaleType = xlLinear
    Y1ScaleType = xlLinear
    X2ScaleType = xlLinear
    Y2ScaleType = xlLinear
    Y1Min = "Auto"
    Y1Max = "Auto"
    X1Min = "Auto"
    X1Max = "Auto"
    Y2Min = "Auto"
    Y2Max = "Auto"
    X2Min = "Auto"
    X2Max = "Auto"
    X1MinimumScaleIsAuto = True
    X1MaximumScaleIsAuto = True
    Y1MinimumScaleIsAuto = True
    Y1MaximumScaleIsAuto = True
    X2MinimumScaleIsAuto = True
    X2MaximumScaleIsAuto = True
    Y2MinimumScaleIsAuto = True
    Y2MaximumScaleIsAuto = True
    X1HasTitle = True
    X2HasTitle = True
    Y1HasTitle = True
    Y2HasTitle = True
    X1Title = "Auto"
    X2Title = "Auto"
    Y1Title = "Auto"
    Y2Title = "Auto"
    SquarePlot = 0
    CrossesAt = -1E+21

    ' Traitement des options de formatage pass�es
    If Not IsMissing(FormattingOptions) Then
        If Len(FormattingOptions & "") > 0 Then ' S'assurer que FormattingOptions n'est pas vide apr�s conversion en String
            Options = Split(FormattingOptions, ";")
            For index = 0 To UBound(Options)
                If InStr(Options(index), "=") > 0 Then ' V�rifie si l'option contient "="
                    sKey = Trim(Split(Options(index), "=")(0))
                    sValue = Trim(Split(Options(index), "=")(1)) ' Assure toi qu'il y a un second �l�ment apr�s le "="

                    Select Case sKey
                        Case "ChartType": chartTypeSetting = CLng(sValue) ' Utilise CLng pour les constantes
                        Case "ChartTitle", "Title": chartTitle = sValue
                        Case "PlotBy": PlotBy = CLng(sValue)
                        Case "FontPolice": FontPolice = sValue
                        Case "TitleFontSize": TitleFontSize = CInt(sValue)
                        Case "XAxisTitleFontSize": XAxisTitleFontSize = CInt(sValue)
                        Case "YAxisTitleFontSize": YAxisTitleFontSize = CInt(sValue)
                        Case "XAxisTicksFontSize": XAxisTicksFontSize = CInt(sValue)
                        Case "YAxisTicksFontSize": YAxisTicksFontSize = CInt(sValue)
                        Case "HasTitle": HasTitle = CBool(sValue)
                        Case "HasLegend": HasLegend = CBool(sValue)
                        Case "X1ShowAxis": X1ShowAxis = CBool(sValue)
                        Case "Y1ShowAxis": Y1ShowAxis = CBool(sValue)
                        Case "X1ShowGridLines", "XShowGridLines": X1ShowGridLines = CBool(sValue)
                        Case "Y1ShowGridLines", "YShowGridLines": Y1ShowGridLines = CBool(sValue)
                        Case "X1HasTitle", "XHasTitle": X1HasTitle = CBool(sValue)
                        Case "X1Title", "XTitle": X1Title = sValue
                        Case "X2Title": X2Title = sValue
                        Case "Y1HasTitle", "YHasTitle": Y1HasTitle = CBool(sValue)
                        Case "Y1Title", "YTitle": Y1Title = sValue
                        Case "Y2HasTitle": Y2HasTitle = CBool(sValue)
                        Case "Y2Title": Y2Title = sValue
                        Case "LegendFontSize": LegendFontSize = CInt(sValue)
                        Case "LegendPosition": LegendPosition = CLng(sValue)
                        Case "LegendInLayout": LegendInLayout = CBool(sValue)
                        Case "PlotAreaLineWeight": PlotAreaLineWeight = CDbl(sValue)
                        Case "SeriesLineWeight": SeriesLineWeight = CDbl(sValue)
                        Case "SeriesLineDashStyle": SeriesLineDashStyle = CLng(sValue)
                        Case "SeriesMarkerStyle": SeriesMarkerStyle = CLng(sValue)
                        Case "HasDataLabels": HasDataLabels = sValue ' "Auto", "True", "False"
                        Case "X1LabelNumberFormat": X1LabelNumberFormat = sValue
                        Case "X2LabelNumberFormat": X2LabelNumberFormat = sValue
                        Case "Y1LabelNumberFormat": Y1LabelNumberFormat = sValue
                        Case "Y2LabelNumberFormat": Y2LabelNumberFormat = sValue
                        Case "X1ScaleType": X1ScaleType = CLng(sValue)
                        Case "Y1ScaleType": Y1ScaleType = CLng(sValue)
                        Case "X2ScaleType": X2ScaleType = CLng(sValue)
                        Case "Y2ScaleType": Y2ScaleType = CLng(sValue)
                        Case "X1Min", "XMin": X1Min = CDbl(sValue): X1MinimumScaleIsAuto = False
                        Case "X1Max", "XMax": X1Max = CDbl(sValue): X1MaximumScaleIsAuto = False
                        Case "Y1Min", "YMin": Y1Min = CDbl(sValue): Y1MinimumScaleIsAuto = False
                        Case "Y1Max", "YMax": Y1Max = CDbl(sValue): Y1MaximumScaleIsAuto = False
                        Case "X2Min": X2Min = CDbl(sValue): X2MinimumScaleIsAuto = False
                        Case "X2Max": X2Max = CDbl(sValue): X2MaximumScaleIsAuto = False
                        Case "Y2Min": Y2Min = CDbl(sValue): Y2MinimumScaleIsAuto = False
                        Case "Y2Max": Y2Max = CDbl(sValue): Y2MaximumScaleIsAuto = False
                        Case "X1MinimumScaleIsAuto": X1MinimumScaleIsAuto = CBool(sValue)
                        Case "X1MaximumScaleIsAuto": X1MaximumScaleIsAuto = CBool(sValue)
                        Case "Y1MinimumScaleIsAuto": Y1MinimumScaleIsAuto = CBool(sValue)
                        Case "Y1MaximumScaleIsAuto": Y1MaximumScaleIsAuto = CBool(sValue)
                        Case "X2MinimumScaleIsAuto": X2MinimumScaleIsAuto = CBool(sValue)
                        Case "X2MaximumScaleIsAuto": X2MaximumScaleIsAuto = CBool(sValue)
                        Case "Y2MinimumScaleIsAuto": Y2MinimumScaleIsAuto = CBool(sValue)
                        Case "Y2MaximumScaleIsAuto": Y2MaximumScaleIsAuto = CBool(sValue)
                        Case "SquarePlot": SquarePlot = CInt(sValue)
                        Case "CrossesAt": CrossesAt = CDbl(sValue)
                        ' Ajoute d'autres cas si n�cessaire
                    End Select
                End If
            Next index
        End If
    End If

    With myChart
        ' Titre du graphique
        .HasTitle = HasTitle
        If .HasTitle Then
            If chartTitle = "Auto" Then
                ' Ne fait rien, garde le titre existant ou vide s'il n'y en a pas
            ElseIf chartTitle <> "" Then ' Applique le titre seulement s'il n'est pas vide
                .chartTitle.Characters.Text = chartTitle
            End If
            On Error Resume Next ' Peut �chouer si .ChartTitle est Nothing (mais HasTitle = True devrait l'�viter)
            .chartTitle.format.TextFrame2.TextRange.Font.Size = TitleFontSize
            On Error GoTo ErrHandler
        End If

        ' Type de graphique.
        If chartTypeSetting <> -1 Then .ChartType = chartTypeSetting ' -1 repr�sente "Auto"

        ' Inverser ligne/colonne (si PlotBy est g�r�)
        If PlotBy <> -1 Then .PlotBy = PlotBy ' -1 repr�sente "Auto"

        ' Police.
        If FontPolice <> "Auto" Then .ChartArea.Font.name = FontPolice

        ' Cadre ext�rieur du PlotArea
        With .PlotArea.format.line
            .ForeColor.RGB = RGB(150, 150, 150)
            .Weight = PlotAreaLineWeight
        End With
        .PlotArea.Interior.ColorIndex = xlNone

        ' Traitement des s�ries : Traits, Marqueurs et Data Labels
        For Each ser In .SeriesCollection
            With ser
                If HasDataLabels <> "Auto" Then
                    If CBool(HasDataLabels) Then ' HasDataLabels = "True"
                        If Not .HasDataLabels Then .ApplyDataLabels Type:=xlDataLabelsShowValue
                    Else ' HasDataLabels = "False"
                        If .HasDataLabels Then .DataLabels.Delete ' Supprime si existant
                    End If
                End If

                .format.line.Visible = msoTrue
                .format.line.Weight = SeriesLineWeight
                .format.line.DashStyle = SeriesLineDashStyle

                If SeriesMarkerStyle <> xlMarkerStyleNone Then
                    .MarkerStyle = SeriesMarkerStyle
                    .MarkerSize = 10
                    .MarkerForegroundColorIndex = .Border.ColorIndex
                    .MarkerBackgroundColorIndex = xlColorIndexNone
                Else
                    .MarkerStyle = xlMarkerStyleNone ' S'assurer qu'il n'y a pas de marqueur si sp�cifi�
                End If
            End With
        Next ser

        ' L�gende.
        .HasLegend = HasLegend
        If .HasLegend Then
            With .Legend
                .IncludeInLayout = LegendInLayout
                If SquarePlot <> 1 Then .Position = LegendPosition
                .format.TextFrame2.TextRange.Font.Size = LegendFontSize
            End With
        End If

        ' Options d'Axes.
        ' Utilisation de On Error Resume Next localement pour g�rer les axes non existants
        ' ou les propri�t�s qui ne s'appliquent pas � tous les types d'axes/graphiques
        For Each axNo In .Axes
            On Error Resume Next ' Active la gestion d'erreurs ici pour les axes
            If axNo.AxisGroup = xlPrimary Then ' Axes primaires (X1 & Y1)
                If axNo.Type = xlValue Then ' Y1 Axis
                    ' Gridlines pour Y1
                    If Y1ShowGridLines Then
                        .SetElement msoElementPrimaryValueGridLinesShow
                        axNo.HasMajorGridlines = True
                        axNo.MajorGridlines.format.line.ForeColor.RGB = RGB(224, 224, 224)
                        axNo.MajorGridlines.format.line.DashStyle = msoLineSysDot
                        axNo.MajorGridlines.format.line.Weight = 1
                    Else
                        .SetElement msoElementPrimaryValueGridLinesNone
                        axNo.HasMajorGridlines = False
                    End If

                    axNo.HasTitle = Y1HasTitle
                    If axNo.HasTitle And Y1Title <> "Auto" Then axNo.AxisTitle.Characters.Text = Y1Title
                    axNo.ScaleType = Y1ScaleType
                    axNo.MinimumScaleIsAuto = Y1MinimumScaleIsAuto
                    axNo.MaximumScaleIsAuto = Y1MaximumScaleIsAuto
                    If Y1Min <> "Auto" Then
                        If Y1ScaleType = xlLinear Then axNo.MinimumScale = CDbl(Y1Min)
                        If Y1ScaleType = xlLogarithmic Then axNo.MinimumScale = CDbl(Y1Min)
                    End If
                    If Y1Max <> "Auto" Then
                        If Y1ScaleType = xlLinear Then axNo.MaximumScale = CDbl(Y1Max)
                        If Y1ScaleType = xlLogarithmic Then axNo.MaximumScale = CDbl(Y1Max)
                    End If
                    axNo.TickLabelPosition = IIf(Y1ShowAxis, xlTickLabelPositionNextToAxis, xlNone)
                    axNo.TickLabels.NumberFormat = Y1LabelNumberFormat
                    axNo.TickLabels.Font.Size = YAxisTicksFontSize
                    axNo.MinorTickMark = IIf(Y1ShowAxis, xlInside, xlNone)
                    axNo.MajorTickMark = IIf(Y1ShowAxis, xlOutside, xlNone)
                    If axNo.HasTitle Then axNo.AxisTitle.format.TextFrame2.TextRange.Font.Size = YAxisTitleFontSize
                    If myChart.ChartType = xlXYScatter Then axNo.Crosses = xlAxisCrossesMinimum ' Crosses only makes sense for XYScatter
                End If

                If axNo.Type = xlCategory Or axNo.Type = xlValue Then ' X1 Axis (peut �tre Category ou Value pour XY)
                    ' Gridlines pour X1
                    If X1ShowGridLines Then
                        .SetElement msoElementPrimaryCategoryGridLinesShow ' Ou msoElementPrimaryValueGridLinesShow
                        axNo.HasMajorGridlines = True
                        axNo.MajorGridlines.format.line.ForeColor.RGB = RGB(224, 224, 224)
                        axNo.MajorGridlines.format.line.DashStyle = msoLineSysDot
                        axNo.MajorGridlines.format.line.Weight = 1
                    Else
                        .SetElement msoElementPrimaryCategoryGridLinesNone ' Ou msoElementPrimaryValueGridLinesNone
                        axNo.HasMajorGridlines = False
                    End If

                    axNo.HasTitle = X1HasTitle
                    If axNo.HasTitle And X1Title <> "Auto" Then axNo.AxisTitle.Characters.Text = X1Title
                    axNo.ScaleType = X1ScaleType
                    axNo.MinimumScaleIsAuto = X1MinimumScaleIsAuto
                    axNo.MaximumScaleIsAuto = X1MaximumScaleIsAuto
                    If X1Min <> "Auto" Then
                        If X1ScaleType = xlLinear Then axNo.MinimumScale = CDbl(X1Min)
                        If X1ScaleType = xlLogarithmic Then axNo.MinimumScale = CDbl(X1Min)
                    End If
                    If X1Max <> "Auto" Then
                        If X1ScaleType = xlLinear Then axNo.MaximumScale = CDbl(X1Max)
                        If X1ScaleType = xlLogarithmic Then axNo.MaximumScale = CDbl(X1Max)
                    End If
                    axNo.TickLabelPosition = IIf(X1ShowAxis, xlTickLabelPositionLow, xlNone)
                    axNo.TickLabels.NumberFormat = X1LabelNumberFormat
                    axNo.TickLabels.Font.Size = XAxisTicksFontSize
                    axNo.MinorTickMark = IIf(X1ShowAxis, xlInside, xlNone)
                    axNo.MajorTickMark = IIf(X1ShowAxis, xlOutside, xlNone)
                    If axNo.HasTitle Then axNo.AxisTitle.format.TextFrame2.TextRange.Font.Size = XAxisTitleFontSize
                    If myChart.ChartType = xlXYScatter Then axNo.CrossesAt = CrossesAt
                End If

            ElseIf axNo.AxisGroup = xlSecondary Then ' Axes secondaires (X2 & Y2)
                If axNo.Type = xlValue Then ' Y2 Axis
                    axNo.HasTitle = Y2HasTitle
                    If axNo.HasTitle And Y2Title <> "Auto" Then axNo.AxisTitle.Characters.Text = Y2Title
                    axNo.ScaleType = Y2ScaleType ' Correction de la faute de frappe
                    axNo.MinimumScaleIsAuto = Y2MinimumScaleIsAuto
                    axNo.MaximumScaleIsAuto = Y2MaximumScaleIsAuto
                    If Y2Min <> "Auto" Then
                        If Y2ScaleType = xlLinear Then axNo.MinimumScale = CDbl(Y2Min)
                        If Y2ScaleType = xlLogarithmic Then axNo.MinimumScale = CDbl(Y2Min)
                    End If
                    If Y2Max <> "Auto" Then
                        If Y2ScaleType = xlLinear Then axNo.MaximumScale = CDbl(Y2Max)
                        If Y2ScaleType = xlLogarithmic Then axNo.MaximumScale = CDbl(Y2Max)
                    End If
                    axNo.TickLabelPosition = xlTickLabelPositionNextToAxis
                    axNo.TickLabels.NumberFormat = Y2LabelNumberFormat
                    axNo.TickLabels.Font.Size = YAxisTicksFontSize
                    axNo.MinorTickMark = xlInside
                    axNo.MajorTickMark = xlOutside
                    If axNo.HasTitle Then axNo.AxisTitle.format.TextFrame2.TextRange.Font.Size = YAxisTitleFontSize
                End If

                If axNo.Type = xlCategory Or axNo.Type = xlValue Then ' X2 Axis
                    axNo.HasTitle = X2HasTitle
                    If axNo.HasTitle And X2Title <> "Auto" Then axNo.AxisTitle.Characters.Text = X2Title
                    axNo.ScaleType = X2ScaleType ' Correction de la faute de frappe
                    axNo.MinimumScaleIsAuto = X2MinimumScaleIsAuto
                    axNo.MaximumScaleIsAuto = X2MaximumScaleIsAuto
                    If X2Min <> "Auto" Then
                        If X2ScaleType = xlLinear Then axNo.MinimumScale = CDbl(X2Min)
                        If X2ScaleType = xlLogarithmic Then axNo.MinimumScale = CDbl(X2Min)
                    End If
                    If X2Max <> "Auto" Then
                        If X2ScaleType = xlLinear Then axNo.MaximumScale = CDbl(X2Max)
                        If X2ScaleType = xlLogarithmic Then axNo.MaximumScale = CDbl(X2Max)
                    End If
                    axNo.TickLabelPosition = xlTickLabelPositionHigh
                    axNo.TickLabels.NumberFormat = X2LabelNumberFormat
                    axNo.TickLabels.Font.Size = XAxisTicksFontSize
                    axNo.MinorTickMark = xlInside
                    axNo.MajorTickMark = xlOutside
                    If axNo.HasTitle Then axNo.AxisTitle.format.TextFrame2.TextRange.Font.Size = XAxisTitleFontSize
                End If
            End If
            ' Reset On Error pour �viter que les erreurs suivantes ne soient ignor�es globalement
            On Error GoTo ErrHandler
            ' Format des indices de texte pour tous les titres d'axe
            If axNo.HasTitle Then
                axNo.AxisTitle.Text = Replace(axNo.AxisTitle.Text, ".mm", ChrW(183) & "mm")
                axNo.AxisTitle.Text = Replace(axNo.AxisTitle.Text, "/mm", ChrW(183) & "mm-1")
            End If
        Next axNo

        ' Format text indices for the chart title
        ChartText_Format myChart ' Cette fonction devrait g�rer les titres des axes aussi

        ' Option : Graphique Carr� centr�.
        If SquarePlot = 1 Then
            If .HasLegend Then
                .Legend.Position = xlRight
                .Legend.format.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
                .Legend.format.Fill.Transparency = 0.15
            End If
            .PlotArea.Position = xlChartElementPositionAutomatic
            .PlotArea.Width = .PlotArea.Height + 0 ' Assure un ratio 1:1. Attention, si le graphique est large, cela le r�duira.
            ' Centrage horizontal du PlotArea
            If .HasLegend Then
                ' Si l�gende � droite, la surface de tra�age prend tout l'espace restant
                ' Pas de centrage manuel n�cessaire si la l�gende g�re l'espace
            Else
                ' Si pas de l�gende, centrer le PlotArea manuellement
                .PlotArea.Left = (.ChartArea.Width / 2) - (.PlotArea.Width / 2)
            End If
        End If

    End With

    ChartFormatting = True
    Exit Function

ErrHandler:
    ' Utilisez la fonction HandleError que nous avons d�finie pr�c�demment
    Call HandleError("Erreur lors du formatage du graphique : " & err.Description)
    ChartFormatting = False
End Function

Public Function ChartText_Format(myChart As Chart) As Boolean
' Formattage des indices et exposants des titres des axes.
    Dim StrToFormat         As String
    Dim toBaselineOffset    As String
    Dim indiceTerms       As Variant
    
    indiceTerms = Array("DS,sat", "DSS", "D,max", "G,inv", "ON ", "OFF ", "th ", "th,lin", "DS ", "GS ", "m ", "m,max ", "B ", "GS,ref", "DS,ref", "sat", "max", "min", ".abs")
    With myChart
        
        For index = 0 To UBound(indiceTerms)
            textLenght = Len(indiceTerms(index))
            
            ' Titre.
            If .HasTitle Then
                textDebut = InStr(1, .chartTitle.Text, indiceTerms(index), vbTextCompare)
                If textDebut Then .chartTitle.format.TextFrame2.TextRange.Characters(textDebut, textLenght).Font.BaselineOffset = -0.25: textDebut = 0
            End If
            ' Axes.
            For Each axNo In .Axes
                If axNo.HasTitle Then
                    AxTitle = axNo.AxisTitle.Text
                    ' [INDICE FORMAT]
                    textDebut = InStr(textDebut + 1, CStr(AxTitle), indiceTerms(index), vbTextCompare)
                    If textDebut Then axNo.AxisTitle.format.TextFrame2.TextRange.Characters(textDebut, textLenght).Font.BaselineOffset = -0.25: textDebut = 0
                End If
            Next
        Next
    End With
    
TitleExponent = True

End Function

Public Function Chart_Create(ChartName As String, wbk As Workbook, Optional rangeData As Variant) As Chart
    ' Creation du graphique.
    ChartName = FalLang.Clear_SubChar(ChartName, ":\/?*[];")
    ChartName = FalLang.Resize_String(ChartName, 31)
    If FalWork.Chart_IsInWorkbook(wbk, ChartName) Then wbk.Charts(ChartName).Delete
    
    
    Dim newChartSheet As Chart
    Set newChartSheet = wbk.Charts.Add
    With newChartSheet
        .Move After:=wbk.Worksheets(Worksheets.count)
        If ChartName <> "" Then .name = ChartName
        .ChartType = xlXYScatterLinesNoMarkers
    End With
    If Not IsMissing(rangeData) And TypeName(rangeData) = "Range" Then
        Chart_SetSource newChartSheet, rangeData
    Else
        Chart_ClearSeries newChartSheet
    End If
    Set Chart_Create = newChartSheet
End Function

Public Function Chart_ClearSeries(myChart As Chart) As Boolean
    ' Assurez-vous que le graphique a des s�ries existantes
    If myChart.SeriesCollection.count > 0 Then
        ' Supprimez les s�ries existantes
        Do While myChart.SeriesCollection.count > 0
            myChart.SeriesCollection(1).Delete
        Loop
    End If
    If myChart.SeriesCollection.count = 0 Then Chart_ClearSeries = True Else Chart_ClearSeries = False
End Function

Public Function Chart_SetSource(myChart As Chart, rangeData As Variant, Optional PlotBy As Integer = 2) As Variant
    If TypeName(rangeData) <> "Range" Then Chart_SetSource = CVErr(2001): Exit Function
    On Error Resume Next
    myChart.SetSourceData Source:=rangeData
    myChart.PlotBy = PlotBy
    Chart_SetSource = True
End Function

Public Function Chart_AddSeries_from_range(myChart As Chart, data_src As Range) As Variant
    Dim seriesRange As Range
    Dim newSeries As series
    Dim i As Integer

    ' Add a series for each column in the data source
    For i = 2 To data_src.Columns.count ' Start from column 2 to skip the first column (XValues)
        If myChart.FullSeriesCollection.count >= 255 Then GoTo ifError  ' Max Series exceeded
        
        If data_src.Cells(1, i).value <> data_src.Cells(1, 1).value Then
            Set newSeries = myChart.SeriesCollection.newSeries
            
            ' Set the series values
            Set seriesRange = data_src.Columns(i).Offset(1, 0).Resize(data_src.Rows.count - 1, 1)
            newSeries.values = seriesRange
            
            ' Set the X-axis values
            Set seriesRange = data_src.Columns(1).Offset(1, 0).Resize(data_src.Rows.count - 1, 1)
            newSeries.XValues = seriesRange
            
            ' Set series names
            newSeries.name = "='" & data_src.parent.name & "'!" & data_src.Cells(1, i).Address
        End If
    Next i
    Chart_AddSeries_from_range = True
    
    Exit Function
    
ifError:
    Chart_AddSeries_from_range = True
End Function

Function Chart_AddSeries_from_a2D(Chart As Chart, Arr2D() As Variant) As Variant
    Dim newSeries       As series
    Dim i               As Integer

    ' Add a series for each element in the array
    For i = LBound(Arr2D) + 1 To UBound(Arr2D)
        If Chart.FullSeriesCollection.count >= 256 Then Exit Function  ' Max Series exceeded
        
        Set newSeries = Chart.SeriesCollection.newSeries
        
        ' Set the series values
        newSeries.values = Application.index(Arr2D, 0, i + 1) ' Use Index to extract column i+1
        
        ' Set the X-axis values (first column of the array)
        newSeries.XValues = Application.index(Arr2D, 0, 1)
        
        ' Set series names
        newSeries.name = Arr2D(1, i + 1) ' Use the appropriate name from the array
    Next i
    Chart_AddSeries_from_a2D = True
End Function

Public Function Chart_ColorSeries(myChart As Chart, Optional nbBeforeChange As Long = 1, Optional nbSeriesByGroup As Integer = -1, Optional ColorStyle As String = "DefautStyle") As Boolean
    On Error GoTo ErrHandler

    Dim currentSeries As series
    Dim aColor As Variant
    Dim lbColor As Long
    Dim ubColor As Long
    Dim iColor As Long ' Index actuel dans la palette de couleurs
    Dim seriesCounterInGroup As Long ' Compte les s�ries au sein du groupe nbSeriesByGroup
    Dim seriesCounterBeforeChange As Long ' Compte les s�ries avant de changer la couleur dans la palette

    ' --- Pr�paration de la palette de couleurs ---
    aColor = ColorPalette(ColorStyle)
    If IsEmpty(aColor) Then ' G�rer le cas o� ColorPalette ne retourne rien (style non trouv�)
        Call HandleError("Style de couleur '" & ColorStyle & "' non trouv� ou palette vide.")
        Chart_ColorSeries = False
        Exit Function
    End If

    lbColor = LBound(aColor)
    ubColor = UBound(aColor)
    iColor = lbColor ' Commence par la premi�re couleur de la palette

    ' --- D�termination du nombre de s�ries par groupe pour la r�initialisation de la palette ---
    Dim effectiveNbSeriesByGroup As Long
    If nbSeriesByGroup = -1 Then
        effectiveNbSeriesByGroup = myChart.FullSeriesCollection.count ' Toutes les s�ries du graphique forment un groupe
    ElseIf nbSeriesByGroup <= 0 Then
        Call HandleError("Le param�tre 'nbSeriesByGroup' doit �tre -1 ou un nombre positif.")
        Chart_ColorSeries = False
        Exit Function
    Else
        effectiveNbSeriesByGroup = nbSeriesByGroup
    End If

    seriesCounterInGroup = 0 ' R�initialise le compteur pour le groupe
    seriesCounterBeforeChange = 0 ' R�initialise le compteur pour le changement de couleur

    ' --- Application des couleurs aux s�ries ---
    For Each currentSeries In myChart.FullSeriesCollection
        ' Applique la couleur de la palette actuelle � la s�rie
        currentSeries.Border.ColorIndex = aColor(iColor)
        ' Note : Si tu veux aussi la couleur de remplissage pour les marqueurs, tu pourrais ajouter:
        ' If currentSeries.HasMarker Then currentSeries.MarkerBackgroundColorIndex = aColor(iColor)
        ' If currentSeries.HasMarker Then currentSeries.MarkerForegroundColorIndex = aColor(iColor)

        seriesCounterInGroup = seriesCounterInGroup + 1
        seriesCounterBeforeChange = seriesCounterBeforeChange + 1

        ' Logique de changement de couleur dans la palette (apr�s nbBeforeChange s�ries)
        If seriesCounterBeforeChange >= nbBeforeChange Then
            iColor = iColor + 1
            seriesCounterBeforeChange = 0 ' R�initialise ce compteur pour le prochain cycle de nbBeforeChange
        End If

        ' Logique de r�initialisation de la palette (apr�s effectiveNbSeriesByGroup s�ries)
        If seriesCounterInGroup >= effectiveNbSeriesByGroup Then
            iColor = lbColor ' Revient � la premi�re couleur de la palette
            seriesCounterInGroup = 0 ' R�initialise le compteur de groupe
        ElseIf iColor > ubColor Then
            ' Si l'index de couleur d�passe la palette, mais le groupe n'est pas encore complet,
            ' cela signifie que la palette est trop courte pour le 'nbBeforeChange'
            ' ou que 'nbSeriesByGroup' est tr�s grand. On boucle simplement la palette.
            iColor = lbColor
        End If
    Next currentSeries

    Chart_ColorSeries = True ' Indique le succ�s
    Exit Function

ErrHandler:
    Call HandleError("Erreur lors de la coloration des s�ries : " & err.Description)
    Chart_ColorSeries = False
End Function


Public Function plot(data_src As Range, series As Integer, nbBySeries As Integer, plot_Options As String, plot_aColor As Variant, ChartName As String) As Chart
' Fonction de tra�age avec diff�rentes options

    ' Configuration par d�faut.
    Dim title, XTitle, YTitle, Xmin, XMax, Ymin, YMax, ChartType, PlotBy, YScaleType, AutoLegend, ShowLegend, AutoColorSerie
    Dim HasTitle As Boolean, HasXTitle As Boolean, HasYTitle As Boolean
    Dim XMinimumScaleIsAuto As Boolean, YMinimumScaleIsAuto As Boolean, XMaximumScaleIsAuto As Boolean, YMaximumScaleIsAuto As Boolean
    Dim sKey            As String
    Dim sValue          As String
    Dim wbk             As Workbook
    Dim wks             As Worksheet

    Set wbk = data_src.parent.parent
    Set wks = data_src.parent

    ChartName = FalLang.Clear_SubChar(ChartName, ":\/?*[];")
    ChartName = FalLang.Resize_String(ChartName, 31)
    ChartType = xlXYScatterLinesNoMarkers
    PlotBy = xlColumns
    YScaleType = xlLinear
    AutoLegend = 0
    ShowLegend = 1
    AutoColorSerie = 0
    XMinimumScaleIsAuto = True
    YMinimumScaleIsAuto = True
    XMaximumScaleIsAuto = True
    YMaximumScaleIsAuto = True
    
    If IsMissing(plot_Options) Then GoTo SkipPlotConfiguration
    Options = Split(plot_Options & ";", ";")
    For index = 0 To UBound(Options)
        If Options(index) Like "*=*" Then
            sKey = Split(Options(index), "=")(0)
            sValue = Split(Options(index), "=")(1)
            Select Case sKey
                Case "ChartType":           ChartType = sValue
                Case "PlotBy":              PlotBy = sValue
                Case "AutoLegend":          AutoLegend = sValue
                Case "ShowLegend":          ShowLegend = sValue
                Case "AutoColorSerie":      AutoColorSerie = sValue
                Case "Title":               title = sValue: HasTitle = True
                Case "XTitle":              XTitle = sValue: HasXTitle = True
                Case "YTitle":              YTitle = sValue: HasYTitle = True
                Case "XMin"::               Xmin = CDbl(sValue): XMinimumScaleIsAuto = False
                Case "XMax":                XMax = CDbl(sValue): XMaximumScaleIsAuto = False
                Case "YMin":                Ymin = CDbl(sValue): YMinimumScaleIsAuto = False
                Case "YMax":                XMax = CDbl(sValue): YMaximumScaleIsAuto = False
                Case "YScaleType":          YScaleType = sValue
            End Select
        End If
    Next
SkipPlotConfiguration:

    ' Creation du graphique.
    If FalWork.Chart_IsInWorkbook(wbk, ChartName) Then wbk.Charts(ChartName).Delete
    Dim newChartSheet   As Chart
    Set newChartSheet = wbk.Charts.Add

    With newChartSheet
        .Move After:=wks
        .name = ChartName
        .ChartType = ChartType
        .SetSourceData Source:=data_src, PlotBy:=CInt(PlotBy)
        ' Titre et titres des axes.
        .SetElement (msoElementChartTitleAboveChart)
        .HasTitle = HasTitle
        .Axes(xlValue, xlPrimary).HasTitle = HasYTitle
        .Axes(xlCategory, xlPrimary).HasTitle = HasXTitle
        If HasTitle Then .chartTitle.Characters.Text = title
        If HasYTitle Then .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = YTitle
        If HasXTitle Then .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = XTitle
        
        ' Axe des abcisses.
        With .Axes(xlCategory, xlPrimary)
            ' Limites.
            .MinimumScaleIsAuto = XMinimumScaleIsAuto
            .MaximumScaleIsAuto = XMaximumScaleIsAuto
            If Not XMinimumScaleIsAuto Then .MinimumScale = Xmin: .CrossesAt = Xmin
            If Not XMaximumScaleIsAuto Then .MaximumScale = XMax
            .ScaleType = xlLinear
        End With
        
        ' Axe des ordonn�es.
        With .Axes(xlValue, xlPrimary)
            ' Echelle Log.
            If YScaleType = "Log" Then .ScaleType = xlLogarithmic Else .ScaleType = xlLinear
            ' Limites.
            .MinimumScaleIsAuto = YMinimumScaleIsAuto
            .MaximumScaleIsAuto = YMaximumScaleIsAuto
            If Not YMinimumScaleIsAuto Then .MinimumScale = Ymin
            If Not YMaximumScaleIsAuto Then .MaximumScale = YMax
        End With
    
        ' Couleur axe principal en fonction du nb de courbes.
        ' Une couleur par mesure
        Dim Limit As Integer, ser As Integer, nbCurves As Integer
        ser = 1
        Var = 1
        rvb = 0
        nbCurves = .SeriesCollection.count
        For l = 1 To nbCurves
            If Var >= 1 + nbCurves Then
            Else
                If AutoColorSerie Then Limit = ser: Var = ser Else Limit = Var + nbBySeries - 1
                For m = Var To Limit Step 1
                    .SeriesCollection(m).Border.ColorIndex = plot_aColor(rvb)
                Next
            End If
            Var = Var + nbBySeries
            If rvb = UBound(plot_aColor) Then rvb = 0 Else rvb = rvb + 1
            ser = ser + 1
            If AutoColorSerie And rvb = series Then rvb = 0
        Next
    
        ' Modification du nom des l�gendes. ???????
        If AutoLegend Then
            lege = 1
            For ser = 1 To nbCurves Step nbBySeries
                .FullSeriesCollection(ser).name = "ID Q[" & wks.Cells(6, 1 + lege).value & ";" & wks.Cells(7, 1 + lege).value & "]"
                lege = lege + 1
            Next
        End If
        
        ' Mise en forme des l�gendes : une seule l�gende par s�rie
        For f = 2 To series + 1
            For ser = 1 To nbBySeries - 1
                .Legend.LegendEntries(f).Delete
            Next
        Next
        
        ' Affichage des l�gendes.
        If ShowLegend = 0 Then .Legend.Delete
    
    End With

    Set plot = newChartSheet
    Exit Function
ifError:
    newChartSheet.Delete
    Set plot = Nothing
End Function

Public Function Chart_YAbsolute(myChart As Chart) As Boolean
    '* @brief This function converts the values of the Y-axis in a chart to their absolute values.
    '* @param myChart The chart object whose Y-axis values need to be converted.
    '* @return Boolean indicating the success of the function.
    Dim cSerie          As Variant
    For Each cSerie In myChart.FullSeriesCollection
        cSerie.values = FalArray.a1D_math_Abs(cSerie.values)
    Next
    Chart_YAbsolute = True
End Function

Public Function Chart_XAbsolute(myChart As Chart) As Boolean
    '* @brief This function converts the values of the X-axis in a chart to their absolute values.
    '* @param myChart The chart object whose X-axis values need to be converted.
    '* @return Boolean indicating the success of the function.
    Dim cSerie          As Variant
    For Each cSerie In myChart.FullSeriesCollection
        cSerie.XValues = FalArray.a1D_math_Abs(cSerie.XValues)
    Next
    Chart_XAbsolute = True
End Function

Public Function Chart_YLog(myChart As Chart) As Boolean
    Chart_YAbsolute myChart
    myChart.Axes(xlValue, xlPrimary).ScaleType = xlLogarithmic
    Chart_YLog = True
End Function

Public Function Chart_XLog(myChart As Chart) As Boolean
    Chart_YAbsolute myChart
    myChart.Axes(xlCategory, xlPrimary).ScaleType = xlLogarithmic
    Chart_XLog = True
End Function

Public Function Chart_Derivate(myChart As Chart, Optional XSample As Long = 3) As Boolean
    '* @brief Cette fonction calcule la d�riv�e des s�ries de donn�es dans un graphique.
    '* @param myChart L'objet de graphique contenant les s�ries de donn�es � d�river.
    '* @param XSample Le nombre d'�chantillons utilis�s pour le calcul de la d�riv�e (valeur par d�faut : 3).
    '* @return Boolean indiquant le succ�s de la fonction.
    Dim cSerie          As Variant
    Dim a2D_X           As Variant
    Dim a2D_Y           As Variant
    Dim a2D_XY          As Variant
    Dim a1D_dY          As Variant
    
    With myChart
        For Each cSerie In .FullSeriesCollection
            a2D_X = FalArray.a1D_To_Columna2D(cSerie.XValues)
            a2D_Y = FalArray.a1D_To_Columna2D(cSerie.values)
            a2D_XY = FalArray.a2D_Merge_ByColumn(a2D_X, a2D_Y)
            a2D_dY = FalArray.a2D_math_Derivate(a2D_XY, XSample, 2)
            a1D_dY = FalArray.a2D_to_a1D_Column(a2D_dY, 2)
            cSerie.values = a1D_dY
            cSerie.name = "d" & cSerie.name
        Next
    End With
    Chart_Derivate = True
End Function

Public Function Delink_ChartData(myChart As Chart) As Boolean
    '* @brief De-links the series data of a given chart.
    '* @param myChart The chart object whose series data needs to be de-linked.
    On Error GoTo ifError
    Dim cSerie          As Variant
    For Each cSerie In myChart.FullSeriesCollection
        cSerie.XValues = cSerie.XValues
        cSerie.values = cSerie.values
        cSerie.name = cSerie.name
    Next
    Delink_ChartData = True
    Exit Function
ifError:
    Delink_ChartData = False
End Function

Public Function SmithChart_Formatting(myChart As Chart, Optional AddBackGround As Boolean = True, Optional FormattingOptions As String = "") As Boolean
    Dim cSerie1         As Variant
    Dim cSerie2         As Variant
    Dim picturePath     As String
    Dim isSmithData     As Boolean
    Dim fmtOptions      As String
    
    With myChart
        ' Identify if there is Real - Imaginary Data in chart.
        For Each cSerie1 In .FullSeriesCollection
            If cSerie1.name Like "*I:*" And Not cSerie1.IsFiltered Then
                For Each cSerie2 In .FullSeriesCollection
                    If Replace(cSerie1.name, "I:", "") = Replace(cSerie2.name, "R:", "") Then
                        isSmithData = True
                        Exit For
                    End If
                Next
            End If
        Next
        If Not isSmithData Then SmithChart_Formatting = False: Exit Function
        
        ' Process the chart.

        fmtOptions = "Y1Min=-1;Y1Max=1;X1Min=-1;X1Max=1;HasTitle=false;HasLegend=false;SquarePlot=1;X1HasTitle=false;Y1HasTitle=false;X1ShowGridLines=false;Y1ShowGridLines=false;X1ShowAxis=false;Y1ShowAxis=false" & _
                        ";X1ScaleType=" & xlLinear & ";Y1ScaleType=" & xlLinear
        ChartFormatting myChart, fmtOptions & ";" & FormattingOptions
    
        For Each cSerie1 In .FullSeriesCollection
            If cSerie1.name Like "*R:*" And Not cSerie1.IsFiltered Then
                For Each cSerie2 In .FullSeriesCollection
                    If Replace(cSerie1.name, "R:", "") = Replace(cSerie2.name, "I:", "") Then
                        cSerie2.XValues = cSerie1.values
                        cSerie1.Delete
                        Exit For
                    End If
                Next
            End If
        Next
        If AddBackGround Then
            picturePath = spreadsheet_MeasX.parent.path & "\SmithChart_BackGround.png"
            If FalFile.FileExist(picturePath) Then
                .PlotArea.format.Fill.Visible = msoTrue
                .PlotArea.format.Fill.UserPicture picturePath
                .PlotArea.format.Fill.TextureTile = msoFalse
                .PlotArea.format.Fill.Transparency = 0.6
            End If
        Else
            .PlotArea.format.Fill.Visible = msoFalse
        End If
    End With
    
    SmithChart_Formatting = True
End Function

Public Function Copy_Chart(myChart As Chart, Optional cpyPosition As String = "Right", Optional offsetPosition As Double = 0) As Variant
    '* @brief Copies a Chart object and returns the copied Chart.
    '* @param myChart The Chart object to be copied.
    '       cpyPosition the relative position for the copied chart.
    '       offsetPosition the relative offset for the position (only for embbeded charts).
    '* @return The copied Chart object.
    On Error GoTo ifError
    Dim parentType  As String
    Dim newChart    As Chart
    Dim wks_des     As Worksheet

    parentType = TypeName(myChart.parent)
    Select Case TypeName(myChart.parent)
        Case "Workbook"
            Select Case UCase(cpyPosition)
                Case "RIGHT", "AFTER": myChart.Copy After:=myChart
                Case "LEFT", "BEFORE": myChart.Copy Before:=myChart
                Case Else: myChart.Copy After:=myChart
            End Select
            Set newChart = ActiveChart
        Case Else
            Set wks_des = myChart.parent.parent
            myChart.parent.Copy
            wks_des.Paste
'            Set newChart = wks_des.Shapes(wks_des.Shapes.count).Chart
            Set newChart = wks_des.ChartObjects(wks_des.ChartObjects.count).Chart
            newChart.parent.Height = myChart.parent.Height
            newChart.parent.Width = myChart.parent.Width
            newChart.parent.Top = myChart.parent.Top
            newChart.parent.Left = myChart.parent.Left
            Select Case UCase(cpyPosition)
                Case "RIGHT": newChart.parent.Left = myChart.parent.Left + myChart.parent.Width + 5 + offsetPosition
                Case "LEFT": newChart.parent.Left = IIf(myChart.parent.Left - myChart.parent.Width - 5 - offsetPosition > 0, myChart.parent.Left - myChart.parent.Width - 5 - offsetPosition, 0)
                Case "DOWN": newChart.parent.Top = myChart.parent.Top + myChart.parent.Height + 5 + offsetPosition
                Case "UP": newChart.parent.Top = IIf(myChart.parent.Top - myChart.parent.Height - 5 - offsetPosition > 0, myChart.parent.Top - myChart.parent.Height - 5 - offsetPosition, 0)
                Case Else: newChart.parent.Left = myChart.parent.Left + myChart.parent.Width + 5 + offsetPosition
            End Select
    End Select

    Set Copy_Chart = newChart
    Exit Function
ifError:
    Copy_Chart = CVErr(2001)
End Function

Public Function Selected_Chart() As Variant
    '* @brief Returns the selected chart object.
    '* @return Variant representing the selected chart object.
    Select Case TypeName(Selection)
        Case "ChartArea", "PlotArea", "Chart": Set Selected_Chart = ActiveChart
        Case "Nothing": Set Selected_Chart = ActiveChart
        Case Else: Selected_Chart = CVErr(2001)
    End Select
End Function

Public Function filter_Series(myChart As Chart, sMatch As String, Optional compareAbsolute As Boolean = False) As Boolean
    ' * @brief Filters series in a chart based on a match condition.
    ' * @param myChart The chart object to filter.
    ' * @param sMatch The string to match against series names.
    ' * @param compareAbsolute Indicates whether an exact match is required. Default is False.
    ' * @return True if the series is filtered, False otherwise.

    Dim i       As Integer

    With myChart
        For i = 1 To .FullSeriesCollection.count
            If compareAbsolute Then
                .FullSeriesCollection(i).IsFiltered = IIf(.FullSeriesCollection(i).name = sMatch, False, True)
            Else
                .FullSeriesCollection(i).IsFiltered = IIf(InStr(.FullSeriesCollection(i).name, sMatch) > 0, False, True)
            End If
        Next i
    End With

End Function

Public Function delete_MatchingSeries(myChart As Chart, sMatch As String, Optional compareAbsolute As Boolean = False) As Boolean
    ' * @brief Filters series in a chart based on a match condition.
    ' * @param myChart The chart object to filter.
    ' * @param sMatch The string to match against series names.
    ' * @param compareAbsolute Indicates whether an exact match is required. Default is False.
    ' * @return True if the series is filtered, False otherwise.

    Dim i       As Integer

    With myChart
        For i = 1 To .FullSeriesCollection.count
            If compareAbsolute Then
                If .FullSeriesCollection(i).name = sMatch Then .FullSeriesCollection(i).Delete: i = i - 1
            Else
                If InStr(.FullSeriesCollection(i).name, sMatch) > 0 Then .FullSeriesCollection(i).Delete: i = i - 1
            End If
            If i + 1 > .FullSeriesCollection.count Then Exit For
        Next i
    End With
    
    delete_MatchingSeries = True
End Function

Public Function delete_UnMatchingSeries(myChart As Chart, sMatch As String, Optional compareAbsolute As Boolean = False) As Boolean
    ' * @brief Filters series in a chart based on a match condition.
    ' * @param myChart The chart object to filter.
    ' * @param sMatch The string to match against series names.
    ' * @param compareAbsolute Indicates whether an exact match is required. Default is False.
    ' * @return True if the series is filtered, False otherwise.

    Dim i       As Integer

    With myChart
        For i = 1 To .FullSeriesCollection.count
            If compareAbsolute Then
                If .FullSeriesCollection(i).name <> sMatch Then .FullSeriesCollection(i).Delete: i = i - 1
            Else
                If InStr(.FullSeriesCollection(i).name, sMatch) = 0 Then .FullSeriesCollection(i).Delete: i = i - 1
            End If
            If i + 1 > .FullSeriesCollection.count Then Exit For
        Next i
        
    End With
    
    delete_UnMatchingSeries = True

End Function

Public Function Resize_ChartFonts(myChart As Chart, Optional fontSize As Double = 14) As Boolean
    myChart.ChartArea.format.TextFrame2.TextRange.Font.Size = fontSize
End Function

Public Function ColorPalette(Optional Style As String = "DefautStyle") As Variant
    ' Palletes de couleurs pour l'attribut ".ColorIndex"
    Select Case Style
        Case "DefautStyle": ColorPalette = Array(3, 5, 4, 7, 46, 45, 44, 43, 41, 11, 49, 1, 56, 48, 23, 22, 21, 18, 17, 40, 39, 38, 37, 36, 35, 34, 33)
        Case "RainbowStyle": ColorPalette = Array(3, 30, 21, 54, 13, 18, 38, 7, 39, 49, 11, 5, 41, 37, 33, 8, 42, 14, 50, 10, 4, 43, 12, 40, 36, 6, 44, 45, 46, 53)
        Case Else: ColorPalette = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56)
    End Select
End Function
