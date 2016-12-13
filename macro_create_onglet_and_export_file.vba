' LDA - 13/12/2016
'Macro pour lire une colonne codeFiltreBaseDeLexport pour chaque valeur dans cette colone : 
' - on va chercher dans l onglet de donnees toutes les valeurs egale dans la colonne codeFiltreBaseDeLexport
' - on va creer un onglet avec le nom de la valeur 
' - on va copier toutes les lignes de l onglet de donnees dont la valeur de la colonne codeFiltreBaseDeLexport correspond
' - on va coller toutes les lignes dans le nouvel onglet
' - on va exporter le nouvel onglet dans un nouveau fichier avec le code et le libelle trouve dans la cellule a cote dans l onglet de parametre


Sub MacroCreateFileByOnglet()
'
' MacroCreateFileByOnglet Macro
'
    'Déclaration variable
    Dim fileSrcName As String
    Dim fileDestName As String
    Dim feuilleParamName As String
    Dim feuilleDataName As String
    Dim feuilleName As String
    Dim socName As String
    Dim filePath As String
    Dim fileExt As String
    Dim Plage As Range
    
    '---------------- Affectation Variable ---------------------------------------
    'nom du fichier de donnees a traiter
    fileSrcName = "fileSrcName.xlsm"
    'nom de l onglet contenant les noms d onglet et de fichier à creer
    feuilleParamName = "Param"
    'nom de l onglet contenant les data a exporter
    feuilleDataName = "Data"
    'information pour export de fichier
    filePath = "C:\Users\XXXXX\Documents\"
    fileExt = ".xlsx"
    '---------------- Fin Affectation Variable ------------------------------------
    'Aller dans l onglet Key projects
    Sheets(feuilleParamName).Select
    'Declaration plage de cellule a regarder
    Set Plage = Range("A2:A399")
    Range("A2").Activate
    For Each Cell In Plage
        'Affectation des variables
        feuilleName = Left(Cell.Value, 30)
        socName = Cell(1, 2).Value
        'creation du nom de fichier a exporter
        fileDestName = feuilleName & " - " & socName & fileExt
        '--------------------- CREATION ONGLET ---------------------------------------------------
        'Ajout de la feuille
        Sheets.Add.Move After:=Sheets(Sheets.Count)
        'renommage de la feuille
        Sheets(Sheets.Count).Name = feuilleName
        'selection de la nouvelle feuille
        Sheets(feuilleName).Select
        '---------------------- COPIE DATA DANS ONGLET --------------------------------------------
        'on va chercher les donnees dans l onglet data
        Sheets("Data").Select
        'on se positionne sur la colonne a filtrer
        Range("Tableau1[[#Headers],[codeFiltreBaseDeLexport]]").Select
        'on filtre les donnees
        ActiveSheet.ListObjects("Tableau1").Range.AutoFilter Field:=3, Criteria1:= _
            feuilleName
        Range("Tableau1[#All]").Select
        Range("Tableau1[[#Headers],[codeFiltreBaseDeLexport]]").Activate
        'on copie les donnees filtrees
        Selection.Copy
        'on selectionne la feuille de destination
        Sheets(feuilleName).Select
        Range("A1").Select
        'on colle les donnees
        ActiveSheet.Paste
        'on selectionne la feuille de data
        Sheets("Data").Select
        'on supprime le filtre
        ActiveSheet.ListObjects("Tableau1").Range.AutoFilter Field:=3
        Application.CutCopyMode = False
        '--------------------- EXPORT ONGLET ---------------------------------------------------
        Range("A2").Select
        Sheets(feuilleName).Select
        Sheets(feuilleName).Copy
        ChDir filePath
        ActiveWorkbook.SaveAs Filename:=filePath & fileDestName, FileFormat:= _
            xlOpenXMLWorkbook, CreateBackup:=False
        'selection du bon classeur a fermer
        Windows(fileDestName).Activate
        'fermeture du classeur
        ActiveWorkbook.Close
        'retour sur le classeur de data
        Windows(fileSrcName).Activate
        'retour sur le bon onglet
        Sheets(feuilleParamName).Select
        '---------------------------------------------------------------------------
    Next
End Sub

