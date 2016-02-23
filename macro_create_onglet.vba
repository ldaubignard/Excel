Sub MacroCreateOnglet()
'
' MacroCreateOnglet Macro
'
    'Déclaration variable
    Dim feuilleName As String
    Dim OwnerTFR As String
    Dim OwnerTME As String
    Dim Plage As Range
    'Déclaration plage de cellule a regarder
    Set Plage = Range("E5:E26")
    Sheets("Key projects").Select
    Range("E5").Select
    For Each Cell In Plage
        'Affectation des variables
        feuilleName = Left(Cell.Value, 30)
        OwnerTFR = Cell(1, 2).Value
        OwnerTME = Cell(1, 3).Value
        'Copie de la cellule
        Selection.Copy
        'MsgBox (feuilleName)
        'Ajout de la feuille
        Sheets.Add.Move After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = feuilleName
        'recuperation modele
        Sheets("model").Select
        Columns("A:E").Select
        Application.CutCopyMode = False
        Selection.Copy
        'appliquer le modele
         Sheets(feuilleName).Select
         Range("A1:E1").Select
         ActiveSheet.Paste
         'Alimentation des champs de chaque feuille créé
         Range("B1").Value = feuilleName
         Range("D3").Value = OwnerTFR
         Range("D4").Value = OwnerTME
    Next
End Sub
