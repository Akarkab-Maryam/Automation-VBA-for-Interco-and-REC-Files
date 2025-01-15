Attribute VB_Name = "Module20"
Sub ModifierFichierExcel()
    Dim chemin As String
    Dim fichierRetour As String
    Dim fichierREC As String
    Dim wbRetour As Workbook
    Dim wbREC As Workbook
    Dim wsREC As Worksheet
    Dim valeurC5 As Variant
    Dim valeurC6 As Variant
    
    ' Demander à l'utilisateur de saisir le chemin du dossier
    chemin = InputBox("Veuillez saisir le chemin du répertoire contenant les fichiers Excel.")
    
    If chemin = "" Then
        MsgBox "Aucun chemin saisi. Le script a été annulé.", vbExclamation
        Exit Sub
    End If
    
    ' Chercher un fichier dont le nom commence par Retour_interco et peut avoir n'importe quelle extension (.xls, .xlsx)
    fichierRetour = Dir(chemin & "\Retour_interco*.*")
    
    If fichierRetour = "" Then
        MsgBox "Le fichier dont le nom commence par 'Retour_interco' n'a pas été trouvé.", vbExclamation
        Exit Sub
    End If
    
    ' Ouvrir le fichier Retour_interco
    Set wbRetour = Workbooks.Open(chemin & "\" & fichierRetour)
    
    ' Copier les valeurs des cellules C5 et C6
    valeurC5 = wbRetour.Sheets(1).Range("C5").Value
    valeurC6 = wbRetour.Sheets(1).Range("C6").Value
    
    wbRetour.Close False ' Fermer le fichier Retour_interco sans sauvegarder
    
    ' Chercher un fichier dont le nom commence par REC et peut avoir n'importe quelle extension (.xls, .xlsx)
    fichierREC = Dir(chemin & "\REC*.xls*")
    
    If fichierREC = "" Then
        MsgBox "Aucun fichier correspondant à 'REC*.xls*' n'a été trouvé.", vbExclamation
        Exit Sub
    End If
    
    ' Ouvrir le fichier REC
    Set wbREC = Workbooks.Open(chemin & "\" & fichierREC)
    
    ' Chercher la feuille "REC"
    On Error Resume Next
    Set wsREC = wbREC.Sheets("REC")
    On Error GoTo 0
    
    If wsREC Is Nothing Then
        MsgBox "La feuille 'REC' n'a pas été trouvée dans le fichier.", vbExclamation
        wbREC.Close False
        Exit Sub
    End If
    
    ' Coller les valeurs dans les cellules F9 et F12
    wsREC.Range("F9").Value = valeurC5
    wsREC.Range("F12").Value = valeurC6
    
    ' Mettre les valeurs collées en rouge
    wsREC.Range("F9").Font.Color = RGB(255, 0, 0)  ' Rouge pour la cellule F9
    wsREC.Range("F12").Font.Color = RGB(255, 0, 0) ' Rouge pour la cellule F12

    ' Écrire "ok" en rouge dans les cellules F10 et F11
    wsREC.Range("F10").Value = "ok"
    wsREC.Range("F11").Value = "ok"
    
    wsREC.Range("F10").Font.Color = RGB(255, 0, 0)  ' Rouge pour F10
    wsREC.Range("F11").Font.Color = RGB(255, 0, 0)  ' Rouge pour F11
    
    ' Sauvegarder et fermer le fichier REC
    wbREC.Save
    wbREC.Close False
    
    MsgBox "Les valeurs ont été copiées avec succès et mises en rouge dans les cellules F9, F10, F11 et F12.", vbInformation
End Sub

