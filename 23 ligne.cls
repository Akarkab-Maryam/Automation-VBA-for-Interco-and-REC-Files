Attribute VB_Name = "Module6"
Sub REC23ligne()
    Dim folderPath As String
    Dim folder As Object
    Dim fileName As String
    Dim fso As Object
    Dim retourIntercoFile As Workbook
    Dim recFile As Workbook
    Dim retourIntercoSheet As Worksheet
    Dim recSheet As Worksheet
    Dim retourIntercoPath As String
    Dim recPath As String
    
    ' Demander à l'utilisateur de sélectionner le répertoire contenant les dossiers
    folderPath = InputBox("Veuillez entrer le chemin du répertoire contenant les dossiers:", "Sélectionner le répertoire")
    
    If folderPath = "" Then
        MsgBox "Aucun répertoire sélectionné", vbExclamation
        Exit Sub
    End If
    
    ' Initialiser le FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' Parcourir tous les dossiers dans le répertoire
    For Each subFolder In folder.Subfolders
        ' Trouver les fichiers Retour_interco et REC dans chaque dossier
        retourIntercoPath = ""
        recPath = ""
        
        For Each file In subFolder.Files
            If InStr(file.Name, "Retour_interco_") = 1 Then
                retourIntercoPath = file.Path
            ElseIf InStr(file.Name, "REC") = 1 Then
                recPath = file.Path
            End If
        Next file
        
        ' Si les deux fichiers sont trouvés
        If retourIntercoPath <> "" And recPath <> "" Then
            ' Ouvrir le fichier Retour_interco
            Set retourIntercoFile = Workbooks.Open(retourIntercoPath)
            Set retourIntercoSheet = retourIntercoFile.Sheets(1)
            
            ' Copier les valeurs de C5, D5, C6, D6
            valeurC5 = retourIntercoSheet.Range("C5").Value
            valeurD5 = retourIntercoSheet.Range("D5").Value
            valeurC6 = retourIntercoSheet.Range("C6").Value
            valeurD6 = retourIntercoSheet.Range("D6").Value
            
            ' Fermer le fichier Retour_interco
            retourIntercoFile.Close SaveChanges:=False
            
            ' Ouvrir le fichier REC
            Set recFile = Workbooks.Open(recPath)
            Set recSheet = recFile.Sheets("REC")
            
            ' Coller les valeurs dans les cellules F10, G10, F23, G23
            recSheet.Range("F10").Value = valeurC5
            recSheet.Range("G10").Value = valeurD5
            recSheet.Range("F23").Value = valeurC6
            recSheet.Range("G23").Value = valeurD6
            
            ' Écrire "OK" dans les cellules F11, F12, F21, F22
            recSheet.Range("F11:F12").Value = "OK"
            recSheet.Range("F21:F22").Value = "OK"
            
            ' Écrire "CONNECTE" dans les cellules F13, F20
            recSheet.Range("F13").Value = "CONNECTE"
            recSheet.Range("F20").Value = "CONNECTE"

            
            ' Mettre les modifications en rouge
            recSheet.Range("F10:G10").Font.Color = vbRed
            recSheet.Range("F23:G23").Font.Color = vbRed
            recSheet.Range("F11:F12").Font.Color = vbRed
            recSheet.Range("F21:F22").Font.Color = vbRed
            recSheet.Range("F13:F20").Font.Color = vbRed
            
            ' Sauvegarder et fermer le fichier REC
            recFile.Save
            recFile.Close SaveChanges:=False
        End If
    Next subFolder
    
    MsgBox "Traitement terminé.", vbInformation
End Sub

