Attribute VB_Name = "Module1"
Sub REC_13_Ligne()
    Dim folderPath As String
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim retourWorkbook As Workbook
    Dim recWorkbook As Workbook
    Dim retourPath As String
    Dim recPath As String
    Dim retourValues As Variant
    Dim ws As Worksheet
    
    ' Demander le chemin du répertoire
    folderPath = InputBox("Entrez le chemin du répertoire contenant les dossiers :")
    If folderPath = "" Then Exit Sub
    
    ' Vérifier si le chemin est valide
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    If folder Is Nothing Then
        MsgBox "Chemin invalide. Veuillez réessayer.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Parcourir chaque sous-dossier
    For Each subFolder In folder.SubFolders
        retourPath = ""
        recPath = ""
        
        ' Chercher le fichier commençant par "Retour_interco"
        For Each file In subFolder.Files
            If Left(file.Name, 14) = "Retour_interco" Then
                retourPath = file.Path
                Exit For
            End If
        Next file
        
        ' Si le fichier "Retour_interco" est trouvé
        If retourPath <> "" Then
            ' Ouvrir le fichier et copier les valeurs
            Set retourWorkbook = Workbooks.Open(retourPath)
            retourValues = Array(retourWorkbook.Sheets(1).Range("C5").Value, retourWorkbook.Sheets(1).Range("C6").Value)
            retourWorkbook.Close False
        Else
            GoTo NextFolder
        End If
        
        ' Chercher le fichier commençant par "REC"
        For Each file In subFolder.Files
            If Left(file.Name, 3) = "REC" Then
                recPath = file.Path
                Exit For
            End If
        Next file
        
        ' Si le fichier "REC" est trouvé
        If recPath <> "" Then
            ' Ouvrir le fichier "REC" et activer la feuille "REC"
            Set recWorkbook = Workbooks.Open(recPath)
            
            ' Vérifier si la feuille "REC" existe dans le fichier
            On Error Resume Next
            Set ws = recWorkbook.Sheets("REC")
            On Error GoTo 0
            
            If Not ws Is Nothing Then
                ' Coller les valeurs dans F9 et F12 de la feuille "REC"
                ws.Range("F9").Value = retourValues(0)
                ws.Range("F12").Value = retourValues(1)
                
                ' Écrire "OK" dans F10 et F11
                ws.Range("F10").Value = "OK"
                ws.Range("F11").Value = "OK"
                
                ' Appliquer la couleur rouge
                ws.Range("F9, F12, F10, F11").Font.Color = RGB(255, 0, 0)
                
                ' Sauvegarder les modifications dans le fichier "REC"
                recWorkbook.Save
            Else
                MsgBox "La feuille 'REC' n'existe pas dans le fichier " & recPath, vbExclamation
            End If
            
            ' Fermer le fichier "REC"
            recWorkbook.Close
        End If
        
NextFolder:
    Next subFolder
    
    MsgBox "Traitement terminé ! Les modifications ont été appliquées dans la feuille 'REC'.", vbInformation
End Sub

