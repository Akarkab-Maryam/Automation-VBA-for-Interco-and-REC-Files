Attribute VB_Name = "Module2"
Sub TraitementSousDossiersEtFichiers()
    Dim fso As Object
    Dim dossierPrincipal As Object
    Dim sousDossier As Object
    Dim fichierRetour As Object
    Dim fichierREC As Object
    Dim wbRetour As Workbook
    Dim wbREC As Workbook
    Dim wsREC As Worksheet
    Dim valeurC5 As Variant
    Dim valeurD5 As Variant
    Dim valeurC6 As Variant
    Dim cheminPrincipal As String

    ' Demander le chemin principal
    cheminPrincipal = InputBox("Veuillez entrer le chemin du r�pertoire principal contenant les sous-dossiers.")

    If cheminPrincipal = "" Then
        MsgBox "Aucun chemin saisi. Le script est annul�.", vbExclamation
        Exit Sub
    End If

    ' V�rifier si le chemin existe
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dossierPrincipal = fso.GetFolder(cheminPrincipal)
    On Error GoTo 0

    If dossierPrincipal Is Nothing Then
        MsgBox "Le chemin sp�cifi� est invalide.", vbExclamation
        Exit Sub
    End If

    ' Parcourir tous les sous-dossiers
    For Each sousDossier In dossierPrincipal.SubFolders
        ' Chercher un fichier Retour_interco dans le sous-dossier
        For Each fichierRetour In sousDossier.Files
            If fichierRetour.Name Like "Retour_interco*" Then
                ' Ouvrir le fichier Retour_interco
                Set wbRetour = Workbooks.Open(fichierRetour.Path)
                valeurC5 = wbRetour.Sheets(1).Range("C5").Value
                valeurD5 = wbRetour.Sheets(1).Range("D5").Value
                valeurC6 = wbRetour.Sheets(1).Range("C6").Value
                wbRetour.Close False
                Exit For
            End If
        Next fichierRetour

        ' Chercher un fichier REC dans le sous-dossier
        For Each fichierREC In sousDossier.Files
            If fichierREC.Name Like "REC*.xls*" Then
                ' Ouvrir le fichier REC
                Set wbREC = Workbooks.Open(fichierREC.Path)

                ' V�rifier l'existence de la feuille REC
                On Error Resume Next
                Set wsREC = wbREC.Sheets("REC")
                On Error GoTo 0

                If Not wsREC Is Nothing Then
                    ' Coller les valeurs
                    With wsREC
                        ' Valeur C5 de Retour_interco coll�e dans F9
                        .Range("F9").Value = valeurC5
                        .Range("F9").Font.Color = RGB(255, 0, 0)

                        ' Valeur D5 coll�e dans G9
                        .Range("G9").Value = valeurD5
                        .Range("G9").Font.Color = RGB(255, 0, 0)

                        ' Valeur C6 coll�e dans F17
                        .Range("F17").Value = valeurC6
                        .Range("F17").Font.Color = RGB(255, 0, 0)

                        ' Cellules avec des valeurs fixes
                        .Range("F10").Value = "OK"
                        .Range("F11").Value = "OK"
                        .Range("F15").Value = "OK"
                        .Range("F16").Value = "OK"
                        .Range("F12").Value = "CONNECTE"
                        .Range("F14").Value = "CONNECTE"

                        ' Appliquer la couleur rouge aux valeurs fixes
                        .Range("F10:F12").Font.Color = RGB(255, 0, 0)
                        .Range("F14:F16").Font.Color = RGB(255, 0, 0)
                    End With

                    ' Sauvegarder et fermer
                    wbREC.Save
                    wbREC.Close False
                Else
                    MsgBox "La feuille 'REC' n'a pas �t� trouv�e dans le fichier : " & fichierREC.Name, vbExclamation
                End If
                Exit For
            End If
        Next fichierREC
    Next sousDossier

    MsgBox "Le traitement de tous les sous-dossiers est termin�.", vbInformation
End Sub

