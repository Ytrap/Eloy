Private Sub CommandButton1_Click()

    ' Déclaration des variables global
    Dim Wbesclave As Workbook
    Dim ValeurPP As Integer
    Dim NbSemaines As Integer
    Dim RefPiece As Variant
    Dim a As Integer
    Dim Count As Integer
    Dim Annee As Integer
    
    Annee = ImportForm.Annee.Value

    Application.ScreenUpdating = True
    
    ' La variable a devient l'année sélectionnée
    a = ImportForm.List.Value
    
    fileExists = Dir("J:\QUALITE\REBUTS\" & Annee & " Rapports rebuts-ppm\Fichiers par mois\" & a & "-" & Annee & ".xlsm")
    
    If fileExists = "" Then
        MsgBox "/!\ Le fichier " & a & "-" & Annee & " n'existe pas"
    Else
        ' Ouverture du fichier source
        Set Wbesclave = Workbooks.Open("J:\QUALITE\REBUTS\" & Annee & " Rapports rebuts-ppm\Fichiers par mois\" & a & "-" & Annee & ".xlsm")
        ' Début de la boucle des semaines du fichier cible
        last = Wbesclave.Worksheets("Feuil1").Cells(8, 28).End(xlDown).Row
        For i = 8 To last
            ValeurPP = Wbesclave.Worksheets("Feuil1").Cells(i, 28).Value
            If ValeurPP <> 0 Then
                NbSemaines = Format(Wbesclave.Worksheets("Feuil1").Cells(i, 1).Value, "WW") - 1
                RefPiece = Wbesclave.Worksheets("Feuil1").Cells(i, 2).Value
                Application.Windows("Weekly scraps Certa.xlsm").Activate
                For n = 7 To 58
                    If Cells(n, 1).Value = NbSemaines And RefPiece = "117924" Then
                        Cells(n, 3).Value = ValeurPP
                        Count = Count + 1
                    ElseIf Cells(n, 1).Value = NbSemaines And RefPiece = "116642" Then
                        Cells(n, 7).Value = ValeurPP
                        Count = Count + 1
                    ElseIf Cells(n, 1).Value = NbSemaines And RefPiece = "116377" Then
                        Cells(n, 9).Value = ValeurPP
                        Count = Count + 1
                    End If
                Next
                        
            End If
            Application.Windows(a & "-" & Annee & ".xlsm").Activate
        Next
            
        Wbesclave.Close
        
        Application.ScreenUpdating = False
        
        Unload Me
        MsgBox Count & " Donnée(s) ont été traitée(s) avec succès!"
    End If
    
End Sub
Private Sub CommandButton2_Click()
    Dim Wbesclave As Workbook
    Dim ValeurPP As Integer
    Dim NbSemaines As Integer
    Dim RefPiece As Variant
    Dim Count As Integer
    Dim Annee As Integer
    
    Annee = ImportForm.Annee.Value

    Application.ScreenUpdating = True
    
    For a = 1 To 12
        fileExists = Dir("J:\QUALITE\REBUTS\" & Annee & " Rapports rebuts-ppm\Fichiers par mois\" & a & "-" & Annee & ".xlsm")
        If fileExists = "" Then
            If ImportForm.Alertes.Value = True Then
                MsgBox "/!\ Le fichier " & a & "-" & Annee & " n'existe pas"
            End If
        Else
            Set Wbesclave = Workbooks.Open("J:\QUALITE\REBUTS\" & Annee & " Rapports rebuts-ppm\Fichiers par mois\" & a & "-" & Annee & ".xlsm")
            last = Wbesclave.Worksheets("Feuil1").Cells(8, 28).End(xlDown).Row
            For i = 8 To last
                ValeurPP = Wbesclave.Worksheets("Feuil1").Cells(i, 28).Value
                If ValeurPP <> 0 Then
                    NbSemaines = Format(Wbesclave.Worksheets("Feuil1").Cells(i, 1).Value, "WW") - 1
                    RefPiece = Wbesclave.Worksheets("Feuil1").Cells(i, 2).Value
                    Application.Windows("Weekly scraps Certa.xlsm").Activate
                    For n = 7 To 58
                        If Cells(n, 1).Value = NbSemaines And RefPiece = "117924" Then
                            Cells(n, 3).Value = ValeurPP
                            Count = Count + 1
                        ElseIf Cells(n, 1).Value = NbSemaines And RefPiece = "116642" Then
                            Cells(n, 7).Value = ValeurPP
                            Count = Count + 1
                        ElseIf Cells(n, 1).Value = NbSemaines And RefPiece = "116377" Then
                            Cells(n, 9).Value = ValeurPP
                            Count = Count + 1
                        End If
                    Next
                        
                End If
                Application.Windows(a & "-" & Annee & ".xlsm").Activate
            Next
            
            Wbesclave.Close
        
            Application.ScreenUpdating = False
        End If
    Next
    
    Unload Me
    MsgBox Count & " Donnée(s) ont été traitée(s) avec succès!"
    
End Sub

Private Sub List_Change()

    ' Mise à jour du boutton d'execution unique selon le mois selectionné
    Select Case ImportForm.List.Value
        Case Is = "1"
            ImportForm.CommandButton1.Caption = "Importer Janvier"
            ImportForm.CommandButton1.Enabled = True
        Case Is = "2"
            ImportForm.CommandButton1.Caption = "Importer Fevrier"
            ImportForm.CommandButton1.Enabled = True
        Case Is = "3"
            ImportForm.CommandButton1.Caption = "Importer Mars"
            ImportForm.CommandButton1.Enabled = True
        Case Is = "4"
            ImportForm.CommandButton1.Caption = "Importer Avril"
            ImportForm.CommandButton1.Enabled = True
        Case Is = "5"
            ImportForm.CommandButton1.Caption = "Importer Mai"
            ImportForm.CommandButton1.Enabled = True
        Case Is = "6"
            ImportForm.CommandButton1.Caption = "Importer Juin"
            ImportForm.CommandButton1.Enabled = True
        Case Is = "7"
            ImportForm.CommandButton1.Caption = "Importer Juillet"
            ImportForm.CommandButton1.Enabled = True
        Case Is = "8"
            ImportForm.CommandButton1.Caption = "Importer Août"
            ImportForm.CommandButton1.Enabled = True
        Case Is = "9"
            ImportForm.CommandButton1.Caption = "Importer Septembre"
            ImportForm.CommandButton1.Enabled = True
        Case Is = "10"
            ImportForm.CommandButton1.Caption = "Importer Octobre"
            ImportForm.CommandButton1.Enabled = True
        Case Is = "11"
            ImportForm.CommandButton1.Caption = "Importer Novembre"
            ImportForm.CommandButton1.Enabled = True
        Case Is = "12"
            ImportForm.CommandButton1.Caption = "Importer Decembre"
            ImportForm.CommandButton1.Enabled = True
        Case Else
            ImportForm.CommandButton1.Caption = "..."
            ImportForm.CommandButton1.Enabled = False
    End Select
End Sub

Private Sub UserForm_Activate()

    ' Ajout des mois dans la liste
    For i = 1 To 12
        ImportForm.List.AddItem (i)
    Next
    ImportForm.List.Value = Month(Now)
    
    ' Ajout des années dans la liste
    For i = 2010 To 2030
        ImportForm.Annee.AddItem (i)
    Next
    ImportForm.Annee.Value = Year(Now)
    
    ' Ajout des moules dans les listes
    For i = 1 To 50
        ImportForm.MouleCAR.AddItem (i)
    Next
    ImportForm.MouleCAR.Value = "3"
    
    For i = 1 To 50
        ImportForm.MouleManifold.AddItem (i)
    Next
    ImportForm.MouleManifold.Value = "10"
    
    For i = 1 To 50
        ImportForm.MouleCAV.AddItem (i)
    Next
    ImportForm.MouleCAV.Value = "4"
    
    ' Activation des alertes par default
    ImportForm.Alertes.Value = True
    
End Sub
    
Public Function GetMoule(num As Integer) As Integer
    GetMoule = 1
End Function
