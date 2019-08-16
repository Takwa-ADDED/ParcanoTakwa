VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Begin VB.MDIForm Frm_Main 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Parcano : Gestion parc automobile"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   1155
   ClientWidth     =   14265
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 ACB_Main 
      Align           =   1  'Align Top
      Height          =   7590
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14265
      _LayoutVersion  =   1
      _ExtentX        =   25162
      _ExtentY        =   13388
      _DataPath       =   ""
      Bands           =   "Frm_Main.frx":0ECA
      Begin VB.Timer Timer1 
         Interval        =   6000
         Left            =   3600
         Top             =   120
      End
      Begin VB.PictureBox IML_List 
         BackColor       =   &H80000005&
         Height          =   480
         Left            =   4320
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   1
         Top             =   120
         Width           =   1200
      End
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
    ACB_Main.Bands("BndEtat").Tools("lblUser").Caption = LStr_NameUser
    Me.Caption = Me.Caption & "- (  ver " & App.Major & "." & App.Minor & "." & App.Revision & " )"
    If UCase(GetSetting("CentraNord", "GestParc", "dbserver")) <> "SRV-SQL1\SRV_PRINCIPAL" Then
        Me.Caption = Me.Caption & " | ver " & App.Major & "." & App.Minor & "." & App.Revision & " |  Base Test"
        ACB_Main.Bands.Item("BndBD").ChildBands.BackColor = QBColor(12)
        ACB_Main.Bands.Item("BndBD").ChildBands.GradientEndColor = QBColor(12)
        ACB_Main.Bands("BndEtat").Tools("lblAffichage").Caption = " Base de données test  ... "
    Else
'        ACB_Main.Bands.Item("BndBD").ChildBands.BackColor = &HC0C0C0
'        ACB_Main.Bands.Item("BndBD").ChildBands.GradientEndColor = &HC0C0C0
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Key
        Case "Nouveau"
            'À faire: Ajouter le code du bouton 'Nouveau'.
            MsgBox "Ajouter le code du bouton 'Nouveau'."
        Case "Propriétés"
            'À faire: Ajouter le code du bouton 'Propriétés'.
            MsgBox "Ajouter le code du bouton 'Propriétés'."
        Case "Rechercher"
            'À faire: Ajouter le code du bouton 'Rechercher'.
            MsgBox "Ajouter le code du bouton 'Rechercher'."
        Case "Enregistrer"
            'À faire: Ajouter le code du bouton 'Enregistrer'.
            MsgBox "Ajouter le code du bouton 'Enregistrer'."
        Case "Rétablir"
            'À faire: Ajouter le code du bouton 'Rétablir'.
            MsgBox "Ajouter le code du bouton 'Rétablir'."
    End Select
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo erreur
    Dim i     As Integer
    Dim Msg
    ' Définit le texte du message.
    Msg = "Voulez-vous vraiment quitter l'application?"
    ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
    If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, "Fin d'application") = vbNo Then
       Cancel = True
    Else
        'Déconnecté la base
        Call LOBJ_CON.Disconnect(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
        Set LOBJ_CON = Nothing
        Set CNB = Nothing
        ' Boucler sur la collection Forms et déchargez
        ' chaque feuille.
         For i = Forms.Count - 1 To 0 Step -1
             Unload Forms(i)
         Next
        End
    End If
Exit Sub
erreur:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Private Sub ACB_Main_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Dim i
On Error GoTo Err
    'parcourir et unload tous les fenètres ouvert
'    For i = Forms.Count - 1 To 0 Step -1
'       If Forms(i).Name <> "Frm_Main" Then
'            Unload Forms(i)
'        End If
'    Next
'
    Select Case Tool.Name
        Case "Btn_Vehicule"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_vehicule", LInt_UserId) = True) Then Frm_Vehicule.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) _
                = vbYes Then
                    If (CHECK_ACCES("Consult_vehicule", LInt_UserId) = True) Then
                        For i = Forms.Count - 1 To 0 Step -1
                           If Forms(i).Name <> "Frm_Main" Then
                                Unload Forms(i)
                            End If
                        Next
                        Frm_Vehicule.Show
                    End If
                End If
            End If
            
        Case "Btn_Station"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Conslt_Fournisseur", LInt_UserId) = True) Then Frm_Station.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Conslt_Fournisseur", LInt_UserId) = True) Then
                        For i = Forms.Count - 1 To 0 Step -1
                           If Forms(i).Name <> "Frm_Main" Then
                                Unload Forms(i)
                            End If
                        Next
                    Frm_Station.Show
                    End If
                End If
            End If
        
        Case "Btn_TypCarburant"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_TC", LInt_UserId) = True) Then FrmCarburant.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_TC", LInt_UserId) = True) Then
                        For i = Forms.Count - 1 To 0 Step -1
                           If Forms(i).Name <> "Frm_Main" Then
                                Unload Forms(i)
                            End If
                        Next
                    FrmCarburant.Show
                    End If
                End If
            End If

        Case "Btn_Personnel"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_personnel", LInt_UserId) = True) Then Frm_Personnel.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                        If (CHECK_ACCES("Consult_personnel", LInt_UserId) = True) Then
                        For i = Forms.Count - 1 To 0 Step -1
                                   If Forms(i).Name <> "Frm_Main" Then
                                        Unload Forms(i)
                                    End If
                                Next
                        Frm_Personnel.Show
                        End If
                End If
            End If
'
        Case "Btn_Boncarburant"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_BC", LInt_UserId) = True) Then FrmAllBonCarburant.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_BC", LInt_UserId) = True) Then
                    For i = Forms.Count - 1 To 0 Step -1
                                       If Forms(i).Name <> "Frm_Main" Then
                                            Unload Forms(i)
                                        End If
                    Next
                    FrmAllBonCarburant.Show
                    End If
                End If
            End If

'
        Case "Btn_BonVidange"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_BV", LInt_UserId) = True) Then FrmBonVidange.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_BV", LInt_UserId) = True) Then
                    For i = Forms.Count - 1 To 0 Step -1
                                           If Forms(i).Name <> "Frm_Main" Then
                                                Unload Forms(i)
                                            End If
                    Next
                    FrmBonVidange.Show
                    End If
                End If
            End If

        Case "Btn_FactureCarburant"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_FF", LInt_UserId) = True) Then frmCreationFacture.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_FF", LInt_UserId) = True) Then
                    For i = Forms.Count - 1 To 0 Step -1
                                               If Forms(i).Name <> "Frm_Main" Then
                                                    Unload Forms(i)
                                                End If
                    Next
                    frmCreationFacture.Show
                    End If
                End If
            End If
'
        Case "Btn_Produits"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_Produit", LInt_UserId) = True) Then Frm_Articles.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_Produit", LInt_UserId) = True) Then
                    For i = Forms.Count - 1 To 0 Step -1
                                                   If Forms(i).Name <> "Frm_Main" Then
                                                        Unload Forms(i)
                                                    End If
                    Next
                    Frm_Articles.Show
                    End If
                End If
            End If
'
        Case "Btn_Alerte"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_Alerte", LInt_UserId) = True) Then Frm_Alertt.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_Alerte", LInt_UserId) = True) Then
                    For i = Forms.Count - 1 To 0 Step -1
                                                       If Forms(i).Name <> "Frm_Main" Then
                                                            Unload Forms(i)
                                                        End If
                    Next
                    Frm_Alertt.Show
                    End If
                End If
            End If
'
        Case "Btn_Utilisateur"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_Utilisateur", LInt_UserId) = True) Then Frm_Utilisateur.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                If (CHECK_ACCES("Consult_Utilisateur", LInt_UserId) = True) Then
                For i = Forms.Count - 1 To 0 Step -1
                                                       If Forms(i).Name <> "Frm_Main" Then
                                                            Unload Forms(i)
                                                        End If
                Next
                Frm_Utilisateur.Show
                End If
            End If
        End If

        Case "Btn_BCReparation"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_BCR", LInt_UserId) = True) Then FrmBCReparation.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_BCR", LInt_UserId) = True) Then
                    For i = Forms.Count - 1 To 0 Step -1
                                                           If Forms(i).Name <> "Frm_Main" Then
                                                                Unload Forms(i)
                                                            End If
                    Next
                    FrmBCReparation.Show
                    End If
                End If
            End If

        Case "Btn_Destination"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_Destination", LInt_UserId) = True) Then Frm_Destination.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_Destination", LInt_UserId) = True) Then
                    For i = Forms.Count - 1 To 0 Step -1
                                                           If Forms(i).Name <> "Frm_Main" Then
                                                                Unload Forms(i)
                                                            End If
                    Next
                    Frm_Destination.Show
                    End If
                End If
            End If

        Case "Btn_PieceReparation"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_PR", LInt_UserId) = True) Then FrmPieceReparation.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_PR", LInt_UserId) = True) Then
                    For i = Forms.Count - 1 To 0 Step -1
                                                               If Forms(i).Name <> "Frm_Main" Then
                                                                    Unload Forms(i)
                                                                End If
                        Next
                        FrmPieceReparation.Show
                        End If
                End If
            End If

        Case "Btn_StatCarburant"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_SC", LInt_UserId) = True) Then Frm_Statistiques.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_PR", LInt_UserId) = True) Then
                        For i = Forms.Count - 1 To 0 Step -1
                                                               If Forms(i).Name <> "Frm_Main" Then
                                                                    Unload Forms(i)
                                                                End If
                        Next
                        Frm_Statistiques.Show
                    End If
                End If
            End If

        Case "Btn_Trafic"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_FT", LInt_UserId) = True) Then Frm_Trafic.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_FT", LInt_UserId) = True) Then
                    For i = Forms.Count - 1 To 0 Step -1
                                                                   If Forms(i).Name <> "Frm_Main" Then
                                                                        Unload Forms(i)
                                                                    End If
                    Next
                    Frm_Trafic.Show
                    End If
                End If
            End If

        Case "Btn_Sup"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_SUp", LInt_UserId) = True) Then Frm_Supervision.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_SUp", LInt_UserId) = True) Then
                        For i = Forms.Count - 1 To 0 Step -1
                                                                           If Forms(i).Name <> "Frm_Main" Then
                                                                                Unload Forms(i)
                                                                            End If
                        Next
                    Frm_Supervision.Show
                    End If
                End If
            End If

        Case "Btn_PrgChauf"
        If Forms.Count = 1 Then
            If (CHECK_ACCES("Consult_PCH", LInt_UserId) = True) Then Frm_PrgChauf.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                If (CHECK_ACCES("Consult_PCH", LInt_UserId) = True) Then
                For i = Forms.Count - 1 To 0 Step -1
                                                                           If Forms(i).Name <> "Frm_Main" Then
                                                                                Unload Forms(i)
                                                                            End If
                        Next
                    Frm_PrgChauf.Show
                    End If
                End If
            End If

        Case "Btn_Conge"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_conge", LInt_UserId) = True) Then Frm_GestionConge.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_conge", LInt_UserId) = True) Then
                        For i = Forms.Count - 1 To 0 Step -1
                                                                       If Forms(i).Name <> "Frm_Main" Then
                                                                            Unload Forms(i)
                                                                        End If
                        Next
                        Frm_GestionConge.Show
                    End If
                End If
             End If

        Case "Btn_PLANNING"
            If Forms.Count = 1 Then
                If (CHECK_ACCES("Consult_PLING", LInt_UserId) = True) Then Frm_PLANNING.Show Else Exit Sub
            Else
                If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    If (CHECK_ACCES("Consult_PLING", LInt_UserId) = True) Then
                        For i = Forms.Count - 1 To 0 Step -1
                                              If Forms(i).Name <> "Frm_Main" Then
                                                     Unload Forms(i)
                                              End If
                         Next
                                Frm_PLANNING.Show
                    End If
                End If
             End If
    End Select
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub


