VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Frm_Main 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Parcano : Gestion parc automobile"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   1155
   ClientWidth     =   8775
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 ACB_Main 
      Align           =   1  'Align Top
      Height          =   7590
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _LayoutVersion  =   1
      _ExtentX        =   15478
      _ExtentY        =   13388
      _DataPath       =   ""
      Bands           =   "MDIForm1.frx":0ECA
      Begin VB.Timer Timer1 
         Interval        =   6000
         Left            =   3600
         Top             =   120
      End
      Begin MSComctlLib.ImageList IML_List 
         Left            =   4320
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   23
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":8A49F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":8B17B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":8BE57
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":8CB33
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":8D80F
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":8E4EB
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":8F1C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":8FEA3
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":90B7F
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9185B
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":92537
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":93213
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":93EEF
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":94BCB
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":958A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9657F
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":97259
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":97F33
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":98B86
               Key             =   "IcoVehicule"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":98EA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":99D92
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9B8E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9D436
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub ACB_Main_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)

On Error GoTo Err

'parcourir et uload tous les fenètres ouvert
For i = Forms.Count - 1 To 0 Step -1
   If Forms(i).Name <> "Frm_Main" Then Unload Forms(i)
Next
    
Select Case Tool.Name
        Case "Btn_Vehicule"
            If (check_acces("Consult_vehicule") = True) Then
                 FrmVehicule.Show
             Else
                Exit Sub
            End If
            
        Case "Btn_Station"
            If (check_acces("Conslt_Fournisseur") = True) Then
                 FrmStation.Show
             Else
                Exit Sub
            End If
        
        Case "Btn_TypCarburant"
            If (check_acces("Consult_TC") = True) Then
                   FrmCarburant.Show
             Else
                Exit Sub
            End If
            
           
        Case "Btn_Personnel"
            If (check_acces("Consult_personnel") = True) Then
                    Frm_Personnel.Show
             Else
                Exit Sub
            End If
            
        Case "Btn_Boncarburant"
            If (check_acces("Consult_BC") = True) Then
                    FrmAllBonCarburant.Show
             Else
                Exit Sub
            End If
            
            
        Case "Btn_BonVidange"
            If (check_acces("Consult_BV") = True) Then
                    FrmBonVidange.Show
             Else
                Exit Sub
            End If
            
        Case "Btn_FactureCarburant"
            If (check_acces("Consult_FF") = True) Then
                    frmCreationFacture.Show
             Else
                Exit Sub
            End If
             
        Case "Btn_Produits"
            If (check_acces("Consult_Produit") = True) Then
                    FrmArticles.Show
             Else
                Exit Sub
            End If
            
            
        Case "Btn_Alerte"
            If (check_acces("Consult_Alerte") = True) Then
                   Frm_Alertt.Show
             Else
                Exit Sub
            End If
            
            
        Case "Btn_Utilisateur"
            If (check_acces("Consult_Utilisateur") = True) Then
                  FrmUtilisateur.Show
             Else
                Exit Sub
            End If
            
            
        Case "Btn_BCReparation"
            If (check_acces("Consult_BCR") = True) Then
                  FrmBCReparation.Show
             Else
                Exit Sub
            End If
           
            
        Case "Btn_Destination"
            If (check_acces("Consult_Destination") = True) Then
                  Frm_Destination.Show
             Else
                Exit Sub
            End If
            
            
        Case "Btn_PieceReparation"
            If (check_acces("Consult_PR") = True) Then
                  FrmPieceReparation.Show
             Else
                Exit Sub
            End If
            
            
        Case "Btn_StatCarburant"
            If (check_acces("Consult_SC") = True) Then
                  FrmStatCarburant.Show
             Else
                Exit Sub
            End If
            
            
        Case "Btn_StatReparation"
            If (check_acces("Consult_SR") = True) Then
                  FrmStatReparation.Show
             Else
                Exit Sub
            End If
            
            
        Case "Btn_StatFT"
            If (check_acces("Consult_ST") = True) Then
                 FrmStatFT.Show
             Else
                Exit Sub
            End If
            
        
        Case "Btn_Trafic"
            If (check_acces("Consult_FT") = True) Then
                 Frm_Trafic.Show
             Else
                Exit Sub
            End If
            
            
        Case "Btn_Sup"
             If (check_acces("Consult_SUp") = True) Then
                 Frm_Supervision.Show
             Else
                Exit Sub
            End If
            
            
        Case "Btn_StatService"
        If (check_acces("Consult_EHS") = True) Then
                 FrmStatService.Show
             Else
                Exit Sub
            End If
            
            
        Case "Btn_PrgChauf"
            If (check_acces("Consult_Utilisateur") = True) Then
                  Frm_PrgChauf.Show
             Else
                Exit Sub
            End If
     
                  
End Select
Exit Sub
Err:
MsgBox Err.Description, vbInformation


End Sub

Private Sub MDIForm_Load()
ACB_Main.Bands("BndEtat").Tools("lblUser").Caption = LStr_NameUser
Me.Caption = " - Parcano - (" & " ver " & App.Major & "." & App.Minor & "." & App.Revision & " )"
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
   Dim i As Integer
   Dim MSG ' Déclare la variable.
   ' Définit le texte du message.
   MSG = "Voulez-vous vraiment quitter l'application?"
   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
   If MsgBox(MSG, vbQuestion + vbYesNo + vbDefaultButton2, "Fin d'application") = vbNo Then
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
   MsgBox Err.Description, 48
End Sub


