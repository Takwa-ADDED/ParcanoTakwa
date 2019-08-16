VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form Frm_MajFT 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Maj Fiche Traffic"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   11880
   Begin SToolBox.SDateBox cda_Sortie 
      Height          =   285
      Left            =   10200
      TabIndex        =   12
      Top             =   2640
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
      BackColor       =   14737632
   End
   Begin SToolBox.STimeBox H_Sorte 
      Height          =   285
      Left            =   11640
      TabIndex        =   8
      Top             =   2640
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   503
      BackColor       =   14737632
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   240
   End
   Begin SToolBox.SCommand SCommand2 
      Height          =   615
      Left            =   10200
      TabIndex        =   4
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BackStyle       =   0
      Caption         =   "Supprimer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16576
   End
   Begin VB.TextBox Txt_KME 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      MaxLength       =   6
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox Txt_KM 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      MaxLength       =   6
      TabIndex        =   1
      Top             =   3840
      Width           =   1815
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   615
      Left            =   11520
      TabIndex        =   2
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BackStyle       =   0
      Caption         =   "Enregistre"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
   End
   Begin VB.TextBox txt_Observation 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8280
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4920
      Width           =   4575
   End
   Begin SToolBox.STimeBox H_Entre 
      Height          =   285
      Left            =   11640
      TabIndex        =   11
      Top             =   3240
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   503
      BackColor       =   14737632
   End
   Begin SToolBox.SDateBox cda_Entre 
      Height          =   285
      Left            =   10200
      TabIndex        =   13
      Top             =   3240
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
      BackColor       =   14737632
   End
   Begin SToolBox.SGrid grid_vehicule 
      Height          =   5415
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   9551
      RowMode         =   -1  'True
      GridLineMode    =   1
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin SToolBox.SGrid grid_Conducteur 
      Height          =   5415
      Left            =   2760
      TabIndex        =   16
      Top             =   1800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   9551
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin SToolBox.SGrid grid_destination 
      Height          =   5415
      Left            =   5400
      TabIndex        =   17
      Top             =   1800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   9551
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin VB.Image Img_alarme 
      Height          =   1320
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1320
   End
   Begin VB.Label Lbl_date 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6240
      TabIndex        =   19
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Lbl_heure 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10080
      TabIndex        =   18
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contrôle Vehicule"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   1560
      TabIndex        =   14
      Top             =   480
      Width           =   2910
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "H.Sortie"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8520
      TabIndex        =   10
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "H.Entrée"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8520
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lbl_Compteur 
      BackStyle       =   0  'Transparent
      Caption         =   "Compteur Actuelle!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Compteur Entrée"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8280
      TabIndex        =   6
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Compteur Sortie"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   3720
      Width           =   2895
   End
End
Attribute VB_Name = "Frm_MajFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Okay As Boolean
Public ii As Integer

Public NumFiche As String
Public Matricule As String
Public Conducteur As String
Public Destination As String
Public CompteurEntre As Long
Public CompteurSortie As Long
Public HeureSortie As String
Public HeureEntre As String
Public Observation As String


Dim thekey As Integer
Dim theshift As Integer
Dim itmX As ListItem

Private Sub selectFT(ByVal vcode As String)
    
    
    Dim SQL As String
Dim rs As New ADODB.Recordset

 SQL = "Select * from Fichetraffic where Numero = " & SQLText(vcode)
 rs.Open SQL, CNB, adOpenDynamic
    If Not rs.EOF Then
      If (Not (IsNull(rs("Vehicule")))) Then Matricule = rs("Vehicule")
      If (Not (IsNull(rs("Conducteur")))) Then Conducteur = rs("Conducteur")
      If (Not (IsNull(rs("Destination")))) Then Destination = rs("Destination")
      If (Not (IsNull(rs("CompteurEntre")))) Then CompteurEntre = rs("CompteurEntre")
      If (Not (IsNull(rs("CompteurSortie")))) Then CompteurSortie = rs("CompteurSortie")
      If (Not (IsNull(rs("HeureEntre")))) Then HeureEntre = rs("HeureEntre")
      If (Not (IsNull(rs("HeureSortie")))) Then HeureSortie = rs("HeureSortie")
      If (Not (IsNull(rs("Observation")))) Then Observation = rs("Observation")
        End If
End Sub

Private Sub CmdSave_Click()

Dim SelectedV As String
Dim SelectedC As String
Dim SelectedD As String
Dim SelectedCE As Long
Dim SelectedCS
Dim Observation As String
Dim HeureSortie As Date
Dim HeureEntre As Date
Dim CompteurFT As Long

Dim SQL As String
Dim rs As New ADODB.Recordset

On Error GoTo Err


        SelectedV = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
        SelectedC = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)
        SelectedD = grid_destination.CellText(grid_destination.SelectedRow, grid_destination.SelectedCol)
        SelectedCE = CStr(Txt_KME.Text)
        SelectedCS = CStr(Txt_KM.Text)
        Observation = txt_Observation.Text
        CompteurFT = Frm_Trafic.CompteurVehicule(SelectedV)
        
        HeureSortie = Format(cda_Sortie.Text, "d/m/yyyy") & " " & Format(H_Sorte.Text, "hh:mm:ss")
         If Not (IsNull(SelectedCE)) And (SelectedCE > 0) Then
        HeureEntre = Format(cda_Entre.Text, "d/m/yyyy") & " " & Format(H_Entre.Text, "hh:mm:ss")
        End If
        
    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    'Modification
        
        CNB.BeginTrans
        SQL = "Update fichetraffic Set "
        SQL = SQL & " Vehicule = " & SQLText(SelectedV)
        SQL = SQL & ", CompteurSortie = " & SelectedCS
        SQL = SQL & ", Conducteur = " & SQLText(SelectedC)
        SQL = SQL & ", Destination = " & SQLText(SelectedD)
        SQL = SQL & ", Observation = " & SQLText(Observation)
        SQL = SQL & ",  HeureSortie = " & SQLText(HeureSortie)
        If Not (IsNull(SelectedCE)) And (SelectedCE > 0) Then
        SQL = SQL & ",  CompteurEntre = " & SelectedCE
        SQL = SQL & ",  HeureEntre = " & SQLText(HeureEntre)
        End If
        SQL = SQL & " where fichetraffic.Numero = " & SQLText(NumFiche)
     CNB.Execute SQL
        If Not (IsNull(SelectedCE)) And (SelectedCE > 0) Then
                SQL = "Update Vehicule Set CompteurFT =" & SelectedCE & " where Matricule=" & SQLText(SelectedV)
                CNB.Execute SQL
        Else
                SQL = "Update Vehicule Set CompteurFT =" & SelectedCS & " where Matricule=" & SQLText(SelectedV)
                CNB.Execute SQL
        End If
    CNB.CommitTrans
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
End If
    
    With Frm_Trafic
    .AfficheExterieur
    .AfficheDepot
    .Affiche_Vehicule
    .Affiche_Conducteur
    .Affiche_Destination
    .Txt_KM.Text = ""
    .txt_Observation.Text = ""
    End With
    
    Unload Me
Exit Sub
Err:
CNB.RollbackTrans
Exit Sub
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdSave_GotFocus()

If grid_vehicule.SelectionCount = 0 Then
   MsgBox "Sélectionner Vehicule      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If

If grid_Conducteur.Enabled = True Then
If grid_Conducteur.SelectionCount = 0 Then
   MsgBox "Sélectionner un Conducteur     ", vbInformation
   grid_Conducteur.SetFocus
   Exit Sub
End If
End If

If grid_destination.Enabled = True Then
If grid_destination.SelectionCount = 0 Then
   MsgBox "Sélectionner la destination     ", vbInformation
   grid_destination.SetFocus
   Exit Sub
End If
End If
If Len(Trim(Txt_KM.Text)) = 0 Then
    MsgBox "Compteur Invalide    ", vbInformation
    Txt_KM.SetFocus
    Exit Sub
End If
If Len(Trim(txt_Observation.Text)) = 0 Then
    txt_Observation.Text = "SANS OBSERVATION"
    Exit Sub
End If

End Sub

Private Sub Form_Load()

Dim dat As Date
 
Dim large As Integer
Dim haut As Integer
large = Screen.Width
haut = Screen.Height
Me.Left = 0
Me.Top = 0
Me.Width = large
Me.Height = haut

Lbl_heure.Caption = Format(Time, "hh:mm:ss")
Timer1.Enabled = True
Timer1.Interval = 1000

Call ViderZone(FrmMajFT)



Call selectFT(NumFiche)

Call Initgrid_Vehicule
Call Affiche_Vehicule
Call Initgrid_Conducteur
Call Affiche_Conducteur
Call Initgrid_Destination
Call Affiche_Destination



Img_alarme.Picture = LoadPicture(App.Path & "\Images\button-green.bmp")
dat = Date
Lbl_date.Caption = dat
dat = Time
Lbl_heure.Caption = dat


If Not (IsNull(CompteurSortie)) Then
Txt_KM.Text = ""
Txt_KME.Text = ""

Txt_KM.Text = CompteurSortie
End If
If Not (IsNull(CompteurEntre)) Then
Txt_KME.Text = CompteurEntre
End If

If Not (IsNull(txt_Observation.Text)) Then
    txt_Observation.Text = Observation
End If

If Not (IsNull(HeureSortie)) Then
    cda_Sortie.Text = Format(HeureSortie, "dd/mm/yyyy")
End If

If Not (IsNull(HeureSortie)) Then
    H_Sorte.Text = Format(HeureSortie, "hh:mm:ss")
End If

If Not (IsNull(HeureEntre)) Then
    cda_Entre.Text = Format(HeureEntre, "dd/mm/yyyy")
End If

If Not (IsNull(HeureEntre)) Then
    H_Entre.Text = Format(HeureEntre, "hh:mm:ss")
End If


End Sub

Private Sub grid_Conducteur_GotFocus()
If grid_vehicule.SelectionCount = 0 Then
   MsgBox "Véhicule Obligatoire      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If

End Sub

Private Sub grid_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"

End Sub



Private Sub grid_destination_GotFocus()
If grid_vehicule.SelectionCount = 0 Then
   MsgBox "Sélectionner Vehicule      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If

If grid_Conducteur.Enabled = True Then
If grid_Conducteur.SelectionCount = 0 Then
   MsgBox "Sélectionner Conducteur      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If
End If

End Sub

Private Sub grid_destination_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"

End Sub



Private Sub grid_vehicule_GotFocus()
If grid_Conducteur.Enabled = False Then grid_Conducteur.Enabled = True
If grid_destination.Enabled = False Then grid_destination.Enabled = True
End Sub

Private Sub grid_vehicule_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub grid_vehicule_LostFocus()
Dim i As Integer
Dim SelectedV As String

If grid_vehicule.SelectionCount <> 0 Then
    SelectedV = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
End If
End Sub


Private Sub SCommand2_Click()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim vcode As String

On Error GoTo Err

    If MsgBox("Confirmez vous la suppression de cette " & vbNewLine & "Fiche", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    SQL = "Delete from FicheTraffic where Numero =" & SQLText(NumFiche)
    CNB.Execute SQL
    End If
    
     With Frm_Trafic
    .AfficheExterieur
    .AfficheDepot
    .Affiche_Vehicule
    .Affiche_Conducteur
    .Affiche_Destination
    .Txt_KM.Text = ""
    .txt_Observation.Text = ""
    End With
    
    Unload Me
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Timer1_Timer()
Lbl_heure = Time
End Sub

Private Sub Txt_KM_GotFocus()
If grid_vehicule.SelectionCount = 0 Then
   MsgBox "Sélectionner Vehicule      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If
If grid_Conducteur.Enabled = True Then
If grid_Conducteur.SelectionCount = 0 Then
   MsgBox "Sélectionner Conducteur      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If
End If

If grid_destination.Enabled = True Then
If grid_destination.SelectionCount = 0 Then
   MsgBox "Sélectionner la destination     ", vbInformation
   grid_destination.SetFocus
   Exit Sub
End If
End If




End Sub

Private Sub Txt_KM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Txt_KM_KeyPress(KeyAscii As Integer)
If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub Txt_KM_LostFocus()

Dim SelectedV As String
Dim compteur As Long

If grid_vehicule.SelectionCount <> 0 Then
If Txt_KM <> "" Then
    SelectedV = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
    compteur = CompteurVehicule(SelectedV)
    lbl_Compteur.Caption = compteur
'Controle de compteur
If Val(Txt_KM.Text) < Val(compteur) Then
    If MsgBox("Nouveau compteur invalid" & vbNewLine & "Voulez vous malgré ça l'accepter.?", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
    Txt_KM.SetFocus
    Exit Sub
    End If
End If

If Val(Txt_KM.Text) - Val(compteur) > 1200 Then
    If MsgBox("Nouveau compteur invalid : Plus que 1200 klm" & vbNewLine & "Vlouez vous malgré ça l'accepter.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
    Txt_KM.SetFocus
    Exit Sub
    End If
End If

End If
End If


End Sub



Private Sub Txt_KME_GotFocus()
If grid_vehicule.SelectionCount = 0 Then
   MsgBox "Sélectionner Vehicule      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If
If grid_Conducteur.Enabled = True Then
If grid_Conducteur.SelectionCount = 0 Then
   MsgBox "Sélectionner Conducteur      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If
End If

If grid_destination.Enabled = True Then
If grid_destination.SelectionCount = 0 Then
   MsgBox "Sélectionner la destination     ", vbInformation
   grid_destination.SetFocus
   Exit Sub
End If
End If
If Len(Trim(Txt_KM.Text)) = 0 Then
    MsgBox "Compteur Invalide    ", vbInformation
    Txt_KM.SetFocus
    Exit Sub
End If


End Sub

Private Sub Txt_KME_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Txt_KME_KeyPress(KeyAscii As Integer)
If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub

Private Sub Txt_KME_LostFocus()
Dim SelectedV As String
Dim compteur As Long

If grid_vehicule.SelectionCount <> 0 Then
If Txt_KME <> "" Then
    SelectedV = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
    compteur = CompteurVehicule(SelectedV)
    lbl_Compteur.Caption = compteur
'Controle de compteur
If Val(Txt_KME.Text) < Val(compteur) Then
    If MsgBox("Nouveau compteur invalid" & vbNewLine & "Voulez vous malgré ça l'accepter.?", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
    Txt_KME.SetFocus
    Exit Sub
    End If
End If

If (Val(Txt_KME.Text) > 0) And (Val(Txt_KME) < Val(Txt_KM)) Then
     If MsgBox("Compteur entrée ne Doit Pas Etre inferieur au compteur de sortie" & vbNewLine & "Voulez vous malgré ça l'accepter.?", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
    Txt_KME.SetFocus
    Exit Sub
End If

If Val(Txt_KM.Text) - Val(compteur) > 1200 Then
    If MsgBox("Nouveau compteur invalid : Plus que 1200 klm" & vbNewLine & "Vlouez vous malgré ça l'accepter.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
    Txt_KM.SetFocus
    Exit Sub
    End If
End If

End If
End If
End If
End Sub

Private Sub txt_Observation_GotFocus()
'If grid_vehicule.SelectionCount = 0 Then
'   MsgBox "Sélectionner Vehicule      ", vbInformation
'   grid_vehicule.SetFocus
'   Exit Sub
'End If
'If grid_Conducteur.Enabled = True Then
'If grid_Conducteur.SelectionCount = 0 Then
'   MsgBox "Sélectionner Conducteur      ", vbInformation
'   grid_vehicule.SetFocus
'   Exit Sub
'End If
'End If
'
'If grid_destination.Enabled = True Then
'If grid_destination.SelectionCount = 0 Then
'   MsgBox "Sélectionner la destination     ", vbInformation
'   grid_destination.SetFocus
'   Exit Sub
'End If
'End If
'
'If Len(Trim(Txt_KM.Text)) = 0 Then
'    MsgBox "Compteur Invalide    ", vbInformation
'    Txt_KM.SetFocus
'    Exit Sub
'End If
End Sub

Private Sub Txt_Observation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Public Sub Initgrid_Vehicule()
With grid_vehicule
    .Redraw = False
    ' Allow the grid to be grouped, but
    ' don't show the grouping box
    .HideGroupingBox = True
    .AllowGrouping = True
    ' Group rows will be shown by
    ' a gradient underline
    .GroupRowBackColor = vbWindowBackground
    .GroupRowForeColor = vbWindowText
    
    .GridLineColor = vbWindowBackground
    .GridFillLineColor = vbWindowBackground
    .GridLines = True
    
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .DrawFocusRectangle = False
    
    .AddColumn "Matricule", "", , , 140
    
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub

Public Sub Affiche_Vehicule()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim j As Integer
Dim Couleur As String

grid_vehicule.ClearRows

SQL = "Select * from vehicule where Actif = 1 order by Matricule"
rs.Open SQL, CNB, adOpenKeyset

If Not rs.EOF Then
    grid_vehicule.Redraw = False
    While Not rs.EOF
        Couleur = "vbRed"
            If (rs("Matricule") = Matricule) Then
                Couleur = "vbGreen"
            End If
        With grid_vehicule
            .AddRow
            If Couleur = "vbRed" Then
                .CellDetails .Rows, .ColumnIndex("Matricule"), rs("Matricule")
            Else
                .CellDetails .Rows, .ColumnIndex("Matricule"), rs("Matricule"), , , vbGreen
            End If
        End With
            rs.MoveNext
    Wend
    grid_vehicule.Redraw = True
 End If


End Sub

Public Sub Initgrid_Conducteur()
With grid_Conducteur
    .Redraw = False
    ' Allow the grid to be grouped, but
    ' don't show the grouping box
    .HideGroupingBox = True
    .AllowGrouping = True
    ' Group rows will be shown by
    ' a gradient underline
    .GroupRowBackColor = vbWindowBackground
    .GroupRowForeColor = vbWindowText
    
    .GridLineColor = vbWindowBackground
    .GridFillLineColor = vbWindowBackground
    .GridLines = True
    
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .DrawFocusRectangle = False
    
    .AddColumn "Libelle", "", , , 140
    
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub

Public Sub Affiche_Conducteur()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim Couleur As String
grid_Conducteur.ClearRows
SQL = "Select * from Personnel where actif=1  order by Libelle"
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    grid_Conducteur.Redraw = False
    While Not rs.EOF
         Couleur = "vbRed"
        
                If (rs("Libelle") = Conducteur) Then
                Couleur = "vbGreen"
            End If
        With grid_Conducteur
            .AddRow
             If Couleur = "vbRed" Then
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Libelle")
            Else
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Libelle"), , , vbGreen
            End If
        End With
        rs.MoveNext
    Wend
    grid_Conducteur.Redraw = True
End If

End Sub

Public Sub Initgrid_Destination()
With grid_destination
    .Redraw = False
    ' Allow the grid to be grouped, but
    ' don't show the grouping box
    .HideGroupingBox = True
    .AllowGrouping = True
    ' Group rows will be shown by
    ' a gradient underline
    .GroupRowBackColor = vbWindowBackground
    .GroupRowForeColor = vbWindowText
    
    .GridLineColor = vbWindowBackground
    .GridFillLineColor = vbWindowBackground
    .GridLines = True
    
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .DrawFocusRectangle = False
    
    .AddColumn "LibelleD", "", , , 140
    
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub


Public Sub Affiche_Destination()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim Couleur As String
grid_destination.ClearRows
SQL = "Select * from Destination where Actif=1  order by Type,Libelle"
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    grid_destination.Redraw = False
    While Not rs.EOF
     Couleur = "vbRed"
        
                If (rs("Libelle") = Destination) Then
                Couleur = "vbGreen"
            End If
        With grid_destination
            .AddRow
            If Couleur = "vbRed" Then
            .CellDetails .Rows, .ColumnIndex("LibelleD"), rs("Libelle")
             Else
             .CellDetails .Rows, .ColumnIndex("LibelleD"), rs("Libelle"), , , vbGreen
             End If
        End With
        rs.MoveNext
    Wend
    grid_destination.Redraw = True
    
    End If

End Sub

Private Function return_Compteur() As Long
Dim rD As New ADODB.Recordset
Dim SQL As String
return_Compteur = 0
SQL = "select Max(Numero) from fichetraffic "
rD.Open SQL, CNB, adOpenKeyset
If Not rD.EOF Then
return_Compteur = rD(0)
End If

rD.Close
End Function

Public Function CompteurVehicule(ByVal vcode As String) As Long
    Dim rD As New ADODB.Recordset
    Dim SQL As String
    SQL = "Select * from vehicule where Matricule = " & SQLText(vcode)
    rD.Open SQL, CNB, adOpenKeyset
    If Not (IsNull(rD("CompteurFT"))) Then
        CompteurVehicule = rD("CompteurFT")
    End If
End Function


