VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form Frm_MajTrafic 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maj Fiche Traffic"
   ClientHeight    =   7560
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic_Supp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   4560
      ScaleHeight     =   825
      ScaleWidth      =   5055
      TabIndex        =   21
      Top             =   840
      Width           =   5055
      Begin VB.Image Cmd_Add 
         Height          =   375
         Left            =   3120
         Picture         =   "Frm_MajTrafic.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Voulez-Vous Réajouter!... "
         BeginProperty Font 
            Name            =   "Footlight MT Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Lbl_Supp 
         BackStyle       =   0  'Transparent
         Caption         =   "Fiche supprime par :"
         BeginProperty Font 
            Name            =   "Footlight MT Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   0
         Width           =   3975
      End
   End
   Begin SToolBox.SDateBox cda_Sortie 
      Height          =   285
      Left            =   9120
      TabIndex        =   3
      Top             =   2040
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
      BackColor       =   14737632
   End
   Begin SToolBox.STimeBox H_Sorte 
      Height          =   285
      Left            =   10560
      TabIndex        =   4
      Top             =   2040
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   503
      BackColor       =   14737632
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   240
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
      Left            =   10080
      MaxLength       =   6
      TabIndex        =   8
      Top             =   3840
      Width           =   1695
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
      Left            =   10080
      MaxLength       =   6
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
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
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4320
      Width           =   4455
   End
   Begin SToolBox.STimeBox H_Entre 
      Height          =   285
      Left            =   10560
      TabIndex        =   6
      Top             =   2640
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   503
      BackColor       =   14737632
   End
   Begin SToolBox.SDateBox cda_Entre 
      Height          =   285
      Left            =   9120
      TabIndex        =   5
      Top             =   2640
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
      BackColor       =   14737632
   End
   Begin SToolBox.SCommand Cmd_Destination 
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   1680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "Destination"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_MajTrafic.frx":11D22
      BackColor       =   16777215
   End
   Begin SToolBox.SCommand Cmd_Vehicule 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "Vehicule"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_MajTrafic.frx":120E0
      BackColor       =   16777215
   End
   Begin SToolBox.SCommand Cmd_Conducteur 
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   1680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "Conducteur"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_MajTrafic.frx":12812
      BackColor       =   16777215
   End
   Begin SToolBox.SGrid grid_vehicule 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   9551
      RowMode         =   -1  'True
      GridLineMode    =   1
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   14737632
      HighlightBackColor=   33023
      NoFocusHighlightBackColor=   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Header          =   0   'False
      HeaderButtons   =   0   'False
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin SToolBox.SGrid grid_Conducteur 
      Height          =   5415
      Left            =   2520
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   9551
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   14737632
      HighlightBackColor=   33023
      NoFocusHighlightBackColor=   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Header          =   0   'False
      HeaderButtons   =   0   'False
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin SToolBox.SGrid grid_destination 
      Height          =   5415
      Left            =   4920
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   9551
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   14737632
      HighlightBackColor=   33023
      NoFocusHighlightBackColor=   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Header          =   0   'False
      HeaderButtons   =   0   'False
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin VB.Image Cmd_Annul 
      Height          =   520
      Left            =   10080
      Picture         =   "Frm_MajTrafic.frx":12B8C
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Image CmdSave 
      Height          =   495
      Left            =   8280
      Picture         =   "Frm_MajTrafic.frx":24CB6
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Image PicBox_Footer 
      Height          =   495
      Left            =   -120
      Picture         =   "Frm_MajTrafic.frx":36F38
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   13575
   End
   Begin VB.Image Img_alarme 
      Height          =   1080
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1200
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   15
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
      Left            =   7440
      TabIndex        =   14
      Top             =   1920
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
      Left            =   7440
      TabIndex        =   13
      Top             =   2520
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
      Left            =   9480
      TabIndex        =   12
      Top             =   1680
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
      Left            =   7440
      TabIndex        =   11
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Image PicBox_Header 
      Height          =   1000
      Left            =   0
      Picture         =   "Frm_MajTrafic.frx":B61F6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
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
      Left            =   7440
      TabIndex        =   9
      Top             =   3120
      Width           =   2895
   End
End
Attribute VB_Name = "Frm_MajTrafic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Matricule As String, Conducteur As String, DESTINATION As String
    Dim SelectedVehic  As String, SelectedCond As String, SelectedDest As String
    Public NumFiche As String
    Public Okay As Boolean
    Public ii As Integer
    Public CompteurEntre As Long
    Public CompteurSortie As Long
    Public HeureSortie As String
    Public HeureEntre As String
    Public Observation As String
    Dim thekey As Integer
    Dim theshift As Integer
    Dim itmX As ListItem
    Dim HEntre As String
    Dim OEntre As String


'~~~~~~~~~~~~~~~~~~~~
    'Mise en Forme~~~
'~~~~~~~~~~~~~~~~~~~~
Private Sub Form_Load()
    Dim dat As Date
    Call ViderZone(Frm_MajTrafic)
    Call Initgrid_Conducteur
    Call Initgrid_Destination
    Call Initgrid_Vehicule
    Call selectFT(NumFiche)
    Call Affiche_Vehicule
    Call Affiche_Conducteur
    Call Affiche_Destination
    Lbl_heure.Caption = Format(Time, "hh:mm:ss")
    Timer1.Enabled = True
    Timer1.Interval = 1000
    Pic_Supp.Visible = False
    Img_alarme.Picture = LoadPicture(App.Path & "\Images\Trafic_V.bmp")
    dat = Date
    Lbl_date.Caption = dat
    dat = Time
    Lbl_heure.Caption = dat
    If Not (IsNull(CompteurSortie)) Then
        Txt_KM.Text = ""
        Txt_KME.Text = ""
        Txt_KM.Text = CompteurSortie
    End If
    If Not (IsNull(CompteurEntre)) Then Txt_KME.Text = CompteurEntre Else Txt_KME.Text = ""
    If Not (IsNull(txt_Observation.Text)) Then txt_Observation.Text = Observation
    If Not (IsNull(HeureSortie)) Then cda_Sortie.Text = Format(HeureSortie, "dd/mm/yyyy")
    If Not (IsNull(HeureSortie)) Then H_Sorte.Text = Format(HeureSortie, "hh:mm:ss")
    If Not (IsNull(HeureEntre)) Then cda_Entre.Text = Format(HeureEntre, "dd/mm/yyyy") Else cda_Entre.Text = ""
    If Not (IsNull(HeureEntre)) Then H_Entre.Text = Format(HeureEntre, "hh:mm:ss") Else H_Entre.Text = ""
    grid_Conducteur.Enabled = True
    grid_destination.Enabled = True
    grid_vehicule.Enabled = True
End Sub
Private Sub Form_Resize()
    Dim WidthForm As Integer
    WidthForm = Frm_Main.ACB_Main.Width
        PicBox_Header.Width = WidthForm
        PicBox_Footer.Width = WidthForm - 2200
        Lbl_heure.Left = WidthForm - 2000
        Lbl_date.Left = WidthForm - 4100
End Sub
Private Sub Cmd_Annul_Click()
    Unload Me
End Sub

'~~~~~~~~~~~~~~~~~
    'ControlBox~~~
'~~~~~~~~~~~~~~~~~
Private Sub grid_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub grid_vehicule_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub grid_destination_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Txt_KM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub Txt_KM_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub Txt_KME_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub Txt_KME_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub Txt_Observation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call CmdSave_Click
End Sub
Private Sub Timer1_Timer()
    Lbl_heure = Time
End Sub
Private Sub EnbDisb(ByVal TYP As Boolean)
    grid_Conducteur.Enabled = TYP
    grid_vehicule.Enabled = TYP
    grid_destination.Enabled = TYP
    Cmd_Annul.Enabled = TYP
    CmdSave.Enabled = TYP
    txt_Observation.Enabled = TYP
    Txt_KME.Enabled = TYP
    Txt_KM.Enabled = TYP
    H_Entre.Enabled = TYP
    cda_Entre.Enabled = TYP
    H_Sorte.Enabled = TYP
    cda_Sortie.Enabled = TYP
End Sub
Private Sub grid_Conducteur_LostFocus()
    If grid_Conducteur.SelectionCount <> 0 Then
        SelectedCond = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, 1)
    End If
End Sub
Private Sub grid_destination_LostFocus()

    If grid_destination.SelectionCount <> 0 Then
        SelectedDest = grid_destination.CellText(grid_destination.SelectedRow, 1)
        If (grid_destination.CellText(grid_destination.SelectedRow, 2) = "KIOSQUE AGIL" Or grid_destination.CellText(grid_destination.SelectedRow, 2) = "KIOSQUE SHELL" _
        Or grid_destination.CellText(grid_destination.SelectedRow, 2) = "REPARATION") And cda_Entre.Text <> "" Then
            If Val(Txt_KME.Text) = 0 Or Txt_KME.Text = "" Then Txt_KME.Text = Txt_KM.Text
        End If
    End If
End Sub
Private Sub grid_vehicule_LostFocus()
    If grid_vehicule.SelectionCount <> 0 Then
        SelectedVehic = grid_vehicule.CellText(grid_vehicule.SelectedRow, 1)
        If Matricule <> SelectedVehic Then Txt_KM.Text = CompteurVehicule(SelectedVehic)
        If Val(Txt_KME.Text) > 0 And Matricule <> SelectedVehic Then
            If CompteurEntre = Val(Txt_KME.Text) Then Txt_KME.Text = "0"
        End If
    End If
End Sub
Private Sub CmdSave_GotFocus()
    '--Sélection Véhicule***
    If SelectedVehic = "" Then
       MsgBox "Sélectionner Vehicule      ", vbInformation
       grid_vehicule.SetFocus
       Exit Sub
    End If
    '--Sélection Conducteur***
    If SelectedCond = "" Then
       MsgBox "Sélectionner un Conducteur     ", vbInformation
       grid_Conducteur.SetFocus
       Exit Sub
    End If
    '--Sélection déstination***
    If SelectedDest = "" Then
       MsgBox "Sélectionner la destination     ", vbInformation
       grid_destination.SetFocus
       Exit Sub
    End If
    If Len(Trim(Txt_KM.Text)) = 0 Or Txt_KM.Text = 0 Then
        MsgBox "Compteur Invalide    ", vbInformation
        Txt_KM.SetFocus
        Exit Sub
    End If
    If Txt_KME.Text = 0 And H_Entre.Text <> "" Then
        MsgBox "Compteur d'entrée Invalide    ", vbInformation
        Txt_KME.SetFocus
        Exit Sub
    End If
    If Len(Trim(txt_Observation.Text)) = 0 Then
        txt_Observation.Text = "SANS OBSERVATION"
        Exit Sub
    End If
End Sub
'~~~~~~~~~~~~~~~~
    'Actualise~~~
'~~~~~~~~~~~~~~~~
Private Sub Cmd_Conducteur_Click()
    Call Affiche_Conducteur
End Sub
Private Sub Cmd_Destination_Click()
    Call Affiche_Destination
End Sub
Private Sub Cmd_Vehicule_Click()
    Call Affiche_Vehicule
End Sub
'~~~~~~~~~~~~~~~~~~~~~~
    'Initialize Grid~~~
'~~~~~~~~~~~~~~~~~~~~~~
Public Sub Initgrid_Vehicule()
    With grid_vehicule
        .Redraw = False
        .HideGroupingBox = True
        .AllowGrouping = True
        .GroupRowBackColor = vbWindowBackground
        .GroupRowForeColor = vbWindowText
        .GridLineColor = vbWindowBackground
        .GridFillLineColor = vbWindowBackground
        .GridLines = True
        .SelectionAlphaBlend = True
        .SelectionOutline = True
        .DrawFocusRectangle = False
        .AddColumn "Code", "", , , , False
        .AddColumn "Matricule", "", , , 140
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Public Sub Initgrid_Conducteur()
    With grid_Conducteur
        .Redraw = False
        .HideGroupingBox = True
        .AllowGrouping = True
        .GroupRowBackColor = vbWindowBackground
        .GroupRowForeColor = vbWindowText
        .GridLineColor = vbWindowBackground
        .GridFillLineColor = vbWindowBackground
        .GridLines = True
        .SelectionAlphaBlend = True
        .SelectionOutline = True
        .DrawFocusRectangle = False
        .AddColumn "Code", "", , , , False
        .AddColumn "Libelle", "", , , 140
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Public Sub Initgrid_Destination()
    With grid_destination
        .Redraw = False
        .HideGroupingBox = True
        .AllowGrouping = True
        .GroupRowBackColor = vbWindowBackground
        .GroupRowForeColor = vbWindowText
        .GridLineColor = vbWindowBackground
        .GridFillLineColor = vbWindowBackground
        .GridLines = True
        .SelectionAlphaBlend = True
        .SelectionOutline = True
        .DrawFocusRectangle = False
        .AddColumn "Code", "", , , , False
        .AddColumn "LibelleD", "", , , 140
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Selectionner Fiche Traffic~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub selectFT(ByVal VCode As String)
    Dim LObj_Find As New Traffic, Lrs_Traffic As New Recordset
On Error GoTo Err
    Set Lrs_Traffic = LObj_Find.GetRow_Traffic_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Traffic.EOF Then
        If (Not (IsNull(Lrs_Traffic("Vehicule")))) Then Matricule = Lrs_Traffic("Vehicule")
        If (Not (IsNull(Lrs_Traffic("Conducteur")))) Then Conducteur = Lrs_Traffic("Conducteur")
        If (Not (IsNull(Lrs_Traffic("Destination")))) Then DESTINATION = Lrs_Traffic("Destination")
        If (Not (IsNull(Lrs_Traffic("CompteurEntre")))) Then CompteurEntre = Lrs_Traffic("CompteurEntre") Else CompteurEntre = 0
        If (Not (IsNull(Lrs_Traffic("CompteurSortie")))) Then CompteurSortie = Lrs_Traffic("CompteurSortie")
        If (Not (IsNull(Lrs_Traffic("HeureEntre")))) Then HeureEntre = Lrs_Traffic("HeureEntre") Else HeureEntre = ""
        If (Not (IsNull(Lrs_Traffic("HeureSortie")))) Then HeureSortie = Lrs_Traffic("HeureSortie")
        If (Not (IsNull(Lrs_Traffic("ObservationEntre")))) Then Observation = Lrs_Traffic("ObservationEntre") Else Observation = "Sans Observation"
        If (Not (IsNull(Lrs_Traffic("HeureEntre")))) Then HEntre = Lrs_Traffic("HeureEntre") Else HEntre = ""
        If (Not (IsNull(Lrs_Traffic("OperateurEntre")))) Then OEntre = Lrs_Traffic("OperateurEntre") Else OEntre = ""
        SelectedVehic = Matricule
        SelectedCond = Conducteur
        SelectedDest = DESTINATION
        If Lrs_Traffic("Supp") = "O" Then
            Dim Lobj_User As New Utilisateur, Lrs_User As New Recordset, UserDeleted As String
            
            Set Lrs_User = Lobj_User.GetRow_UserByCode(ErrNumber, ErrDescription, ErrSourceDetail, Lrs_Traffic("userdelete"), CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            If Not Lrs_User.EOF Then UserDeleted = Lrs_User("nomprn") Else UserDeleted = ""
            Pic_Supp.Visible = True
            Lbl_Supp.Caption = "Fichie supprimer par : " & UserDeleted
            Call EnbDisb(False)
        Else
            Pic_Supp.Visible = False
            Call EnbDisb(True)
        End If
    End If
    Set Lrs_Traffic = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'~~~~~~~~~~~~~~~~~~~~
    'Affiche Grids~~~
'~~~~~~~~~~~~~~~~~~~~
Public Sub Affiche_Vehicule()
    Dim LObj_Find As New VEHICULE, Lrs_Vehicule As New Recordset, Couleur As String
On Error GoTo Err
    grid_vehicule.ClearRows
    Set Lrs_Vehicule = LObj_Find.GetAllActifVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Vehicule.EOF Then
        grid_vehicule.Redraw = False
        While Not Lrs_Vehicule.EOF
            Couleur = "vbRed"
            If (Lrs_Vehicule("code") = Matricule) Then
                Couleur = "vbGreen"
            End If
            With grid_vehicule
                .AddRow
                If Couleur = "vbRed" Then
                    .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Vehicule("Code")
                    .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Vehicule("Matricule")
                Else
                    .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Vehicule("Code")
                    .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Vehicule("Matricule"), , , vbGreen
                End If
            End With
            Lrs_Vehicule.MoveNext
        Wend
        grid_vehicule.Redraw = True
     End If
     Set Lrs_Vehicule = Nothing
     
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub Affiche_Conducteur()
    Dim LObj_Find As New Conducteur, Lrs_Conducteur As New Recordset, Couleur As String
On Error GoTo Err
    grid_Conducteur.ClearRows
    Set Lrs_Conducteur = LObj_Find.GetAll_ConducteursActif(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Conducteur.EOF Then
        grid_Conducteur.Redraw = False
        While Not Lrs_Conducteur.EOF
             Couleur = "vbRed"
            If (Lrs_Conducteur("code") = Conducteur) Then
                Couleur = "vbGreen"
            End If
            With grid_Conducteur
                .AddRow
                 If Couleur = "vbRed" Then
                    .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Conducteur("Code")
                    .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Conducteur("Libelle")
                Else
                    .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Conducteur("Code")
                    .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Conducteur("Libelle"), , , vbGreen
                End If
            End With
            Lrs_Conducteur.MoveNext
        Wend
        grid_Conducteur.Redraw = True
    End If
    Set Lrs_Conducteur = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub Affiche_Destination()
    Dim LObj_Find As New DESTINATION, Lrs_Destination As New Recordset, Couleur As String
On Error GoTo Err
    grid_destination.ClearRows
    Set Lrs_Destination = LObj_Find.GetAll_DestinationActif(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Destination.EOF Then
        grid_destination.Redraw = False
        While Not Lrs_Destination.EOF
            Couleur = "vbRed"
            If (Lrs_Destination("numero") = DESTINATION) Then
                Couleur = "vbGreen"
            End If
            With grid_destination
                .AddRow
                If Couleur = "vbRed" Then
                    .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Destination("numero")
                    .CellDetails .Rows, .ColumnIndex("LibelleD"), Lrs_Destination("Libelle")
                Else
                    .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Destination("numero")
                    .CellDetails .Rows, .ColumnIndex("LibelleD"), Lrs_Destination("Libelle"), , , vbGreen
                End If
            End With
            Lrs_Destination.MoveNext
        Wend
        grid_destination.Redraw = True
    End If
    Set Lrs_Destination = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'~~~~~~~~~~~~~~~~~
    'Ré-ajouter~~~
'~~~~~~~~~~~~~~~~~
Private Sub Cmd_Add_Click()
    Dim LObj_Traffic As New Traffic, VCode As String
On Error GoTo Err
    If (CHECK_ACCES("Supp_FT", LInt_UserId) = False) Then
        MsgBox "Ré-ajoutation n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour ajouter Traffic", vbExclamation
        Exit Sub
    Else
        If MsgBox("Confirmez vous l'ajoutation de cette " & vbNewLine & "Fiche", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
            Call LObj_Traffic.Delete_Add_Traffic(ErrNumber, ErrDescription, ErrSourceDetail, NumFiche, "N", LInt_UserIdMaj, CNB)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
            Set LObj_Traffic = Nothing
            MsgBox "Fiche ajoutée...          ", vbInformation, App.ProductName
            EnbDisb (True)
            Pic_Supp.Visible = False
        End If
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Function Return_LibDest(ByVal cdDest As String) As String

Dim Lrs_Find As New Recordset, Lobj_FindD As New DESTINATION

Return_LibDest = ""
Set Lrs_Find = Lobj_FindD.GetRow_Destination_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, cdDest, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
Set Lobj_FindD = Nothing
If Not Lrs_Find.EOF Then
    Return_LibDest = Lrs_Find.Fields("Libelle")
End If
Lrs_Find.Close
End Function

'~~~~~~~~~~~~~~~~~~~~~
    'Enregistrement~~~
'~~~~~~~~~~~~~~~~~~~~~
Private Sub CmdSave_Click()
    Dim LOBJ_ACCES As New Utilisateur, LObj_Traffic As New Traffic, Lobj_Vehicule As New VEHICULE, Lrs_Traffic As New Recordset
    Dim SelectedCE As Long, SelectedCS As Long, CompteurFT As Long
    Dim Observation As String, HeureSortie As Date, HeureEntre As Date
On Error GoTo Err
    If (CHECK_ACCES("Maj_FT", LInt_UserIdMaj) = False) Then
        MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-être pas des autorisations nécessaires pour Modifier Traffic", vbExclamation
        Exit Sub
    Else
        Call grid_destination_LostFocus
        Call grid_Conducteur_LostFocus
        Call grid_vehicule_LostFocus

        If cda_Entre.Text < cda_Sortie.Text And cda_Entre.Text <> "" Then
            MsgBox "Date d'entrée non valide!....          ", vbExclamation, App.ProductName
            Exit Sub
        End If
        If SelectedVehic = "" Then
           MsgBox "Sélectionner Véhicule      ", vbInformation
           grid_vehicule.SetFocus
           Exit Sub
        End If
        '--Sélection Conducteur***
        If SelectedCond = "" Then
           MsgBox "Sélectionner un Conducteur     ", vbInformation
           grid_Conducteur.SetFocus
           Exit Sub
        End If
        '--Sélection déstination***
        If SelectedDest = "" Then
           MsgBox "Sélectionner la destination     ", vbInformation
           grid_destination.SetFocus
           Exit Sub
        End If
        If Val(Txt_KME.Text) >= (Val(Txt_KM.Text) + 1200) Then
            MsgBox "Compteur invalide!....       " & vbCr & "Plus de 1200 Km" & vbCr & "Ancien Compteur est: " & Val(Txt_KM.Text) & " KM" & vbCr & vbCr & "Vérifier le compteur saisie", vbCritical, App.ProductName
            Exit Sub
        End If
        If (Len(Trim(Txt_KM.Text)) = 0 Or Val(Txt_KM.Text) = 0) And (SelectedVehic <> "00038" And SelectedVehic <> "0014") Then
            MsgBox "Compteur de sortie Invalide    ", vbInformation
            Txt_KM.SetFocus
            Exit Sub
        End If
        If (Val(Txt_KME.Text) = 0 Or Len(Trim(Txt_KME.Text)) = 0) And (H_Entre.Text <> "") _
        And (SelectedVehic <> "00038" And SelectedVehic <> "0014") Then
            MsgBox "Compteur d'entrée Invalide    ", vbInformation
            Txt_KME.SetFocus
            Exit Sub
        End If
        
        If (Val(Txt_KME.Text) < Val(Txt_KM.Text)) And (H_Entre.Text <> "") Then
            MsgBox "Vérifier les compteurs d'entrée et de sortie ", vbInformation
            Exit Sub
        End If
        If (Val(Txt_KME.Text) = Val(Txt_KM.Text)) And (SelectedVehic <> "00038" And SelectedVehic <> "0014") _
          And (Return_LibDest(SelectedDest) <> "KIOSQUE AGIL" And Return_LibDest(SelectedDest) <> "KIOSQUE SHELL" _
          And Return_LibDest(SelectedDest) <> "REPARATION") Then
            MsgBox "Vérifier les compteurs d'entrée et de sortie ", vbInformation
            Exit Sub
        End If
        If Len(Trim(txt_Observation.Text)) = 0 Then
            txt_Observation.Text = "SANS OBSERVATION"
        End If

        If Txt_KME.Text <> "" And Val(Txt_KME.Text) <> 0 Then SelectedCE = CStr(Txt_KME.Text)
        SelectedCS = CStr(Txt_KM.Text)
        CompteurFT = CompteurVehicule(SelectedVehic)
        HeureSortie = Format(cda_Sortie.Text, "d/m/yyyy") & " " & Format(H_Sorte.Text, "hh:mm:ss")
        If Not (IsNull(SelectedCE)) And (SelectedCE > 0) Then HeureEntre = Format(cda_Entre.Text, "d/m/yyyy") & " " & Format(Replace(H_Entre.Text, "_", "0"), "hh:mm:ss")
        If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
            'Modification
            Set Lrs_Traffic = CreateEmptyRS_Traffic()
            With Lrs_Traffic
                .AddNew
                .Fields("Vehicule") = SelectedVehic
                .Fields("CompteurSortie") = SelectedCS
                .Fields("Conducteur") = SelectedCond
                .Fields("Destination") = SelectedDest
                .Fields("HeureSortie") = Format(HeureSortie, "dd/mm/yyyy hh:mm:ss")
                .Fields("CompteurEntre") = SelectedCE
                If Not (IsNull(SelectedCE)) And (SelectedCE > 0) Then .Fields("HeureEntre") = Format(HeureEntre, "dd/mm/yyyy hh:mm:ss")
                .Fields("userupdate") = LInt_UserIdMaj
                .Fields("OperateurEntre") = LStr_NameUser
                .Fields("Observation") = txt_Observation.Text
            End With
            Call LObj_Traffic.UpDate_Traffic(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Traffic, NumFiche, SelectedCE, SelectedDest)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
            Set LObj_Traffic = Nothing
            Set Lrs_Traffic = Nothing
            If Not (IsNull(SelectedCE)) And (SelectedCE > 0) Then
                Call Lobj_Vehicule.UpdateCompteurFT_Vehicule(ErrNumber, ErrDescription, ErrSourceDetail, SelectedVehic, SelectedCE, CNB)
                If ErrNumber <> 0 Then
                    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                    ErrNumber = 0
                    Exit Sub
                End If
                Set Lobj_Vehicule = Nothing
            Else
                Set Lobj_Vehicule = New VEHICULE
                Call Lobj_Vehicule.UpdateCompteurFT_Vehicule(ErrNumber, ErrDescription, ErrSourceDetail, SelectedVehic, SelectedCS, CNB)
                If ErrNumber <> 0 Then
                    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                    ErrNumber = 0
                    Exit Sub
                End If
                Set Lobj_Vehicule = Nothing
            End If
            MsgBox "Enregistrement terminé avec succé  ", vbInformation, App.ProductName
        End If
        With Frm_Trafic
            .AfficheExterieur
            .AfficheDepot
            .Affiche_Vehicule
            .Affiche_Conducteur
            .Affiche_Destination
            .Txt_KM.Text = ""
            .Form_Initialize
            .cmd_r_Click
        End With
        Unload Me
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~
    'Vérifier Compteur~~~
'~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Txt_KM_LostFocus()
    Dim Compteur As Long
On Error GoTo Err
    If grid_vehicule.SelectionCount <> 0 Then
        If Txt_KM <> "" Then
            SelectedVehic = grid_vehicule.CellText(grid_vehicule.SelectedRow, 1)
            Compteur = CompteurVehicule(SelectedVehic)
            lbl_Compteur.Caption = Compteur
            If Val(Txt_KM.Text) - Val(Compteur) > 1200 Then
                If MsgBox("Nouveau compteur invalid : Plus que 1200 klm" & vbNewLine & "Vlouez vous malgré ça l'accepter.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                    Txt_KM.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub Txt_KME_LostFocus()
    Dim Compteur As Long
On Error GoTo Err
    If grid_vehicule.SelectionCount <> 0 Then
        If Txt_KME <> "" Or Val(Txt_KME.Text) > 0 Then
            SelectedVehic = grid_vehicule.CellText(grid_vehicule.SelectedRow, 1)
            Compteur = CompteurVehicule(SelectedVehic)
            lbl_Compteur.Caption = Compteur
            If (Val(Txt_KME.Text) > 0) And (Val(Txt_KME) < Val(Txt_KM)) Then
                If MsgBox("Compteur entrée ne doit pas être inférieur au compteur de sortie" & vbNewLine & "Voulez vous malgré ça l'accepter.?", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
                    Txt_KME.SetFocus
                    Exit Sub
                End If
                If Val(Txt_KM.Text) - Val(Compteur) > 1200 Then
                    If MsgBox("Nouveau compteur invalid : Plus que 1200 klm" & vbNewLine & "Vlouez vous malgré ça l'accepter.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                        Txt_KM.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub


