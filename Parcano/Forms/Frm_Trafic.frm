VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Trafic 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gestion de Traffic Auto"
   ClientHeight    =   11430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14490
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Frm_Trafic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11430
   ScaleWidth      =   14490
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   30000
      Left            =   4440
      Top             =   1200
   End
   Begin VB.TextBox TXT_PWD 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   20
      Top             =   1200
      Width           =   2775
   End
   Begin MSComctlLib.ListView LSV_Exterieur 
      Height          =   2175
      Left            =   120
      TabIndex        =   18
      Top             =   6000
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Matricule"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Conducteur"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Destination"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Heure.S"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Heure.E"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "CPT.S"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "CPT.E"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Distance (KM)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Diff.Dist"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Durée"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Diff.Durée"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Op. Sortie"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Op. Entre"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   5400
      Top             =   240
   End
   Begin SToolBox.SCommand Cmd_Destination 
      Height          =   375
      Left            =   5400
      TabIndex        =   17
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
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
      BackColor       =   16777215
   End
   Begin SToolBox.SGrid Grid_Destination 
      Height          =   3495
      Left            =   5400
      TabIndex        =   16
      Top             =   2160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6165
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      NoFocusHighlightBackColor=   14737632
      AlternateRowBackColor=   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin SToolBox.SCommand Cmd_Conducteur 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
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
      Picture         =   "Frm_Trafic.frx":000C
      BackColor       =   16777215
   End
   Begin SToolBox.SCommand Cmd_Vehicule 
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
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
      Picture         =   "Frm_Trafic.frx":0386
      BackColor       =   16777215
   End
   Begin SToolBox.SGrid grid_vehicule 
      Height          =   3495
      Left            =   2760
      TabIndex        =   12
      Top             =   2160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6165
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   -2147483644
      ForeColor       =   0
      NoFocusHighlightBackColor=   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin SToolBox.SGrid grid_Conducteur 
      Height          =   3495
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6165
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      NoFocusHighlightBackColor=   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin VB.PictureBox Pic_confirmation 
      BackColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   8040
      ScaleHeight     =   3675
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   2160
      Width           =   3375
      Begin VB.TextBox Txt_KM 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   0
         MaxLength       =   6
         TabIndex        =   0
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Lbl_Vehicule 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         TabIndex        =   11
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Lbl_Conducteur 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Lbl_CmptSt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Compteur"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl_Compteur 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Compteur Actuelle !!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2655
      End
      Begin VB.Image Img_Conducteur 
         Height          =   705
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   855
      End
      Begin VB.Image Img_Vehicule 
         Height          =   735
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Lab_Distination 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         TabIndex        =   6
         Top             =   2640
         Width           =   3975
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   0
         Picture         =   "Frm_Trafic.frx":0AB8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3255
      End
      Begin VB.Image Image5 
         Height          =   495
         Left            =   0
         Picture         =   "Frm_Trafic.frx":7DCD2
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   3255
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   240
   End
   Begin MSComctlLib.ListView Lsv_Depot 
      Height          =   3855
      Left            =   4395
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vehicule a l'interieure"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label LBL_Access 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   22
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Lbl_user 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User Connecté :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   21
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label PWD 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   19
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image CmdSave 
      Height          =   3735
      Left            =   11400
      Picture         =   "Frm_Trafic.frx":FCF90
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image CmdPrint 
      Height          =   500
      Left            =   8400
      Picture         =   "Frm_Trafic.frx":12331E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image cmd_r 
      Height          =   480
      Left            =   6240
      Picture         =   "Frm_Trafic.frx":134988
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image SCommand1 
      Height          =   500
      Left            =   10560
      Picture         =   "Frm_Trafic.frx":14431E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image Img_alarme 
      Height          =   840
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   840
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   120
      Picture         =   "Frm_Trafic.frx":155D90
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   13095
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   120
      Picture         =   "Frm_Trafic.frx":157B16
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   7815
   End
   Begin VB.Label Lbl_heure 
      Alignment       =   1  'Right Justify
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
      Left            =   10200
      TabIndex        =   3
      Top             =   0
      Width           =   1455
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
      Left            =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contrôle Vehicule"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2970
   End
   Begin VB.Image PicBox_Header 
      Height          =   1000
      Left            =   0
      Picture         =   "Frm_Trafic.frx":15989C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "Frm_Trafic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mise à jour    Le 17/10/2017---
'Mise à jour    Le 04/12/2017---
'===============================
Option Explicit
    Dim Operation()                                 As String
    Dim itmX                                        As ListItem
    Dim VEHICULE                                    As String
    Dim Conducteur                                  As String
    Dim DESTINATION                                 As String
    Dim HrSortie                                    As String
    Public MsgError                                 As String
    Dim Access_code As String
Dim Access_Name As String

Private Sub Form_Activate()
If Frm_MajTrafic.Visible = False Then TXT_PWD.SetFocus

End Sub

Private Sub Form_Load()
    grid_Conducteur.Enabled = False
    Call AfficheExterieur
    Call AfficheDepot
    Call Initgrid_Conducteur
    Call Affiche_Conducteur
    Call Initgrid_Vehicule
    Call Affiche_Vehicule
    Call Initgrid_Destination
    Call Affiche_Destination
    Call grid_Conducteur_ColumnClick(3)
    Call grid_vehicule_ColumnClick(3)
    Call Cmd_Conducteur_Click
    grid_vehicule.Enabled = False
    Grid_Destination.Enabled = False
End Sub
Private Sub Form_Resize()
    Dim WidthForm As Integer, HeightForm As Integer, ScreenWidth As Integer, ScreenHeight As Integer, ScreenResoultion As String
    WidthForm = Me.Width
    ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    ScreenResoultion = ScreenWidth & "x" & ScreenHeight
    WidthForm = Me.Width
    HeightForm = Me.Height
    PicBox_Header.Width = WidthForm
    SCommand1.Left = WidthForm - 2500
    CmdPrint.Left = WidthForm - 4600
    cmd_r.Left = WidthForm - 6730
    CmdSave.Left = WidthForm - 1250
    Pic_confirmation.Left = WidthForm - 4600
    LSV_Exterieur.Height = HeightForm / 2.4
    LSV_Exterieur.Width = WidthForm - 200
    Image3.Width = WidthForm - 200
    Image3.Top = HeightForm + 450
    If ScreenResoultion = "1024x768" Then
        LSV_Exterieur.Height = HeightForm / 2
        grid_Conducteur.Width = 3500
        grid_vehicule.Width = 3500
        grid_vehicule.Left = WidthForm - 11800
        Cmd_Vehicule.Left = WidthForm - 11800
        Grid_Destination.Width = 3500
        Grid_Destination.Left = WidthForm - 8250
        Cmd_Destination.Left = WidthForm - 8250
    End If
End Sub
Public Sub Form_Initialize()
    Dim dat As Date
On Error GoTo Err
    grid_vehicule.Enabled = False
    Grid_Destination.Enabled = False
    Txt_KM.Enabled = False
    CmdSave.Enabled = False
    Lab_Distination.Caption = ""
    Lbl_heure.Caption = Format(Time, "hh:mm:ss")
    Timer1.Enabled = True
    Timer1.Interval = 1000
On Error Resume Next
    Img_alarme.Picture = LoadPicture(App.Path & "\Images\Trafic_V.bmp")
On Error GoTo Err
    dat = Date
    Lbl_date.Caption = UCase(Format(Now, "dddd-dd-mm-yyyy"))
    dat = Time
    Lbl_heure.Caption = dat
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub
'========== Initial ControlBox***
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
        .AddColumn "Matricule", "", , , 220
        .AddColumn "couleur", "couleur", , , False, , , , , , , CCLSortNumeric
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
        .AddColumn "Libelle", "", , , 230
        .AddColumn "couleur", "couleur", , , False, , , , , , , CCLSortNumeric
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Public Sub Initgrid_Destination()
    With Grid_Destination
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
        .AddColumn "LibelleD", "", , , 230
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
'========== GET Details
'========== Affiche Voitures en exterieure***
Public Sub AfficheExterieur()
    Dim LObj_Find                                       As New Traffic
    Dim Lrs_Find                                        As New Recordset
    Dim DateSys                                         As Date
    Dim min                                             As Long
    Dim heur                                            As Long
    Dim Dur                                             As Long
    Dim temp                                            As String
    Dim Name_Tab                                        As String
    Dim i                                               As Integer
On Error GoTo Err
    MouseOn
    DateSys = Date
    LSV_Exterieur.ListItems.Clear
    Name_Tab = "Fichetraffic"
    '========== Voitures en exterieure de plus d'un jours***
    Set LObj_Find = New Traffic
    Set Lrs_Find = LObj_Find.GetAll_TrafficVehiculeExterieur(ErrNumber, ErrDescription, ErrSourceDetail, Name_Tab, DateSys, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        MouseOff
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
            Set itmX = LSV_Exterieur.ListItems.Add(, , "")
            itmX.SubItems(1) = CStr(Lrs_Find.Fields("Numero").Value)
            If Not IsNull(Lrs_Find("abrevVeh")) Then itmX.SubItems(2) = Lrs_Find("abrevVeh") Else itmX.SubItems(2) = Lrs_Find("MatriculeVehic")
            If Not IsNull(Lrs_Find("abrevCond")) Then itmX.SubItems(3) = Lrs_Find("abrevCond") Else itmX.SubItems(3) = Lrs_Find("LibelleCond")
            itmX.SubItems(4) = Lrs_Find("LibelleDest")
            itmX.SubItems(5) = Format(Lrs_Find("heureSortie"), "hh:mm")
            If Not IsNull(Lrs_Find("HeureENtre")) Then itmX.SubItems(6) = Format(Lrs_Find("HeureENtre"), "hh:mm")
            If Not IsNull(Lrs_Find("CompteurSortie")) Then itmX.SubItems(7) = Lrs_Find("CompteurSortie")
            If Not IsNull(Lrs_Find("CompteurEntre")) Then
                itmX.SubItems(8) = Lrs_Find("CompteurEntre")
                itmX.SubItems(9) = Val(Lrs_Find("CompteurEntre")) - Val(Lrs_Find("CompteurSortie"))
                itmX.SubItems(10) = Val(itmX.SubItems(9)) - Val(Lrs_Find("MaxCompteur"))
            End If
            If IsNull(Lrs_Find("HeureENtre")) Then
                Dur = 0
                heur = 0
                min = 0
                temp = ""
                Dur = Abs(DateDiff("n", Lrs_Find("HeureSortie"), Now))
                heur = Dur \ 60
                min = Dur - (heur * 60)
                temp = CStr(heur) & ":" & CStr(min)
                itmX.SubItems(11) = Format(temp, "hh:mm")
                itmX.SubItems(12) = ""
            End If
            itmX.SubItems(13) = Get_AbrevPerso(Lrs_Find("OperateurSortie"))
            If Not IsNull(Lrs_Find("OperateurEntre")) Then itmX.SubItems(14) = Get_AbrevPerso(Lrs_Find("OperateurEntre"))
            If Val(Lrs_Find("MaxCompteur")) < Val(itmX.SubItems(9)) Then
                itmX.ListSubItems(9).ForeColor = vbRed
                itmX.ListSubItems(10).ForeColor = vbRed
            End If
            Lrs_Find.MoveNext
        Wend
    End If
    Set Lrs_Find = Nothing
    '========== Detail des fiches d'aujourdhui***
    Set LObj_Find = New Traffic
    Set Lrs_Find = LObj_Find.GetAll_TrafficByDateSys(ErrNumber, ErrDescription, ErrSourceDetail, Name_Tab, DateSys, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        MouseOff
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
            Dur = 0
            heur = 0
            min = 0
            temp = ""
            Set itmX = LSV_Exterieur.ListItems.Add(, , "")
            itmX.SubItems(1) = CStr(Lrs_Find("Numero"))
            If Not IsNull(Lrs_Find("abrevVeh")) Then itmX.SubItems(2) = Lrs_Find("abrevVeh") Else itmX.SubItems(2) = Lrs_Find("MatriculeVehic")
            If Not IsNull(Lrs_Find("abrevCond")) Then itmX.SubItems(3) = Lrs_Find("abrevCond") Else itmX.SubItems(3) = Lrs_Find("LibelleCond")
            itmX.SubItems(4) = Lrs_Find("libelleDest")
            itmX.SubItems(5) = Format(Lrs_Find("heureSortie"), "hh:mm")
            If Not IsNull(Lrs_Find("HeureENtre")) Then itmX.SubItems(6) = Format(Lrs_Find("HeureENtre"), "hh:mm")
            If Not IsNull(Lrs_Find("CompteurSortie")) Then itmX.SubItems(7) = Lrs_Find("CompteurSortie")
            If Not IsNull(Lrs_Find("CompteurEntre")) Then
                itmX.SubItems(8) = Lrs_Find("CompteurEntre")
                itmX.SubItems(9) = Val(Lrs_Find("CompteurEntre")) - Val(Lrs_Find("CompteurSortie"))
                itmX.SubItems(10) = Val(itmX.SubItems(9)) - Val(Lrs_Find("MaxCompteur"))
            End If
            If Not IsNull(Lrs_Find("HeureENtre")) Then
                '========== Calcule de durée***
                Dur = Abs(DateDiff("n", Lrs_Find("HeureSortie"), Lrs_Find("HeureEntre")))
                heur = Dur \ 60
                min = Dur - (heur * 60)
                temp = CStr(heur) & ":" & CStr(min)
                itmX.SubItems(11) = Format(temp, "hh:mm")
            Else
                '========== Calcule de durée
                Dur = Abs(DateDiff("n", Lrs_Find("HeureSortie"), Now))
                heur = Dur \ 60
                min = Dur - (heur * 60)
                temp = CStr(heur) & ":" & CStr(min)
                itmX.SubItems(11) = Format(temp, "hh:mm")
            End If
            '========== Calcul différence entre durée max et durée du voyage***
            If Dur >= 1440 Then
                Dim XChp() As String, XChps() As String
                Dur = 0
                XChps = Split(Format(Lrs_Find("MaxDuree"), "hh:mm"), ":")
                XChp = Split(Format(itmX.SubItems(11), "hh:mm"), ":")
                heur = Abs(Val(XChps(0)) - Val(XChp(0)))
                min = Abs(Val(XChps(1)) - Val(XChp(1)))
                If min > 59 Then
                    min = (heur * 60) + min
                    heur = min / 60
                    min = min - (heur * 60)
                End If
            Else
                Dur = 0
                Dur = DateDiff("n", Format(Lrs_Find("MaxDuree"), "hh:mm"), Format(itmX.SubItems(11), "hh:mm"))
                heur = Dur \ 60
                min = Dur - (heur * 60)
            End If
            temp = CStr(heur) & ":" & CStr(min)
            itmX.SubItems(12) = Format(temp, "hh:mm")
            itmX.SubItems(13) = Get_AbrevPerso(Lrs_Find("OperateurSortie"))
            If Not IsNull(Lrs_Find("OperateurEntre")) Then itmX.SubItems(14) = Get_AbrevPerso(Lrs_Find("OperateurEntre"))
            If Val(Lrs_Find("MaxCompteur")) < Val(itmX.SubItems(9)) Then
                itmX.ListSubItems(9).ForeColor = vbRed
                itmX.ListSubItems(10).ForeColor = vbRed
            ElseIf Val(Lrs_Find("MinCompteur")) > Val(itmX.SubItems(9)) Then
                itmX.ListSubItems(9).ForeColor = &H8000&
                itmX.ListSubItems(10).ForeColor = &H8000&
            Else
                itmX.SubItems(10) = ""
            End If
            If Val(Lrs_Find("MaxDuree")) < Val(itmX.SubItems(11)) Then
                itmX.ListSubItems(11).ForeColor = &H80&
                itmX.ListSubItems(12).ForeColor = &H80&
            Else
                itmX.ListSubItems(12) = ""
            End If
            Lrs_Find.MoveNext
        Wend
    End If
    Set Lrs_Find = Nothing
    MouseOff
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
    MouseOff
End Sub
'========== Affiche Voitures en depot***
Public Sub AfficheDepot()
    Dim LObj_Find                                   As New VEHICULE
    Dim Lrs_Find                                    As New Recordset
    Dim DateSys                                     As Date
    Dim i                                           As Integer
    Dim J                                           As Integer
On Error GoTo Err
    MouseOn
    DateSys = Date
    Lsv_Depot.ListItems.Clear
    Set Lrs_Find = LObj_Find.GetMatricule_Vehicules(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        MouseOff
        Exit Sub
    End If
    Set LObj_Find = Nothing
    While Not Lrs_Find.EOF
        Set itmX = Lsv_Depot.ListItems.Add(, , "")
        itmX.SubItems(1) = CStr(Lrs_Find("Matricule"))
        Lrs_Find.MoveNext
    Wend
    For J = 1 To LSV_Exterieur.ListItems.Count
        For i = 1 To Lsv_Depot.ListItems.Count - 1
            If (Get_AbrevVeh(Lsv_Depot.ListItems(i).SubItems(1)) = LSV_Exterieur.ListItems(J).SubItems(2)) And (Len(LSV_Exterieur.ListItems(J).SubItems(6)) = 0) Then Lsv_Depot.ListItems.Remove (i)
        Next
    Next
    Set Lrs_Find = Nothing
    MouseOff
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
    MouseOff
End Sub
Public Sub Affiche_Conducteur()
    Dim LObj_Find                               As New Conducteur
    Dim Lrs_Find                                As New Recordset
    Dim i                                       As Integer
    Dim Couleur                                 As String
On Error GoTo Err
    grid_Conducteur.ClearRows
    MouseOn
    Set Lrs_Find = LObj_Find.GetAll_ConducteursActif(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        MouseOff
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        grid_Conducteur.Redraw = False
        While Not Lrs_Find.EOF
            '========== Définir couleur***
            Couleur = "vbGreen"
            For i = 1 To LSV_Exterieur.ListItems.Count
                If (Get_AbrevPerso(Lrs_Find("Libelle")) = LSV_Exterieur.ListItems(i).SubItems(3) And LSV_Exterieur.ListItems(i).SubItems(6) = "" And LSV_Exterieur.ListItems(i).SubItems(4) <> "REPARATION") Then Couleur = "vbRed"
            Next i
            With grid_Conducteur
                .AddRow
                If Couleur = "vbRed" Then
                    .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find("Code")
                    .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find("Libelle"), , , &H8080FF
                    .CellDetails .Rows, .ColumnIndex("couleur"), 2
                Else
                    If (Lrs_Find("disponible") = "O") Then
                        .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find("Code")
                        .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find("Libelle"), , , &HC0FFC0
                        .CellDetails .Rows, .ColumnIndex("couleur"), 3
                    Else
                        .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find("Code")
                        .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find("Libelle"), , , &H80FFFF
                        .CellDetails .Rows, .ColumnIndex("couleur"), 1
                    End If
                End If
            End With
            Lrs_Find.MoveNext
        Wend
        grid_Conducteur.Redraw = True
     End If
    grid_Conducteur.SelectedRow = 1
    Set Lrs_Find = Nothing
    MouseOff
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub Affiche_Vehicule()
    Dim LObj_Find                                           As New VEHICULE
    Dim Lrs_Find                                            As New Recordset
    Dim Couleur                                             As String
On Error GoTo Err
    MouseOn
    grid_vehicule.ClearRows
    Set Lrs_Find = LObj_Find.GetAllActifVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        MouseOff
        Exit Sub
    End If
    Set LObj_Find = Nothing
    ReDim Operation(1)
    If Not Lrs_Find.EOF Then
        grid_vehicule.Redraw = False
        While Not Lrs_Find.EOF
            Couleur = "vbRed"
            Operation = ReturnOperation(Lrs_Find("code"))
            If Operation(0) = "S" Then Couleur = "vbGreen"
            With grid_vehicule
                .AddRow
                If Couleur = "vbRed" Then
                    .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find("Code")
                    .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Find("Matricule"), , , &H8080FF
                    .CellDetails .Rows, .ColumnIndex("couleur"), 2
                Else
                    If (Lrs_Find("disponible") = "O") Then
                        .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find("Code")
                        .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Find("Matricule"), , , &HC0FFC0
                        .CellDetails .Rows, .ColumnIndex("couleur"), 3
                    Else
                        .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find("Code")
                        .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Find("Matricule"), , , &H80FFFF
                        .CellDetails .Rows, .ColumnIndex("couleur"), 1
                    End If
                End If
            End With
            Lrs_Find.MoveNext
        Wend
        grid_vehicule.Redraw = True
     End If
    grid_vehicule.SelectedRow = 1
    Set Lrs_Find = Nothing
    MouseOff
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub Affiche_Destination()
    Dim LObj_Find                                       As New DESTINATION
    Dim Lrs_Find                                        As New Recordset
On Error GoTo Err
    MouseOn
    Grid_Destination.ClearRows
    Set Lrs_Find = LObj_Find.GetAll_DestinationActif(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        MouseOff
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        Grid_Destination.Redraw = False
        While Not Lrs_Find.EOF
            With Grid_Destination
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find("Numero")
                .CellDetails .Rows, .ColumnIndex("LibelleD"), Lrs_Find("Libelle")
            End With
            Lrs_Find.MoveNext
        Wend
        Grid_Destination.Redraw = True
    End If
    Grid_Destination.SelectedRow = 1
    Set Lrs_Find = Nothing
    MouseOff
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'========== Control Box***
'~~~~~~~~~~~~~~~~~~~~~~~
    'SGrid Conducteur~~~
'~~~~~~~~~~~~~~~~~~~~~~~
Public Sub grid_Conducteur_ColumnClick(ByVal lCol As Long)
    Dim sTag As String, i As Long
    With grid_Conducteur.SortObject
        .Clear
        .SortColumn(1) = lCol
        sTag = grid_Conducteur.ColumnTag(lCol)
        If (sTag = "") Then
            sTag = "DESC"
            .SortOrder(1) = CCLOrderAscending
        Else
            sTag = ""
            .SortOrder(1) = CCLOrderAscending
        End If
        grid_Conducteur.ColumnTag(lCol) = sTag
        Select Case grid_Conducteur.ColumnKey(lCol)
            Case "couleur"
                 .SortType(1) = CCLSortNumeric
        End Select
    End With
    Screen.MousePointer = vbHourglass
    grid_Conducteur.Sort
    Screen.MousePointer = vbDefault
End Sub
Private Sub grid_Conducteur_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    Dim CodeCond As String, LObj_Find As New Conducteur, Lrs_Find As New Recordset
On Error GoTo Err
    If (CHECK_ACCES("MAJ_Disp", Access_code) = True) Then
        'Afféctation Var. Conducteur***
        Conducteur = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)
        CodeCond = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, 1)
        Set LObj_Find = New Conducteur
        Set Lrs_Find = LObj_Find.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, Conducteur, CNB)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
        'Si Conducteur en Service maintenant et couleur en Rouge***
        If (grid_Conducteur.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &H8080FF) Then
            MsgBox "Conducteur déjà en service maintenant..." & vbCr & "En Mission...", vbInformation, "Parcano..."
            Exit Sub
        End If
        If (Not Lrs_Find.EOF) Then
            If Lrs_Find("disponible") = "O" Then
                If MsgBox(Conducteur & "    |->    EN-Service   >>>>>>>>   HORS-Service?", vbYesNo + vbDefaultButton2 + vbCritical) = vbYes Then
                    Call SaveDispo(CodeCond, "HS")
                    Set LObj_Find = New Conducteur
                    Call LObj_Find.Update_Desp_Conducteur(ErrNumber, ErrDescription, ErrSourceDetail, "N", Conducteur, CNB)
                    If ErrNumber <> 0 Then
                        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                        ErrNumber = 0
                        Exit Sub
                    End If
                    Set LObj_Find = Nothing
                End If
            Else
                If MsgBox(Conducteur & "    |->    HORS-Service   >>>>>>>>   EN_Service ?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                    Call SaveDispo(CodeCond, "ES")
                    Set LObj_Find = New Conducteur
                    Call LObj_Find.Update_Desp_Conducteur(ErrNumber, ErrDescription, ErrSourceDetail, "O", Conducteur, CNB)
                    If ErrNumber <> 0 Then
                        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                        ErrNumber = 0
                        Exit Sub
                    End If
                    Set LObj_Find = Nothing
                End If
            End If
        End If
        Set Lrs_Find = Nothing
        Call cmd_r_Click
    Else
        MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-être pas des autorisations nécessaires pour modifier ou ré-ajouter un conducteur", vbInformation, "Parcano..."
        Exit Sub
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub grid_Conducteur_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    Dim LObj_Find As New Conducteur, Lrs_Find As New Recordset, LObj_FindV As New VEHICULE
    Dim Pic_Conducteur As String, compteurSt As Long, i As Integer, Pic_Vehicule As String
 On Error GoTo Err
    Conducteur = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)
    '-- Photo Conducteur***
    Set Lrs_Find = LObj_Find.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, Conducteur, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then Pic_Conducteur = Lrs_Find("PicBox")
    Set Lrs_Find = Nothing
    If (grid_Conducteur.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &HC0FFC0) And (grid_vehicule.CellBackColor(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol) = &HC0FFC0) Then
        Grid_Destination.Enabled = True
    Else
        Grid_Destination.Enabled = False
    End If
    If (grid_Conducteur.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &H80FFFF) Then
        grid_vehicule.Enabled = False
        Grid_Destination.Enabled = False
    Else
        grid_vehicule.Enabled = True
    End If
    Lbl_CmptSt.Caption = ""
    Txt_KM.Text = compteurSt
    Lbl_Vehicule.Caption = ""
    Lab_Distination.Caption = ""
    Txt_KM.Enabled = False
    CmdSave.Enabled = False
    If (grid_Conducteur.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &H8080FF) Then
        grid_vehicule.Enabled = False
        Grid_Destination.Enabled = False
        For i = 1 To LSV_Exterieur.ListItems.Count
            If (Get_AbrevPerso(Conducteur) = LSV_Exterieur.ListItems(i).SubItems(3)) And (LSV_Exterieur.ListItems(i).SubItems(6) = "") And (LSV_Exterieur.ListItems(i).SubItems(4) <> "REPARATION") Then
                VEHICULE = Get_LibAbrevVeh(LSV_Exterieur.ListItems(i).SubItems(2))
                compteurSt = Trim(LSV_Exterieur.ListItems(i).SubItems(7))
                Lbl_CmptSt.Caption = Trim(LSV_Exterieur.ListItems(i).SubItems(7))
                DESTINATION = LSV_Exterieur.ListItems(i).SubItems(4)
                HrSortie = LSV_Exterieur.ListItems(i).SubItems(5)
            End If
        Next i
        '-- Code Vehicule***
        Set Lrs_Find = LObj_FindV.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VEHICULE)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_FindV = Nothing
        If Not Lrs_Find.EOF Then Pic_Vehicule = Lrs_Find("PicBox")
        Set Lrs_Find = Nothing
        Txt_KM.Enabled = True
        Lbl_Vehicule.Caption = VEHICULE
        Lab_Distination.Caption = DESTINATION
        Call Txt_KM_GotFocus
        On Error Resume Next
            Img_Vehicule.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Vehicule\" & Pic_Vehicule)
        On Error GoTo Err
            CmdSave.Enabled = True
    End If
On Error Resume Next
    If grid_vehicule.Enabled = True Then
        Img_alarme.Picture = LoadPicture(App.Path & "\Images\Trafic_V.bmp")
    Else
        Img_alarme.Picture = LoadPicture(App.Path & "\Images\Trafic_R.bmp")
    End If
    Img_Conducteur.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Personnel\" & Pic_Conducteur)
    Lbl_Conducteur.Caption = Conducteur
On Error GoTo Err
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
'~~~~~~~~~~~~~~~~~~~~~~
    'SGrid Véhiculer~~~
'~~~~~~~~~~~~~~~~~~~~~~
Public Sub grid_vehicule_ColumnClick(ByVal lCol As Long)
    Dim sTag As String, i As Long
    With grid_vehicule.SortObject
        .Clear
        .SortColumn(1) = lCol
        sTag = grid_vehicule.ColumnTag(lCol)
        If (sTag = "") Then
            sTag = "DESC"
            .SortOrder(1) = CCLOrderAscending
        Else
            sTag = ""
            .SortOrder(1) = CCLOrderAscending
        End If
        grid_vehicule.ColumnTag(lCol) = sTag
        Select Case grid_vehicule.ColumnKey(lCol)
             Case "couleur"
                 .SortType(1) = CCLSortNumeric
        End Select
    End With
    Screen.MousePointer = vbHourglass
    grid_vehicule.Sort
    Screen.MousePointer = vbDefault
End Sub
Private Sub grid_vehicule_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    Call MajDisp
End Sub
Private Sub grid_vehicule_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    Dim LObj_Find As New VEHICULE, Lrs_Find As New Recordset, compteurSt As Long
    Dim i As Integer, Pic_Vehicule As String
On Error GoTo Err
    VEHICULE = grid_vehicule.CellText(grid_vehicule.SelectedRow, 2)
    '-- Code Vehicule***
    Set Lrs_Find = LObj_Find.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VEHICULE)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        Pic_Vehicule = Lrs_Find("PicBox")
        If Not IsNull(Lrs_Find("CompteurFT")) Then compteurSt = Val(Lrs_Find("CompteurFT"))
    End If
    Set Lrs_Find = Nothing

    If (grid_Conducteur.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &HC0FFC0) And (grid_vehicule.CellBackColor(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol) = &HC0FFC0) Then
        Grid_Destination.Enabled = True
    Else
        Grid_Destination.Enabled = False
    End If
    Lbl_CmptSt.Caption = ""
    Txt_KM.Text = compteurSt
    Txt_KM.Enabled = False
    CmdSave.Enabled = False
    If (grid_Conducteur.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &H8080FF) And (grid_vehicule.CellBackColor(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol) = &H8080FF) Then
        For i = 1 To LSV_Exterieur.ListItems.Count
            If (Get_AbrevVeh(VEHICULE) = LSV_Exterieur.ListItems(i).SubItems(2)) Then
                Conducteur = Get_LibAbrevPerso(LSV_Exterieur.ListItems(i).SubItems(3))
                compteurSt = LSV_Exterieur.ListItems(i).SubItems(7)
            End If
        Next i
        Txt_KM.Enabled = True
        Lbl_CmptSt.Caption = compteurSt
        CmdSave.Enabled = True
    End If
    If (grid_vehicule.CellBackColor(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol) = &H8080FF) Then
        For i = 1 To LSV_Exterieur.ListItems.Count
            If (Get_AbrevVeh(VEHICULE) = LSV_Exterieur.ListItems(i).SubItems(2)) And LSV_Exterieur.ListItems(i).SubItems(6) = "" And LSV_Exterieur.ListItems(i).SubItems(4) = "REPARATION" Then
                compteurSt = LSV_Exterieur.ListItems(i).SubItems(7)
                Txt_KM.Enabled = True
                Lbl_CmptSt.Caption = compteurSt
                CmdSave.Enabled = True
            End If
        Next i
    End If
On Error Resume Next
    Img_Vehicule.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Vehicule\" & Pic_Vehicule)
    Lbl_Vehicule.Caption = VEHICULE
On Error GoTo Err
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~
    'SGrid Destination~~~
'~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub grid_destination_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    Dim DESTINATION As String
    DESTINATION = Grid_Destination.CellText(Grid_Destination.SelectedRow, Grid_Destination.SelectedCol)
    CmdSave.Enabled = True
    Lab_Distination.Caption = DESTINATION
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Mise à jour la disponibilite de conducteur & Véhicule~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub SaveDispo(ByVal Conducteur As String, ByVal Operation As String)
    Dim LObj_Find As New Conducteur, Lrs_Find As Recordset
    Dim LInt_NumCompteur As Long, Numero As String
On Error GoTo Err
    'Si Operation = En-Service
    If Operation = "ES" Then
        'Update Ancien Ligne
        Call LObj_Find.Update_DispPers_In_DispoPerso(ErrNumber, ErrDescription, ErrSourceDetail, Conducteur, "Hors-Service", CNB)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
        'Créer Nouvelle Ligne
        LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_DispoPerso")
        If ErrNumber <> 0 Then
           MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
           ErrNumber = 0
           Exit Sub
        End If
        'Insertion enregistrement assiette
        Numero = Format(LInt_NumCompteur, "00000")
       'Insertion enregistrement
        Set Lrs_Find = CreateEmptyRS_DispoPerso()
        With Lrs_Find
            .AddNew
            .Fields("Numero") = Numero
            .Fields("Conducteur") = Conducteur
            .Fields("Etat") = "En-Service"
            .Fields("HDebut") = Format(Now, "dd/mm/yyyy hh:mm:ss")
        End With
        Set LObj_Find = New Conducteur
        Call LObj_Find.Save_DispoPerso(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Find)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
        Set Lrs_Find = Nothing
    'Si Operation = Hors-Service
    ElseIf Operation = "HS" Then
        'Update Ancien Ligne
        Call LObj_Find.Update_DispPers_In_DispoPerso(ErrNumber, ErrDescription, ErrSourceDetail, Conducteur, "En-Service", CNB)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
        'Créer Nouvelle Ligne
        LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_DispoPerso")
        If ErrNumber <> 0 Then
           MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
           ErrNumber = 0
           Exit Sub
        End If
        'Insertion enregistrement assiette
        Numero = Format(LInt_NumCompteur, "00000")
        'Insertion enregistrement
        Set Lrs_Find = CreateEmptyRS_DispoPerso()
        With Lrs_Find
            .AddNew
            .Fields("Numero") = Numero
            .Fields("Conducteur") = Conducteur
            .Fields("Etat") = "Hors-Service"
            .Fields("HDebut") = Format(Now, "dd/mm/yyyy hh:mm:ss")
        End With
        Set LObj_Find = New Conducteur
        Call LObj_Find.Save_DispoPerso(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Find)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
        Set Lrs_Find = Nothing
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub MajDisp()
    Dim LObj_Find As New VEHICULE, Lrs_Find As New Recordset
On Error GoTo Err
    VEHICULE = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
    Set Lrs_Find = LObj_Find.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VEHICULE)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If (Not Lrs_Find.EOF) Then
        If Lrs_Find("disponible") = "O" Then
            If MsgBox("EN-Service -> " & VEHICULE & " -> HORS-Service ?", vbYesNo + vbDefaultButton2 + vbCritical) = vbYes Then
                Set LObj_Find = New VEHICULE
                Call LObj_Find.Update_Disp_Vehicule(ErrNumber, ErrDescription, ErrSourceDetail, "N", VEHICULE, CNB)
                If ErrNumber <> 0 Then
                    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                    ErrNumber = 0
                    Exit Sub
                End If
                Set LObj_Find = Nothing
            End If
        Else
            If MsgBox("HORS-Service -> " & VEHICULE & " -> EN-SERVICE ?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                Call LObj_Find.Update_Disp_Vehicule(ErrNumber, ErrDescription, ErrSourceDetail, "O", VEHICULE, CNB)
                If ErrNumber <> 0 Then
                    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                    ErrNumber = 0
                    Exit Sub
                End If
                Set LObj_Find = Nothing
            End If
        End If
    End If
    Set Lrs_Find = Nothing
    Call cmd_r_Click
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub SaveUpdate()
    Dim Lrs_Find            As New Recordset
    Dim LObj_FindV          As New VEHICULE
    Dim LObj_FindT          As New Traffic
    Dim Code_vehicule       As String
    Dim NumeroFiche         As String
    Dim NumeroTxt           As String
    Dim Operation()         As String

On Error GoTo Err
    '================
    '-- CODE VEHICULE
    '================
    Set Lrs_Find = LObj_FindV.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VEHICULE)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_FindV = Nothing
    If Not Lrs_Find.EOF Then Code_vehicule = Lrs_Find("Code")
    Set Lrs_Find = Nothing
    '===============================
    'GET OPERATION & NUMERO DE FICHE
    '===============================
    NumeroFiche = "Auto"
    ReDim Operation(1)
    Operation = ReturnOperation(Code_vehicule)
    If (Operation(0) = "E") Then
        NumeroFiche = Operation(1)
    End If
    NumeroTxt = Format(CStr(NumeroFiche), "00000")
    
    '========================
    'INSERTION ENREGISTREMENT
    '========================
    Set Lrs_Find = New Recordset
    Set Lrs_Find = CreateEmptyRS_Traffic()
    With Lrs_Find
        .AddNew
        .Fields("HeureEntre") = Format(Now, "dd/mm/yyyy hh:mm:ss")
        .Fields("CompteurEntre") = Txt_KM.Text
        .Fields("OperateurEntre") = Access_Name
        .Fields("Observation") = "Sans observation"
    End With
    Call LObj_FindT.UpDate_Traffic_VE(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Find, NumeroTxt)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_FindT = Nothing
    Set Lrs_Find = Nothing
    Set LObj_FindV = New VEHICULE
    Call LObj_FindV.UpdateCompteurFT_Vehicule(ErrNumber, ErrDescription, ErrSourceDetail, VEHICULE, Txt_KM.Text, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_FindV = Nothing
    MsgBox "Enregistrement terminé avec succé!...", vbInformation, App.ProductName
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~
    'Enregistre Traffic~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub CmdSave_Click()
    Dim Lobj_FindD As New DESTINATION, LObj_FindV As New VEHICULE, LObj_FindC As New Conducteur, LObj_FindT As Traffic, Lrs_Find As New Recordset
    Dim Code_Conducteur As String, Code_Destination As String, Code_vehicule As String
    Dim LInt_NumCompteur As Long, NumeroTxt As String, Kilometrage As Long, NumeroFiche As String
    Dim heure As Date, Compteur As Long, i As Integer, Msg As VbMsgBoxResult
    Dim maxCmpt As Long, min As Long, alert As String
    Dim heur As Long, Dur As Long, temp As String
On Error GoTo Err
    maxCmpt = 0
    If (CHECK_ACCES("Ins_FT", Access_code) = False) Then
        MsgBox "Enregistrement n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-être pas des autorisations nécessaires pour ajouter ou modifier un Traffic", vbExclamation, App.ProductName
    Else
        heure = Now
        Dur = 0
        heur = 0
        min = 0
        temp = ""
        alert = ""
        '====================
        'VEHICULE ENTRE***
        '====================
        If Txt_KM.Enabled = True Then
            If Val(Txt_KM.Text) >= (Val(Lbl_CmptSt.Caption) + 1200) Then
                MsgBox "Compteur invalide!....       " & Val(Txt_KM.Text) & " KM           " & vbCr & "Plus de 1200 Km" & vbCr & "Ancien Compteur est: " & Val(Lbl_CmptSt.Caption) & " KM" & vbCr & vbCr & "Vérifier le compteur saisie", vbCritical, App.ProductName
                Exit Sub
            End If
            If (Txt_KM = "" Or Txt_KM = "0") And (Lbl_Vehicule.Caption <> "M.PEUG.ROUGE" And Lbl_Vehicule.Caption <> "M.PEUG.BLEU" And Lbl_Vehicule.Caption <> "OVETTO DEUX 2") Then
                MsgBox "Entrer le compteur de sortie...", vbExclamation, App.ProductName
                Exit Sub
            End If
            If (Txt_KM <> "" And Txt_KM <> "0" And _
                (Lbl_Vehicule.Caption <> "M.PEUG.ROUGE" Or Lbl_Vehicule.Caption <> "M.PEUG.BLEU" Or Lbl_Vehicule.Caption <> "OVETTO DEUX 2")) Or _
                (Txt_KM <> "" And (Lbl_Vehicule.Caption = "M.PEUG.ROUGE" Or Lbl_Vehicule.Caption = "M.PEUG.BLEU" Or Lbl_Vehicule.Caption = "OVETTO DEUX 2")) Then
                    
                    For i = 1 To LSV_Exterieur.ListItems.Count
                        If LSV_Exterieur.ListItems(i).SubItems(6) = "" Then
                            If (Get_AbrevPerso(Conducteur) = LSV_Exterieur.ListItems(i).SubItems(3)) And (LSV_Exterieur.ListItems(i).SubItems(6) = "") And (LSV_Exterieur.ListItems(i).SubItems(4) <> "REPARATION") Then
                                VEHICULE = Get_LibAbrevVeh(LSV_Exterieur.ListItems(i).SubItems(2))
                                Compteur = LSV_Exterieur.ListItems(i).SubItems(7)
                                DESTINATION = LSV_Exterieur.ListItems(i).SubItems(4)
                            End If
                        End If
                    Next i
                    For i = 1 To LSV_Exterieur.ListItems.Count
                        If (Get_AbrevVeh(VEHICULE) = LSV_Exterieur.ListItems(i).SubItems(2)) And LSV_Exterieur.ListItems(i).SubItems(6) = "" And LSV_Exterieur.ListItems(i).SubItems(4) = "REPARATION" Then
                            Compteur = LSV_Exterieur.ListItems(i).SubItems(7)
                            DESTINATION = LSV_Exterieur.ListItems(i).SubItems(4)
                            VEHICULE = Get_LibAbrevVeh(LSV_Exterieur.ListItems(i).SubItems(2))
                        End If
                    Next i
                    If Val(Txt_KM.Text) < Val(Compteur) Then
                        MsgBox "Compteur invalide..." & vbNewLine & "Vérifier le compteur saisie", vbExclamation, App.ProductName
                        Exit Sub
                    End If
                    If (Lbl_Vehicule.Caption <> "M.PEUG.ROUGE" And Lbl_Vehicule.Caption <> "M.PEUG.BLEU" And Lbl_Vehicule.Caption <> "OVETTO DEUX 2") Then
                        If Val(Txt_KM.Text) = Val(Compteur) And (DESTINATION <> "KIOSQUE AGIL" And DESTINATION <> "KIOSQUE SHELL") Then
                            MsgBox "les compteurs d'entrée et de sortie sont égaux!...", vbExclamation, App.ProductName
                            Txt_KM.SetFocus
                            Exit Sub
                        End If
                    End If
                    '===============================
                    '--VERIFIER COMPTEUR DESTINATION et durée
                    '===============================
                    Set Lrs_Find = Lobj_FindD.GetRow_Destination_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, DESTINATION, CNB)
                    If ErrNumber <> 0 Then
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                        ErrNumber = 0
                        Exit Sub
                    End If
                    Set Lobj_FindD = Nothing
                    If Not Lrs_Find.EOF Then
                        'calcul de la durée du voyage
                        Dur = Abs(DateDiff("n", HrSortie, Format(Now, "hh:mm")))
                        heur = Dur \ 60
                        min = Dur - (heur * 60)
                        temp = CStr(heur) & ":" & CStr(min)
                        temp = Format(temp, "hh:mm")
                        
                        'calcul différence entre durée max et durée du voyage
                        Dur = DateDiff("n", Format(Lrs_Find.Fields("MaxDuree"), "hh:mm"), temp)
                        heur = Dur \ 60
                        min = Dur - (heur * 60)
                        temp = CStr(heur) & ":" & CStr(min)
                        temp = Format(temp, "hh:mm")

                        ' Si dépassement de la distance ou durée max afficher alerte pour saisir observation
                        If (Lbl_Vehicule.Caption <> "M.PEUG.ROUGE" And Lbl_Vehicule.Caption <> "M.PEUG.BLEU" And Lbl_Vehicule.Caption <> "OVETTO DEUX 2") _
                         And (DESTINATION <> "KIOSQUE AGIL" And DESTINATION <> "KIOSQUE SHELL" And DESTINATION <> "REPARATION") _
                         And (Val(Txt_KM.Text) > (Val(Lrs_Find("MaxCompteur")) + Val(Compteur))) Then
                            alert = "Dépassement de la distance Maximum!... " & vbNewLine
                         End If
                         If (Dur > 0) Then
                            alert = alert & "Dépassement de la durée Maximum!... "
                         End If
                         If alert <> "" Then
                            maxCmpt = Val(Lrs_Find("MaxCompteur"))
                                With Frm_ObsTraffic
                                    .ALERTE_ = alert
                                    .VEHICULE_ = VEHICULE
                                    .COMPT_ENTRE_ = Txt_KM.Text
                                    .Lbl_CmptEtr = Txt_KM.Text
                                    .Lbl_CmptSort = Val(Compteur)
                                    .Lbl_DistMax = maxCmpt
                                    .Lbl_Diff = (Val(Txt_KM.Text) - Val(Compteur)) - maxCmpt
                                    .HEURE_ = heure
                                    .Lbl_DrMax.Caption = Lrs_Find.Fields("MaxDuree")
                                    .Lbl_DifDr = Format(temp, "hh:mm")
                                    .Show vbModal
                                End With
                        Else
                            Call SaveUpdate
                        End If
'
                    Else
                        Set Lrs_Find = Nothing
                        MsgBox "Destination Invalide!...", vbExclamation, App.ProductName
                        Exit Sub
                    End If
                    Set Lrs_Find = Nothing
            End If
        ElseIf Txt_KM.Enabled = False Then
        '========================
        'Sélection déstination
        '========================
            If Grid_Destination.SelectionCount = 0 Then
                MsgBox "Sélectionner la destination!...", vbInformation, App.ProductName
                Grid_Destination.SetFocus
                Exit Sub
            End If
            Conducteur = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)
            VEHICULE = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
            DESTINATION = Grid_Destination.CellText(Grid_Destination.SelectedRow, Grid_Destination.SelectedCol)
        '========================
        '-- Code Conducteur***
        '========================
            Set Lrs_Find = LObj_FindC.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, Conducteur, CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            Set LObj_FindC = Nothing
            If Not Lrs_Find.EOF Then Code_Conducteur = Lrs_Find("Code")
            Set Lrs_Find = Nothing
        '========================
        '-- Code Vehicule***
        '========================
            Set Lrs_Find = LObj_FindV.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VEHICULE)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            Set LObj_FindV = Nothing
            If Not Lrs_Find.EOF Then Code_vehicule = Lrs_Find("Code")
            Set Lrs_Find = Nothing
        '========================
        '-- Code Destination***
        '========================
            Set Lrs_Find = Lobj_FindD.GetRow_Destination_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, DESTINATION, CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            Set Lobj_FindD = Nothing
            If Not Lrs_Find.EOF Then Code_Destination = Lrs_Find("Numero")
            Set Lrs_Find = Nothing
            Msg = MsgBox("Confirmez vous l'enregistrement", vbYesNoCancel + vbDefaultButton2 + vbInformation)
            If Msg = vbCancel Then
                Call Affiche_Vehicule
                Call Affiche_Conducteur
                Call Affiche_Destination
                Txt_KM.Text = ""
                grid_Conducteur.SetFocus
                Call grid_vehicule_ColumnClick(2)
                Call grid_Conducteur_ColumnClick(2)
                Exit Sub
            ElseIf Msg = vbNo Then
                Exit Sub
            Else
                Kilometrage = CompteurVehicule(Code_vehicule)
                LInt_NumCompteur = return_Compteur() + 1
                NumeroTxt = Format(LInt_NumCompteur, "00000")
                '========================
                'Insertion enregistrement
                '========================
                Set Lrs_Find = CreateEmptyRS_Traffic()
                With Lrs_Find
                    .AddNew
                    .Fields("Numero") = NumeroTxt
                    .Fields("Vehicule") = Code_vehicule
                    .Fields("CompteurSortie") = Kilometrage
                    .Fields("Conducteur") = Code_Conducteur
                    .Fields("Destination") = Code_Destination
                    .Fields("HeureSortie") = Format(heure, "dd/mm/yyyy hh:mm:ss")
                    .Fields("OperateurSortie") = Access_Name
                    .Fields("userinsert") = Access_code
                End With
                Set LObj_FindT = New Traffic
                Call LObj_FindT.Save_Traffic(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Find)
                If ErrNumber <> 0 Then
                    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                    ErrNumber = 0
                    Exit Sub
                End If
                Set LObj_FindT = Nothing
                Set Lrs_Find = Nothing
                MsgBox "Enregistrement terminé avec succé!...", vbInformation, App.ProductName
            End If
        Else
            MsgBox "Enregistrement invalide, Vérifier les coordonnées", vbExclamation, App.ProductName
            Exit Sub
        End If
        cmd_r_Click
        Cmd_Conducteur_Click
        Lbl_Conducteur = ""
        Lbl_Vehicule = ""
        Lab_Distination = ""
        Img_Conducteur.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Personnel\user.jpg")
        Img_Vehicule.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Vehicule\car.jpg")
        Grid_Destination.Enabled = False
        grid_vehicule.Enabled = False
        CmdSave.Enabled = False
        grid_Conducteur.Enabled = False
        TXT_PWD.Text = ""
        TXT_PWD.Enabled = True
        LBL_Access.Caption = ""
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer2_Timer()
Dim i As Integer, heur As Integer, min As Integer
Dim XChp() As String, XCount As Integer

For i = 1 To LSV_Exterieur.ListItems.Count
    If LSV_Exterieur.ListItems(i).SubItems(6) = "" Then
        XChp = Split(LSV_Exterieur.ListItems(i).SubItems(11), ":")
        heur = XChp(0)
        min = XChp(1) + 1
        If min > 59 Then
            min = (heur * 60) + min
            heur = min / 60
            min = min - (heur * 60)
        End If
        LSV_Exterieur.ListItems(i).SubItems(11) = Format(CStr(heur) & ":" & CStr(min), "hh:mm")
    End If
Next i
End Sub

Private Sub Timer3_Timer()
cmd_r_Click
        Cmd_Conducteur_Click
        Lbl_Conducteur = ""
        Lbl_Vehicule = ""
        Lab_Distination = ""
        Img_Conducteur.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Personnel\user.jpg")
        Img_Vehicule.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Vehicule\car.jpg")
        Grid_Destination.Enabled = False
        grid_vehicule.Enabled = False
        CmdSave.Enabled = False
        grid_Conducteur.Enabled = False
        TXT_PWD.Text = ""
        TXT_PWD.Enabled = True
        LBL_Access.Caption = ""
        If Frm_MajTrafic.Visible = False Then TXT_PWD.SetFocus
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Initialise ControlBox~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Txt_KM_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub Txt_KM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call CmdSave_Click
End Sub
Private Sub SCommand1_Click()
    Frm_Alertt.Show
End Sub
Private Sub CmdPrint_Click()
    On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "Compteurs"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub Timer1_Timer()
     Lbl_date.Caption = UCase(Format(Now, "dddd-dd-mm-yyyy"))
    Lbl_heure = Time
End Sub
Private Sub grid_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyReturn Then
        If (grid_Conducteur.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &H8080FF) Then
            Conducteur = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)
            Txt_KM.SetFocus
        ElseIf (grid_Conducteur.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &HC0FFC0) Then
            Conducteur = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)
            grid_vehicule.SetFocus
            '==================================================================
            VEHICULE = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
            Call grid_vehicule_SelectionChange(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
            '===================================================================
        End If
    End If
End Sub
Private Sub grid_destination_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyReturn Then
        DESTINATION = Grid_Destination.CellText(Grid_Destination.SelectedRow, 2)
        'CmdSave.SetFocus
    End If
End Sub
Private Sub grid_vehicule_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyReturn Then
        If (grid_vehicule.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &H8080FF) Or (grid_vehicule.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &H80FFFF) Then
           Exit Sub
        ElseIf (grid_vehicule.CellBackColor(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol) = &HC0FFC0) Then
            VEHICULE = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
            Grid_Destination.SetFocus
            '==================================================================
            DESTINATION = Grid_Destination.CellText(Grid_Destination.SelectedRow, Grid_Destination.SelectedCol)
            Call grid_destination_SelectionChange(Grid_Destination.SelectedRow, Grid_Destination.SelectedCol)
            '===================================================================
        End If
    End If
End Sub
'~~~~~~~~~~~~~~~~
    'Actualise~~~
'~~~~~~~~~~~~~~~~
Public Sub cmd_r_Click()
    Txt_KM.Enabled = False
    Txt_KM.Text = ""
    Call Affiche_Vehicule
    Call Affiche_Conducteur
    Call Affiche_Destination
    Call AfficheExterieur
    Call AfficheDepot
    Call grid_vehicule_ColumnClick(3)
    Call grid_Conducteur_ColumnClick(3)
End Sub
Private Sub Cmd_Conducteur_Click()
    Call Affiche_Conducteur
    Call grid_Conducteur_ColumnClick(3)
End Sub
Private Sub Cmd_Destination_Click()
    Call Affiche_Destination
End Sub
Private Sub Cmd_Vehicule_Click()
    Call Affiche_Vehicule
    Call grid_vehicule_ColumnClick(3)
End Sub
'~~~~~~~~~~~~~~~~~~~~~~
    'Afficher Detail~~~
'~~~~~~~~~~~~~~~~~~~~~~
Private Sub LSV_Exterieur_DblClick()
    Dim i As Integer
On Error GoTo Err
    Unload Frm_MajTrafic
    Unload Frm_FindView
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_BCarb
    Unload FrmFind_Fils
    With Frm_ControlePwd
        i = LSV_Exterieur.SelectedItem.Index
        .VCode = LSV_Exterieur.ListItems(i).SubItems(1)
        .Sible = "MajTraffic"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'retourne l'abréviation d'un personnel
Private Function Get_AbrevPerso(ByVal personnel As String) As String

Dim LOBJ_Pers As New personnel
Dim rs As New Recordset

Get_AbrevPerso = personnel
Set rs = LOBJ_Pers.Get_Abrev(ErrNumber, ErrDescription, ErrSourceDetail, CNB, personnel)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
Set LOBJ_Pers = Nothing
If Not rs.EOF Then
    If Not IsNull(rs(0)) Then Get_AbrevPerso = rs(0)
End If
End Function
'retourne nom du personnel par son abréviation
Private Function Get_LibAbrevPerso(ByVal personnel As String) As String

Dim LOBJ_Pers As New personnel
Dim rs As New Recordset

Get_LibAbrevPerso = personnel
Set rs = LOBJ_Pers.Get_LibAbrev(ErrNumber, ErrDescription, ErrSourceDetail, CNB, personnel)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
Set LOBJ_Pers = Nothing
If Not rs.EOF Then
    Get_LibAbrevPerso = rs(0)
End If
End Function
'Retourne l'abréviation d'un véhicule
Private Function Get_AbrevVeh(ByVal Vehic As String) As String

Dim LOBJ_Veh As New VEHICULE
Dim rs As New Recordset

Get_AbrevVeh = Vehic
Set rs = LOBJ_Veh.Get_Abrev(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Vehic)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
Set LOBJ_Veh = Nothing
If Not rs.EOF Then
    If Not IsNull(rs(0)) Then Get_AbrevVeh = rs(0)
End If
End Function

Private Function Get_LibAbrevVeh(ByVal Vehic As String) As String

Dim LOBJ_Veh As New VEHICULE
Dim rs As New Recordset

Get_LibAbrevVeh = Vehic
Set rs = LOBJ_Veh.Get_MatricAbrev(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Vehic)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
Set LOBJ_Veh = Nothing
If Not rs.EOF Then
    Get_LibAbrevVeh = rs(0)
End If
End Function




'~~~~~~~~~~~~~~~~~~~~~~~
    'TextBox Compteur~~~
'~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Txt_KM_LostFocus()
    Dim Compteur As Long, i As Integer
On Error GoTo Err
    'Verification de Compteur
    If grid_Conducteur.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &H8080FF Then
        Conducteur = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)
        If Txt_KM <> "" Then
            For i = 1 To LSV_Exterieur.ListItems.Count
                If (Get_AbrevPerso(Conducteur) = LSV_Exterieur.ListItems(i).SubItems(3)) And (LSV_Exterieur.ListItems(i).SubItems(6) = "") And (LSV_Exterieur.ListItems(i).SubItems(4) <> "REPARATION") Then
                    Compteur = LSV_Exterieur.ListItems(i).SubItems(7)
                End If
            Next i
            For i = 1 To LSV_Exterieur.ListItems.Count
                If (Get_AbrevVeh(VEHICULE) = LSV_Exterieur.ListItems(i).SubItems(2)) And LSV_Exterieur.ListItems(i).SubItems(6) = "" And LSV_Exterieur.ListItems(i).SubItems(4) = "REPARATION" Then
                    Compteur = LSV_Exterieur.ListItems(i).SubItems(7)
                End If
            Next i
            If Val(Txt_KM.Text) - Val(Compteur) > 1200 Then
                MsgBox ("Nouveau compteur invalid : Plus que 1200 klm")
                Txt_KM.SetFocus
            Exit Sub
            End If
        End If
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub Txt_KM_GotFocus()
    Dim i As Integer, Compteur As Long
On Error GoTo Err
    If grid_Conducteur.CellBackColor(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol) = &H8080FF Or grid_vehicule.CellBackColor(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol) = &H8080FF Then
        For i = 1 To LSV_Exterieur.ListItems.Count
            If (Get_AbrevPerso(Conducteur) = LSV_Exterieur.ListItems(i).SubItems(3)) And (LSV_Exterieur.ListItems(i).SubItems(6) = "") And (LSV_Exterieur.ListItems(i).SubItems(4) <> "REPARATION") Then
                Compteur = LSV_Exterieur.ListItems(i).SubItems(7)
            End If
        Next i
        For i = 1 To LSV_Exterieur.ListItems.Count
            If (Get_AbrevVeh(VEHICULE) = LSV_Exterieur.ListItems(i).SubItems(2)) And LSV_Exterieur.ListItems(i).SubItems(6) = "" And LSV_Exterieur.ListItems(i).SubItems(4) = "REPARATION" Then
                Compteur = LSV_Exterieur.ListItems(i).SubItems(7)
            End If
        Next i
        
        If Len(CStr(Compteur)) = 6 Then
            Txt_KM.Text = Left(CStr(Compteur), 2)
        ElseIf Len(CStr(Compteur)) = 5 Then
            Txt_KM.Text = Left(CStr(Compteur), 1)
        Else
            Txt_KM.Text = 0
        End If
    End If
    Txt_KM.SelStart = Len(Txt_KM.Text)
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub

Public Sub TXT_PWD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then


Dim i As Integer
'Dim Access_code As Integer
'Dim Access_Name As String

'Connexion à la base
Dim strConnect As String
Dim w, k
Dim SQL As String
Dim rs As New ADODB.Recordset




    'Verifier si l'utilisateur est actif
    w = UCase(TXT_PWD.Text)
    k = ""
    For i = 1 To Len(w)
        k = k & Asc(Mid(w, i, 1))
    Next
    
    rs.CursorLocation = adUseClient
    SQL = "select * from Utilisateur where MP like " & SQLText(UCase(k))
    rs.Open SQL, CNB, adOpenKeyset, adLockPessimistic
    If Not rs.EOF Then
        
        Access_code = rs.Fields("Code")
        Access_Name = rs.Fields("NomPrn")
        
        
        
        If (CHECK_ACCES("Consult_FT", Access_code) = True) Then
        grid_Conducteur.Enabled = True
        LBL_Access.Caption = Access_Name
        TXT_PWD.Enabled = False
        
         Timer3.Enabled = True
        Timer3.Interval = 30000
       
        
        Else: Exit Sub
       
        End If
    Else
        MsgBox "Mot passe incorrect .?  ", vbExclamation
        With TXT_PWD
            .SetFocus
            .Text = ""
            .SelStart = 0
            .SelLength = Len(TXT_PWD.Text)
        End With
    End If

    Set rs = Nothing

End If

End Sub

