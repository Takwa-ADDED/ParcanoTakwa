VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultPR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5775
      TabIndex        =   28
      Top             =   3480
      Width           =   5775
      Begin SToolBox.SDateBox cda_Create 
         Height          =   285
         Left            =   4560
         TabIndex        =   29
         Tag             =   "M"
         Top             =   0
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Text            =   ""
         Enabled         =   0   'False
         BackColor       =   14737632
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Piece :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   3240
         TabIndex        =   31
         Top             =   0
         Width           =   1230
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date Operation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3615
      Left            =   11640
      ScaleHeight     =   3615
      ScaleWidth      =   615
      TabIndex        =   15
      Top             =   6480
      Width           =   615
      Begin SToolBox.SCommand SCommand2 
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         Enabled         =   0   'False
         BackStyle       =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FrmConsultPR.frx":0000
      End
      Begin SToolBox.SCommand SCommand3 
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         Enabled         =   0   'False
         BackStyle       =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FrmConsultPR.frx":0182
      End
      Begin SToolBox.SCommand SCommand4 
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         Enabled         =   0   'False
         BackStyle       =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FrmConsultPR.frx":04D5
      End
   End
   Begin VB.TextBox txt_Numero 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   465
      Left            =   1680
      TabIndex        =   14
      Tag             =   "M"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.ComboBox cbo_typePiece 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmConsultPR.frx":0657
      Left            =   1680
      List            =   "FrmConsultPR.frx":0661
      TabIndex        =   13
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Tex_RSP 
      Enabled         =   0   'False
      Height          =   495
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "00,00"
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txt_ref 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CheckBox chk_paye 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   10
      Top             =   5520
      Width           =   255
   End
   Begin VB.TextBox txt_BCReparation 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txt_Timbre 
      Enabled         =   0   'False
      Height          =   495
      Left            =   8640
      TabIndex        =   8
      Text            =   "00,400"
      Top             =   1920
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1095
      ScaleWidth      =   4695
      TabIndex        =   1
      Top             =   5160
      Width           =   4695
      Begin VB.TextBox txt_ville 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txt_adresse 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txt_rsocial 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ville :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   975
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   645
         TabIndex        =   6
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Raison sociale :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   0
         Width           =   1380
      End
   End
   Begin VB.TextBox txt_MatriculeStation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      Tag             =   "M"
      Top             =   4680
      Width           =   2895
   End
   Begin MSComctlLib.ListView Lsv_Detail 
      Height          =   3735
      Left            =   0
      TabIndex        =   19
      Top             =   6480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6588
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
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Désignation"
         Object.Width           =   3617
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qte"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Véhicule"
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "PU.HT"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Remise (%)"
         Object.Width           =   3177
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tot.HT"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "TVA (%)"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Prix.TTC"
         Object.Width           =   1766
      EndProperty
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   10680
      TabIndex        =   20
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      Enabled         =   0   'False
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmConsultPR.frx":067D
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   9960
      TabIndex        =   21
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      Enabled         =   0   'False
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmConsultPR.frx":07FF
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   10320
      TabIndex        =   22
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      Enabled         =   0   'False
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmConsultPR.frx":0B52
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   9600
      TabIndex        =   23
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      Enabled         =   0   'False
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmConsultPR.frx":0EA5
   End
   Begin SToolBox.SCommand CmdPrint 
      Height          =   495
      Left            =   11160
      TabIndex        =   24
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      Enabled         =   0   'False
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmConsultPR.frx":1027
   End
   Begin MSComctlLib.ListView Lsv_Toto 
      Height          =   1455
      Left            =   6000
      TabIndex        =   25
      Top             =   3840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2566
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tot.Brut"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tot.RL"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "TOT.RP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tot.HT"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tot.TVA"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tot.TTC"
         Object.Width           =   1853
      EndProperty
   End
   Begin SToolBox.SCommand SCommand1 
      Height          =   375
      Left            =   4800
      TabIndex        =   26
      Top             =   2400
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Enabled         =   0   'False
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmConsultPR.frx":137A
      ButtonType      =   1
   End
   Begin SToolBox.SDateBox cda_Operation 
      Height          =   285
      Left            =   1680
      TabIndex        =   27
      Tag             =   "M"
      Top             =   3480
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
   End
   Begin SToolBox.SCommand SCommand5 
      Height          =   375
      Left            =   4800
      TabIndex        =   32
      Top             =   2880
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Enabled         =   0   'False
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmConsultPR.frx":16CD
      ButtonType      =   1
   End
   Begin SToolBox.SCommand CmdFindStation 
      Height          =   375
      Left            =   4560
      TabIndex        =   33
      Top             =   4740
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Enabled         =   0   'False
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmConsultPR.frx":1A20
      ButtonType      =   1
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmConsultPR.frx":1D73
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro pièce"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   0
      TabIndex        =   43
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pièce de reception"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   3840
      TabIndex        =   42
      Top             =   240
      Width           =   2625
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Type de pièce"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   0
      TabIndex        =   41
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Remise sur pièce (%) :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   6000
      TabIndex        =   40
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAUX EN (DT)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   7080
      TabIndex        =   39
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Référence"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   0
      TabIndex        =   38
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Payé: O/N"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   5880
      TabIndex        =   37
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "BC Reparation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Timbre fiscal :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   6000
      TabIndex        =   35
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Fournisseur 
      BackStyle       =   0  'Transparent
      Caption         =   "Station "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   0
      TabIndex        =   34
      Top             =   4680
      Width           =   735
   End
End
Attribute VB_Name = "FrmConsultPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Okayy As Boolean

Dim itmX As ListItem

Dim thekey As Integer
Dim theshift As Integer
Private Sub cbo_typePiece_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub




Private Sub CmdAdd_Click()

On Error GoTo Err

Dim rQ As New ADODB.Recordset
'Dim SQL As String
'SQL = "Select * from utilisateur where INS_BC = 1 and code= " & LInt_UserId
'rQ.Open SQL, CNB, adOpenDynamic
'If rQ.EOF Then
'    rQ.Close
'    MsgBox "Accès refusé.", vbExclamation
'    Exit Sub
'End If


Okayy = False
Picture2.Enabled = True


If Lsv_Detail.ListItems.Count > 0 And txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Call ViderZone(FrmPieceReparation)
Tex_RSP.Text = "00,00"
txt_Timbre.Text = "00,400"
Tex_RSP.Text = Format(Tex_RSP.Text, "#0.00")

txt_Numero.Text = "Auto"
cda_Create.Text = Date
cda_Operation.Text = Date
cbo_typePiece.SetFocus

Lsv_Detail.ListItems.Clear
Lsv_Toto.ListItems.Clear

'Crémentation
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim vcode As String

On Error GoTo Err

'    If PIC_NFACT.Visible = True Then
'    MsgBox "Maj impossible", vbInformation
'    Exit Sub
'    End If

    If txt_Numero.Text = "Auto" Then
        If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
        Else
        txt_Numero.SetFocus
        Exit Sub
        End If
    End If
    
'    If txt_Numero.Text <> "Auto" Then
'        Dim rQ As New ADODB.Recordset
'        SQL = "Select * from utilisateur where MAJ_BC = 1 and code= " & LInt_UserId
'        rQ.Open SQL, CNB, adOpenDynamic
'        If rQ.EOF Then
'            rQ.Close
'            MsgBox "Accès refusé.", vbExclamation
'            Exit Sub
'        End If
'    End If
    
    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & " pièce de réparation ", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    vcode = txt_Numero.Text
    SQL = "Delete from AssPieceReparation where Numero =" & SQLText(vcode)
    CNB.Execute SQL
    SQL = "Delete from DetailPieceReparation where Numero =" & SQLText(vcode)
    CNB.Execute SQL
    Call ViderZone(FrmPieceReparation)
    Lsv_Detail.ListItems.Clear
    Lsv_Toto.ListItems.Clear
    txt_Numero.SetFocus
    End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub


Private Sub CmdFind_Click()
On Error GoTo Err
If Lsv_Detail.ListItems.Count > 0 And txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

If Okayy = True Then
    If MsgBox("Annuler la maj en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
End If
If cbo_typePiece.Text = "" Then
    MsgBox ("Chasissez un type de piece")
    cbo_typePiece.SetFocus
    Exit Sub
    End If
    
 If cbo_typePiece.Text = "Bon Livraison" Then
Unload FrmFind
With FrmFind
    .StrSource = "BLPieceReparation"
    .Show
End With
End If
 If cbo_typePiece.Text = "Facture" Then
Unload FrmFind
With FrmFind
    .StrSource = "FacturePieceReparation"
    .Show
End With
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub


Private Sub cmdFindNumero_Click()
On Error Resume Next
Unload FrmFind
If cbo_typePiece.Text = "BL" Then
With FrmFind
    .StrSource = "BL Reparation"
    .Show
End With
Else
    With FrmFind
    .StrSource = "Facture Reparation"
    .Show
End With
End If
End Sub



Private Sub CmdPrint_Click()
'On Error GoTo Err
If txt_Numero.Text = "" Then Exit Sub
If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
    Exit Sub
    Else
    txt_Numero.SetFocus
    Exit Sub
    End If
End If
If txt_Numero.Text = "Auto" Then Exit Sub
If MsgBox("Imprimer ce bon        ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Dim F As Form
    Set F = New Frm_Rpt_Apercus
    With F
        .numero = txt_Numero.Text
        Call .PrintOutAndApercu_PieceReparation(0)
        .Show
    End With
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdSave_Click()

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim ii As Long
Dim LInt_NumCompteur As Long
Dim energie As String
Dim vcode


If Left(CheckMandatory(FrmPieceReparation), 1) = 1 Then
   Exit Sub
End If

If Lsv_Detail.ListItems.Count = 0 Then
    MsgBox "Veuillez saisir details ", vbInformation
    Exit Sub
End If

  For ii = 1 To Lsv_Detail.ListItems.Count
    If (Lsv_Detail.ListItems(ii).SubItems(8) = "00,00") Then
       MsgBox "Veuillez saisir details ", vbInformation
    Exit Sub
    Exit For
    End If
 Next
 
On Error GoTo Err

    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    
    'Delete l'enregistrement
    vcode = txt_Numero.Text
    CNB.BeginTrans
    SQL = "Delete from AssPieceReparation where Numero =" & SQLText(vcode)
    CNB.Execute SQL
    SQL = "Delete from DetailPieceReparation where Numero =" & SQLText(vcode)
    CNB.Execute SQL
    
    If vcode = "Auto" Then
    LInt_NumCompteur = Return_Compteur() + 1
    
    'Insertion enregistrement assiette
    txt_Numero.Text = Format(LInt_NumCompteur, "00000")
    End If
    
    
    'Insertion enregistrement details
    For ii = 1 To Lsv_Detail.ListItems.Count
    SQL = "Insert into DetailPieceReparation  (Numero, Designation,Qte,Vehicule,PUHT,Remise,TVA) values ("

    SQL = SQL & SQLText(txt_Numero.Text)
    SQL = SQL & "," & SQLText(Lsv_Detail.ListItems(ii).SubItems(1))
    SQL = SQL & "," & Val(Lsv_Detail.ListItems(ii).SubItems(2))
    SQL = SQL & "," & SQLText(Lsv_Detail.ListItems(ii).SubItems(3))
    SQL = SQL & "," & Replace((Lsv_Detail.ListItems(ii).SubItems(4)), ",", ".")
    SQL = SQL & "," & Replace(Lsv_Detail.ListItems(ii).SubItems(5), ",", ".")
    SQL = SQL & "," & Replace(Lsv_Detail.ListItems(ii).SubItems(7), ",", ".")
    SQL = SQL & ")"
    CNB.Execute SQL
    Lsv_Detail.ListItems(ii).Text = txt_Numero.Text
    Next
    
    'insertion assiette
    SQL = "Insert into AssPieceReparation  (Numero, Type,refPiece,DatePiece,DateOperation, Fournisseur,RemisePiece, TotTTC, Payement,Timbre) values ("
    SQL = SQL & SQLText(txt_Numero.Text)
    SQL = SQL & "," & SQLText(cbo_typePiece.Text)
    SQL = SQL & "," & SQLText(txt_ref.Text)
    SQL = SQL & "," & SQLText(cda_Create.Text)
    SQL = SQL & "," & SQLText(cda_Operation.Text)
    SQL = SQL & "," & SQLText(txt_MatriculeStation.Text)
    SQL = SQL & "," & Replace(Tex_RSP.Text, ",", ".")
    SQL = SQL & "," & Replace(Lsv_Toto.ListItems(1).SubItems(5), ",", ".")
    SQL = SQL & "," & chk_paye.Value
    SQL = SQL & "," & Replace(txt_Timbre.Text, ",", ".")
    SQL = SQL & ")"
    CNB.Execute SQL

    CNB.CommitTrans
    Okayy = False
    
     If MsgBox("Enregistrement terminé avec succé  " & vbNewLine & "Imprimer ce bon        ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Dim F As Form
    Set F = New Frm_Rpt_Apercus
    With F
        .numero = txt_Numero.Text
        Call .PrintOutAndApercu_PieceReparation(0)
        .Show
    End With
    End If
    End If
'    txt_Numero.SetFocus
    Call ViderZone(FrmPieceReparation)
    Lsv_Detail.ListItems.Clear
    Lsv_Toto.ListItems.Clear
Exit Sub
Err:
CNB.RollbackTrans
MsgBox Err.Description, vbInformation
End Sub
Private Sub Command1_Click()
MsgBox Format(5, "000000")
End Sub


Private Sub Form_Load()
On Error GoTo Err
cda_Create.Text = Date
cda_Operation.Text = Date
Me.Move 0, 0
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub




Private Sub Form_Unload(Cancel As Integer)
On Error GoTo erreur
   Dim i As Integer
   Dim MSG ' Déclare la variable.
   ' Définit le texte du message.
   MSG = "Voulez-vous vraiment quitter?"
   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
   If MsgBox(MSG, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
      Cancel = True
   Else
   Unload Me
   End If
   
   Exit Sub
erreur:
   MsgBox Err.Description, 48

End Sub

Private Sub Lsv_Detail_DblClick()
Dim i
Dim ii

On Error GoTo Err

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If
With FrmSaisiePieceReparation
    .Okay = False
    .ii = Lsv_Detail.SelectedItem.Index
    i = Lsv_Detail.SelectedItem.Index
    .txt_Numero.Text = txt_Numero.Text
    .Txt_Designation.Text = Lsv_Detail.ListItems(i).SubItems(1)
    .cbo_Matricule.Text = Lsv_Detail.ListItems(i).SubItems(3)
    .txt_Qte.Text = Lsv_Detail.ListItems(i).SubItems(2)
    .txt_PUHT.Text = Lsv_Detail.ListItems(i).SubItems(4)
    .Txt_Remise.Text = Lsv_Detail.ListItems(i).SubItems(5)
    .txt_TotHT.Text = Lsv_Detail.ListItems(i).SubItems(6)
    .Txt_tva.Text = Lsv_Detail.ListItems(i).SubItems(7)
    .txt_ttc.Text = Lsv_Detail.ListItems(i).SubItems(8)
    .Show
    
End With

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub




Private Sub SCommand1_Click()
On Error GoTo Err
If Lsv_Detail.ListItems.Count > 0 And txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

If Okayy = True Then
    If MsgBox("Annuler la maj en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
End If
If cbo_typePiece.Text = "" Then
    MsgBox ("Chasissez un type de piece")
    cbo_typePiece.SetFocus
    Exit Sub
    End If
    
 If cbo_typePiece.Text = "Bon Livraison" Then
Unload FrmFind
With FrmFind
    .StrSource = "BLPieceReparation"
    .Show
End With
End If
 If cbo_typePiece.Text = "Facture" Then
Unload FrmFind
With FrmFind
    .StrSource = "FacturePieceReparation"
    .Show
End With
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation



End Sub

Private Sub SCommand2_Click()

Dim i As Integer
On Error GoTo Err
If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If

If Lsv_Detail.ListItems.Count <= 0 Then Exit Sub
Okayy = True
If MsgBox("Confirmez vous la suppression de la ligne en cour.?", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
i = Lsv_Detail.SelectedItem.Index
Lsv_Detail.ListItems.Remove i
Call AppCalcul
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub SCommand3_Click()
Dim i

On Error GoTo Err

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If

If Lsv_Detail.ListItems.Count <= 0 Then Exit Sub
Okayy = True
With FrmSaisiePieceReparation
    .Okay = False
    .ii = Lsv_Detail.SelectedItem.Index
    i = Lsv_Detail.SelectedItem.Index
    .txt_Numero.Text = txt_Numero.Text
    .Txt_Designation.Text = Lsv_Detail.ListItems(i).SubItems(1)
    .cbo_Matricule.Text = Lsv_Detail.ListItems(i).SubItems(3)
    .txt_Qte.Text = Lsv_Detail.ListItems(i).SubItems(2)
    .txt_PUHT.Text = Lsv_Detail.ListItems(i).SubItems(4)
    .Txt_Remise.Text = Lsv_Detail.ListItems(i).SubItems(5)
    .txt_TotHT.Text = Lsv_Detail.ListItems(i).SubItems(6)
    .Txt_tva.Text = Lsv_Detail.ListItems(i).SubItems(7)
    .txt_ttc.Text = Lsv_Detail.ListItems(i).SubItems(8)
    .Show
End With
Err:
Exit Sub
MsgBox Err.Description, vbInformation
End Sub

Private Sub SCommand4_Click()
 If txt_Numero.Text = "" Then
        If Len(Trim(txt_Numero.Text)) = 0 Then
            MsgBox "N° bon obligatoire      ", vbInformation
            txt_Numero.SetFocus
            Exit Sub
        End If
        End If
        
  If txt_MatriculeStation.Text = "" Then
        If Len(Trim(txt_MatriculeStation.Text)) = 0 Then
            MsgBox "Station obligatoire      ", vbInformation
          
            Exit Sub
        End If
      End If
    
On Error GoTo Err
     Okayy = True
    If txt_Numero.Text = "" Then
        If Len(Trim(txt_Numero.Text)) = 0 Then
            MsgBox "N° bon obligatoire      ", vbInformation
            txt_Numero.SetFocus
            Exit Sub
        End If
      
  
        
        With FrmSaisiePieceReparation
            .txt_Numero.Text = Me.txt_Numero.Text
            .Okay = True
            .Show
        End With
    Else
        With FrmSaisiePieceReparation
            .txt_Numero.Text = Me.txt_Numero.Text
            .Okay = True
            .Show
        End With
    End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub






Private Sub SCommand5_Click()

On Error GoTo Err

If cbo_typePiece.Text = "" Then
    MsgBox ("Chasissez un type de piece")
    cbo_typePiece.SetFocus
    Exit Sub
    End If
    
 If cbo_typePiece.Text = "Bon Livraison" Then
Unload FrmFind
With FrmFind
    .StrSource = "FIndBCReparation"
    .Show
End With

End If
 If cbo_typePiece.Text = "Facture" Then
    MsgBox ("Tu peut pas tranferer une facture")
    cbo_typePiece.SetFocus
    Exit Sub
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub Tex_RSP_GotFocus()
If Lsv_Detail.ListItems.Count = 0 Then
    MsgBox "Veuillez saisir details avant ", vbInformation
    Else
    Tex_RSP.Locked = False
End If
End Sub


Private Sub Tex_RSP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call AppCalcul
SendKeys "{tab}"
End If
End Sub


Private Sub Tex_RSP_LostFocus()
Call AppCalcul
Tex_RSP.Text = Format(Tex_RSP.Text, "#0.00")
End Sub

Private Sub txt_BCReparation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_MatriculeStation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Numero_GotFocus()

On Error GoTo Err

    Call ViderZone(FrmPieceReparation)
    
    Lsv_Detail.ListItems.Clear
    
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub txt_Numero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub txt_Numero_LostFocus()

On Error GoTo Err

If Len(Trim(txt_Numero.Text)) > 0 Then
Call AfficheRow(txt_Numero.Text)
End If

If Len(Trim(txt_Numero.Text)) > 0 Then
txt_MatriculeStation.Enabled = True
CmdFindStation.Enabled = True
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub
Public Sub AfficheRow(ByVal vcode As String)

Dim TotHTBrut As Double
Dim TotTTC As Double
Dim Fcode As String
Dim Qte As Double
Dim PUHT As Double
Dim Remise As Double
Dim tva As Double

Dim SQL As String
Dim rs As New ADODB.Recordset

Call ViderZone(FrmPieceReparation)
Lsv_Detail.ListItems.Clear
Lsv_Toto.ListItems.Clear

'Assiette
SQL = "Select * from AssPieceReparation where Numero = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    Fcode = rs("Fournisseur")
    txt_Numero.Text = rs("Numero")
    If Not IsNull(rs("Type")) Then cbo_typePiece.Text = rs("Type")
    If Not IsNull(rs("refPiece")) Then txt_ref.Text = rs("refPiece")
    If Not IsNull(rs("DatePiece")) Then cda_Create.Text = rs("DatePiece")
    If Not IsNull(rs("DateOperation")) Then cda_Operation.Text = rs("DateOperation")
    If Not IsNull(rs("Fournisseur")) Then Call AfficheRow_Station(rs("Fournisseur"))
    If Not IsNull(rs("RemisePiece")) Then Tex_RSP.Text = rs("RemisePiece")
    If Not IsNull(rs("timbre")) Then txt_Timbre.Text = rs("Timbre")
    If Not IsNull(rs("Payement")) Then chk_paye = rs("Payement")
   Else
    MsgBox "Numéro bon introuvable", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If

rs.Close


'Detail
SQL = " Select * from DetailPieceReparation WHERE   Numero = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then

    While Not rs.EOF
    
    TotTTC = 0
    TotHTBrut = 0
    Qte = 0
    PUHT = 0
    Remise = 0
    tva = 0
    
    Qte = rs("Qte")
    PUHT = rs("PUHT")
    Remise = rs("Remise")
    tva = rs("tva")
    
    TotHTBrut = FrmSaisiePieceReparation.Return_TotHT(Qte, PUHT, Remise)
    TotTTC = TotHTBrut + (TotHTBrut * (tva / 100))
            
            Set itmX = Lsv_Detail.ListItems.Add(, , CStr(txt_Numero.Text))
            itmX.SubItems(1) = rs("Designation")
            itmX.SubItems(2) = rs("Qte")
            itmX.SubItems(3) = rs("Vehicule")
            itmX.SubItems(4) = Format(rs("PUHT"), "#,##0.000")
            itmX.SubItems(5) = Format(rs("Remise"), "#0.00")
            itmX.SubItems(6) = Format(TotHTBrut, "#,##0.000")
            itmX.SubItems(7) = Format(rs("tva"), "#0.00")
            itmX.SubItems(8) = Format(TotTTC, "#,##0.000")
            
            
        rs.MoveNext
    Wend
    Call AppCalcul
End If



rs.Close

End Sub
Public Sub AppCalcul()
Dim i
Dim ii As Integer
Dim Tot_brut As Double
Dim Tot_remise As Double
Dim Tot_ht As Double
Dim Tot_tva As Double
Dim tot_ttc As Double
Dim val_RP As Double
Dim Timbre As Double

'Totaux
Tot_brut = 0
Tot_remise = 0
Tot_ht = 0
Tot_tva = 0
tot_ttc = 0
val_RP = 0
Timbre = 0

'Lignes
Dim brut As Double
Dim Remise As Double
Dim ht As Double
Dim tva As Double
Dim ttc As Double

Dim pu As Double
Dim Qte As Integer

Lsv_Toto.ListItems.Clear

For ii = 1 To Lsv_Detail.ListItems.Count
    brut = 0
    Remise = 0
    ht = 0
    tva = 0
    ttc = 0
    
    pu = 0
    Qte = 0
    
    
    'totale Brut
    
    Qte = Lsv_Detail.ListItems(ii).SubItems(2)
    pu = Lsv_Detail.ListItems(ii).SubItems(4)
    brut = Qte * pu
    Tot_brut = Tot_brut + brut

    'totale ht
    ht = Lsv_Detail.ListItems(ii).SubItems(6)
    Tot_ht = Tot_ht + ht
    
    'Totale remise
    Remise = brut - ht
    Tot_remise = Tot_remise + Remise
    
    'Totale ttc
    ttc = Lsv_Detail.ListItems(ii).SubItems(8)
    tot_ttc = tot_ttc + ttc
    
    'totale tva
    
    tva = Lsv_Detail.ListItems(ii).SubItems(7)
    
    Tot_tva = Tot_tva + ((ht * tva) / 100)
    

    
    
Next
tot_ttc = tot_ttc - (tot_ttc * CDbl(Replace(Tex_RSP.Text, ".", ",")) / 100) + CDbl(Replace(txt_Timbre.Text, ".", ","))
val_RP = tot_ttc * Val(Tex_RSP.Text) / 100


Set itmX = Lsv_Toto.ListItems.Add(, , Format(Tot_brut, "#,##0.000"))
        itmX.SubItems(1) = Format(Tot_remise, "#,##0.000")
        itmX.SubItems(2) = Format(val_RP, "#,##0.000")
        itmX.SubItems(3) = Format(Tot_ht, "#,##0.000")
        itmX.SubItems(4) = Format(Tot_tva, "#,##0.000")
        itmX.SubItems(5) = Format(tot_ttc, "#,##0.000")
 

End Sub


Private Function Return_Compteur() As Long
Dim rD As New ADODB.Recordset
Dim SQL As String
Return_Compteur = 0
SQL = "select Max(Numero) from AssPieceReparation "
rD.Open SQL, CNB, adOpenKeyset
If Not rD.EOF Then
Return_Compteur = rD(0)
End If
rD.Close
End Function

Private Sub txt_ref_GotFocus()
If Len(Trim(cbo_typePiece.Text)) = 0 Then
    MsgBox "Type Piece Obligatoire      ", vbInformation
    cbo_typePiece.SetFocus
    Exit Sub
End If


End Sub



Private Sub txt_ref_LostFocus()
If txt_ref.Text = "" Then
txt_ref.Text = "Sans Ref"
End If
End Sub


Public Sub AfficheRow_BCR(ByVal vcode As String)

Dim SQL As String
Dim rs As New ADODB.Recordset
Lsv_Detail.ListItems.Clear


SQL = "SELECT * from detailBCReparation "
SQL = SQL & " WHERE    detailBCReparation.Numero = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    While Not rs.EOF
            Set itmX = Lsv_Detail.ListItems.Add(, , CStr(txt_Numero.Text))
            txt_BCReparation.Text = rs("Numero")
            itmX.SubItems(1) = rs("Désignation")
            itmX.SubItems(2) = rs("qté")
            itmX.SubItems(3) = rs("Vehicule")
            itmX.SubItems(4) = "00,00"
            itmX.SubItems(5) = "00,00"
            itmX.SubItems(6) = "00,00"
            itmX.SubItems(7) = "00,00"
            itmX.SubItems(8) = "00,00"
            
        rs.MoveNext
    Wend
End If
rs.Close

End Sub

Private Sub txt_Timbre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call AppCalcul
SendKeys "{tab}"
End If
End Sub

Private Sub txt_Timbre_LostFocus()
Call AppCalcul
txt_Timbre.Text = Format(txt_Timbre.Text, "#,##0.000")

End Sub

Public Sub AfficheRow_Station(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from station where Actif=1 And code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_MatriculeStation.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
End If

End Sub


