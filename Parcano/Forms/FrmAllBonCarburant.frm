VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAllBonCarburant 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bon carburant"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAllBonCarburant.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   14670
   Begin MSComctlLib.ListView Lsv_Client 
      Height          =   3375
      Left            =   240
      TabIndex        =   44
      Top             =   5280
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5953
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
      Appearance      =   0
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CodeMat"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Immatriculation"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Energie"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "A.Compteur"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "N.Compteur"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "N.Litre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Prix.Litre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Valeur.TTC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "uht"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "TVA"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "KM.Parcouru"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Consommation par cent KM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Observ"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Anomali"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComCtl2.DTPicker dateOp 
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   148635649
      CurrentDate     =   42875
   End
   Begin VB.PictureBox Pict_user 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8400
      ScaleHeight     =   495
      ScaleWidth      =   3495
      TabIndex        =   39
      Top             =   1080
      Width           =   3495
      Begin VB.Label Lbl_UserSaisi 
         BackStyle       =   0  'Transparent
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
         Left            =   1320
         TabIndex        =   41
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Lbl_user 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BC saisi par :"
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
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.ComboBox cbo_MatriculeStation 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   2520
      Width           =   2655
   End
   Begin SToolBox.SBiCombo Cbo_Conducteur 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   4680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7440
      ScaleHeight     =   495
      ScaleWidth      =   3615
      TabIndex        =   35
      Top             =   4440
      Width           =   3615
      Begin VB.TextBox txt_Valeur 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   13
         Tag             =   "M"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total TTC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   0
         TabIndex        =   36
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   5280
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   7440
      ScaleHeight     =   1905
      ScaleWidth      =   3705
      TabIndex        =   34
      Top             =   2400
      Width           =   3735
      Begin SToolBox.SGrid Grid_Recherche 
         Height          =   2055
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3625
         RowMode         =   -1  'True
         GridLines       =   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   16777215
         GridLineColor   =   16777215
         GridFillLineColor=   16777215
         GroupRowBackColor=   -2147483624
         GroupRowForeColor=   192
         AlternateRowBackColor=   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         DisableIcons    =   -1  'True
         StretchLastColumnToFit=   -1  'True
      End
   End
   Begin VB.PictureBox PIC_NFACT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3360
      ScaleHeight     =   495
      ScaleWidth      =   5055
      TabIndex        =   31
      Top             =   960
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ce bon est inseré dans une facture N° : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   33
         Top             =   120
         Width           =   3660
      End
      Begin VB.Label LBL_NFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "250000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3840
         TabIndex        =   32
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   11040
      ScaleHeight     =   3375
      ScaleWidth      =   615
      TabIndex        =   28
      Top             =   5280
      Width           =   615
      Begin SToolBox.SCommand Cmd_SuppLign 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
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
         Picture         =   "FrmAllBonCarburant.frx":0ECA
      End
      Begin SToolBox.SCommand Cmd_Modif 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
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
         Picture         =   "FrmAllBonCarburant.frx":104C
      End
      Begin SToolBox.SCommand Cmd_NewSaisi 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
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
         Picture         =   "FrmAllBonCarburant.frx":139F
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3480
      ScaleHeight     =   375
      ScaleWidth      =   6615
      TabIndex        =   26
      Top             =   1800
      Width           =   6615
      Begin VB.Label cda_Create 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   42
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Operation :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   37
         Top             =   0
         Width           =   1380
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date création :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   6735
      TabIndex        =   22
      Top             =   3120
      Width           =   6735
      Begin VB.TextBox txt_NBC 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txt_ville 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txt_adresse 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txt_rsocial 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Lbl_nbc 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "  Nbr BC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   43
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ville :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   25
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   24
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Raison sociale :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   1290
      End
   End
   Begin VB.TextBox txt_Numero 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   435
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "M"
      Top             =   1680
      Width           =   1215
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   10680
      TabIndex        =   14
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
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
      Picture         =   "FrmAllBonCarburant.frx":1521
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   9960
      TabIndex        =   17
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
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
      Picture         =   "FrmAllBonCarburant.frx":16A3
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   10320
      TabIndex        =   16
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
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
      Picture         =   "FrmAllBonCarburant.frx":19F6
   End
   Begin SToolBox.SCommand cmdFindNumero 
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   1680
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      Picture         =   "FrmAllBonCarburant.frx":1D49
      ButtonType      =   1
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   9600
      TabIndex        =   18
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
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
      Picture         =   "FrmAllBonCarburant.frx":209C
   End
   Begin SToolBox.SCommand CmdPrint 
      Height          =   495
      Left            =   11160
      TabIndex        =   15
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
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
      Picture         =   "FrmAllBonCarburant.frx":221E
   End
   Begin SToolBox.SCommand CmdFindStation 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2520
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      Picture         =   "FrmAllBonCarburant.frx":2571
      ButtonType      =   1
   End
   Begin SToolBox.SCommand CmdFindConducteur 
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      Top             =   4680
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      Picture         =   "FrmAllBonCarburant.frx":28C4
      ButtonType      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bon de sortie carburant"
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
      Left            =   360
      TabIndex        =   38
      Top             =   360
      Width           =   3855
   End
   Begin VB.Image PicBox_Header 
      Height          =   1575
      Left            =   -120
      Picture         =   "FrmAllBonCarburant.frx":2C17
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   12615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conducteur :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   270
      Left            =   120
      TabIndex        =   30
      Top             =   4680
      Width           =   1410
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Station "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro bon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "FrmAllBonCarburant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Okayy As Boolean
Dim itmX As ListItem
Dim LOBJ_BonCarburant As BonCarburant
Dim indexSelect As Integer

Private Sub ExistData(ByVal cbo As ComboBox)
On Error GoTo Err

    Dim RCount As Integer, i As Integer, Existe As Boolean, tcbo As String
    
    RCount = cbo.ListCount
    tcbo = cbo.Text
    
    For i = 0 To RCount - 1
        cbo.ListIndex = i
        If tcbo = cbo.Text Then
            Existe = True
            Exit For
        Else
            Existe = False
        End If
    Next i
    If i = RCount Then
        If Existe = False Then
            MsgBox "Saisie non Valide!...     ", vbExclamation, App.ProductName
            cbo.Text = ""
            txt_rsocial.Text = ""
            txt_adresse.Text = ""
            txt_ville.Text = ""
            Exit Sub
        End If
    End If
    Existe = False
    
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub

' Fonction retournant le code du conducteur à partir de son libelle

Private Function RET_CODE_CONDUCTEUR(VCode As String) As String

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset
' Initialisation
RET_CODE_CONDUCTEUR = ""
Set LOBJ_Personnel = New personnel

Set rs = LOBJ_Personnel.GetCODE_CONDUCTEUR(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If

If Not rs.EOF Then
    RET_CODE_CONDUCTEUR = rs("Code")
End If
rs.Close
 
End Function

'Fonction retourne le prix de l'energie à partie du libelle

Private Function RET_PRIX_ENERGIE(Libelle As String) As Double

Dim LOBJ_Energie As Energie
Dim rs As New Recordset
' Initialisation
RET_PRIX_ENERGIE = 0

Set LOBJ_Energie = New Energie
Set rs = LOBJ_Energie.Get_PrixEnergie(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Libelle)
If Not rs.EOF Then
    RET_PRIX_ENERGIE = rs("Prix")
End If
rs.Close
End Function

'Fonction retourne le nombre des bons de carburant du station à partie de son code

Private Function Return_NBC(VCode As String) As Long

Dim LOBJ_Station As Station
Dim rs As New Recordset
' Initialisation
Return_NBC = 0

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.Get_NBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)

If Not rs.EOF Then
    Return_NBC = rs("numbc")
End If
rs.Close
End Function

Private Sub cbo_conducteur_LostFocus()
Call ExistDonnee(cbo_Conducteur)
End Sub

Private Sub cbo_MatriculeStation_LostFocus()
If Len(Trim(cbo_MatriculeStation.Text)) > 0 Then Call AfficheRow_Station(cbo_MatriculeStation.Text)

End Sub

Private Sub Cmd_NewSaisi_GotFocus()
Dim KeyCode As Integer
If KeyCode = vbKeyReturn Then Call Cmd_NewSaisi_Click
End Sub

'Recherche des anciens bonsCarburant

Private Sub CmdFind_Click()

On Error GoTo Err

If Lsv_Client.ListItems.Count > 0 And txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

If Okayy = True Then   'Modification d'un bon est en cours
    If MsgBox("Annuler le MAJ en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If
txt_nbc.Visible = True
Lbl_nbc.Visible = True
CmdDelete.Enabled = True
Unload FrmFind
With FrmFind_BCarb
    .Show vbModal
End With
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

'Afficher liste des conducteurs

Private Sub cmdFindConducteur_Click()

On Error GoTo Err
If txt_Numero.Text = "" Then
    Exit Sub
Else
Okayy = True
Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Personnel_Conducteur"
    .Show vbModal
End With
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

'Afficher liste des anciens bons de carburant
'Même liste à afficher que celle du cmdFind_click
Private Sub cmdFindNumero_Click()

On Error Resume Next
If Lsv_Client.ListItems.Count > 0 And txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

If Okayy = True Then   'Modification d'un bon est en cours
    If MsgBox("Annuler le MAJ en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If
txt_nbc.Visible = True
Lbl_nbc.Visible = True
Unload FrmFind
With FrmFind_BCarb
    .Show vbModal
End With
End Sub

'Afficher liste des stations

Private Sub CmdFindStation_Click()

On Error GoTo Err

If txt_Numero.Text <> "" Then
    Unload FrmFind_Fils
    FrmFind_Fils.StrSource = "Station carburant"  ' selon StrSource la liste de recherche s'affiche
    FrmFind_Fils.Show vbModal

End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

' Impression du bon de carburant

Private Sub CmdPrint_Click()

Dim F As Form
On Error GoTo Err

If txt_Numero.Text = "" Then Exit Sub

'Annuler modification d'un bon en cours
If Okayy = True Then
    If MsgBox("Annuler le MAJ en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
End If

'Si un bon est en cours de saisie pas encore enregistrer
If Lsv_Client.ListItems.Count > 0 And txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        txt_Numero.SetFocus
        Exit Sub
    End If
End If

' Impression d'un bonC existant mais pas un bon en cours de saisie
If txt_Numero.Text = "Auto" Then Exit Sub

If MsgBox("Imprimer ce bon   ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Set F = New Frm_Rpt_Apercus  'Mettre le bon à imprimer de cette forme de fichier
    With F
        .Numero = txt_Numero.Text
        'Afficher la fichier CrystalReport avant l'impression
        Call .PrintOutAndApercu_BC(0)
        .Show
    End With
End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

'Enregistrement soit de la modification du bon soit de la saisie du nouveau bon

Private Sub CmdSave_Click()

Dim F As Form
Dim LOBJ_Personnel As personnel
Dim rs As New Recordset

On Error GoTo Err

' Ce bon est inseré dans une facture N°: bon payé à ne pas ni supprimer ni modifier
If PIC_NFACT.Visible = True Then
    MsgBox "MAJ impossible, ce bon est inséré dans un facture. ", vbInformation
    Exit Sub
End If

' Vérifier si les champs de la forme sont remplis ou non
If Left(CheckMandatory(FrmAllBonCarburant), 1) = 1 Then ' tout les champs vides
   Exit Sub
End If

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_CONDUCTEUR(ErrNumber, ErrDescription, ErrSourceDetail, CNB, cbo_Conducteur.FirstValue)
If ErrNumber <> 0 Then
   MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
   ErrNumber = 0
   Exit Sub
End If
If rs.EOF Then
    MsgBox "Ce conducteur n'existe pas ", vbInformation
    cbo_Conducteur.SetFocus
    Exit Sub
    rs.Close
End If

'Vérifier s'il n'y a pas de détails ajoutés dans la liste Lsv_Client
If Lsv_Client.ListItems.Count = 0 Then
    MsgBox "Veuillez saisir les details ", vbInformation
End If

If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub

'---------------- MAJ d'un ancien BC ----------------

If txt_Numero.Text <> "Auto" And txt_Numero.Text <> "" Then
    'Vérification du droit de MAJ du bon de carburant pour cet utilisateur
        If (CHECK_ACCES("MAJ_BC", LInt_UserId) = False) Then
            MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier un Bon de Carburant", vbExclamation
            Exit Sub
        End If
        
    txt_nbc.Text = CStr(Return_NBC(cbo_MatriculeStation.Text))
    Call Modifier_BC
End If

'---------------- Insertion d'un nouveau BC ----------------

If txt_Numero.Text = "Auto" Then
    Call Ajouter_BC
End If

Okayy = False

'Afficher le bon à imprimer après l'enregistrement
If MsgBox("Enregistrement terminé avec succé  " & vbNewLine & "Imprimer ce bon        ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Set F = New Frm_Rpt_Apercus
    With F
        .Numero = txt_Numero.Text
        Call .PrintOutAndApercu_BC(0)
        .Show
    End With
End If
    Call ViderZone(FrmAllBonCarburant)
    Grid_Recherche.ClearRows
    Lsv_Client.ListItems.Clear
    txt_nbc.Visible = True
    Lbl_nbc.Visible = True
    cbo_MatriculeStation.Enabled = True
    CmdFindStation.Enabled = True
    cmdFindConducteur.Enabled = True
    cbo_Conducteur.Enabled = True
    'txt_Numero.SetFocus
    'Grid_Recherche.ClearRows
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

'Ajout d'un nouveau BC

Private Sub Ajouter_BC()

Dim LOBJ_Station As Station
Dim Lobj_Vehicule As VEHICULE
Dim LRs_NewRecord As Recordset
Dim LInt_NumCompteur As Long
Dim i As Long

'Incrementation du compteur du numero du bon de carburant
LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "Boncarburant")
If ErrNumber <> 0 Then
   MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
   ErrNumber = 0
   Exit Sub
End If

'Insertion enregistrement assiette
txt_Numero.Text = Format(LInt_NumCompteur, "00000")
Set LOBJ_Station = New Station
'Incrémenter le nombre des bons de carburant pour cette station
Call LOBJ_Station.UpdateNBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, 1, cbo_MatriculeStation.Text)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
txt_nbc.Text = CStr(Return_NBC(cbo_MatriculeStation.Text))
'Changer le CompteurCarburant par le nouveau compteur pour chaque véhicule ajouté
'dans les détails du bon
Set Lobj_Vehicule = New VEHICULE
For i = 1 To Lsv_Client.ListItems.Count
'Modifier l'affectation de Ancompteur  par compteurCarburant dans la table véhicule
    Call Lobj_Vehicule.Update_Compt(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Val(Lsv_Client.ListItems(i).SubItems(6)), Lsv_Client.ListItems(i).SubItems(2))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
Next
'Insertion l'assiette et les details du BC
Call Insert_AssBC
Call Insert_DetailBC

End Sub

'Modification d'un ancien BC

Private Sub Modifier_BC()

Dim Lobj_Vehicule As VEHICULE
Dim VCode
Dim i As Integer
'Enregistrement du MAJ du BC
Set LOBJ_BonCarburant = New BonCarburant
VCode = txt_Numero.Text
'Supprimer les details de ce BC
Call LOBJ_BonCarburant.Delete_DetailBON(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set Lobj_Vehicule = New VEHICULE
For i = 1 To Lsv_Client.ListItems.Count
'Modifier l'affectation de Ancompteur  par compteurCarburant dans la table véhicule
    Call Lobj_Vehicule.Update_Compt(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Val(Lsv_Client.ListItems(i).SubItems(6)), Lsv_Client.ListItems(i).SubItems(2))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
Next
'MAJ de l'assiette du BC : Conducteur , date opération et les totaux
Call Update_AssBC

' Insertion de nouveau des details du BC
Call Insert_DetailBC

End Sub

'Update assiette BC

Private Sub Update_AssBC()

Dim LRs_NewRecord As New Recordset
Dim LOBJ_BonCarburant As BonCarburant

Set LOBJ_BonCarburant = New BonCarburant
Set LRs_NewRecord = CreateEmptyRS_AssBC()
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = txt_Numero.Text
    .Fields("DateDoc") = CDate(cda_Create.Caption)
    .Fields("Heure") = Format(Time, "hh:mm")
    .Fields("Station") = cbo_MatriculeStation.Text
    .Fields("Conducteur") = cbo_Conducteur.FirstValue
    .Fields("valeur") = CDbl(txt_Valeur.Text)   'Total TTC
    .Fields("nbc") = Val(txt_nbc.Text)
    .Fields("dateop") = CDate(dateOp.Value)
    .Fields("UserUpdate") = LInt_UserId
End With

Call LOBJ_BonCarburant.Update_AssBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

End Sub

'Insertion du bon du commande dans la base
Private Sub Insert_AssBC()

Dim LRs_NewRecord As New Recordset

Set LOBJ_BonCarburant = New BonCarburant
Set LRs_NewRecord = CreateEmptyRS_AssBC()
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = txt_Numero.Text
    .Fields("DateDoc") = CDate(cda_Create.Caption)
    .Fields("Heure") = Format(Time, "hh:mm")
    .Fields("Station") = cbo_MatriculeStation.Text
    .Fields("Conducteur") = cbo_Conducteur.FirstValue
    .Fields("valeur") = CDbl(txt_Valeur.Text)   'Total TTC
    .Fields("nbc") = Val(txt_nbc.Text)
    .Fields("dateop") = CDate(dateOp.Value)
    .Fields("UserInsert") = LInt_UserId
    
End With

Call LOBJ_BonCarburant.Insert_AssBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

End Sub

'Insertion des détails du bon du commande dans la base
Private Sub Insert_DetailBC()

Dim LRs_NewRecord As New Recordset
Dim ii As Integer

Set LOBJ_BonCarburant = New BonCarburant
Set LRs_NewRecord = CreateEmptyRS_DetailBC()

For ii = 1 To Lsv_Client.ListItems.Count
    With LRs_NewRecord
        .AddNew
        .Fields("Numero") = txt_Numero.Text   'Numero BC
        .Fields("Vehicule") = Lsv_Client.ListItems(ii).SubItems(2) 'Immatriculation
        .Fields("Energie") = Lsv_Client.ListItems(ii).SubItems(4)
        .Fields("CompteurCarburant") = Val(Lsv_Client.ListItems(ii).SubItems(6))
        .Fields("litre") = CDbl(Lsv_Client.ListItems(ii).SubItems(7))
        .Fields("prixLitre") = CDbl(Lsv_Client.ListItems(ii).SubItems(8))
        .Fields("prixht") = CDbl(Lsv_Client.ListItems(ii).SubItems(10))
        .Fields("tva") = CDbl(Lsv_Client.ListItems(ii).SubItems(11))
        .Fields("Observation") = Lsv_Client.ListItems(ii).SubItems(14)
        .Fields("AnomalieConsom") = CDbl(Lsv_Client.ListItems(ii).SubItems(15))
    End With
    Lsv_Client.ListItems(ii).Text = txt_Numero.Text
Next
Call LOBJ_BonCarburant.Insert_DetailBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing
End Sub

'Saisir un nouveau bon de carburant

Private Sub CmdAdd_Click()


On Error GoTo Err

txt_nbc.Visible = False
Lbl_nbc.Visible = False


' Vérifier les droits d'accès de l'utilisateur : s'il a le droit d'ajouter un nouveau bon.
    If (CHECK_ACCES("InS_BC", LInt_UserId) = False) Then
        MsgBox "Insertion n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Ajouter un Bon de Carburant", vbExclamation
        Exit Sub
    End If

Timer1.Enabled = False
PIC_NFACT.Visible = False
Okayy = False

If Lsv_Client.ListItems.Count > 0 And txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Call ViderZone(FrmAllBonCarburant)
CmdDelete.Enabled = False
CmdSave.Enabled = True
cbo_MatriculeStation.Enabled = True
CmdFindStation.Enabled = True      ' Choisir une station parmis la liste
cmdFindConducteur.Enabled = True   ' Choisir un conducteur parmis la liste
cbo_Conducteur.Enabled = True
Lbl_UserSaisi.Caption = LStr_NameUser
txt_Numero.Text = "Auto"
txt_Valeur.Text = "0,000"
cda_Create.Caption = Date
dateOp.Value = Date
'dateOp.SetFocus

Cmd_SuppLign.Enabled = True  'Pour supprimer une ligne de commande
Cmd_Modif.Enabled = True  'Pour modifier une ligne d'un bon de carburant existant
Cmd_NewSaisi.Enabled = True  'Pour saisir une nouvelle ligne du bon de carburant : FrmSaisieBoncarburant
' vider la liste d'affichage du detail du bon de carburant et le Grid récapitulatif.
Lsv_Client.ListItems.Clear
Grid_Recherche.ClearRows
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

'Suppression d'un bon du carburant ou annulation du saisie d'un nouveau bon

Private Sub CmdDelete_Click()

Dim LOBJ_Station As Station
Dim Lobj_Vehicule As VEHICULE
Dim LOBJ_BonCarburant As BonCarburant
Dim LOBJ_Personnel As personnel
Dim rs As New Recordset
Dim VCode As String
Dim NBC As Long
Dim i As Integer
Dim numMax

On Error GoTo Err

' Ce bon est inseré dans une facture N°: bon payé donc on peut pas ni le supprimer ni modifier
If PIC_NFACT.Visible = True Then
    MsgBox "Maj impossible, Ce bon est inseré dans une facture. ", vbInformation
    Exit Sub
End If

'Annulation de la création d'un nouveau bon
If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        txt_Numero.SetFocus
        Exit Sub
    End If
End If

' Vérifier si l'utilisateur à le droit de supprimer un bon suivant son code
If txt_Numero.Text <> "Auto" Then
If (CHECK_ACCES("Supp_BC", LInt_UserId) = False) Then
        MsgBox "Suppression n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Supprimer un Bon de Carburant", vbExclamation
        Exit Sub
    End If
End If
    
If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "bon de carburant", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
    Exit Sub
End If

VCode = txt_Numero.Text  ' numero du bon

Set LOBJ_Station = New Station
' MAJ du nombre des bon de la station : supprimer 1
Call LOBJ_Station.UpdateNBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, -1, cbo_MatriculeStation.Text)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
NBC = Return_NBC(cbo_MatriculeStation.Text)

'si on supprime le dernier BC pour un vehicule on doit remettre le compteurCarb à l'ancienne valeur
Set LOBJ_BonCarburant = New BonCarburant
Set Lobj_Vehicule = New VEHICULE
    'MAJ CompteurCarburant de chaque véhicule inclu dans ce bon : remmetre l'ancien compteur
For i = 1 To Lsv_Client.ListItems.Count
    Set rs = LOBJ_BonCarburant.Get_MaxNumBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lsv_Client.ListItems(i).SubItems(2))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs.EOF Then
       numMax = rs("maxNum")
    End If
    rs.Close
    If Val(txt_Numero.Text) = Val(numMax) Then
        Call Lobj_Vehicule.Update_Compt(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lsv_Client.ListItems(i).SubItems(5), Lsv_Client.ListItems(i).SubItems(2))
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
    End If
Next

' Suppression de l'assiette et les details du bon  set Supp='O'
Call LOBJ_BonCarburant.Delete_DetBCBySupp(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

Call LOBJ_BonCarburant.Delete_AssBON(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LInt_UserId, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

Call ViderZone(FrmAllBonCarburant)
Lsv_Client.ListItems.Clear
Grid_Recherche.ClearRows
txt_Numero.SetFocus

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub dateOp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"   'CmdFindStation
End Sub

Private Sub Form_Load()

On Error GoTo Err
CmdDelete.Enabled = False
CmdSave.Enabled = False
CmdPrint.Enabled = False
Me.Width = 11715
Me.Height = 8625
Me.Move 0, 0
cda_Create.Caption = Date
dateOp.Value = Date
txt_nbc.Visible = False
Lbl_nbc.Visible = False
Call Affiche_Personnel_SBCombo(cbo_Conducteur) ' Charger liste des conducteurs dans SBiComboBox
Call Affiche_StatCarb_Combo(cbo_MatriculeStation)
Call InitialGrid  'Initialisation du Grid
Me.WindowState = 2

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub Form_Resize()
    Dim WidthForm As Integer
    WidthForm = Frm_Main.ACB_Main.Width
        PicBox_Header.Width = WidthForm - 1000
        CmdAdd.Left = WidthForm - 5500
        CmdDelete.Left = WidthForm - 5100
        CmdFind.Left = WidthForm - 4700
        CmdSave.Left = WidthForm - 4300
        CmdPrint.Left = WidthForm - 3900
End Sub

Private Sub Form_Unload(Cancel As Integer)

'On Error GoTo erreur
'   Dim i As Integer
'   Dim Msg ' Déclare la variable.
'   ' Définit le texte du message.
'   Msg = "Voulez-vous vraiment quitter?"
'   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
'   If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'      Cancel = True
'   Else
'      Unload Me
'   End If
'
'   Exit Sub
'erreur:
'   MsgBox Err.Description, 48

End Sub


'Double click sur liste des details
'Afficher pour modifier ligne detail du bon à partir du FrmSaisieBoncarburant
'puis la ré-afficher dans la liste Lsv_Client
Private Sub Lsv_Client_DblClick()

Dim i
Dim ii

On Error GoTo Err

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If

If PIC_NFACT.Visible = True Or Timer1.Enabled = True Then Exit Sub  'Bon facturé
If Lsv_Client.ListItems.Count <= 0 Then Exit Sub ' Pas de details insérés

With FrmSaisieBoncarburant
    .Okay = False
    .ii = Lsv_Client.SelectedItem.Index
    i = Lsv_Client.SelectedItem.Index
    .txt_Numero.Text = txt_Numero.Text
    .cda_Create.Text = cda_Create.Caption
'    .Cbo_Matricule.Text = CStr(Lsv_Client.ListItems(i).SubItems(2) & " - " & Lsv_Client.ListItems(i).SubItems(3))
    .AfficheRow_Vehicule_sansPrix (Lsv_Client.ListItems(i).SubItems(2))
    .txt_Ncompteur.Text = Lsv_Client.ListItems(i).SubItems(6)
    .txt_Compteur.Text = Lsv_Client.ListItems(i).SubItems(5)
    .txt_NbreLitre.Text = Lsv_Client.ListItems(i).SubItems(7)
    .txt_prixLitre.Text = Lsv_Client.ListItems(i).SubItems(8)
    .txt_Valeur.Text = Lsv_Client.ListItems(i).SubItems(9)
    .txt_ht.Caption = Lsv_Client.ListItems(i).SubItems(10)
    .txt_tva.Caption = Lsv_Client.ListItems(i).SubItems(11)
    .LBL_DIF_COMP.Caption = Val(.txt_Ncompteur.Text) - Val(.txt_Compteur.Text) & " KM "
    .Lbl_Consommation.Caption = Lsv_Client.ListItems(i).SubItems(13)
    .Txt_Observ.Text = Lsv_Client.ListItems(i).SubItems(14)
    .Lbl_anomaliConso.Caption = Lsv_Client.ListItems(i).SubItems(15)
    .Anomalie
    .cbo_Matricule.Enabled = False
    .cmdFindMatricule.Enabled = False
    .Show vbModal
End With

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

'Suppression d'une ligne de la liste des details du bon

Private Sub Cmd_SuppLign_Click()

On Error GoTo Err
'Si pas de bon sélectionné ou pas de bon en cours de saisie
If Len(Trim(txt_Numero.Text)) = 0 Then  'Trim : Renvoie une copie d'une chaîne sans espaces à gauche ni à droite
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If

If indexSelect <> 0 Then
'Liste de details du bon est vide
    If Lsv_Client.ListItems.Count <= 0 Then Exit Sub
    Okayy = True
    If MsgBox("Confirmez vous la suppression de la ligne en cours.?", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    '    I = Lsv_Client.SelectedItem.Index  ' indice de la ligne de detail sélectionné
        Lsv_Client.ListItems.Remove indexSelect     'Supprimer la ligne de la liste
        Call AppCalcul                    'Refaire les calculs
        Call Get_Details
    End If
Else
    MsgBox "Sélectionner une ligne à supprimer !...       ", vbExclamation, App.ProductName
    Exit Sub
End If
indexSelect = 0
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

'Modification d'une ligne de detail

Private Sub Cmd_Modif_Click()

On Error GoTo Err

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If
If indexSelect <> 0 Then
    If Lsv_Client.ListItems.Count <= 0 Then Exit Sub
    Okayy = True
    'Afficher le contenu de la ligne dans les champs textes pour la modifier
    With FrmSaisieBoncarburant
        .Okay = False
        .ii = Lsv_Client.SelectedItem.Index
        .txt_Numero.Text = txt_Numero.Text
        .cda_Create.Text = cda_Create.Caption
        .cbo_Matricule.Text = CStr(Lsv_Client.ListItems(indexSelect).SubItems(2) & " - " & Lsv_Client.ListItems(indexSelect).SubItems(3))
        .AfficheRow_Vehicule_sansPrix (Lsv_Client.ListItems(indexSelect).SubItems(2))
        .txt_Ncompteur.Text = Lsv_Client.ListItems(indexSelect).SubItems(6)
        .txt_NbreLitre.Text = Lsv_Client.ListItems(indexSelect).SubItems(7)
        .txt_prixLitre.Text = Lsv_Client.ListItems(indexSelect).SubItems(8)
        .txt_Valeur.Text = Lsv_Client.ListItems(indexSelect).SubItems(9)
        .LBL_DIF_COMP.Caption = Val(.txt_Ncompteur.Text) - Val(.txt_Compteur.Text) & " KM "
        .Lbl_Consommation.Caption = Lsv_Client.ListItems(indexSelect).SubItems(13)
        .Txt_Observ.Text = Lsv_Client.ListItems(indexSelect).SubItems(14)
        .Lbl_anomaliConso.Caption = Lsv_Client.ListItems(indexSelect).SubItems(15)
        .Anomalie
        .cbo_Matricule.Enabled = False
        .cmdFindMatricule.Enabled = False
        .Show vbModal
    End With
Else
    MsgBox "Sélectionner une ligne à modifier!...       ", vbExclamation, App.ProductName
    Exit Sub
End If
indexSelect = 0
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

'Saisir une nouvelle ligne de detail à partir de la forme FrmSaisieBoncarburant

Private Sub Cmd_NewSaisi_Click()

'Vérification du num du bon
If txt_Numero.Text = "" Then
    If Len(Trim(txt_Numero.Text)) = 0 Then
        MsgBox "N° bon obligatoire      ", vbInformation
        txt_Numero.SetFocus
        Exit Sub
    End If
End If
 'Vérification de la station
If cbo_MatriculeStation.Text = "" Or cbo_MatriculeStation.Text = " " Then
    If Len(Trim(cbo_MatriculeStation.Text)) = 0 Then
        MsgBox "Station obligatoire      ", vbInformation
        cbo_MatriculeStation.SetFocus
        Exit Sub
    End If
End If
'Vérification du conducteur
If cbo_Conducteur.Text = "" Then
    If Len(Trim(cbo_Conducteur.Text)) = 0 Then
        MsgBox "Conducteur obligatoire      ", vbInformation
        cbo_Conducteur.SetFocus
        Exit Sub
    End If
End If

On Error GoTo Err
Okayy = True

    With FrmSaisieBoncarburant
        .txt_Numero.Text = Me.txt_Numero.Text
        .cda_Create.Text = Me.cda_Create.Caption
        .Okay = True
        .Show vbModal
    End With
FrmSaisieBoncarburant.cbo_Matricule.Enabled = True
FrmSaisieBoncarburant.cmdFindMatricule.Enabled = True
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

'Charger les details de la station choisi dans les champs textes associés

Public Sub AfficheRow_Station(ByVal VCode As String)

Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.GetStatByCodeLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    cbo_MatriculeStation.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
Else
    MsgBox "Code introuvable!", vbInformation, App.ProductName
    cbo_MatriculeStation.Text = ""
End If

End Sub

Private Sub cbo_MatriculeStation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cbo_Conducteur.SetFocus
End Sub

'Vider tout les champs de la forme en cliquant sur le champs text txt_Numero

Private Sub txt_Numero_GotFocus()

On Error GoTo Err

Call ViderZone(FrmAllBonCarburant)
Lsv_Client.ListItems.Clear
cbo_MatriculeStation.Enabled = True
CmdFindStation.Enabled = True
cmdFindConducteur.Enabled = True
cbo_Conducteur.Enabled = True

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub txt_Numero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cbo_MatriculeStation.SetFocus
End Sub

'Charger les données concernant ce bon de numero txt_Numero

Private Sub txt_Numero_LostFocus()

On Error GoTo Err

If Len(Trim(txt_Numero.Text)) > 0 Then Call AfficheRow(txt_Numero.Text)

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub
Private Sub EnabDis(ByVal bl As Boolean)

CmdDelete.Enabled = bl
CmdSave.Enabled = bl
CmdPrint.Enabled = bl
Cmd_SuppLign.Enabled = bl
Cmd_Modif.Enabled = bl
Cmd_NewSaisi.Enabled = bl
cbo_MatriculeStation.Enabled = bl
CmdFindStation.Enabled = bl
cmdFindConducteur.Enabled = bl
cbo_Conducteur.Enabled = bl
txt_Numero.Enabled = bl
Lsv_Client.Enabled = bl
End Sub

'Charger les données concernant ce bon par Dbclick sur la liste affiche dans FrmFind

Public Sub AfficheRow(ByVal VCode As String)

Dim rs As New Recordset
Dim rs1 As New Recordset
Dim LOBJ_BonCarburant As BonCarburant
Dim i

Lsv_Client.ListItems.Clear
Set LOBJ_BonCarburant = New BonCarburant

Set rs = LOBJ_BonCarburant.Get_AssBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    If rs("Supp") = "O" Then
        MsgBox "Ce bon Carburant a été supprimé par " & Get_NameUserByCode(rs("UserDelete")), vbInformation
        Call EnabDis(False)
    Else
        Call EnabDis(True)
        cbo_MatriculeStation.Enabled = False
        CmdFindStation.Enabled = False
    End If
    'Charge de l'assiette du bonCarburant
    If Not IsNull(rs("Numero")) Then txt_Numero.Text = rs("Numero")
    If Not IsNull(rs("STATION")) Then cbo_MatriculeStation.Text = rs("STATION")
    If Not IsNull(rs("DATEDOC")) Then cda_Create.Caption = rs("DATEDOC")
    If Not IsNull(rs("CONDUCTEUR")) Then cbo_Conducteur.Text = rs("CONDUCTEUR")

    If Not IsNull(rs("VALEUR")) Then txt_Valeur.Text = CStr(Format(rs("VALEUR"), "#,##0.000"))
    If Not IsNull(rs("NBC")) Then txt_nbc.Text = rs("NBC")
    If Not IsNull(rs("dateOp")) Then dateOp.Value = rs("dateOp")
    If Not IsNull(rs("UserInsert")) Then Lbl_UserSaisi.Caption = Get_NameUserByCode(rs("UserInsert"))
    Call AfficheRow_Station(rs("STATION"))
    Call AfficheRow_Conducteur(rs("CONDUCTEUR"))
    
    If rs("Transf") = "O" Then
        LBL_NFact.Caption = rs("NumFact")
        PIC_NFACT.Visible = True
        Call EnabDis(False)
        Call Timer1_Timer
    Else
        Timer1.Enabled = False
        PIC_NFACT.Visible = False
    End If
End If
rs.Close

'Charger les details du bon dans la liste LSV_client
Set rs = LOBJ_BonCarburant.Get_DetailBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
i = 0
If Not rs.EOF Then
    While Not rs.EOF
        i = Lsv_Client.ListItems.Count + 1
        Set itmX = Lsv_Client.ListItems.Add(, , CStr(txt_Numero.Text))
        If Not IsNull(cda_Create.Caption) Then itmX.SubItems(1) = CStr(cda_Create.Caption)
        
        If Not IsNull(rs("Vehicule")) Then
            itmX.SubItems(2) = CStr(rs("Vehicule"))

            Set rs1 = LOBJ_BonCarburant.Get_AnComptCar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode, rs("Vehicule"))
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
            If Not rs1.EOF Then
                If Not IsNull(rs1("maxCpt")) Then itmX.SubItems(5) = CStr(rs1("maxCpt"))
            Else
                itmX.SubItems(5) = "0"
            End If
        End If
        
        If Not IsNull(rs("Matricule")) Then itmX.SubItems(3) = CStr(rs("Matricule"))
        If Not IsNull(rs("Energie")) Then itmX.SubItems(4) = CStr(rs("Energie"))
        If Not IsNull(rs("CompteurCarburant")) Then itmX.SubItems(6) = CStr(rs("CompteurCarburant"))
        If Not IsNull(rs("Litre")) Then itmX.SubItems(7) = CStr(Format(rs("Litre"), "#,##0.00"))
        itmX.SubItems(8) = CStr(Format(rs("prixLitre"), "#,##0.000"))
        itmX.SubItems(9) = Format(rs("Litre") * rs("prixLitre"), "#,##0.000")
        itmX.SubItems(10) = Format(rs("prixht"), "#,##0.000")
        itmX.SubItems(11) = Format(rs("tva"), "#0.00")
        itmX.SubItems(12) = Val(CStr(rs("CompteurCarburant"))) - Val(itmX.SubItems(5)) & " KM "
        itmX.SubItems(13) = CStr(Format(Calcule_Consommation(Val(itmX.SubItems(7)), Val(itmX.SubItems(12))), "#,##0.00")) & " L/100km"
        If Not IsNull(rs("Observation")) Then itmX.SubItems(14) = CStr(rs("Observation"))
        If Not IsNull(rs("AnomalieConsom")) Then
            itmX.SubItems(15) = CStr(rs("AnomalieConsom"))
            If CDbl(rs("AnomalieConsom")) >= 2 Then
                With Lsv_Client.ListItems(i)
                    .Bold = True
                    .ListSubItems(3).Bold = True
                    .ListSubItems(6).Bold = True
                    .ListSubItems(7).Bold = True
                    .ListSubItems(13).Bold = True
                    .ListSubItems(15).Bold = True
                    .ForeColor = &HFF&
                    .ListSubItems(3).ForeColor = &HFF&
                    .ListSubItems(6).ForeColor = &HFF&
                    .ListSubItems(7).ForeColor = &HFF&
                    .ListSubItems(13).ForeColor = &HFF&
                    .ListSubItems(15).ForeColor = &HFF&
                End With
            End If
        End If
        rs.MoveNext
    Wend
End If
rs.Close

Call Get_Details
End Sub

'Calcul de la consommation de l'energie suivant le kilométrage

Public Function Calcule_Consommation(ByVal NbLitre As Long, ByVal Kilometrage As Long) As Double
    
    If Kilometrage <= 0 Then
        Calcule_Consommation = 0
    Else
        Calcule_Consommation = (NbLitre * 100) / Kilometrage
    End If
End Function

'Charger la liste des noms des conducteurs actifs

Public Sub AfficheRow_Conducteur(ByVal VCode As String)

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_CONDUCTEUR(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Libelle")) Then cbo_Conducteur.Text = rs("code") & "  -  " & rs("Libelle")
Else
    MsgBox "Code introuvable!", vbInformation, App.ProductName
    cbo_Conducteur.SetFocus
End If
End Sub

'Calcul des nombres de litre et totals de prix des litres pour tout les véhicules dans ce bon

Public Sub AppCalcul()

Dim ii
Dim Valeur As Double
Dim NBL As Double
Dim VALLIGNE As Double

'Parcourir la liste pour calculer les sommes
For ii = 1 To Lsv_Client.ListItems.Count
    NBL = NBL + (Lsv_Client.ListItems(ii).SubItems(7))  'Litre
    VALLIGNE = Lsv_Client.ListItems(ii).SubItems(9)     'Litre * prixLitre
    Valeur = Valeur + VALLIGNE
Next
txt_Valeur.Text = Format(Valeur, "#,##0.000")

End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = True
Timer1.Interval = 600

If PIC_NFACT.Visible = True Then
    PIC_NFACT.Visible = False
Else
    PIC_NFACT.Visible = True
End If
End Sub

'Afficher La somme des prix et de litres pour chaque type d'énergie dans Grid_Recherche

Public Sub Get_Details()

Dim LOBJ_Energie As Energie
Dim rs As New Recordset
Dim Energi As String
Dim Litre As Double
Dim ttc As Double
Dim ii As Integer

Grid_Recherche.ClearRows
Set LOBJ_Energie = New Energie
Set rs = LOBJ_Energie.Get_Energie(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        Litre = 0
        ttc = 0
        For ii = 1 To Lsv_Client.ListItems.Count
            Energi = Lsv_Client.ListItems(ii).SubItems(4)  'TypeEnergie = libellé
            If Energi = rs.Fields("Libelle") Then
                'Somme des Litre pour chaque type d'énergie à part
                Litre = Litre + Lsv_Client.ListItems(ii).SubItems(7)
                'Somme Litre * prixLitre  pour chaque type d'énergie à part
                ttc = ttc + Lsv_Client.ListItems(ii).SubItems(9)
            End If
        Next
        'Afficher les calculs effectué dans Grid_Recherche
        If Litre <> 0 Then
            With Grid_Recherche
                .AddRow
                .CellDetails .Rows, 1, CDbl(Litre)
                .CellDetails .Rows, 2, rs.Fields("Libelle")
                .CellDetails .Rows, 3, Format(ttc, "#,##0.000")
            End With
        End If
    
    rs.MoveNext
    Wend
End If
rs.Close
End Sub

Private Sub InitialGrid()

With Grid_Recherche
    .Redraw = False
      .AddColumn "Volume", "Volume", eSortType:=CCLSortStringNoCase
      .AddColumn "Eenergie", "Energie"
      .AddColumn "Valeur", "Valeur"
      .AddColumn "Q", ""
      .BackColor = RGB(225, 237, 226)
      .StretchLastColumnToFit = True
      .Redraw = True
    End With
End Sub

Private Sub Cbo_Conducteur_GotFocus()

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
End If
Okayy = True

End Sub
Private Sub Lsv_Client_Click()
    If Lsv_Client.ListItems.Count > 0 Then indexSelect = Lsv_Client.SelectedItem.Index
End Sub

Private Sub cbo_MatriculeStation_Click()
If Len(Trim(cbo_MatriculeStation.Text)) > 0 Then Call AfficheRow_Station(cbo_MatriculeStation.Text)

End Sub

