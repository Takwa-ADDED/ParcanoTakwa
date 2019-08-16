VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "STOOLBOX.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form FrmConsultBV 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Détails bon vidange"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_Numero 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   1920
      TabIndex        =   3
      Tag             =   "M"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.PictureBox PIC_NFACT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5640
      ScaleHeight     =   495
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ce bon est inseré dans une facture N° : "
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
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   3300
      End
      Begin VB.Label LBL_NFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1250"
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
         Left            =   3720
         TabIndex        =   1
         Top             =   120
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   9570
      Left            =   0
      ScaleHeight     =   9570
      ScaleWidth      =   11715
      TabIndex        =   7
      Top             =   1440
      Width           =   11715
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4440
         ScaleHeight     =   375
         ScaleWidth      =   3015
         TabIndex        =   38
         Top             =   480
         Width           =   3015
         Begin SToolBox.SDateBox cda_Create 
            Height          =   285
            Left            =   1680
            TabIndex        =   39
            Tag             =   "M"
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date  :"
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
            Left            =   960
            TabIndex        =   40
            Top             =   120
            Width           =   540
         End
      End
      Begin MSComctlLib.ListView grid 
         Height          =   2175
         Left            =   720
         TabIndex        =   36
         Top             =   5280
         Width           =   9495
         _ExtentX        =   16748
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
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Numero"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Produit"
            Object.Width           =   4657
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "THT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "TVA"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "PrixTTC"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.TextBox txt_MatriculeStation 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1920
         TabIndex        =   34
         Tag             =   "M"
         Top             =   3240
         Width           =   2415
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   240
         ScaleHeight     =   1455
         ScaleWidth      =   4695
         TabIndex        =   27
         Top             =   3720
         Width           =   4695
         Begin VB.TextBox txt_ville 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   30
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txt_adresse 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   29
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txt_rsocial 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   28
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ville  :"
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
            TabIndex        =   33
            Top             =   1200
            Width           =   480
         End
         Begin VB.Label Label13 
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
            TabIndex        =   32
            Top             =   720
            Width           =   780
         End
         Begin VB.Label Label14 
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
            TabIndex        =   31
            Top             =   240
            Width           =   1290
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1815
         ScaleWidth      =   10695
         TabIndex        =   12
         Top             =   1200
         Width           =   10695
         Begin VB.ComboBox Cbo_Conducteur 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   8400
            TabIndex        =   45
            Tag             =   "M"
            Top             =   0
            Width           =   2295
         End
         Begin VB.TextBox txt_compteur 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8400
            TabIndex        =   17
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txt_Type 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   16
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txt_libelle 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   15
            Top             =   0
            Width           =   2775
         End
         Begin VB.TextBox txt_Energie 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   14
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txt_KlmVidange 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   13
            Tag             =   "M"
            Top             =   1440
            Width           =   2775
         End
         Begin SToolBox.SDateBox cda_FinAssur 
            Height          =   285
            Left            =   8400
            TabIndex        =   18
            Top             =   960
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_FinVisite 
            Height          =   285
            Left            =   8400
            TabIndex        =   19
            Top             =   1440
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Label Lbl_cond 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Conducteur :"
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
            Left            =   6720
            TabIndex        =   46
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date fin assurance :"
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
            Left            =   6720
            TabIndex        =   26
            Top             =   960
            Width           =   1665
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date fin visite :"
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
            Left            =   6720
            TabIndex        =   25
            Top             =   1440
            Width           =   1260
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type :"
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
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Width           =   510
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compteur :"
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
            Left            =   6720
            TabIndex        =   23
            Top             =   480
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Energie :"
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
            Left            =   240
            TabIndex        =   22
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Matricule :"
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
            Left            =   240
            TabIndex        =   21
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NB KM Vidange :"
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
            Left            =   240
            TabIndex        =   20
            Top             =   1560
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6600
         ScaleHeight     =   375
         ScaleWidth      =   3015
         TabIndex        =   9
         Top             =   4200
         Width           =   3015
         Begin VB.TextBox txt_Valeur 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            TabIndex        =   10
            Tag             =   "M"
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valeur :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   11
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.TextBox txt_Matricule 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1920
         TabIndex        =   8
         Tag             =   "M"
         Top             =   600
         Width           =   2295
      End
      Begin SToolBox.SDateBox dateOp 
         Height          =   285
         Left            =   9120
         TabIndex        =   41
         Tag             =   "M"
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Text            =   ""
         Enabled         =   0   'False
         BackColor       =   14737632
      End
      Begin VB.Label Lbl_User 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   6960
         TabIndex        =   48
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bon de vidange saisi par :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   47
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Lbl_Stat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Station :"
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
         Left            =   240
         TabIndex        =   44
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Numéro bon"
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
         Left            =   240
         TabIndex        =   43
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Operation:"
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
         Left            =   7680
         TabIndex        =   42
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Immatriculation"
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
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bon de sortie vidange"
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
      Left            =   600
      TabIndex        =   37
      Top             =   480
      Width           =   3540
   End
   Begin VB.Image PicBox_Header 
      Height          =   1455
      Left            =   0
      Picture         =   "FrmConsultBV.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14415
   End
   Begin VB.Label Label15 
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
      Left            =   840
      TabIndex        =   6
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conducteur"
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
      Height          =   240
      Left            =   420
      TabIndex        =   5
      Top             =   3960
      Width           =   1125
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type Vidange"
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
      Height          =   240
      Left            =   255
      TabIndex        =   4
      Top             =   6000
      Width           =   1305
   End
End
Attribute VB_Name = "FrmConsultBV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Afficher assiette bon de vidange
Public Sub AfficheRow(ByVal VCode As String)

Dim LOBJ_Bv As BonVidange
Dim rs As New Recordset

Call ViderZone(FrmConsultBV)
grid.ListItems.Clear
Set LOBJ_Bv = New BonVidange
Set rs = LOBJ_Bv.Get_BV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    txt_Numero.Text = rs("Numero")
    If Not IsNull(rs("VEHICULE")) Then txt_Matricule.Text = rs("VEHICULE")
    If Not IsNull(rs("DATEDOC")) Then cda_Create.Text = rs("DATEDOC")
    If Not IsNull(rs("dateop")) Then dateOp.Text = rs("dateop")
    If Not IsNull(rs("VALEUR")) Then txt_Valeur.Text = Format(rs("VALEUR"), "#,##0.000")
    If Not IsNull(rs("NBKLMvid")) Then txt_KlmVidange.Text = rs("NBKLMvid")
    If Not IsNull(rs("UserInsert")) Then Lbl_user.Caption = Get_NameUserByCode(rs("UserInsert"))
    
    Call AfficheRow_Vehicule(rs("VEHICULE"))
    Call AfficheRow_Station(rs("STATION"))
    Call AfficheRow_Conducteur(rs("CONDUCTEUR"))
    Call AfficheRow_Lubrifiant_BV(txt_Numero.Text)
    
    If rs("Transf") = "O" Then
        LBL_NFact.Caption = rs("NumFact")
        PIC_NFACT.Visible = True
  
    Else
        PIC_NFACT.Visible = False
    End If
End If
rs.Close

End Sub

'Afficher détails bon de vidange
Public Sub AfficheRow_Lubrifiant_BV(ByVal VCode As String)

Dim LOBJ_DetBV As BonVidange
Dim rs As New Recordset
Dim som As Double

som = 0
Set LOBJ_DetBV = New BonVidange
Set rs = LOBJ_DetBV.Get_Lub_BV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
            Set itmX = grid.ListItems.Add(, , CStr(rs("Numero")))
            itmX.SubItems(1) = CStr(rs("Libelle"))
            itmX.SubItems(2) = CStr(Format(rs("THT"), "#,##0.000"))
            itmX.SubItems(3) = CStr(Format(rs("TVA"), "#,##0.000"))
            itmX.SubItems(4) = CStr(Format(rs("prixTTC"), "#,##0.000"))
        If Not IsNull(rs("prixTTC")) Then
        som = som + CDbl(rs("prixTTC"))
        End If
        rs.MoveNext
    Wend
End If
rs.Close
txt_Valeur.Text = Format(som, "#,##0.000")

End Sub

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
    txt_MatriculeStation.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
Else
    MsgBox "Code introuvable", vbInformation
    txt_MatriculeStation.SetFocus
    Exit Sub
End If
rs.Close
End Sub

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
    If Not IsNull(rs("Libelle")) Then cbo_conducteur.Text = rs("Libelle")
End If
rs.Close

End Sub

Public Sub AfficheRow_Vehicule(ByVal VCode As String)

Dim LOBJ_Vehi As VEHICULE
Dim rs As New Recordset

Set LOBJ_Vehi = New VEHICULE
Set rs = LOBJ_Vehi.GetVehiculeByCode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    If Not IsNull(rs("Matricule")) Then txt_libelle.Text = rs("Matricule")
    If Not IsNull(rs("marque")) Then txt_Type.Text = rs("TYPE")
    If Not IsNull(rs("Energie")) Then txt_Energie.Text = rs("Energie")
    If Not IsNull(rs("CompteurVidange")) Then txt_compteur.Text = rs("CompteurVidange")
    If Not IsNull(rs("DAteFinAssur")) Then cda_FinAssur.Text = rs("DAteFinAssur")
    If Not IsNull(rs("DAteFinVisite")) Then cda_FinVisite.Text = rs("DAteFinVisite")
Else
    MsgBox "Code introuvable", vbInformation
    Exit Sub
End If
rs.Close

End Sub

Private Sub Form_Load()
PicBox_Header.Width = Me.Width
Me.Height = 9210
Me.Width = 11715
Me.Move 0, 0
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'
'On Error GoTo erreur
'   Dim i As Integer
'   Dim MSG ' Déclare la variable.
'   ' Définit le texte du message.
'   MSG = "Voulez-vous vraiment quitter?"
'   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
'   If MsgBox(MSG, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'      Cancel = True
'   Else
'   Unload Me
'   End If
'
'   Exit Sub
'erreur:
'   MsgBox Err.Description, 48
'End Sub

