VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPieceReparation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Pièce de Reparation"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12720
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10395
   ScaleWidth      =   12720
   Begin MSComctlLib.ListView Lsv_Detail 
      Height          =   3615
      Left            =   120
      TabIndex        =   63
      Top             =   6600
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6376
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Désignation"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Qte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Véhicule"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "PU.HT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Remise(%)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tot.HT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "TVA(%)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "TOT.TTC.Ligne"
         Object.Width           =   2540
      EndProperty
   End
   Begin SToolBox.SCommand Cmd_FindTypP 
      Height          =   375
      Left            =   3360
      TabIndex        =   43
      Top             =   960
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
      Picture         =   "FrmPieceReparation.frx":0000
      ButtonType      =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   12615
      TabIndex        =   24
      Top             =   960
      Width           =   12615
      Begin MSComctlLib.ListView Lsv_Toto 
         Height          =   1815
         Left            =   5400
         TabIndex        =   64
         Top             =   3480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3201
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "a"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "T.HT.Brut"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "T.Ren.Ligne"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "T.Ren.Piece"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "T.HT.Net"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "T.TVA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "T.TTC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox Pict_TRP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   7200
         ScaleHeight     =   2025
         ScaleWidth      =   4305
         TabIndex        =   58
         Top             =   840
         Width           =   4335
         Begin VB.TextBox Txt_TvaMO 
            Height          =   285
            Left            =   2520
            TabIndex        =   13
            Text            =   "0"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txt_Timbre 
            Height          =   285
            Left            =   2520
            TabIndex        =   15
            Text            =   "0"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox Tex_RSP 
            Height          =   285
            Left            =   2520
            TabIndex        =   14
            Text            =   "0"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox Txt_PMainOeuvre 
            Height          =   285
            Left            =   2520
            TabIndex        =   12
            Text            =   "0"
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Lbl_TvaMO 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TVA M.Oeuvre"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   600
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
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   1560
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
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label Lbl_MOeuvre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Prix main d'oeuvre :"
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
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.PictureBox Pict_Creat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   3135
         TabIndex        =   57
         Top             =   600
         Width           =   3135
         Begin SToolBox.SOptionButton op_creat 
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   2
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "Transfert"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SToolBox.SOptionButton op_creat 
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            BackStyle       =   0
            Value           =   1
            Caption         =   "Création"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox Pict_Transf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3735
         TabIndex        =   56
         Top             =   1560
         Width           =   3735
         Begin VB.OptionButton Opt_Fact 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Facture"
            Height          =   255
            Left            =   1920
            TabIndex        =   6
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton Opt_PRecep 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pièce réception"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.PictureBox Pict_Type 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   5055
         TabIndex        =   54
         Top             =   1440
         Width           =   5055
         Begin VB.ComboBox cbo_typePiece 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            ItemData        =   "FrmPieceReparation.frx":0353
            Left            =   1560
            List            =   "FrmPieceReparation.frx":0363
            TabIndex        =   4
            Top             =   0
            Width           =   2775
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
            TabIndex        =   55
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.PictureBox Pict_stat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1680
         ScaleHeight     =   375
         ScaleWidth      =   3375
         TabIndex        =   52
         Top             =   3480
         Width           =   3375
         Begin VB.ComboBox cbo_MatriculeStation 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2775
         End
         Begin SToolBox.SCommand CmdFindStation 
            Height          =   375
            Left            =   2880
            TabIndex        =   53
            Top             =   0
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
            Picture         =   "FrmPieceReparation.frx":0394
            ButtonType      =   1
         End
      End
      Begin VB.PictureBox Pict_BCRep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   5175
         TabIndex        =   48
         Top             =   1080
         Width           =   5175
         Begin VB.PictureBox Pict_txtBCR 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1560
            ScaleHeight     =   375
            ScaleWidth      =   2775
            TabIndex        =   51
            Top             =   0
            Width           =   2775
            Begin VB.TextBox txt_BCReparation 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   375
               Left            =   0
               TabIndex        =   3
               Top             =   0
               Width           =   2775
            End
         End
         Begin SToolBox.SCommand Cmd_FinBCRep 
            Height          =   375
            Left            =   4440
            TabIndex        =   49
            Top             =   0
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
            Picture         =   "FrmPieceReparation.frx":06E7
            ButtonType      =   1
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
            TabIndex        =   50
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.PictureBox Pict_User 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3840
         ScaleHeight     =   495
         ScaleWidth      =   3735
         TabIndex        =   45
         Top             =   120
         Width           =   3735
         Begin VB.Label Lbl_UserSaisi 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1800
            TabIndex        =   47
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label Lbl_user 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pièce saisie par :"
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
            Left            =   120
            TabIndex        =   46
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.PictureBox Pict_date 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3480
         ScaleHeight     =   375
         ScaleWidth      =   3615
         TabIndex        =   36
         Top             =   2280
         Width           =   3615
         Begin MSComCtl2.DTPicker cda_Operation 
            Height          =   375
            Left            =   1800
            TabIndex        =   7
            Top             =   0
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   160301057
            CurrentDate     =   42877
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.TextBox txt_Numero 
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
         ForeColor       =   &H000040C0&
         Height          =   435
         Left            =   1680
         TabIndex        =   0
         Tag             =   "M"
         Top             =   0
         Width           =   1575
      End
      Begin VB.TextBox txt_ref 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   2880
         Width           =   2775
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   4695
         TabIndex        =   28
         Top             =   3960
         Width           =   4695
         Begin VB.TextBox txt_ville 
            Height          =   315
            Left            =   1560
            TabIndex        =   31
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txt_adresse 
            Height          =   315
            Left            =   1560
            TabIndex        =   30
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txt_rsocial 
            Height          =   315
            Left            =   1560
            TabIndex        =   29
            Top             =   0
            Width           =   2775
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
            Left            =   0
            TabIndex        =   34
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
            Left            =   0
            TabIndex        =   33
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
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   1380
         End
      End
      Begin VB.PictureBox PIC_NFACT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   7800
         ScaleHeight     =   495
         ScaleWidth      =   5055
         TabIndex        =   25
         Top             =   0
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
            Left            =   120
            TabIndex        =   27
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
            TabIndex        =   26
            Top             =   120
            Width           =   720
         End
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   1230
      End
      Begin VB.Label cda_Create 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1440
         TabIndex        =   35
         Top             =   2280
         Width           =   1695
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
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAUX EN (DT)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   7920
         TabIndex        =   41
         Top             =   3000
         Width           =   2895
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
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Width           =   1575
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
         Left            =   120
         TabIndex        =   39
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label NumBC 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   4560
         TabIndex        =   38
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3615
      Left            =   11640
      ScaleHeight     =   3615
      ScaleWidth      =   615
      TabIndex        =   17
      Top             =   6480
      Width           =   615
      Begin SToolBox.SCommand Cmd_SuppL 
         Height          =   495
         Left            =   120
         TabIndex        =   18
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
         Picture         =   "FrmPieceReparation.frx":0A3A
      End
      Begin SToolBox.SCommand Cmd_MdfL 
         Height          =   495
         Left            =   120
         TabIndex        =   19
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
         Picture         =   "FrmPieceReparation.frx":0BBC
      End
      Begin SToolBox.SCommand Cmd_SaisiL 
         Height          =   495
         Left            =   120
         TabIndex        =   10
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
         Picture         =   "FrmPieceReparation.frx":0F0F
      End
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   11400
      TabIndex        =   16
      Top             =   240
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
      Picture         =   "FrmPieceReparation.frx":1091
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   10440
      TabIndex        =   20
      Top             =   240
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
      Picture         =   "FrmPieceReparation.frx":1213
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   10920
      TabIndex        =   21
      Top             =   240
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
      Picture         =   "FrmPieceReparation.frx":1566
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   9960
      TabIndex        =   22
      Top             =   240
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
      Picture         =   "FrmPieceReparation.frx":18B9
   End
   Begin SToolBox.SCommand CmdPrint 
      Height          =   495
      Left            =   11880
      TabIndex        =   23
      Top             =   240
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
      Picture         =   "FrmPieceReparation.frx":1A3B
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pièce de Réparation"
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
      TabIndex        =   44
      Top             =   360
      Width           =   3225
   End
   Begin VB.Image PicBox_Header 
      Height          =   1215
      Left            =   0
      Picture         =   "FrmPieceReparation.frx":1D8E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "FrmPieceReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Okayy As Boolean
Dim itmX As ListItem
Dim thekey As Integer
Dim theshift As Integer

Private Sub cbo_MatriculeStation_LostFocus()
 Call ExistData(cbo_MatriculeStation)
End Sub

Private Sub Cmd_SaisiL_GotFocus()
Dim KeyCode As Integer
If KeyCode = vbKeyReturn Then Call Cmd_SaisiL_Click
End Sub

Private Sub CmdAdd_Click()

On Error GoTo Err

Dim LOBJ_Personnel As personnel

Set LOBJ_Personnel = New personnel
' Vérifier les droits d'accès de l'utilisateur : s'il a le droit d'ajouter un nouveau bon.
If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Ins_PR", LInt_UserId) Then
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If
Okayy = False

If Lsv_Detail.ListItems.Count > 0 Or txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Call ViderZone(FrmPieceReparation)
Pict_Creat.Visible = True
PIC_NFACT.Visible = False
Timer1.Enabled = False
op_creat(0).Value = vbChecked
    Pict_BCRep.Visible = True
        Opt_PRecep.Value = False
        Opt_Fact.Value = False
Pict_Type.Visible = False
Pict_Transf.Visible = False
Cmd_FindTypP.Visible = False
Picture2.Enabled = False
Lbl_UserSaisi.Caption = LStr_NameUser
txt_Numero.Text = "Auto"
CmdDelete.Enabled = True
CmdSave.Enabled = True
CmdPrint.Enabled = True
cda_Create.Caption = Date
cda_Operation.Value = Date
Tex_RSP.Text = 0
txt_Timbre.Text = 0
Txt_PMainOeuvre.Text = 0
Lsv_Detail.ListItems.Clear
Lsv_Detail.Enabled = False
Lsv_Toto.ListItems.Clear

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()

Dim VCode As String
Dim LOBJ_Personnel As personnel
Dim LOBJ_PRepar As PieceReparation

On Error GoTo Err

If PIC_NFACT.Visible = True Then
    MsgBox "Maj impossible", vbInformation
    Exit Sub
End If

If txt_Numero.Text = "" Then
    MsgBox "Aucune pièce de réception n'est choisit pour la suppression", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If

If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        txt_Numero.SetFocus
        Exit Sub
    End If
End If
Set LOBJ_Personnel = New personnel
If txt_Numero.Text <> "Auto" Then
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "MAJ_PR", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
End If
    
Set LOBJ_PRepar = New PieceReparation
If MsgBox("Confirmez vous supprimer cette " & vbNewLine & " pièce de réparation ", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub

VCode = txt_Numero.Text
Call LOBJ_PRepar.Delete_DetailPRepaBySup(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

Call LOBJ_PRepar.Delete_AssPRep(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode, LInt_UserId)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Call ViderZone(FrmPieceReparation)
Lsv_Detail.ListItems.Clear
Lsv_Toto.ListItems.Clear
Pict_Creat.Visible = False
Pict_Type.Visible = True
Pict_stat.Enabled = False
Picture2.Enabled = False
Pict_BCRep.Visible = False
Pict_Transf.Visible = False
Pict_TRP.Enabled = False
Picture1.Enabled = True
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdFind_Click()

On Error GoTo Err

If Lsv_Detail.ListItems.Count > 0 Or txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
ElseIf Okayy = True Then
    If MsgBox("Annuler le maj en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If
Pict_stat.Enabled = False
Picture2.Enabled = False
PIC_NFACT.Visible = False
Pict_Creat.Visible = False
Pict_Transf.Visible = False
Pict_BCRep.Visible = False
Pict_Type.Visible = True
Lsv_Detail.Enabled = False
cbo_typePiece.Text = cbo_typePiece.List(0)
Picture1.Enabled = True
Cmd_FindTypP.Visible = True

If cbo_typePiece.Text = "Piece Reception" Then
    Unload FrmFind
    With FrmFind
        .StrSource = "BLPieceReparation"
        .Show vbModal
    End With
ElseIf cbo_typePiece.Text = "Facture" Then
    Unload FrmFind
    With FrmFind
        .StrSource = "FacturePieceReparation"
        .Show vbModal
    End With
Else
    Unload FrmFind
    With FrmFind
        .StrSource = "AllPieceReparation"
        .Show vbModal
    End With
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdFindStation_Click()

On Error GoTo Err
If txt_Numero.Text <> "" Then
    Unload FrmFind_Fils
    With FrmFind_Fils
        .StrSource = "Station PR"
        .Show vbModal
    End With
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdPrint_Click()

Dim TotHTBrut As Double
Dim TotRemLigne As Double
Dim TotHtNet As Double
Dim TotTva  As Double
Dim RemiseP As Double
Dim TotTTC As Double
Dim MainOeuvre As Double
Dim TvaMOeuvre As Double

If txt_Numero.Text = "" Then Exit Sub
If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        txt_Numero.SetFocus
        Exit Sub
    End If
End If

If MsgBox("Imprimer ce bon de pièce de réparation ? ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    TotHTBrut = Lsv_Toto.ListItems(1).ListSubItems(1)
    TotRemLigne = Lsv_Toto.ListItems(1).ListSubItems(2)
    RemiseP = Lsv_Toto.ListItems(1).ListSubItems(3)
    TotHtNet = Lsv_Toto.ListItems(1).ListSubItems(4)
    TotTva = Lsv_Toto.ListItems(1).ListSubItems(5)
    TotTTC = Lsv_Toto.ListItems(1).ListSubItems(6)
    MainOeuvre = CDbl(Txt_PMainOeuvre.Text)
    TvaMOeuvre = CDbl(Txt_TvaMO.Text)

    Dim F As Form
    Set F = New Frm_Rpt_Apercus
    With F
        .Numero = txt_Numero.Text
        Call .PrintOutAndApercu_PieceRepa(0, TotHTBrut, TotRemLigne, RemiseP, TotHtNet, TotTva, TotTTC, MainOeuvre, TvaMOeuvre)
        .Show
    End With
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdSave_Click()

Dim LOBJ_Personnel As personnel
Dim LOBJ_PieceRep As PieceReparation
Dim ii As Long
Dim VCode
Dim TotHTBrut As Double
Dim TotRemLigne As Double
Dim TotHtNet As Double
Dim TotTva  As Double
Dim RemiseP As Double
Dim TotTTC As Double
Dim MainOeuvre As Double
Dim TvaMOeuvre As Double

If Left(CheckMandatory(FrmPieceReparation), 1) = 1 Then
   Exit Sub
End If

If Lsv_Detail.ListItems.Count = 0 Then
    MsgBox "Veuillez saisir les details de la reparation", vbInformation
    Exit Sub
End If

If PIC_NFACT.Visible = True Then
    MsgBox "Pièce facturée , MAJ impossible ", vbInformation
    Exit Sub
End If

For ii = 1 To Lsv_Detail.ListItems.Count
    If (Lsv_Detail.ListItems(ii).SubItems(8) = 0 Or Lsv_Detail.ListItems(ii).SubItems(8) = "") Then
        MsgBox "Veuillez vérifier les details saisis ", vbInformation
        Exit Sub
        Exit For
    End If
Next

If CDate(cda_Operation.Value) > Date Then
    MsgBox "Vérifier la date d'opération", vbInformation
    cda_Operation.SetFocus
    Exit Sub
End If
'Vérifier type de la pièce à créer ou transférer
If (op_creat(0).Value = vbChecked) Then  'transfert
   If Opt_PRecep.Value = False And Opt_Fact.Value = False Then
        MsgBox " Vous devez choisir type de la pièce", vbInformation
        Exit Sub
    End If
ElseIf (op_creat(1).Value = vbChecked) Then 'création
    If cbo_typePiece.Text = "" Then
        MsgBox " Vous devez choisir type de la pièce", vbInformation
        Exit Sub
    End If
End If
'On Error GoTo Err

If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
VCode = txt_Numero.Text
If VCode = "Auto" Or VCode = "" Then
    Call AjoutPR
End If

If VCode <> "Auto" And VCode <> "" Then
    Set LOBJ_Personnel = New personnel
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "MAJ_PR", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
    Call ModifierPR
End If
'MAJ Bon de commande transféré
Set LOBJ_PieceRep = New PieceReparation
If Not (IsNull(NumBC.Caption)) Then
    Call LOBJ_PieceRep.Update_Trans(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt_Numero.Text, NumBC.Caption)
        If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
End If

Okayy = False
    
If MsgBox("Enregistrement terminé avec succé  " & vbNewLine & " Imprimer ce bon        ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Dim F As Form
    TotHTBrut = Lsv_Toto.ListItems(1).ListSubItems(1)
    TotRemLigne = Lsv_Toto.ListItems(1).ListSubItems(2)
    RemiseP = Lsv_Toto.ListItems(1).ListSubItems(3)
    TotHtNet = Lsv_Toto.ListItems(1).ListSubItems(4)
    TotTva = Lsv_Toto.ListItems(1).ListSubItems(5)
    TotTTC = Lsv_Toto.ListItems(1).ListSubItems(6)
    MainOeuvre = CDbl(Txt_PMainOeuvre.Text)
    TvaMOeuvre = CDbl(Txt_TvaMO.Text)
    
    Set F = New Frm_Rpt_Apercus
    With F
        .Numero = txt_Numero.Text
        Call .PrintOutAndApercu_PieceRepa(0, TotHTBrut, TotRemLigne, RemiseP, TotHtNet, TotTva, TotTTC, MainOeuvre, TvaMOeuvre)
        .Show
    End With
End If

Lsv_Detail.Enabled = False
Cmd_FindTypP.Visible = True
Pict_Creat.Visible = False
Pict_Type.Visible = True
Pict_stat.Enabled = False
PIC_NFACT.Visible = False
Picture1.Enabled = True
Picture2.Enabled = False
Pict_BCRep.Visible = False
Pict_Transf.Visible = False
Pict_TRP.Enabled = False
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub AjoutPR()

Dim LOBJ_PieceRepar As PieceReparation
Dim LRs_NewRecord As New Recordset
Dim LInt_NumCompteur As Long
Dim ttc As Double

LInt_NumCompteur = return_Compteur() + 1
'Insertion enregistrement assiette
txt_Numero.Text = Format(LInt_NumCompteur, "00000")
ttc = CDbl(Lsv_Toto.ListItems(1).SubItems(6))
Set LOBJ_PieceRepar = New PieceReparation
Set LRs_NewRecord = CreateEmptyRS_AssPRepar()
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = txt_Numero.Text
    If (op_creat(1).Value = vbChecked) Or Pict_Type.Visible = True Then
        .Fields("Type") = cbo_typePiece.Text
    ElseIf (op_creat(0).Value = vbChecked) Or Pict_Transf.Visible = True Then
        If Opt_PRecep.Value = True Then
            .Fields("Type") = "Piece Reception"
        ElseIf Opt_Fact.Value = True Then
            .Fields("Type") = "Facture"
        End If
    End If
    .Fields("DatePiece") = CDate(cda_Create.Caption)
    .Fields("RemisePiece") = CDbl(Replace(Tex_RSP.Text, ".", ","))
    .Fields("TotTTC") = CDbl(Replace(ttc, ".", ","))
    .Fields("DateOperation") = CDate(cda_Operation.Value)
    .Fields("refPiece") = txt_ref.Text
    .Fields("Timbre") = CDbl(Replace(txt_Timbre.Text, ".", ","))
    .Fields("Fournisseur") = cbo_MatriculeStation.Text
    .Fields("PrixMOeuvre") = CDbl(Replace(Txt_PMainOeuvre.Text, ".", ","))
    .Fields("TVA_MOeuvre") = CDbl(Replace(Txt_TvaMO.Text, ".", ","))
    .Fields("UserInsert") = LInt_UserId
End With
Call LOBJ_PieceRepar.Insert_AssPieceRepar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

Call insert_DetPieceRep

End Sub

Private Sub insert_DetPieceRep()

Dim LOBJ_PieceRepar As PieceReparation
Dim LRs_NewRecord As New Recordset
Dim ii As Integer

Set LOBJ_PieceRepar = New PieceReparation
Set LRs_NewRecord = CreateEmptyRS_DetPRepar()

For ii = 1 To Lsv_Detail.ListItems.Count
    With LRs_NewRecord
        .AddNew
        .Fields("Numero") = txt_Numero.Text
        .Fields("Designation") = Lsv_Detail.ListItems(ii).SubItems(1)
        .Fields("Qte") = Val(Lsv_Detail.ListItems(ii).SubItems(2))
        .Fields("Vehicule") = Lsv_Detail.ListItems(ii).SubItems(3)
        .Fields("PUHT") = CDbl(Lsv_Detail.ListItems(ii).SubItems(4))
        .Fields("Remise") = CDbl(Lsv_Detail.ListItems(ii).SubItems(5))
        .Fields("TVA") = CDbl(Lsv_Detail.ListItems(ii).SubItems(7))
    End With
    Lsv_Detail.ListItems(ii).Text = txt_Numero.Text
Next
Call LOBJ_PieceRepar.Insert_DetPieceRepar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing
End Sub

'Modification de la pièce de réparation
Private Sub ModifierPR()

Dim LOBJ_PieceRepar As PieceReparation
Dim LRs_NewRecord As New Recordset
Dim ttc As Double

ttc = CDbl(Lsv_Toto.ListItems(1).SubItems(6))
Set LOBJ_PieceRepar = New PieceReparation
Set LRs_NewRecord = CreateEmptyRS_AssPRepar()
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = txt_Numero.Text
    If (op_creat(1).Value = vbChecked) Then
        .Fields("Type") = cbo_typePiece.Text
    ElseIf (op_creat(0).Value = vbChecked) Then
        If Opt_PRecep.Value = True Then
            .Fields("Type") = "Piece Reception"
        ElseIf Opt_Fact.Value = True Then
            .Fields("Type") = "Facture"
        End If
    End If
    .Fields("DatePiece") = CDate(cda_Create.Caption)
    .Fields("RemisePiece") = CDbl(Replace(Tex_RSP.Text, ".", ","))
    .Fields("TotTTC") = CDbl(Replace(ttc, ".", ","))
    .Fields("DateOperation") = CDate(cda_Operation.Value)
    .Fields("refPiece") = txt_ref.Text
    .Fields("Timbre") = CDbl(Replace(txt_Timbre.Text, ".", ","))
    .Fields("Fournisseur") = cbo_MatriculeStation.Text
    .Fields("PrixMOeuvre") = CDbl(Replace(Txt_PMainOeuvre.Text, ".", ","))
    .Fields("TVA_MOeuvre") = CDbl(Replace(Txt_TvaMO.Text, ".", ","))
    .Fields("UserUpdate") = LInt_UserId
End With
Call LOBJ_PieceRepar.Update_PieceRep(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

Call LOBJ_PieceRepar.Delete_DetailPRepa(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt_Numero.Text)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

Call insert_DetPieceRep

End Sub

Private Sub Form_Load()

On Error GoTo Err
CmdDelete.Enabled = False
CmdSave.Enabled = False
CmdPrint.Enabled = False
cda_Create.Caption = Date
cda_Operation.Value = Date
Me.WindowState = 2
Call Affiche_StatRep_Combo(cbo_MatriculeStation)
cbo_typePiece.Text = cbo_typePiece.List(0)
Pict_BCRep.Visible = False
Pict_Transf.Visible = False
Pict_Creat.Visible = False
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
'   Unload Me
'   End If
'
'   Exit Sub
'erreur:
'   MsgBox Err.Description, 48

End Sub

'Afficher les détails de la pièce dans FrmSaisiePieceReparation afin de les modifier
Private Sub Lsv_Detail_DblClick()

Dim i
Dim ii
Dim LOBJ_Prod As Produit_Lubrifiant
Dim rs As New Recordset
Dim vprix
Dim vtva
On Error GoTo Err

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If
i = Lsv_Detail.SelectedItem.Index
Set LOBJ_Prod = New Produit_Lubrifiant
If Lsv_Detail.ListItems(i).SubItems(4) = "0" Or Lsv_Detail.ListItems(i).SubItems(4) = "" Then
    Set rs = LOBJ_Prod.Get_ProdLubByLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lsv_Detail.ListItems(i).SubItems(1))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs.EOF Then
        vprix = rs("prixht")
        vtva = rs("tva")
    End If
    rs.Close
Else
    vprix = Lsv_Detail.ListItems(i).SubItems(4)
    vtva = Lsv_Detail.ListItems(i).SubItems(7)
End If

With FrmSaisiePieceReparation
    .Okay = False
    .ii = Lsv_Detail.SelectedItem.Index
    .txt_Numero.Text = txt_Numero.Text
    .Txt_Designation.Text = Lsv_Detail.ListItems(i).SubItems(1)
    .cbo_Matricule.Text = Lsv_Detail.ListItems(i).SubItems(3)
    .txt_Qte.Text = Lsv_Detail.ListItems(i).SubItems(2)
    .txt_PUHT.Text = vprix 'Lsv_Detail.ListItems(i).SubItems(4)
    .Txt_Remise.Text = Lsv_Detail.ListItems(i).SubItems(5)
    .txt_TotHT.Text = Lsv_Detail.ListItems(i).SubItems(6)
    .Txt_tva.Text = vtva 'Lsv_Detail.ListItems(i).SubItems(7)
    .txt_ttc.Text = Lsv_Detail.ListItems(i).SubItems(8)
    .Show vbModal
    
End With

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

'Boutons radio Création et transfert
Private Sub op_creat_Changed(Index As Integer, Value As CheckBoxConstants)
 
    If (op_creat(0).Value = vbChecked) Then
        Pict_BCRep.Visible = True
        Opt_PRecep.Value = False
        Opt_Fact.Value = False
        Pict_Type.Visible = False
        Pict_stat.Enabled = False
    ElseIf (op_creat(1).Value = vbChecked) Then
        Pict_BCRep.Visible = False
        Pict_stat.Enabled = True
        Pict_Transf.Visible = False
        Pict_Type.Visible = True
        cbo_typePiece.Text = cbo_typePiece.List(0)
        Picture2.Enabled = True
        Pict_stat.Enabled = True
    End If
End Sub

'en cliquant sur l'un des boutons radio Création et transfert
Private Sub op_creat_Click(Index As Integer)

    If (op_creat(0).Value = vbChecked) Then
        Pict_BCRep.Visible = True
        Opt_PRecep.Value = False
        Opt_Fact.Value = False
        Pict_Type.Visible = False
        Pict_stat.Enabled = False
    ElseIf (op_creat(1).Value = vbChecked) Then
        Pict_BCRep.Visible = False
        Pict_Transf.Visible = False
        Pict_Type.Visible = True
        cbo_typePiece.Text = cbo_typePiece.List(0)
        Picture2.Enabled = True
        Pict_stat.Enabled = True
    End If
End Sub

Private Sub op_creat_GotFocus(Index As Integer)

    If (op_creat(0).Value = vbChecked) Then
        Pict_BCRep.Visible = True
        Opt_PRecep.Value = False
        Opt_Fact.Value = False
        Pict_Type.Visible = False
        Pict_stat.Enabled = False
    ElseIf (op_creat(1).Value = vbChecked) Then
        Pict_BCRep.Visible = False
        Pict_Transf.Visible = False
        Pict_Type.Visible = True
        cbo_typePiece.Text = cbo_typePiece.List(0)
        Picture2.Enabled = True
        Pict_stat.Enabled = True
    End If
End Sub

'Afficher liste des pièce (par Type pièce du cbo_typePiece)
Private Sub Cmd_FindTypP_Click()

On Error GoTo Err
If Lsv_Detail.ListItems.Count > 0 And txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

ElseIf Okayy = True Then
    If MsgBox("Annuler le maj en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
End If
Pict_Creat.Visible = False
Pict_stat.Enabled = False
Picture2.Enabled = False
PIC_NFACT.Visible = False
Pict_Transf.Visible = False
Pict_BCRep.Visible = False
Picture1.Enabled = True
Cmd_FindTypP.Visible = True
Pict_Type.Visible = True
Lsv_Detail.ListItems.Clear
Lsv_Detail.Enabled = False
Lsv_Toto.ListItems.Clear
Lbl_UserSaisi.Caption = ""
txt_Numero.Text = ""

 'Afficher Piece Reception
If cbo_typePiece.Text = "Piece Reception" Then
    Unload FrmFind
    With FrmFind
        .StrSource = "BLPieceReparation"
        .Show vbModal
    End With
End If

'Afficher Facture
If cbo_typePiece.Text = "Facture" Then
    Unload FrmFind
    With FrmFind
        .StrSource = "FacturePieceReparation"
        .Show vbModal
    End With
End If

'Avoir
If cbo_typePiece.Text = "Avoir" Then
    Unload FrmFind
    With FrmFind
        .StrSource = "AvoirPieceReparation"
        .Show vbModal
    End With
End If

'Bon de Retour
If cbo_typePiece.Text = "Bon Retour" Then
    Unload FrmFind
    With FrmFind
        .StrSource = "BRPieceReparation"
        .Show vbModal
    End With
End If
If Pict_Type.Visible = False Or cbo_typePiece.Text = "" Then
    Unload FrmFind
    With FrmFind
        .StrSource = "AllPieceReparation"
        .Show vbModal
    End With
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

'suppression d'une ligne de la liste des détails
Private Sub Cmd_SuppL_Click()

Dim i As Integer
On Error GoTo Err
If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If

If Lsv_Detail.ListItems.Count <= 0 Then Exit Sub
Okayy = True
If MsgBox("Confirmez vous la suppression de la ligne selectionné.?", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    i = Lsv_Detail.SelectedItem.Index
    Lsv_Detail.ListItems.Remove i
    Call AppCalcul
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

'Modification de la ligne selectionné en affichant les détails dans FrmSaisiePieceReparation
Private Sub Cmd_MdfL_Click()

Dim i
Dim ii
Dim LOBJ_Prod As Produit_Lubrifiant
Dim rs As New Recordset
Dim vprix As String
Dim vtva As String

On Error GoTo Err

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If

If Lsv_Detail.ListItems.Count <= 0 Then Exit Sub
If Lsv_Detail.SelectedItem.Selected = 0 Then
    MsgBox "Selectionner le detail de la pièce à modifier ", vbInformation
    Exit Sub
End If

i = Lsv_Detail.SelectedItem.Index
Set LOBJ_Prod = New Produit_Lubrifiant
If Lsv_Detail.ListItems(i).SubItems(4) = 0 Or Lsv_Detail.ListItems(i).SubItems(4) = "" Then
    Set rs = LOBJ_Prod.Get_ProdLubByLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lsv_Detail.ListItems(i).SubItems(1))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs.EOF Then
        vprix = rs("prixht")
        vtva = rs("tva")
    End If
    rs.Close
Else
    vprix = Lsv_Detail.ListItems(i).SubItems(4)
    vtva = Lsv_Detail.ListItems(i).SubItems(7)
End If

Okayy = True
With FrmSaisiePieceReparation
    .Okay = False
    .ii = Lsv_Detail.SelectedItem.Index
    .txt_Numero.Text = txt_Numero.Text
    .Txt_Designation.Text = Lsv_Detail.ListItems(i).SubItems(1)
    .cbo_Matricule.Text = Lsv_Detail.ListItems(i).SubItems(3)
    .txt_Qte.Text = Lsv_Detail.ListItems(i).SubItems(2)
    .txt_PUHT.Text = vprix
    .Txt_Remise.Text = Lsv_Detail.ListItems(i).SubItems(5)
    .txt_TotHT.Text = Lsv_Detail.ListItems(i).SubItems(6)
    .Txt_tva.Text = vtva
    .txt_ttc.Text = Lsv_Detail.ListItems(i).SubItems(8)
    .Show vbModal
End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

'Saisir une nouvelle ligne de détail
Private Sub Cmd_SaisiL_Click()

On Error GoTo Err
Okayy = True

If txt_Numero.Text = "" Then
    If Len(Trim(txt_Numero.Text)) = 0 Then
        MsgBox "N° bon obligatoire      ", vbInformation
        Exit Sub
    End If
End If
        
If cbo_MatriculeStation.Text = "" Or cbo_MatriculeStation.Text = " " Then
    If Len(Trim(cbo_MatriculeStation.Text)) = 0 Then
        MsgBox "Station obligatoire      ", vbInformation
        Exit Sub
    End If
End If
If Pict_Transf.Visible = True Then
    If Opt_Fact.Value = False And Opt_PRecep.Value = False Then
        MsgBox "Vous devez choisir le type de la pièce", vbInformation
        Exit Sub
    End If
End If

If txt_Numero.Text = "" Then
    If Len(Trim(txt_Numero.Text)) = 0 Then
        MsgBox "N° bon obligatoire      ", vbInformation
        txt_Numero.SetFocus
        Exit Sub
    End If
    With FrmSaisiePieceReparation
        .txt_Numero.Text = Me.txt_Numero.Text
        .Okay = True
        .Show vbModal
    End With
Else
    With FrmSaisiePieceReparation
        .txt_Numero.Text = Me.txt_Numero.Text
        .Okay = True
        .Show vbModal
    End With
End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

'Afficher liste des BCReparation
Private Sub Cmd_FinBCRep_Click()

On Error GoTo Err
   
 '"Piece Reception"
Unload FrmFind
With FrmFind
    .StrSource = "FIndBCReparation"
    .Show vbModal
End With
 
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

'Afficher l'assiette et les détails d'une pièce de réparation lors d'un DBclick dans Frid du FrmFind
Public Sub AfficheRow(ByVal VCode As String)

Dim TotHTBrut As Double
Dim TotTTC As Double
Dim Fcode As String
Dim Qte As Double
Dim PUHT As Double
Dim Remise As Double
Dim tva As Double
Dim LOBJ_PieceRep As PieceReparation
Dim rs As New Recordset
Dim Transf As Boolean

Transf = False
Call ViderZone(FrmPieceReparation)
Lsv_Detail.ListItems.Clear
Lsv_Toto.ListItems.Clear
Lbl_UserSaisi.Caption = ""
txt_Numero.Text = ""
'Assiette
Set LOBJ_PieceRep = New PieceReparation
Set rs = LOBJ_PieceRep.Get_AssPieceReparation(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
If rs("Supp") = "O" Then
    MsgBox "Pièce réparation supprimée", vbInformation
    Call ViderZone(FrmPieceReparation)
    txt_Numero.SetFocus
    Exit Sub
End If
    'Charge
    Fcode = rs("Fournisseur")
    txt_Numero.Text = rs("Numero")
    If Not IsNull(rs("Type")) Then cbo_typePiece.Text = rs("Type")
    If Not IsNull(rs("refPiece")) Then txt_ref.Text = rs("refPiece")
    If Not IsNull(rs("DatePiece")) Then cda_Create.Caption = rs("DatePiece")
    If Not IsNull(rs("DateOperation")) Then cda_Operation.Value = rs("DateOperation")
    If Not IsNull(rs("Fournisseur")) Then Call AfficheRow_Station(rs("Fournisseur"))
    If Not IsNull(rs("RemisePiece")) Then Tex_RSP.Text = rs("RemisePiece")
    If Not IsNull(rs("timbre")) Then txt_Timbre.Text = rs("Timbre")
    If Not IsNull(rs("PrixMOeuvre")) Then Txt_PMainOeuvre.Text = rs("PrixMOeuvre")
    If Not IsNull(rs("TVA_MOeuvre")) Then Txt_TvaMO.Text = rs("TVA_MOeuvre")
    If Not IsNull(rs("UserInsert")) Then Lbl_UserSaisi.Caption = Get_NameUserByCode(rs("UserInsert"))
    If Not IsNull(rs("NumFact")) Then
        LBL_NFact.Caption = rs("NumFact")
        PIC_NFACT.Visible = True
        Call Timer1_Timer
        Transf = True
        Pict_Type.Enabled = False
        Lsv_Detail.Enabled = False
        Picture2.Enabled = False
        Pict_TRP.Enabled = False
    Else
        Pict_Type.Enabled = True
        Pict_TRP.Enabled = True
        Lsv_Detail.Enabled = True
        Timer1.Enabled = False
        Picture2.Enabled = True
    End If
Else
    Timer1.Enabled = False
    MsgBox "Numéro bon introuvable", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If
rs.Close
'Details
Set rs = LOBJ_PieceRep.Get_DetPieceReparation(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
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
        itmX.SubItems(5) = Format(rs("Remise"), "#,##0.00")
        itmX.SubItems(6) = Format(TotHTBrut, "#,##0.000")
        itmX.SubItems(7) = Format(rs("tva"), "#,##0.00")
        itmX.SubItems(8) = Format(TotTTC, "#,##0.000")
            
        rs.MoveNext
    Wend
    Call AppCalcul
End If
If Transf = True Then
    Picture1.Enabled = False
End If
rs.Close

End Sub

Public Sub AppCalcul()

Dim ii As Integer
'Ligne de pièce
Dim pu As Double
Dim Qte As Double
Dim HTBrutLigne As Double
Dim RemiseL As Double
Dim HTNetLigne As Double
Dim tvaLigne As Double
Dim ttcLigne As Double

'Totaux Pièce
Dim TotHTBrut As Double
Dim TotRemLigne As Double
Dim TotHtNet As Double
Dim TotTva  As Double
Dim ValRemP As Double
Dim TotTTCSansRP As Double
Dim RemiseP As Double
Dim TotHtNetPiece As Double
Dim Timbre As Double
Dim TotTTC As Double
Dim MainOeuvre As Double
Dim Tva_MOeuvre As Double
'Intit totaux
TotHTBrut = 0
TotRemLigne = 0
RemiseP = 0
TotHtNet = 0
TotTva = 0
ValRemP = 0
Timbre = 0
TotTTC = 0
MainOeuvre = 0
Tva_MOeuvre = 0
Lsv_Toto.ListItems.Clear

For ii = 1 To Lsv_Detail.ListItems.Count
  'Intit Lignes
    pu = 0
    Qte = 0
    HTBrutLigne = 0
    RemiseL = 0
    HTNetLigne = 0
    tvaLigne = 0
    ttcLigne = 0
    
    'TotHTBrut
    Qte = Lsv_Detail.ListItems(ii).SubItems(2)
    pu = Lsv_Detail.ListItems(ii).SubItems(4)
    HTBrutLigne = Qte * pu
    TotHTBrut = TotHTBrut + HTBrutLigne
    
    'TotRemLigne
    If Lsv_Detail.ListItems(ii).SubItems(5) <> "" Then
        RemiseL = Lsv_Detail.ListItems(ii).SubItems(5)
        TotRemLigne = TotRemLigne + (HTBrutLigne * RemiseL / 100)
    End If
    'TotHtNet
    HTNetLigne = HTBrutLigne - (HTBrutLigne * RemiseL / 100)  'Lsv_Detail.ListItems(ii).SubItems(6)
    RemiseP = RemiseP + (HTNetLigne * CDbl(Tex_RSP.Text) / 100)
    HTNetLigne = HTNetLigne - (HTNetLigne * CDbl(Tex_RSP.Text) / 100)
    TotHtNet = TotHtNet + HTNetLigne
    
    'TotTva
    tvaLigne = Lsv_Detail.ListItems(ii).SubItems(7)
    TotTva = TotTva + (HTNetLigne * tvaLigne / 100)

Next

'Main d'oeuvre
If Txt_PMainOeuvre.Text = "" Then Txt_PMainOeuvre.Text = "0"
Txt_PMainOeuvre.Text = Format(Txt_PMainOeuvre.Text, "##0.000")
MainOeuvre = CDbl(Txt_PMainOeuvre.Text)

'TotHtNet et brut de toute piece
'Ajouter prix main d'oeuvre brut
TotHtNet = TotHtNet + (MainOeuvre - (MainOeuvre * CDbl(Tex_RSP.Text) / 100))
TotHTBrut = TotHTBrut + MainOeuvre

'RemiseP : TotHtNet + main d'oeuvre
If Tex_RSP.Text = "" Then Tex_RSP.Text = "0"
Tex_RSP.Text = Format(Tex_RSP.Text, "##0.00")
RemiseP = RemiseP + (MainOeuvre * CDbl(Tex_RSP.Text) / 100) 'Appliquer remise de la pièce sur tout les produits et main d'oeuvre

If Txt_TvaMO.Text = "" Then Txt_TvaMO.Text = "0"
Txt_TvaMO.Text = Format(Txt_TvaMO.Text, "##0.00")
Tva_MOeuvre = CDbl(Txt_TvaMO.Text)
Tva_MOeuvre = (MainOeuvre - (MainOeuvre * CDbl(Tex_RSP.Text) / 100)) * Tva_MOeuvre / 100
MainOeuvre = MainOeuvre + Tva_MOeuvre 'Prix main d'oeuvre avec tva

'TotTva de toute la pièce
TotTva = TotTva + Tva_MOeuvre

'Timbre
If txt_Timbre.Text = "" Then txt_Timbre.Text = "0"
txt_Timbre.Text = Format(txt_Timbre.Text, "##0.000")
Timbre = CDbl(txt_Timbre.Text)

'TotHtNetPiece
TotHtNetPiece = TotHTBrut - TotRemLigne - RemiseP

'TotTTC

TotTTC = TotHtNetPiece + TotTva + Timbre

Set itmX = Lsv_Toto.ListItems.Add(, , CStr(Format(TotHTBrut, "#,##0.000")))
    itmX.SubItems(1) = CStr(Format(TotHTBrut, "#,##0.000"))
    itmX.SubItems(2) = CStr(Format(TotRemLigne, "#,##0.000"))
    itmX.SubItems(3) = CStr(Format(RemiseP, "#,##0.000"))
    itmX.SubItems(4) = CStr(Format(TotHtNetPiece, "#,##0.000"))
    itmX.SubItems(5) = CStr(Format(TotTva, "#,##0.000"))
    itmX.SubItems(6) = CStr(Format(TotTTC, "#,##0.000"))

End Sub

'Incrementation du compteur de la table piece de réparatio lors d'un nouveau saisi
Private Function return_Compteur() As Long

Dim rD As New Recordset
Dim LOBJ_PieceRep As PieceReparation

return_Compteur = 0
Set LOBJ_PieceRep = New PieceReparation
Set rD = LOBJ_PieceRep.Get_MaxNum(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
If Not rD.EOF Then
    return_Compteur = rD(0)
End If
rD.Close
End Function

'Afficher l'assiette et les détails d'un bon de commande de réparation en DBclick sur Grid de FrmFind (Liste des BC de réparation)
Public Sub AfficheRow_BCR(ByVal VCode As String)

Dim LOBJ_BCRepa As BCReparation
Dim LOBJ_Prod As Produit_Lubrifiant
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim ttc
Dim trv As Boolean

Call ViderZone(FrmPieceReparation)
Lsv_Detail.ListItems.Clear
Lsv_Toto.ListItems.Clear
Lbl_UserSaisi.Caption = ""
cda_Operation.Value = Date
Pict_Transf.Visible = True
Pict_Creat.Visible = False
trv = False

Set LOBJ_Prod = New Produit_Lubrifiant
Set LOBJ_BCRepa = New BCReparation
'Assiette du BC de réparation
Set rs = LOBJ_BCRepa.Get_AssBRepar(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
'si le bon de commande est transféré afficher le numero de la pièce de réception
If Not rs.EOF Then
    If rs("transf") = "O" Then
        txt_Numero.Text = rs("NumPR")
    Else
        txt_Numero.Text = "Auto"
    End If
    cda_Create.Caption = rs("DateCreation")
    cbo_MatriculeStation.Text = rs("Fournisseur")
    If Not IsNull(rs("UserInsert")) Then Lbl_UserSaisi.Caption = rs("UserInsert")
End If
rs.Close

'Détails du BC de réparation
Set rs = LOBJ_BCRepa.Get_DetBRepar(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        Set itmX = Lsv_Detail.ListItems.Add(, , CStr(rs("Numero")))
        txt_BCReparation.Text = rs("Numero")
        itmX.SubItems(1) = rs("Désignation")
        itmX.SubItems(2) = rs("Qté")
        itmX.SubItems(3) = rs("Vehicule")
        Set rs1 = LOBJ_Prod.Get_ProdLubByLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, itmX.SubItems(1))
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        If Not rs1.EOF Then
            itmX.SubItems(4) = Format(rs1("prixht"), "#,##0.000")
            itmX.SubItems(5) = 0
            itmX.SubItems(6) = CDbl(rs1("prixht")) * Val(rs("qté"))
            itmX.SubItems(7) = rs1("tva")
            ttc = CDbl(itmX.SubItems(6)) + CDbl(itmX.SubItems(6)) * CDbl(rs1("tva")) / 100
            itmX.SubItems(8) = Format(ttc, "#,##0.000")
         Else
         'un nouveau produit n'est pas inséré dans la base
            trv = True
            itmX.SubItems(4) = Format(0, "#,##0.000")
            itmX.SubItems(5) = Format(0, "#,##0.000")
            itmX.SubItems(6) = Format(0, "#,##0.000")
            itmX.SubItems(7) = Format(0, "#,##0.000")
            itmX.SubItems(8) = Format(0, "#,##0.000")
         End If
         rs1.Close
        rs.MoveNext
    Wend
End If
rs.Close
Call AppCalcul
If trv = True Then MsgBox "Un nouveau produit est trouvé, vous devez saisir ses détails "
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
    cbo_MatriculeStation.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
End If

End Sub

Private Sub cbo_typePiece_GotFocus()
If (op_creat(0).Value = vbUnchecked) And (op_creat(1).Value = vbUnchecked) Then
    MsgBox "Type de créaion invalide.", vbExclamation
    op_creat(0).SetFocus
    Exit Sub
End If
End Sub

Private Sub cbo_typePiece_Click()
Lsv_Detail.Enabled = True
Picture2.Enabled = True
If Lsv_Detail.ListItems.Count > 0 Then Pict_TRP.Enabled = True
End Sub

Private Sub cbo_typePiece_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cbo_typePiece_LostFocus()
If (cbo_typePiece.Text = "Piece Reception") Then txt_Timbre.Text = 0
If (cbo_typePiece.Text = "Facture") Then txt_Timbre.Text = 0
Lsv_Detail.Enabled = True
Picture2.Enabled = True
If Lsv_Detail.ListItems.Count > 0 Then Pict_TRP.Enabled = True
End Sub

Private Sub Opt_Fact_Click()
Picture2.Enabled = True
Lsv_Detail.Enabled = True
If Lsv_Detail.ListItems.Count > 0 Then Pict_TRP.Enabled = True
End Sub

Private Sub Opt_PRecep_Click()
Picture2.Enabled = True
Lsv_Detail.Enabled = True
If Lsv_Detail.ListItems.Count > 0 Then Pict_TRP.Enabled = True
End Sub


Private Sub Tex_RSP_GotFocus()
Tex_RSP.SelStart = 0
Tex_RSP.SelLength = Len(Tex_RSP.Text)

End Sub

Private Sub Tex_RSP_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) Like "." Then KeyAscii = 44
If Not (Chr(KeyAscii) Like "[0123456789,]") And KeyAscii <> 13 And KeyAscii <> 8 Or InStr(Tex_RSP.Text, ",") <> 0 And Chr(KeyAscii) = "," Then
    KeyAscii = 0
End If
End Sub

Private Sub Txt_PMainOeuvre_GotFocus()
Txt_PMainOeuvre.SelStart = 0
Txt_PMainOeuvre.SelLength = Len(Txt_PMainOeuvre.Text)
End Sub

Private Sub Txt_PMainOeuvre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Txt_PMainOeuvre.Text = Format(Txt_PMainOeuvre.Text, "##0.000")
        Call AppCalcul
        SendKeys "{tab}"
    End If
End Sub

Private Sub Txt_PMainOeuvre_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) Like "." Then KeyAscii = 44
If Not (Chr(KeyAscii) Like "[0123456789,]") And KeyAscii <> 13 And KeyAscii <> 8 Or InStr(Txt_PMainOeuvre.Text, ",") <> 0 And Chr(KeyAscii) = "," Then
    KeyAscii = 0
End If
End Sub

Private Sub Txt_PMainOeuvre_LostFocus()
If Txt_PMainOeuvre.Text = "" Then Txt_PMainOeuvre.Text = 0
Txt_PMainOeuvre.Text = Format(Txt_PMainOeuvre, "##0.000")
Call AppCalcul
End Sub

Private Sub txt_Timbre_GotFocus()
txt_Timbre.SelStart = 0
txt_Timbre.SelLength = Len(txt_Timbre.Text)

End Sub

Private Sub txt_Timbre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txt_Timbre.Text = Format(txt_Timbre, "##0.000")
    Call AppCalcul
    SendKeys "{tab}"
End If
End Sub

Private Sub txt_Timbre_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) Like "." Then KeyAscii = 44
If Not (Chr(KeyAscii) Like "[0123456789,]") And KeyAscii <> 13 And KeyAscii <> 8 Or InStr(txt_Timbre.Text, ",") <> 0 And Chr(KeyAscii) = "," Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_Timbre_LostFocus()
If txt_Timbre.Text = "" Then txt_Timbre.Text = 0
txt_Timbre.Text = Format(txt_Timbre, "##0.000")
Call AppCalcul
End Sub

Private Sub txt_ref_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_ref_LostFocus()
If txt_ref.Text = "" Then
    txt_ref.Text = "Sans Ref"
End If
End Sub

Private Sub Tex_RSP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Tex_RSP.Text = Format(Tex_RSP.Text, "##0.00")
    Call AppCalcul
    SendKeys "{tab}"
End If
End Sub

Private Sub Tex_RSP_LostFocus()
If Tex_RSP.Text = "" Then Tex_RSP.Text = 0
 Tex_RSP.Text = Format(Tex_RSP.Text, "##0.00")
Call AppCalcul
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

Private Sub txt_BCReparation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cbo_MatriculeStation_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub cbo_MatriculeStation_Click()
If Len(Trim(cbo_MatriculeStation.Text)) > 0 Then Call AfficheRow_Station(cbo_MatriculeStation.Text)

End Sub

Private Sub txt_Numero_LostFocus()

On Error GoTo Err

If Len(Trim(txt_Numero.Text)) > 0 Then
    Call AfficheRow(txt_Numero.Text)
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

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

Private Sub Txt_TvaMO_GotFocus()
Txt_TvaMO.SelStart = 0
Txt_TvaMO.SelLength = Len(Txt_TvaMO.Text)
End Sub

Private Sub Txt_TvaMO_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Txt_TvaMO.Text = Format(Txt_TvaMO.Text, "##0.00")
    Call AppCalcul
    SendKeys "{tab}"
End If
End Sub

Private Sub Txt_TvaMO_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) Like "." Then KeyAscii = 44
If Not (Chr(KeyAscii) Like "[0123456789,]") And KeyAscii <> 13 And KeyAscii <> 8 Or InStr(Txt_TvaMO.Text, ",") <> 0 And Chr(KeyAscii) = "," Then
    KeyAscii = 0
End If
End Sub

Private Sub Txt_TvaMO_LostFocus()
If Txt_TvaMO.Text = "" Then Txt_TvaMO.Text = 0
Txt_TvaMO.Text = Format(Txt_TvaMO, "##0.00")
Call AppCalcul
End Sub
