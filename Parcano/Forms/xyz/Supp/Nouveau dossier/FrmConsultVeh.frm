VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmConsultVeh 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Véhicule"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11310
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
   ScaleHeight     =   8730
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer T_VID 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7320
      Top             =   120
   End
   Begin VB.Timer T_TAX 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   120
   End
   Begin VB.Timer T_VIS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   120
   End
   Begin VB.Timer T_ASS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5160
      Top             =   120
   End
   Begin VB.TextBox txt_Matricule 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   1800
      TabIndex        =   0
      Tag             =   "M"
      Top             =   1560
      Width           =   2295
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   1560
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
      Picture         =   "FrmConsultVeh.frx":0000
      ButtonType      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   60
      TabIndex        =   2
      Top             =   2160
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Carte grise et Caractéristiques technique."
      TabPicture(0)   =   "FrmConsultVeh.frx":0353
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Papier et alerte"
      TabPicture(1)   =   "FrmConsultVeh.frx":036F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Autres"
      TabPicture(2)   =   "FrmConsultVeh.frx":038B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Papier et alerte"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   -74880
         TabIndex        =   49
         Top             =   600
         Width           =   10215
         Begin VB.TextBox txt_FourAssur 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   52
            Tag             =   "M"
            Top             =   960
            Width           =   4095
         End
         Begin VB.TextBox txt_AgenceAssur 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   51
            Tag             =   "M"
            Top             =   1560
            Width           =   4095
         End
         Begin VB.TextBox txt_NumAssur 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   50
            Tag             =   "M"
            Top             =   360
            Width           =   4080
         End
         Begin SToolBox.SDateBox cda_DebuAssur 
            Height          =   285
            Left            =   1560
            TabIndex        =   53
            Tag             =   "M"
            Top             =   2160
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_FinAssur 
            Height          =   285
            Left            =   3720
            TabIndex        =   54
            Tag             =   "M"
            Top             =   2160
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_DebutVesite 
            Height          =   285
            Left            =   1575
            TabIndex        =   55
            Top             =   2760
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_FinVisite 
            Height          =   285
            Left            =   3720
            TabIndex        =   56
            Top             =   2760
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_debuttax 
            Height          =   285
            Left            =   1560
            TabIndex        =   57
            Top             =   3360
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_fintax 
            Height          =   285
            Left            =   3720
            TabIndex        =   58
            Top             =   3360
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
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
            Left            =   3000
            TabIndex        =   70
            Top             =   3360
            Width           =   315
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Taxe du :"
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
            Left            =   570
            TabIndex        =   69
            Top             =   3360
            Width           =   765
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Visite du :"
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
            Left            =   600
            TabIndex        =   68
            Top             =   2760
            Width           =   810
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
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
            Left            =   3000
            TabIndex        =   67
            Top             =   2760
            Width           =   315
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
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
            Left            =   3000
            TabIndex        =   66
            Top             =   2160
            Width           =   315
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Assu du :"
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
            Left            =   600
            TabIndex        =   65
            Top             =   2160
            Width           =   750
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agence :"
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
            Left            =   600
            TabIndex        =   64
            Top             =   1560
            Width           =   720
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Assureur :"
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
            Left            =   600
            TabIndex        =   63
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° police :"
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
            Left            =   600
            TabIndex        =   62
            Top             =   360
            Width           =   825
         End
         Begin VB.Label LBL_ALERT_ASSURANCE 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Bell MT"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   5040
            TabIndex        =   61
            Top             =   2160
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label LBL_ALERT_VISITE 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Bell MT"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   5040
            TabIndex        =   60
            Top             =   2760
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label LBL_ALERT_TAXE 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Bell MT"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   5040
            TabIndex        =   59
            Top             =   3360
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Im_ass 
            Height          =   240
            Left            =   5280
            Picture         =   "FrmConsultVeh.frx":03A7
            Stretch         =   -1  'True
            Top             =   2160
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Im_Vis 
            Height          =   240
            Left            =   5280
            Picture         =   "FrmConsultVeh.frx":06B1
            Stretch         =   -1  'True
            Top             =   2760
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Im_tax 
            Height          =   240
            Left            =   5280
            Picture         =   "FrmConsultVeh.frx":09BB
            Stretch         =   -1  'True
            Top             =   3360
            Visible         =   0   'False
            Width           =   195
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   11055
         Begin VB.TextBox Txt_NvCpt 
            Height          =   405
            Left            =   6240
            TabIndex        =   78
            Top             =   4560
            Width           =   2415
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   4440
            ScaleHeight     =   570
            ScaleWidth      =   4275
            TabIndex        =   74
            Top             =   5040
            Width           =   4275
            Begin VB.TextBox txt_DerCompteurV 
               Height          =   315
               Left            =   1800
               TabIndex        =   75
               Tag             =   "M"
               Top             =   120
               Width           =   2415
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Der.compteur vid"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   240
               TabIndex        =   76
               Top             =   120
               Width           =   1365
               WordWrap        =   -1  'True
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   7200
            ScaleHeight     =   495
            ScaleWidth      =   3855
            TabIndex        =   73
            Top             =   5520
            Width           =   3855
         End
         Begin VB.TextBox txt_Obs 
            Appearance      =   0  'Flat
            Height          =   915
            Left            =   1920
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   5160
            Width           =   2295
         End
         Begin VB.TextBox txt_Charge 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6240
            MaxLength       =   30
            TabIndex        =   21
            Tag             =   "M"
            Top             =   3600
            Width           =   2415
         End
         Begin VB.TextBox txt_PTRA 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6240
            MaxLength       =   30
            TabIndex        =   20
            Tag             =   "M"
            Top             =   3120
            Width           =   2415
         End
         Begin VB.TextBox txt_PoidVide 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6240
            MaxLength       =   30
            TabIndex        =   19
            Tag             =   "M"
            Top             =   2640
            Width           =   2415
         End
         Begin VB.TextBox txt_PTAC 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6240
            MaxLength       =   30
            TabIndex        =   18
            Tag             =   "M"
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txt_NbrEssieux 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6240
            MaxLength       =   30
            TabIndex        =   17
            Tag             =   "M"
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox txt_Cylindre 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6240
            MaxLength       =   30
            TabIndex        =   16
            Tag             =   "M"
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox txt_Plassedebout 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "M"
            Top             =   4320
            Width           =   2295
         End
         Begin VB.TextBox txt_PlasseAssise 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "M"
            Top             =   3840
            Width           =   2295
         End
         Begin VB.TextBox txt_Carrosserie 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   13
            Tag             =   "M"
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox txt_Genre 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "M"
            Top             =   2880
            Width           =   2295
         End
         Begin VB.TextBox txt_TYPCOM 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "M"
            Top             =   2400
            Width           =   2295
         End
         Begin VB.TextBox txt_Marque 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "M"
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txt_Puissance 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "M"
            Top             =   720
            Width           =   2415
         End
         Begin VB.ComboBox Cbo_Energie 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6240
            TabIndex        =   8
            Tag             =   "M"
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txt_KlmVidange 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6240
            TabIndex        =   7
            Tag             =   "M"
            Top             =   4080
            Width           =   2415
         End
         Begin VB.TextBox txt_libelle 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   6
            Tag             =   "M"
            Top             =   120
            Width           =   2295
         End
         Begin VB.TextBox txt_Type 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   5
            Tag             =   "M"
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox txt_Nserie 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   4
            Tag             =   "M"
            Top             =   1560
            Width           =   2295
         End
         Begin SToolBox.SDateBox cda_DateCircul 
            Height          =   285
            Left            =   1920
            TabIndex        =   23
            Tag             =   "M"
            Top             =   2040
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SCommand cmdFindenergie 
            Height          =   495
            Left            =   8760
            TabIndex        =   24
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
            Picture         =   "FrmConsultVeh.frx":0CC5
            ButtonType      =   1
         End
         Begin SToolBox.SDateBox cda_dateSortie 
            Height          =   285
            Left            =   1920
            TabIndex        =   25
            Top             =   4800
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Label Lbl_NvCpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Compteur traffic :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4680
            TabIndex        =   77
            Top             =   4680
            Width           =   1575
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observation_ _ _:"
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
            TabIndex        =   48
            Top             =   5280
            Width           =   1485
         End
         Begin VB.Label LBL_VIDANGE 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Il vous reste 180 klm"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   735
            Left            =   9480
            TabIndex        =   47
            Top             =   4680
            Visible         =   0   'False
            Width           =   1575
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date session _ _ _:"
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
            TabIndex        =   46
            Top             =   4800
            Width           =   1575
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Charge utile :"
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
            Left            =   4680
            TabIndex        =   45
            Top             =   3720
            Width           =   1110
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cylindree :"
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
            Left            =   4680
            TabIndex        =   44
            Top             =   1320
            Width           =   885
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Poids a vide :"
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
            Left            =   4680
            TabIndex        =   43
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P.T.A.C. :"
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
            Left            =   4680
            TabIndex        =   42
            Top             =   2280
            Width           =   705
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre essieux :"
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
            Left            =   4680
            TabIndex        =   41
            Top             =   1800
            Width           =   1440
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P.T.R.A :"
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
            Left            =   4680
            TabIndex        =   40
            Top             =   3240
            Width           =   675
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Place debout _ _ _:"
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
            TabIndex        =   39
            Top             =   4440
            Width           =   1590
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plasses assises _ _:"
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
            TabIndex        =   38
            Top             =   3960
            Width           =   1635
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carrosserie _ _ _ _:"
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
            TabIndex        =   37
            Top             =   3480
            Width           =   1620
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Genre _ _ _ _ _ _ _:"
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
            TabIndex        =   36
            Top             =   3000
            Width           =   1605
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type commerciale _:"
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
            TabIndex        =   35
            Top             =   2520
            Width           =   1755
         End
         Begin VB.Image Im_Vid 
            Height          =   240
            Left            =   9000
            Picture         =   "FrmConsultVeh.frx":1018
            Stretch         =   -1  'True
            Top             =   5160
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label LBL_ALERT_VIDANGE 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Bell MT"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   8760
            TabIndex        =   34
            Top             =   5160
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puissance :"
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
            Left            =   4680
            TabIndex        =   33
            Top             =   840
            Width           =   930
         End
         Begin VB.Label Label7 
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
            Left            =   4680
            TabIndex        =   32
            Top             =   4200
            Width           =   1320
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
            Left            =   4680
            TabIndex        =   31
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Matricule_ _ _ _ _ _ :"
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
            TabIndex        =   30
            Top             =   240
            Width           =   1740
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type_ _ _ _ _ _  _ _:"
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
            TabIndex        =   29
            Top             =   1200
            Width           =   1665
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date 1er circulation :"
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
            TabIndex        =   28
            Top             =   2040
            Width           =   1755
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° serie du type _ _:"
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
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marque_ _ _ _ _  _ _:"
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
            TabIndex        =   26
            Top             =   720
            Width           =   1740
         End
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Immatriculation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1740
      TabIndex        =   72
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche véhicule"
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
      Left            =   360
      TabIndex        =   71
      Top             =   240
      Width           =   2025
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmConsultVeh.frx":1322
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "FrmConsultVeh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flap
Dim flap1
Dim flap2
Dim flap3
Public AA As String

Public Sub AfficheRow(ByVal vcode As String)

Dim LOBJ_BonVidange As BonVidange
Dim LOBJ_Vehicule As Vehicule
Dim LOBJ_Lub As Lubrifiant
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim AA As Long

Call ViderZone(FrmConsultVeh)
Set LOBJ_Vehicule = New Vehicule
AA = 0
Set rs = LOBJ_Vehicule.GetVehicule(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vcode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    If Not IsNull(rs("Matricule")) Then
        txt_libelle.Text = rs("Matricule")
        Txt_NvCpt.Text = CompteurVehicule(rs("Matricule"))
    End If
    If Not IsNull(rs("marque")) Then txt_Type.Text = rs("TYPE")
    If Not IsNull(rs("puissance")) Then txt_Marque.Text = rs("marque")
    If Not IsNull(rs("Matricule")) Then txt_Puissance.Text = rs("puissance")
    If Not IsNull(rs("Energie")) Then Cbo_Energie.Text = rs("Energie")
'    If Not IsNull(rs("compteur")) Then txt_compteur.Text = rs("compteur")
    If Not IsNull(rs("NBKLMvid")) Then txt_KlmVidange.Text = rs("NBKLMvid")

    If Not IsNull(rs("NumSerie")) Then txt_Nserie.Text = rs("NumSerie")
    If Not IsNull(rs("DateCircul")) Then cda_DateCircul.Text = rs("DateCircul")
    If Not IsNull(rs("NumAssur")) Then txt_NumAssur.Text = rs("NumAssur")
    If Not IsNull(rs("FournisAssur")) Then txt_FourAssur.Text = rs("FournisAssur")
    If Not IsNull(rs("AgenceAssur")) Then txt_AgenceAssur.Text = rs("AgenceAssur")
    
    If Not IsNull(rs("DateDebAssur")) And rs("DateDebAssur") <> "01/01/1900" Then cda_DebuAssur.Text = rs("DateDebAssur")
    If Not IsNull(rs("DateFinAssur")) And rs("DateFinAssur") <> "01/01/1900" Then cda_FinAssur.Text = rs("DAteFinAssur")
    Set LOBJ_BonVidange = New BonVidange
    Set rs1 = LOBJ_BonVidange.Get_DerBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("Code"))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs1.EOF Then
        If Not IsNull(rs1("CompteurVidange")) Then txt_DerCompteurV.Text = rs1("CompteurVidange")
    End If
    rs1.Close
    '
    If Not IsNull(rs("genre")) Then txt_Genre.Text = rs("genre")
    If Not IsNull(rs("Carrosserie")) Then txt_Carrosserie.Text = rs("Carrosserie")
    If Not IsNull(rs("PlaceAssis")) Then txt_PlasseAssise.Text = rs("PlaceAssis")
    If Not IsNull(rs("PlaceDebout")) Then txt_Plassedebout.Text = rs("PlaceDebout")
    If Not IsNull(rs("Cylindre")) Then txt_Cylindre.Text = rs("Cylindre")
    If Not IsNull(rs("NbrEssieux")) Then txt_NbrEssieux.Text = rs("NbrEssieux")
    If Not IsNull(rs("PTAC")) Then txt_PTAC.Text = rs("PTAC")
    If Not IsNull(rs("PTRA")) Then txt_PTRA.Text = rs("PTRA")
    If Not IsNull(rs("Charge")) Then txt_Charge.Text = rs("Charge")
    If Not IsNull(rs("PoidsVide")) Then txt_PoidVide.Text = rs("PoidsVide")
    If Not IsNull(rs("TypeComm")) Then txt_TYPCOM.Text = rs("TypeComm")


    'Alert assurance
    If rs("DateFinAssur") <> "01/01/1900" And DateAdd("d", 5, Date) >= rs("DateFinAssur") Then
        LBL_ALERT_ASSURANCE.Visible = True
        T_ASS.Enabled = True
        Im_ass.Visible = True
    Else
        LBL_ALERT_ASSURANCE.Visible = False
        T_ASS.Enabled = False
        Im_ass.Visible = False
    End If
    'fin assurance
    If Not IsNull(rs("DateDebVisite")) And rs("DateDebVisite") <> "01/01/1900" Then cda_DebutVesite.Text = rs("DateDebVisite")
    If Not IsNull(rs("DAteFinVisite")) And rs("DAteFinVisite") <> "01/01/1900" Then cda_FinVisite.Text = rs("DAteFinVisite")
    'Alert visite
    If rs("DAteFinVisite") <> "01/01/1900" And DateAdd("d", 5, Date) >= rs("DAteFinVisite") Then
        LBL_ALERT_VISITE.Visible = True
        T_VIS.Enabled = True
        Im_Vis.Visible = True
    Else
        LBL_ALERT_VISITE.Visible = False
        T_VIS.Enabled = False
        Im_Vis.Visible = False
    End If
    'fin visite
    If Not IsNull(rs("DateDebTax")) And rs("DateDebTax") <> "01/01/1900" Then cda_debuttax.Text = rs("DateDebTax")
    If Not IsNull(rs("DateFinTax")) And rs("DateFinTax") <> "01/01/1900" Then cda_fintax.Text = rs("DateFinTax")
    'Alert tax
    If rs("DateFinTax") <> "01/01/1900" And DateAdd("d", 5, Date) >= rs("DateFinTax") Then
        LBL_ALERT_TAXE.Visible = True
        T_TAX.Enabled = True
        Im_tax.Visible = True
    Else
        LBL_ALERT_TAXE.Visible = False
        T_TAX.Enabled = False
        Im_tax.Visible = False
    End If
    'fin tax
    'Alert vidange
    If txt_DerCompteurV.Text <> 0 Or txt_DerCompteurV.Text <> "" Then
        AA = Val(Txt_NvCpt.Text) - Val(txt_DerCompteurV.Text)
    If AA + 500 >= rs("NBKLMvid") Then
        If AA > rs("NBKLMvid") Then
            LBL_VIDANGE.Caption = "Vous avez dépassé le nbr de klm de vidange par  : " & AA - rs("NBKLMvid")
        Else
            LBL_VIDANGE.Caption = "Il vous reste " & rs("NBKLMvid") - AA & " klm pour le nouveau vidange"
        End If
        LBL_ALERT_VIDANGE.Visible = True
        LBL_VIDANGE.Visible = True
        T_VID.Enabled = True
        Im_Vid.Visible = True
    Else
        LBL_ALERT_VIDANGE.Visible = False
        LBL_VIDANGE.Visible = False
        Im_Vid.Visible = False
        T_VID.Enabled = False
    End If
    End If
    'Fin
    
    If Not IsNull(rs("DateSortie")) Then cda_dateSortie.Text = rs("DateSortie")
    If Not IsNull(rs("Obs")) Then txt_Obs.Text = rs("Obs")

    
Else
    MsgBox "Code introuvable", vbInformation

End If
rs.Close

End Sub

'Dernier compteur du véhicule en entrant (ficheTraffic)
Public Function CompteurVehicule(ByVal vcode As String) As String

Dim LOBJ_Vehicule As Vehicule
Dim rs1 As New Recordset
Dim Name_Tab As String

CompteurVehicule = "0"
Set LOBJ_Vehicule = New Vehicule
Name_Tab = "FicheTraffic_" & Year(Date)
Set rs1 = LOBJ_Vehicule.Get_DerCompt(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Name_Tab, vcode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
If Not rs1.EOF Then
    If Not IsNull(rs1("maxCpt")) Then CompteurVehicule = rs1("maxCpt")
End If
rs1.Close
    
End Function

Private Sub Affiche_Energie()

Dim LOBJ_Energie As Energie
Dim rs As New Recordset

Set LOBJ_Energie = New Energie
Set rs = LOBJ_Energie.Get_Energ(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        With Cbo_Energie
            .AddItem rs("libelle")
        End With
        rs.MoveNext
    Wend
End If
End Sub

Private Sub Cbo_Energie_GotFocus()

If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub Cbo_Energie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_marque_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_marque_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cda_DateCircul_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub cda_dateSortie_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub cda_DebuAssur_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub cda_DebutVesite_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub cda_FinAssur_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub cda_FinVisite_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub cmdFindenergie_Click()
Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Energie"
    .Show vbModal
End With
End Sub

Private Sub cmdFindMatricule_Click()
Unload FrmFind
With FrmFind
    .StrSource = "Véhicule"
    .Show vbModal
End With
End Sub

Private Sub Form_Load()

On Error GoTo Err

Me.Width = 10560
Me.Height = 8565
Me.Move 0, 0

Call Affiche_Energie
'Call Affiche_Lubrifiant
SSTab1.TabCaption(1) = "Papiers et alertes"

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub Form_Resize()

On Error Resume Next

Image1.Width = Me.Width
CmdSave.Left = Me.Width - 700
CmdFind.Left = Me.Width - 1100
CmdDelete.Left = Me.Width - 1500
CmdAdd.Left = Me.Width - 1900
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo erreur
   Dim i As Integer
   Dim MSG ' Déclare la variable.
   ' Définit le texte du message.
   MSG = "Voulez-vous vraiment quitter?"
   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
   If MsgBox(MSG, vbQuestion + vbYesNo + vbDefaultButton2, Label1.Caption) = vbNo Then
      Cancel = True
   Else
   Unload FrmConsultVeh
   End If
   
   Exit Sub
erreur:
   MsgBox Err.Description, 48
End Sub

Private Sub T_ASS_Timer()
    If flap = 0 Then
        LBL_ALERT_ASSURANCE.Visible = True
        Im_ass.Visible = True
        SSTab1.TabCaption(1) = "....................."
        
        flap = 1
    Else
        LBL_ALERT_ASSURANCE.Visible = False
        SSTab1.TabCaption(1) = "Papiers et alertes"
        Im_ass.Visible = False
        flap = 0
    End If
End Sub

Private Sub T_TAX_Timer()
    If flap2 = 0 Then
        LBL_ALERT_TAXE.Visible = True
        Im_tax.Visible = True
        SSTab1.TabCaption(1) = "....................."
        flap2 = 1
    Else
        LBL_ALERT_TAXE.Visible = False
        Im_tax.Visible = False
        SSTab1.TabCaption(1) = "Papiers et alertes"
        flap2 = 0
    End If
End Sub

Private Sub T_VID_Timer()
    If flap3 = 0 Then
        LBL_ALERT_VIDANGE.Visible = True
        Im_Vid.Visible = True
        LBL_VIDANGE.Visible = True
        flap3 = 1
    Else
        LBL_ALERT_VIDANGE.Visible = False
        Im_Vid.Visible = False
        LBL_VIDANGE.Visible = False
        flap3 = 0
    End If
End Sub

Private Sub T_VIS_Timer()
    If flap1 = 0 Then
        LBL_ALERT_VISITE.Visible = True
        Im_Vis.Visible = True
        SSTab1.TabCaption(1) = "....................."
        flap1 = 1
    Else
        LBL_ALERT_VISITE.Visible = False
        Im_Vis.Visible = False
        SSTab1.TabCaption(1) = "Papiers et alertes"
        flap1 = 0
    End If
End Sub

Private Sub txt_AgenceAssur_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_AgenceAssur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Txt_NvCpt_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub Txt_NvCpt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Txt_NvCpt_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub

Private Sub txt_FourAssur_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_FourAssur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_KlmVidange_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_KlmVidange_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_KlmVidange_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub

Private Sub txt_libelle_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub Txt_Libelle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Matricule_GotFocus()

Call ViderZone(FrmConsultVeh)

LBL_ALERT_ASSURANCE.Visible = False
LBL_ALERT_VIDANGE.Visible = False
LBL_ALERT_VISITE.Visible = False
LBL_ALERT_TAXE.Visible = False

T_ASS.Enabled = False
T_TAX.Enabled = False
T_VID.Enabled = False
T_VIS.Enabled = False

End Sub

Public Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
If KeyCode = vbKeyRight Then Call AfficheRow(txt_Matricule.Text)

End Sub

Private Sub txt_Matricule_LostFocus()

On Error GoTo Err

If Len(Trim(txt_Matricule.Text)) > 0 Then Call AfficheRow(txt_Matricule.Text)

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub txt_Nserie_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_Nserie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_NumAssur_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_NumAssur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Obs_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_Obs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Puissance_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_Puissance_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Type_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_Type_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

'Private Sub Affiche_Lubrifiant()
'
'Dim LOBJ_Lubr As Lubrifiant
'Dim rs As New Recordset
'
'Set LOBJ_Lubr = New Lubrifiant
'Set rs = LOBJ_Lubr.Get_Lubrifiant(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
'If ErrNumber <> 0 Then
'    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
'    ErrNumber = 0
'    Exit Sub
'End If
'If Not rs.EOF Then
'
'    While Not rs.EOF
'        With Cbo_Vidange
'            .AddItem rs("libelle")
'        End With
'        rs.MoveNext
'    Wend
'
'End If
'End Sub
