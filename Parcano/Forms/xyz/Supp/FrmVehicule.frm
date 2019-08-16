VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmVehicule 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Parcano"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmVehicule.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   12060
   Begin VB.Timer T_VID 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   120
   End
   Begin VB.Timer T_TAX 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   120
   End
   Begin VB.Timer T_VIS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   120
   End
   Begin VB.Timer T_ASS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   120
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   1200
      Width           =   615
      _ExtentX        =   1085
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
      Picture         =   "FrmVehicule.frx":0ECA
      ButtonType      =   1
   End
   Begin VB.TextBox txt_Matricule 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
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
      Left            =   2160
      TabIndex        =   0
      Tag             =   "M"
      Top             =   1200
      Width           =   2295
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   10440
      TabIndex        =   1
      Top             =   360
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
      Picture         =   "FrmVehicule.frx":121D
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   9600
      TabIndex        =   3
      Top             =   360
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
      Picture         =   "FrmVehicule.frx":139F
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      Top             =   360
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
      Picture         =   "FrmVehicule.frx":16F2
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   9240
      TabIndex        =   4
      Top             =   360
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
      Picture         =   "FrmVehicule.frx":1A45
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   60
      TabIndex        =   10
      Top             =   1800
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabPicture(0)   =   "FrmVehicule.frx":1BC7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Papier et alerte"
      TabPicture(1)   =   "FrmVehicule.frx":1BE3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6375
         Left            =   0
         TabIndex        =   33
         Top             =   480
         Width           =   11895
         Begin VB.TextBox txt_Nserie 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   60
            Tag             =   "M"
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox txt_Type 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   59
            Tag             =   "M"
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox txt_libelle 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   58
            Tag             =   "M"
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox Cbo_Energie 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5880
            TabIndex        =   57
            Tag             =   "M"
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txt_Puissance 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5880
            MaxLength       =   10
            TabIndex        =   56
            Tag             =   "M"
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txt_Marque 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   55
            Tag             =   "M"
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox txt_TYPCOM 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   54
            Tag             =   "M"
            Top             =   3000
            Width           =   2295
         End
         Begin VB.TextBox txt_Genre 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   53
            Tag             =   "M"
            Top             =   3600
            Width           =   2295
         End
         Begin VB.TextBox txt_Carrosserie 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   52
            Tag             =   "M"
            Top             =   4080
            Width           =   2295
         End
         Begin VB.TextBox txt_PlasseAssise 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   51
            Tag             =   "M"
            Top             =   4680
            Width           =   2295
         End
         Begin VB.TextBox txt_Plassedebout 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   50
            Tag             =   "M"
            Top             =   5280
            Width           =   2295
         End
         Begin VB.TextBox txt_Cylindre 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   49
            Tag             =   "M"
            Top             =   1560
            Width           =   1935
         End
         Begin VB.TextBox txt_NbrEssieux 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   48
            Tag             =   "M"
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox txt_PTAC 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   47
            Tag             =   "M"
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox txt_PoidVide 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   46
            Tag             =   "M"
            Top             =   3360
            Width           =   1935
         End
         Begin VB.TextBox txt_PTRA 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   45
            Tag             =   "M"
            Top             =   3960
            Width           =   1935
         End
         Begin VB.TextBox txt_Charge 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   44
            Tag             =   "M"
            Top             =   4560
            Width           =   1935
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1095
            Left            =   8520
            ScaleHeight     =   1095
            ScaleWidth      =   3255
            TabIndex        =   43
            Top             =   5040
            Width           =   3255
         End
         Begin VB.TextBox txt_Obs 
            Appearance      =   0  'Flat
            Height          =   915
            Left            =   5880
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   5160
            Width           =   2295
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   2010
            Left            =   8400
            ScaleHeight     =   2010
            ScaleWidth      =   3375
            TabIndex        =   35
            Top             =   1080
            Width           =   3375
            Begin VB.TextBox txt_DerCompteurV 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1920
               TabIndex        =   38
               Tag             =   "M"
               Top             =   1560
               Width           =   1455
            End
            Begin VB.TextBox Txt_NvCpt 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   1920
               TabIndex        =   37
               Top             =   0
               Width           =   1455
            End
            Begin VB.TextBox txt_BC 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1920
               TabIndex        =   36
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Compteur vidange :"
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
               TabIndex        =   41
               Top             =   1680
               Width           =   1725
               WordWrap        =   -1  'True
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
               Left            =   0
               TabIndex        =   40
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label7 
               BackColor       =   &H80000009&
               Caption         =   "Compteur Carburant :"
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
               Left            =   0
               TabIndex        =   39
               Top             =   840
               Width           =   2055
            End
         End
         Begin VB.CheckBox ch_Actif 
            BackColor       =   &H80000009&
            Caption         =   "Actif O/ N"
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
            Left            =   8520
            TabIndex        =   34
            Top             =   240
            Width           =   1455
         End
         Begin SToolBox.SDateBox cda_DateCircul 
            Height          =   285
            Left            =   1800
            TabIndex        =   61
            Tag             =   "M"
            Top             =   2520
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SCommand cmdFindenergie 
            Height          =   495
            Left            =   7800
            TabIndex        =   62
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
            Picture         =   "FrmVehicule.frx":1BFF
            ButtonType      =   1
         End
         Begin SToolBox.SDateBox cda_dateSortie 
            Height          =   285
            Left            =   1800
            TabIndex        =   63
            Top             =   5880
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marque :"
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
            TabIndex        =   85
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° serie du type :"
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
            TabIndex        =   84
            Top             =   1920
            Width           =   1440
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   83
            Top             =   2520
            Width           =   1755
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            TabIndex        =   82
            Top             =   1440
            Width           =   510
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   885
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   80
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   79
            Top             =   960
            Width           =   930
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
            Left            =   10560
            TabIndex        =   78
            Top             =   3720
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Im_Vid 
            Height          =   240
            Left            =   10800
            Picture         =   "FrmVehicule.frx":1F52
            Stretch         =   -1  'True
            Top             =   3720
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type commerciale :"
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
            TabIndex        =   77
            Top             =   3120
            Width           =   1650
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Genre :"
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
            TabIndex        =   76
            Top             =   3720
            Width           =   600
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carrosserie :"
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
            TabIndex        =   75
            Top             =   4200
            Width           =   1065
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plasses assises :"
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
            TabIndex        =   74
            Top             =   4680
            Width           =   1380
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Place debout :"
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
            TabIndex        =   73
            Top             =   5280
            Width           =   1185
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   72
            Top             =   4080
            Width           =   675
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   71
            Top             =   2280
            Width           =   1440
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   70
            Top             =   2880
            Width           =   705
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   69
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   68
            Top             =   1680
            Width           =   885
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   67
            Top             =   4680
            Width           =   1110
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date cession :"
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
            TabIndex        =   66
            Top             =   5880
            Width           =   1170
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
            Height          =   1335
            Left            =   8760
            TabIndex        =   65
            Top             =   3600
            Visible         =   0   'False
            Width           =   1455
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observation"
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
            Left            =   4440
            TabIndex        =   64
            Top             =   5280
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Papier et alerte"
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   -74940
         TabIndex        =   11
         Top             =   480
         Width           =   10215
         Begin VB.TextBox txt_NumAssur 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   14
            Tag             =   "M"
            Top             =   360
            Width           =   4080
         End
         Begin VB.TextBox txt_AgenceAssur 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   13
            Tag             =   "M"
            Top             =   1800
            Width           =   4095
         End
         Begin VB.TextBox txt_FourAssur 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   12
            Tag             =   "M"
            Top             =   1080
            Width           =   4095
         End
         Begin SToolBox.SDateBox cda_DebuAssur 
            Height          =   285
            Left            =   1560
            TabIndex        =   15
            Tag             =   "M"
            Top             =   2400
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_FinAssur 
            Height          =   285
            Left            =   3480
            TabIndex        =   16
            Tag             =   "M"
            Top             =   2400
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_DebutVesite 
            Height          =   285
            Left            =   1560
            TabIndex        =   17
            Top             =   3120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_FinVisite 
            Height          =   285
            Left            =   3480
            TabIndex        =   18
            Top             =   3120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_debuttax 
            Height          =   285
            Left            =   1560
            TabIndex        =   19
            Top             =   3840
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_fintax 
            Height          =   285
            Left            =   3480
            TabIndex        =   20
            Top             =   3840
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Image Im_tax 
            Height          =   240
            Left            =   5040
            Picture         =   "FrmVehicule.frx":225C
            Stretch         =   -1  'True
            Top             =   3840
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Im_Vis 
            Height          =   240
            Left            =   5040
            Picture         =   "FrmVehicule.frx":2566
            Stretch         =   -1  'True
            Top             =   3120
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Im_ass 
            Height          =   240
            Left            =   5040
            Picture         =   "FrmVehicule.frx":2870
            Stretch         =   -1  'True
            Top             =   2400
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
            Left            =   4800
            TabIndex        =   32
            Top             =   3840
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
            Left            =   4800
            TabIndex        =   31
            Top             =   3120
            Visible         =   0   'False
            Width           =   195
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
            Left            =   4800
            TabIndex        =   30
            Top             =   2400
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label Label15 
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
            Left            =   120
            TabIndex        =   29
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label16 
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
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label17 
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
            Left            =   120
            TabIndex        =   27
            Top             =   1800
            Width           =   720
         End
         Begin VB.Label Label18 
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
            Left            =   120
            TabIndex        =   26
            Top             =   2520
            Width           =   750
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            Height          =   195
            Left            =   3000
            TabIndex        =   25
            Top             =   2400
            Width           =   300
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            Height          =   195
            Left            =   3000
            TabIndex        =   24
            Top             =   3120
            Width           =   300
         End
         Begin VB.Label Label24 
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
            Left            =   120
            TabIndex        =   23
            Top             =   3120
            Width           =   810
         End
         Begin VB.Label Label26 
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
            Left            =   120
            TabIndex        =   22
            Top             =   3840
            Width           =   765
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            Height          =   195
            Left            =   3000
            TabIndex        =   21
            Top             =   3840
            Width           =   300
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche véhicule"
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
      Left            =   840
      TabIndex        =   9
      Top             =   480
      Width           =   2355
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
      Left            =   480
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image PicBox_Header 
      Height          =   1455
      Left            =   0
      Picture         =   "FrmVehicule.frx":2B7A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Der.compteur vid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4440
      TabIndex        =   7
      Top             =   6360
      Width           =   1365
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Compteur Carburant"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   720
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   2580
   End
End
Attribute VB_Name = "FrmVehicule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flap
Dim flap1
Dim flap2
Dim flap3
Public AA As String

Public Sub AfficheRow(ByVal vCode As String)

Dim LOBJ_BonVidange As BonVidange
Dim LObj_vehicule As Vehicule
Dim LOBJ_Personnel As Personnel
Dim LOBJ_Lub As Lubrifiant
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim AA As Long

On Error GoTo Err
Call ViderZone(Frm_Vehicule)

Set LOBJ_Personnel = New Personnel
'Verifier le droit d'accès pour insertion d'un bonVidange
If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Maj_Compt", LInt_UserId) Then
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If

Set LObj_vehicule = New Vehicule
AA = 0
Set rs = LObj_vehicule.GetVehicule(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vCode)
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
        Txt_NvCpt.Text = CompteurVehicule(rs("Matricule")) 'dernier compt saisi dans la fiche trafic
    End If
    If Not IsNull(rs("marque")) Then txt_Type.Text = rs("TYPE")
    If Not IsNull(rs("puissance")) Then txt_Marque.Text = rs("marque")
    If Not IsNull(rs("Matricule")) Then txt_Puissance.Text = rs("puissance")
    If Not IsNull(rs("Energie")) Then Cbo_Energie.Text = rs("Energie")

    If Not IsNull(rs("NumSerie")) Then txt_Nserie.Text = rs("NumSerie")
    If Not IsNull(rs("DateCircul")) Then cda_DateCircul.Text = rs("DateCircul")
    If Not IsNull(rs("NumAssur")) Then txt_NumAssur.Text = rs("NumAssur")
    If Not IsNull(rs("FournisAssur")) Then txt_FourAssur.Text = rs("FournisAssur")
    If Not IsNull(rs("AgenceAssur")) Then txt_AgenceAssur.Text = rs("AgenceAssur")
    
    If Not IsNull(rs("DateDebAssur")) And rs("DateDebAssur") <> "01/01/1900" Then cda_DebuAssur.Text = rs("DateDebAssur")
    If Not IsNull(rs("DateFinAssur")) And rs("DateFinAssur") <> "01/01/1900" Then cda_FinAssur.Text = rs("DAteFinAssur")
    
    If Not IsNull(rs("CompteurCarburant")) Then txt_BC.Text = rs("CompteurCarburant")
    If Not IsNull(rs("CompteurFT")) Then
        If Val(Txt_NvCpt.Text) < Val(rs("CompteurFT")) Then
            Txt_NvCpt.Text = rs("CompteurFT")
'le compteur saisi dans la fiche traffic > à celui saissi dans véhicule (Compteur vidange)
        End If
    End If
    
    Set LOBJ_BonVidange = New BonVidange
    Set rs1 = LOBJ_BonVidange.Get_DerBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("Code"))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs1.EOF Then
        If Not IsNull(rs1("CompteurVidange")) Then txt_DerCompteurV.Text = rs1("CompteurVidange")
        'Alert vidange
        If txt_DerCompteurV.Text <> 0 Or txt_DerCompteurV.Text <> "" Then
            AA = Val(Txt_NvCpt.Text) - Val(txt_DerCompteurV.Text)
            If AA + 500 >= rs1("NBKLMvid") Then
                If AA > rs1("NBKLMvid") Then
                    LBL_VIDANGE.Caption = "Vous avez dépassé le nbr de klm de vidange par  : " & AA - rs1("NBKLMvid")
                Else
                    LBL_VIDANGE.Caption = "Il vous reste " & rs1("NBKLMvid") - AA & " klm pour le nouveau vidange"
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
    If Not IsNull(rs("Actif")) Then ch_Actif.Value = rs("Actif")
    
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
    
    If Not IsNull(rs("DateSortie")) Then cda_dateSortie.Text = rs("DateSortie")
    If Not IsNull(rs("Obs")) Then txt_Obs.Text = rs("Obs")
Else
    MsgBox "Code introuvable", vbInformation
    txt_Matricule.SetFocus
End If
rs.Close
Exit Sub

Err:
MsgBox Err.Description, vbInformation
End Sub

'Dernier compteur du véhicule en entrant (ficheTraffic)
Public Function CompteurVehicule(ByVal vCode As String) As String

Dim LObj_vehicule As Vehicule
Dim rs1 As New Recordset

CompteurVehicule = "0"
Set LObj_vehicule = New Vehicule
Set rs1 = LObj_vehicule.Get_DerCompt(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vCode)
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

Private Sub Cbo_Energie_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub Cbo_Energie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cda_DateCircul_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cda_dateSortie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_BC_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_BV_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_Carrosserie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Charge_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Cylindre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_FT_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_Genre_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Cda_DateCarteGris_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
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

Private Sub CmdAdd_Click()

Dim LOBJ_Personnel As Personnel
On Error GoTo Err

Set LOBJ_Personnel = New Personnel
'Verifier le droit d'accès pour insertion d'un bonVidange
If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Ins_Vehicule", LInt_UserId) Then
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If

If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Call ViderZone(Frm_Vehicule)
txt_BV.Text = 0
txt_BC.Text = 0
txt_FT.Text = 0
txt_Matricule.Text = "Auto"
txt_libelle.SetFocus
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()

Dim rs As New Recordset
Dim vCode
Dim LOBJ_Personnel As Personnel
Dim LObj_vehicule As Vehicule
On Error GoTo Err

Set LOBJ_Personnel = New Personnel
If txt_Matricule.Text <> "Auto" Then
    'Verifier le droit d'accès pour insertion d'un bonVidange
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Supp_vehicule", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
End If
If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        txt_Matricule.SetFocus
        Exit Sub
    End If
End If

vCode = txt_Matricule.Text
Set LObj_vehicule = New Vehicule
If MsgBox("Confirmez vous la suppression de cette " & vbNewLine & "véhicule", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    Call LObj_vehicule.Delete_Vehicule(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    txt_Matricule.SetFocus
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdFind_Click()

On Error Resume Next

If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Unload FrmFind
With FrmFind
    .StrSource = "Véhicule"
    .Show vbModal
End With

End Sub

Private Sub cmdFindenergie_Click()

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Energie"
    .Show vbModal
End With
End Sub

Private Sub cmdFindMarque_Click()
Unload FrmFind
With FrmFind
    .StrSource = "Marque"
    .Show vbModal
End With
End Sub

Private Sub cmdFindMatricule_Click()
Unload FrmFind
With FrmFind_Actif
    .StrSource = "Véhicule"
    .Show vbModal
End With
End Sub

Private Sub cmdFindVidange_Click()

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Lubrifiant"
    .Show vbModal
End With

End Sub

Private Sub CmdSave_Click()

Dim LOBJ_Personnel As Personnel
Dim LInt_NumCompteur As Long
On Error GoTo Err

  If Left(CheckMandatory(Frm_Vehicule), 1) = 1 Then
     Exit Sub
  End If

  If CDate(cda_FinAssur.Text) < CDate(cda_DebuAssur.Text) Then
      MsgBox "Date fin invalide", vbInformation
      cda_FinAssur.SetFocus
      Exit Sub
  End If
  
  If CDate(cda_FinVisite.Text) < CDate(cda_DebuAssur.Text) Then
      MsgBox "Date fin invalide", vbInformation
      cda_FinVisite.SetFocus
      Exit Sub
  End If
  If CDate(cda_fintax.Text) < CDate(cda_debuttax.Text) Then
      MsgBox "Date fin invalide", vbInformation
      cda_fintax.SetFocus
      Exit Sub
  End If
  
If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
  Set LOBJ_Personnel = New Personnel
If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
    'Verifier le droit d'accès pour MAJ d'un vehicule
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Maj_vehicule", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
    Call Modif_Veh

ElseIf txt_Matricule.Text = "Auto" Then
    LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_Vehicule")
    If ErrNumber <> 0 Then
       MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
       ErrNumber = 0
       Exit Sub
    End If
    Set LObj_Compteur = Nothing
    'Insertion enregistrement assiette
    txt_Matricule.Text = Format(LInt_NumCompteur, "00000")
    Call Ajout_Vehicule
End If
  
  MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
  txt_Matricule.SetFocus
  
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub Ajout_Vehicule()

Dim LObj_vehicule As Vehicule
Dim LRs_NewRecord As New Recordset

Set LObj_vehicule = New Vehicule
Set LRs_NewRecord = Remplir_Recordset

Call LObj_vehicule.Insert_Veh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing
End Sub

Private Function Remplir_Recordset() As Recordset

Dim LRs_NewRecord As New Recordset

Set LRs_NewRecord = CreateEmptyRS_Vehicule

With LRs_NewRecord
    .AddNew
    .Fields("Code") = txt_Matricule.Text
    .Fields("Matricule") = txt_libelle.Text
    .Fields("TYPE") = txt_Type.Text
    .Fields("Marque") = txt_Marque.Text
    .Fields("Puissance") = txt_Puissance.Text
    .Fields("Energie") = Cbo_Energie.Text
    .Fields("genre") = txt_Genre.Text
    .Fields("Carrosserie") = txt_Carrosserie.Text
    .Fields("PlaceAssis") = txt_PlasseAssise.Text
    .Fields("PlaceDebout") = txt_Plassedebout.Text
    .Fields("Cylindre") = txt_Cylindre.Text
    .Fields("NbrEssieux") = txt_NbrEssieux.Text
    .Fields("PTAC") = txt_PTAC.Text
    .Fields("PTRA") = txt_PTRA.Text
    .Fields("Charge") = txt_Charge.Text
    .Fields("PoidsVide") = txt_PoidVide.Text
    .Fields("TypeComm") = txt_TYPCOM.Text
    .Fields("NumSerie") = txt_Nserie.Text
    .Fields("DateCircul") = CDate(cda_DateCircul.Text)
    .Fields("NumAssur") = txt_NumAssur.Text
    .Fields("FournisAssur") = txt_FourAssur.Text
    .Fields("AgenceAssur") = txt_AgenceAssur.Text
    .Fields("DateDebAssur") = CDate(cda_DebuAssur.Text)
    .Fields("DAteFinAssur") = CDate(cda_FinAssur.Text)
    .Fields("DateDebVisite") = CDate(cda_DebutVesite.Text)
    .Fields("DAteFinVisite") = CDate(cda_FinVisite.Text)
    .Fields("DateDebTax") = CDate(cda_debuttax.Text)
    .Fields("DateFinTax") = CDate(cda_fintax.Text)
    .Fields("DateSortie") = CDate(cda_dateSortie.Text)
    .Fields("Obs") = txt_Obs.Text
    .Fields("CompteurVidange") = txt_BV.Text
    .Fields("CompteurCarburant") = txt_BC.Text
    .Fields("CompteurFT") = txt_FT.Text
    .Fields("actif") = ch_Actif.Value
    .Fields("Disponible") = "O"
End With
Set Remplir_Recordset = LRs_NewRecord
End Function
 
Private Sub Modif_Veh()
 
Dim LObj_vehicule As Vehicule
Dim LRs_NewRecord As New Recordset

Set LObj_vehicule = New Vehicule
Set LRs_NewRecord = Remplir_Recordset

Call LObj_vehicule.Update_Veh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

End Sub

Private Sub Form_Load()
On Error GoTo Err

Me.Width = 10560
Me.Height = 8565
Me.Move 0, 0
Me.WindowState = 2
Call Affiche_Energie
If AA <> "" Then Call AfficheRow(AA)
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub


Private Sub Form_Resize()

On Error Resume Next
Dim WidthForm As Integer
WidthForm = Frm_Main.ACB_Main.Width
PicBox_Header.Width = WidthForm - 1000
CmdAdd.Left = WidthForm - 5500
CmdDelete.Left = WidthForm - 5100
CmdFind.Left = WidthForm - 4700
CmdSave.Left = WidthForm - 4300
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
   Unload Frm_Vehicule
   End If
   
   Exit Sub
erreur:
   MsgBox Err.Description, 48
End Sub

Private Sub T_ASS_Timer()
    If flap = 0 Then
        LBL_ALERT_ASSURANCE.Visible = True
        Im_ass.Visible = True
        flap = 1
    Else
        LBL_ALERT_ASSURANCE.Visible = False
        Im_ass.Visible = False
        flap = 0
    End If
End Sub

Private Sub T_TAX_Timer()
    If flap2 = 0 Then
        LBL_ALERT_TAXE.Visible = True
        Im_tax.Visible = True
        flap2 = 1
    Else
        LBL_ALERT_TAXE.Visible = False
        Im_tax.Visible = False
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
        flap1 = 1
    Else
        LBL_ALERT_VISITE.Visible = False
        Im_Vis.Visible = False
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

Private Sub txt_compteur_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° immatriculation obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_compteur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_compteur_KeyPress(KeyAscii As Integer)

On Error Resume Next
If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then
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

Call ViderZone(Frm_Vehicule)

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
End Sub

Private Sub txt_Matricule_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_Matricule_LostFocus()

On Error GoTo Err

If Len(Trim(txt_Matricule.Text)) > 0 Then Call AfficheRow(txt_Matricule.Text)

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub txt_NbrEssieux_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
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

Private Sub Txt_NvCpt_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
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

Private Sub txt_PlasseAssise_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_PlasseAssise_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_Plassedebout_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Plassedebout_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_PoidVide_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_PTAC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_PTRA_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub txt_TYPCOM_KeyDown(KeyCode As Integer, Shift As Integer)
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
