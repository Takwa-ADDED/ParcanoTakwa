VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Vehicule 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Parcano"
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Vehicule.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   13620
   Begin VB.CommandButton Cmd_Alertes 
      Caption         =   "Alertes"
      Height          =   615
      Left            =   9120
      TabIndex        =   119
      Top             =   480
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   7920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      TabIndex        =   26
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
      Picture         =   "Frm_Vehicule.frx":0ECA
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   27
      Top             =   1800
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13996
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
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
      TabPicture(0)   =   "Frm_Vehicule.frx":121D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl_Supp"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Cmd_Supp"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Papier et alerte"
      TabPicture(1)   =   "Frm_Vehicule.frx":1239
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7455
         Left            =   0
         TabIndex        =   56
         Top             =   360
         Width           =   11895
         Begin MSComctlLib.ListView grid 
            Height          =   1455
            Left            =   8040
            TabIndex        =   118
            Top             =   720
            Width           =   3135
            _ExtentX        =   5530
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
            Appearance      =   0
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Numero"
               Object.Width           =   1588
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Lubrifiant"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "THT"
               Object.Width           =   1588
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "TVA"
               Object.Width           =   1588
            EndProperty
         End
         Begin SToolBox.SDateBox cda_DateSortie 
            Height          =   285
            Left            =   1800
            TabIndex        =   88
            Top             =   5520
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.TextBox txt_klmVidange 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   117
            Tag             =   "M"
            Top             =   3120
            Width           =   1935
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   7920
            ScaleHeight     =   585
            ScaleWidth      =   4065
            TabIndex        =   92
            Top             =   6600
            Width           =   4095
            Begin VB.Label Label42 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ": Champ(s) Obligatoire..."
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   360
               TabIndex        =   96
               Top             =   0
               Width           =   2055
            End
            Begin VB.Label Label38 
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Perpetua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label37 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ": De 2 à 8 Lettres Maximum / Sans Espace..."
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   360
               TabIndex        =   94
               Top             =   240
               Width           =   3615
            End
            Begin VB.Label Label8 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Perpetua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.TextBox Txt_Abr 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   2
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox Cbo_Lub 
            Height          =   315
            Left            =   8760
            TabIndex        =   25
            Top             =   360
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker cda_DateCircul 
            Height          =   375
            Left            =   1920
            TabIndex        =   6
            Top             =   2640
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            Format          =   94830593
            CurrentDate     =   42857
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Séléctionner une photo ..."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3135
            Left            =   8280
            TabIndex        =   83
            Top             =   2400
            Width           =   3495
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   375
               Left            =   240
               ScaleHeight     =   315
               ScaleWidth      =   3075
               TabIndex        =   84
               Top             =   240
               Width           =   3135
               Begin VB.TextBox Txt_Photo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  BeginProperty Font 
                     Name            =   "MS Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   0
                  MaxLength       =   30
                  TabIndex        =   85
                  Tag             =   "M"
                  Top             =   0
                  Width           =   3015
               End
            End
            Begin SToolBox.SCommand Cmd_Photo 
               Height          =   375
               Left            =   1440
               TabIndex        =   29
               Top             =   2640
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               Caption         =   "Parcourir..."
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   8421504
            End
            Begin VB.Image Img_Vehicule 
               Height          =   1575
               Left            =   240
               Picture         =   "Frm_Vehicule.frx":1255
               Stretch         =   -1  'True
               Top             =   840
               Width           =   1935
            End
         End
         Begin VB.TextBox txt_NbrEssieux 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "M"
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txt_PTAC 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   16
            Tag             =   "M"
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox txt_PoidVide 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   17
            Tag             =   "M"
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox txt_PTRA 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   18
            Tag             =   "M"
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox txt_Charge 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   19
            Tag             =   "M"
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txt_Cylindre 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "M"
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txt_Puissance 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   13
            Tag             =   "M"
            Top             =   6480
            Width           =   2295
         End
         Begin VB.TextBox txt_Nserie 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   5
            Tag             =   "M"
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox txt_Type 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   4
            Tag             =   "M"
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox txt_libelle 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   1
            Tag             =   "M"
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox Cbo_Energie 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Tag             =   "M"
            Top             =   6000
            Width           =   2055
         End
         Begin VB.TextBox txt_Marque 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   3
            Tag             =   "M"
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox txt_TYPCOM 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   7
            Tag             =   "M"
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox txt_Genre 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   8
            Tag             =   "M"
            Top             =   3600
            Width           =   2295
         End
         Begin VB.TextBox txt_Carrosserie 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   9
            Tag             =   "M"
            Top             =   4080
            Width           =   2295
         End
         Begin VB.TextBox txt_PlasseAssise 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "M"
            Top             =   4560
            Width           =   2295
         End
         Begin VB.TextBox txt_Plassedebout 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "M"
            Top             =   5040
            Width           =   2295
         End
         Begin VB.TextBox txt_Obs 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   1635
            Left            =   5880
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   3600
            Width           =   2175
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1530
            Left            =   4320
            ScaleHeight     =   1530
            ScaleWidth      =   3975
            TabIndex        =   57
            Top             =   5400
            Width           =   3975
            Begin VB.TextBox txt_DerCompteurV 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               TabIndex        =   23
               Tag             =   "M"
               Top             =   960
               Width           =   1935
            End
            Begin VB.TextBox Txt_NvCpt 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               TabIndex        =   21
               Top             =   0
               Width           =   1935
            End
            Begin VB.TextBox txt_BC 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               TabIndex        =   22
               Top             =   480
               Width           =   1935
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
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   120
               TabIndex        =   60
               Top             =   1080
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
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label7 
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H00808080&
               Height          =   375
               Left            =   120
               TabIndex        =   58
               Top             =   600
               Width           =   2055
            End
         End
         Begin VB.CheckBox ch_Actif 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   5160
            TabIndex        =   24
            Top             =   6840
            Width           =   1455
         End
         Begin SToolBox.SCommand cmdFindenergie 
            Height          =   495
            Left            =   3840
            TabIndex        =   28
            Top             =   6000
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
            Picture         =   "Frm_Vehicule.frx":532D
            ButtonType      =   1
         End
         Begin SToolBox.SCommand Cmd_ok 
            Height          =   375
            Left            =   11280
            TabIndex        =   89
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Caption         =   "OK"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   33023
         End
         Begin SToolBox.SCommand Cmd_Annul 
            Height          =   375
            Left            =   11280
            TabIndex        =   90
            Top             =   960
            Width           =   495
            _ExtentX        =   873
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
            Picture         =   "Frm_Vehicule.frx":5680
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "KM Vidange :"
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
            Left            =   4440
            TabIndex        =   116
            Top             =   3240
            Width           =   1155
         End
         Begin VB.Label Label65 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " *"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   840
            TabIndex        =   111
            Top             =   6000
            Width           =   255
         End
         Begin VB.Label Label64 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " *"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1560
            TabIndex        =   110
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label63 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " *"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1680
            TabIndex        =   109
            Top             =   2640
            Width           =   255
         End
         Begin VB.Label Label57 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " *"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   960
            TabIndex        =   108
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   4150
            TabIndex        =   98
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label43 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " *"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1190
            TabIndex        =   97
            Top             =   760
            Width           =   255
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Abréviation :"
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
            TabIndex        =   91
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label Lbl_Lubr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Lubrifiant :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7920
            TabIndex        =   87
            Top             =   360
            Width           =   975
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
            TabIndex        =   82
            Top             =   1320
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
            TabIndex        =   81
            Top             =   2280
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
            TabIndex        =   80
            Top             =   2760
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
            TabIndex        =   79
            Top             =   1800
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
            TabIndex        =   78
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
            Left            =   120
            TabIndex        =   77
            Top             =   6080
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
            Left            =   120
            TabIndex        =   76
            Top             =   6520
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
            Left            =   11040
            TabIndex        =   75
            Top             =   5880
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Im_Vid 
            Height          =   240
            Left            =   11400
            Picture         =   "Frm_Vehicule.frx":59D3
            Stretch         =   -1  'True
            Top             =   5880
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
            TabIndex        =   74
            Top             =   3240
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
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   71
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
            TabIndex        =   70
            Top             =   5160
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
            TabIndex        =   69
            Top             =   2280
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
            TabIndex        =   68
            Top             =   720
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
            TabIndex        =   67
            Top             =   1320
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
            TabIndex        =   66
            Top             =   1800
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
            TabIndex        =   65
            Top             =   360
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
            TabIndex        =   64
            Top             =   2760
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
            TabIndex        =   63
            Top             =   5600
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
            Height          =   1095
            Left            =   8280
            TabIndex        =   62
            Top             =   5520
            Visible         =   0   'False
            Width           =   2535
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
            TabIndex        =   61
            Top             =   3600
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Papier et alerte"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7215
         Left            =   -75000
         TabIndex        =   43
         Top             =   360
         Width           =   11895
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   7920
            ScaleHeight     =   345
            ScaleWidth      =   4065
            TabIndex        =   99
            Top             =   6840
            Width           =   4095
            Begin VB.Label Label50 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Perpetua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   101
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label45 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ": Champ(s) Obligatoire..."
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   360
               TabIndex        =   100
               Top             =   0
               Width           =   2055
            End
         End
         Begin MSComCtl2.DTPicker cda_DebuAssur 
            Height          =   375
            Left            =   1680
            TabIndex        =   33
            Top             =   2880
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632256
            Format          =   147783681
            CurrentDate     =   42858
         End
         Begin MSComCtl2.DTPicker cda_DebutVesite 
            Height          =   375
            Left            =   1680
            TabIndex        =   35
            Top             =   3600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632256
            Format          =   147783681
            CurrentDate     =   42858
         End
         Begin MSComCtl2.DTPicker cda_FinAssur 
            Height          =   375
            Left            =   4080
            TabIndex        =   34
            Top             =   2880
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632256
            Format          =   147783681
            CurrentDate     =   42858
         End
         Begin MSComCtl2.DTPicker cda_fintax 
            Height          =   375
            Left            =   4080
            TabIndex        =   38
            Top             =   4320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632256
            Format          =   147783681
            CurrentDate     =   42858
         End
         Begin MSComCtl2.DTPicker cda_debuttax 
            Height          =   375
            Left            =   1680
            TabIndex        =   37
            Top             =   4320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632256
            Format          =   147783681
            CurrentDate     =   42858
         End
         Begin MSComCtl2.DTPicker cda_FinVisite 
            Height          =   375
            Left            =   4080
            TabIndex        =   36
            Top             =   3600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632256
            Format          =   147783681
            CurrentDate     =   42858
         End
         Begin VB.TextBox txt_NumAssur 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   1695
            MaxLength       =   50
            TabIndex        =   30
            Tag             =   "M"
            Top             =   720
            Width           =   4080
         End
         Begin VB.TextBox txt_AgenceAssur 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   32
            Tag             =   "M"
            Top             =   2160
            Width           =   4095
         End
         Begin VB.TextBox txt_FourAssur 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   31
            Tag             =   "M"
            Top             =   1440
            Width           =   4095
         End
         Begin VB.Label Label56 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " *"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   960
            TabIndex        =   107
            Top             =   4320
            Width           =   255
         End
         Begin VB.Label Label55 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1080
            TabIndex        =   106
            Top             =   3600
            Width           =   255
         End
         Begin VB.Label Label54 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " *"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   960
            TabIndex        =   105
            Top             =   2880
            Width           =   255
         End
         Begin VB.Label Label53 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   960
            TabIndex        =   104
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label52 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " *"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1080
            TabIndex        =   103
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label51 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " *"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1200
            TabIndex        =   102
            Top             =   720
            Width           =   255
         End
         Begin VB.Image Im_tax 
            Height          =   240
            Left            =   6000
            Picture         =   "Frm_Vehicule.frx":5CDD
            Stretch         =   -1  'True
            Top             =   4320
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Im_Vis 
            Height          =   240
            Left            =   6000
            Picture         =   "Frm_Vehicule.frx":5FE7
            Stretch         =   -1  'True
            Top             =   3600
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Im_ass 
            Height          =   240
            Left            =   6000
            Picture         =   "Frm_Vehicule.frx":62F1
            Stretch         =   -1  'True
            Top             =   2880
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
            Left            =   5760
            TabIndex        =   55
            Top             =   4320
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
            Left            =   5760
            TabIndex        =   54
            Top             =   3600
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
            Left            =   5760
            TabIndex        =   53
            Top             =   2880
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contrat N°  :"
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
            TabIndex        =   52
            Top             =   840
            Width           =   1020
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
            Left            =   240
            TabIndex        =   51
            Top             =   1560
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
            Left            =   240
            TabIndex        =   50
            Top             =   2160
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
            Left            =   240
            TabIndex        =   49
            Top             =   3000
            Width           =   750
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
            Left            =   3585
            TabIndex        =   48
            Top             =   3000
            Width           =   315
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
            Left            =   3585
            TabIndex        =   47
            Top             =   3720
            Width           =   315
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
            Left            =   240
            TabIndex        =   46
            Top             =   3720
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
            Left            =   240
            TabIndex        =   45
            Top             =   4440
            Width           =   765
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
            Left            =   3585
            TabIndex        =   44
            Top             =   4440
            Width           =   315
         End
      End
      Begin VB.Image Cmd_Supp 
         Height          =   375
         Left            =   10440
         Picture         =   "Frm_Vehicule.frx":65FB
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Lbl_Supp 
         BackStyle       =   0  'Transparent
         Caption         =   "=> Véhicule est supprimé, Voulez-Vous ré-ajouter?..."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5400
         TabIndex        =   86
         Top             =   0
         Width           =   5535
      End
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   12000
      TabIndex        =   113
      Top             =   600
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
      Picture         =   "Frm_Vehicule.frx":1831D
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   11280
      TabIndex        =   114
      Top             =   600
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
      Picture         =   "Frm_Vehicule.frx":1849F
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   11640
      TabIndex        =   115
      Top             =   600
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
      Picture         =   "Frm_Vehicule.frx":187F2
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   10920
      TabIndex        =   112
      Top             =   600
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
      Picture         =   "Frm_Vehicule.frx":18B45
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
      TabIndex        =   42
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
      TabIndex        =   41
      Top             =   1320
      Width           =   1575
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
      TabIndex        =   40
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
      TabIndex        =   39
      Top             =   3360
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image PicBox_Header 
      Height          =   1455
      Left            =   0
      Picture         =   "Frm_Vehicule.frx":18CC7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15735
   End
End
Attribute VB_Name = "Frm_Vehicule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim StrPicture  As String
    Dim ChangPic    As Boolean
    Public AA       As String
    Dim thekey      As Integer
    Dim theshift    As Integer
    Dim flap
    Dim flap1
    Dim flap2
    Dim flap3

Private Sub Cmd_Alertes_Click()
Frm_gestion_alertes.Show
End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    SSTab1.Tab = 0
    Cmd_Supp.Visible = False
    CmdSave.Enabled = False
    Lbl_Supp.Visible = False
    CmdDelete.Enabled = False
    Call InitialiseDate
    Call Affiche_Energie
    If AA <> "" Then Call AfficheRow(AA)
    Call Affiche_Lubrif_Combo(Cbo_Lub)
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Dim WidthForm   As Integer
    WidthForm = Frm_Main.ACB_Main.Width
    PicBox_Header.Width = WidthForm - 1000
    Cmd_Alertes.Left = WidthForm - 7300
    CmdAdd.Left = WidthForm - 5500
    CmdDelete.Left = WidthForm - 5100
    CmdFind.Left = WidthForm - 4700
    CmdSave.Left = WidthForm - 4300
End Sub
Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo Err
'    If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) _
'        = vbNo Then Cancel = True Else Unload Frm_Vehicule
'Exit Sub
'Err:
'   MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub



Private Sub Picture5_Click()

End Sub

Private Sub txt_klmVidange_KeyPress(KeyAscii As Integer)
If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txt_Nserie_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° immatriculation obligatoire      ", vbInformation
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_NumAssur_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° immatriculation obligatoire      ", vbInformation
        txt_Matricule.SetFocus
    End If
End Sub

Private Sub Txt_NvCpt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Cbo_Lub.SetFocus
End Sub

Private Sub txt_Obs_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° immatriculation obligatoire      ", vbInformation
        txt_Matricule.SetFocus
    End If
End Sub

Private Sub txt_Puissance_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° immatriculation obligatoire      ", vbInformation
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_Type_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° immatriculation obligatoire      ", vbInformation
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_AgenceAssur_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° immatriculation obligatoire      ", vbInformation
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_FourAssur_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° immatriculation obligatoire      ", vbInformation
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_Libelle_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° immatriculation obligatoire      ", vbInformation
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub Cbo_Energie_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° immatriculation obligatoire      ", vbInformation
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_marque_GotFocus()
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
Private Sub Txt_Abr_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° immatriculation obligatoire      ", vbInformation
        txt_Matricule.SetFocus
    End If
End Sub

Private Sub txt_AgenceAssur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cda_DebuAssur.SetFocus
End Sub
Private Sub txt_FourAssur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_AgenceAssur.SetFocus
End Sub
Public Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    On Error GoTo Err
    If Len(Trim(txt_Matricule.Text)) > 0 And Trim(txt_Matricule.Text) <> "Auto" Then Call AfficheRow(txt_Matricule.Text)
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
    End If
End Sub
Private Sub Txt_Libelle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Txt_Abr.SetFocus
    End Sub
Private Sub txt_Nserie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cda_DateCircul.SetFocus
End Sub
Private Sub txt_NumAssur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_FourAssur.SetFocus
End Sub
Private Sub txt_NbrEssieux_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_PTAC.SetFocus
End Sub
Private Sub txt_Obs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Txt_NvCpt.SetFocus
End Sub
Private Sub txt_PlasseAssise_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_Plassedebout.SetFocus
End Sub
Private Sub txt_Plassedebout_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cda_DateSortie.SetFocus
End Sub
Private Sub txt_PoidVide_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_PTRA.SetFocus
End Sub
Private Sub txt_PTAC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_PoidVide.SetFocus
End Sub
Private Sub txt_PTRA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_Charge.SetFocus
End Sub
Private Sub txt_Puissance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_Cylindre.SetFocus
End Sub
Private Sub txt_TYPCOM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_Genre.SetFocus
End Sub
Private Sub txt_Type_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_Nserie.SetFocus
End Sub
Private Sub Cbo_Energie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_Puissance.SetFocus
End Sub
Private Sub cda_DateCircul_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_TYPCOM.SetFocus
End Sub
Private Sub cda_dateSortie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Cbo_Energie.SetFocus
End Sub
Private Sub txt_Carrosserie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_PlasseAssise.SetFocus
End Sub
Private Sub txt_Charge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_klmVidange.SetFocus
End Sub

Private Sub txt_KlmVidange_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then txt_Obs.SetFocus
End Sub
Private Sub txt_Cylindre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_NbrEssieux.SetFocus
End Sub
Private Sub txt_Genre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_Carrosserie.SetFocus
End Sub
Private Sub txt_marque_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_Type.SetFocus
End Sub
Private Sub Txt_Abr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt_Marque.SetFocus
End Sub
Private Sub Txt_Abr_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Len(Txt_Abr.Text) > 7 And KeyAscii <> 8 And KeyAscii <> 127 Then KeyAscii = 0
    If Not (Chr(KeyAscii) Like "[0123456789AZERTYUIOPQSDFGHJKLMWXCVBNazertyuiopqsdfghjklmwxcvbn.]") _
        And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Txt_NvCpt_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txt_PlasseAssise_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txt_Plassedebout_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txt_Matricule_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txt_BC_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
'# ControlBox
Private Sub txt_libelle_Change()
    Txt_Photo = txt_libelle.Text
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
Private Sub txt_Matricule_LostFocus()
On Error GoTo Err
    If Len(Trim(txt_Matricule.Text)) > 0 And Trim(txt_Matricule.Text) <> "Auto" Then Call AfficheRow(txt_Matricule.Text)
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
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
Private Sub EnbDisb(ByVal TYP As Boolean)
   ' txt_Matricule.Enabled = TYP
    txt_libelle.Enabled = TYP
    Txt_Abr.Enabled = TYP
    txt_Marque.Enabled = TYP
    txt_Type.Enabled = TYP
    txt_Nserie.Enabled = TYP
    cda_DateCircul.Enabled = TYP
    txt_TYPCOM.Enabled = TYP
    txt_Genre.Enabled = TYP
    txt_Carrosserie.Enabled = TYP
    txt_PlasseAssise.Enabled = TYP
    txt_Plassedebout.Enabled = TYP
    cda_DateSortie.Enabled = TYP
    Cbo_Energie.Enabled = TYP
    cmdFindenergie.Enabled = TYP
    txt_Puissance.Enabled = TYP
    txt_Cylindre.Enabled = TYP
    txt_NbrEssieux.Enabled = TYP
    txt_PTAC.Enabled = TYP
    txt_PoidVide.Enabled = TYP
    txt_PTRA.Enabled = TYP
    txt_Charge.Enabled = TYP
    txt_klmVidange.Enabled = TYP
    txt_Obs.Enabled = TYP
    ch_Actif.Enabled = TYP
    Picture2.Enabled = TYP
    Cmd_Photo.Enabled = TYP
    CmdDelete.Enabled = TYP
    CmdSave.Enabled = TYP
    cda_DebuAssur.Enabled = TYP
    cda_FinAssur.Enabled = TYP
    cda_DebutVesite.Enabled = TYP
    cda_FinVisite.Enabled = TYP
    cda_debuttax.Enabled = TYP
    cda_fintax.Enabled = TYP
    txt_NumAssur.Enabled = TYP
    txt_FourAssur.Enabled = TYP
    txt_AgenceAssur.Enabled = TYP
    Img_Vehicule.Enabled = TYP
End Sub
'# initialise DateBox***
Private Sub InitialiseDate()
    cda_DebuAssur.Value = Date
    cda_FinAssur.Value = Date
    cda_DebutVesite.Value = Date
    cda_FinVisite.Value = Date
    cda_debuttax.Value = Date
    cda_fintax.Value = Date
    cda_DateCircul.Value = Date
End Sub
'# Afficher Energie***
Private Sub Cbo_Energie_LostFocus()
    Dim LObj_Find   As New Energie
    Dim Lrs_Find    As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.Get_EnergByLiborCod(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Cbo_Energie.Text)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Lrs_Find.EOF Then
        MsgBox "Energie inexistante, vérifier votre saisie. ", vbInformation, App.ProductName
        Cbo_Energie.SetFocus
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Private Sub Affiche_Energie()
    Dim LObj_Find    As New Energie
    Dim Lrs_Find        As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.Get_Energ(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
            With Cbo_Energie
                .AddItem Lrs_Find("libelle")
            End With
            Lrs_Find.MoveNext
        Wend
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'# Selectionner Photo***
Private Sub Img_Vehicule_Click()
    SelectPicture
End Sub
Private Sub Cmd_Photo_Click()
    SelectPicture
End Sub
Private Sub SelectPicture()

    With CDlg
   
        .DialogTitle = "Séléctionner Photo.."
        .FileName = ""
        .Filter = "Image (*.jpg; *.bmp)|*.jpg; *.bmp"
        .ShowOpen
        If Len(Trim(.FileName)) < 1 Then Exit Sub
        Txt_Photo.Text = CDlg.FileName
        Img_Vehicule.Picture = LoadPicture(.FileName)
        ChangPic = True
    End With
End Sub
'# Nouveau***
Private Sub CmdAdd_Click()
On Error GoTo Err
    If (CHECK_ACCES("Ins_Vehicule", LInt_UserId) = False) Then
        MsgBox "Insertion n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Ajouter Véhicule", vbExclamation, "Parcano..."
        Exit Sub
    End If
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
    End If
    
    EnbDisb (True)
    Call ViderZone(Frm_Vehicule)
    Img_Vehicule.Picture = LoadPicture("\\srv-files\Centrano\Image Parcano\Vehicule\car.jpg")
    Call InitialiseDate
    Cmd_Supp.Visible = False
    Lbl_Supp.Visible = False
    txt_BC.Text = 0
    Txt_NvCpt.Text = 0
    txt_DerCompteurV.Text = 0
    txt_klmVidange.Text = 0
    grid.ListItems.Clear
    txt_Matricule.Text = "Auto"
    ch_Actif.Value = 1
    txt_libelle.SetFocus
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'# Suppression***
Private Sub CmdDelete_Click()
    Dim LObj_Find      As New VEHICULE
    Dim VCode           As String
On Error GoTo Err
    If txt_Matricule.Text <> "Auto" Then
        If (CHECK_ACCES("Supp_vehicule", LInt_UserId) = False) Then
            MsgBox "Suppression n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Supprimer Véhicule", vbExclamation
            Exit Sub
        End If
    End If
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
            Exit Sub
        Else
            txt_Matricule.SetFocus
            Exit Sub
        End If
    End If
    VCode = txt_Matricule.Text
    If MsgBox("Confirmez vous la suppression de cette " & vbNewLine & "véhicule", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
        Call LObj_Find.Delete_Add_Vehicule(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "O", LInt_UserId, VCode)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
        MsgBox "Vehicule Supprimer avec succes!...", vbInformation, App.ProductName
        txt_Matricule.SetFocus
        Call InitialiseDate
    On Error Resume Next
        Img_Vehicule.Picture = LoadPicture("\\srv-files\Centrano\Image Parcano\Vehicule\car.jpg")
    On Error GoTo Err
        Call EnbDisb(True)
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'# Ré-ajouter***
Private Sub Cmd_Supp_Click()
    Dim VCode       As String
    Dim LObj_Find   As New VEHICULE
On Error GoTo Err
    If txt_Matricule.Text <> "Auto" Then
        If (CHECK_ACCES("Supp_vehicule", LInt_UserId) = False) Then
            MsgBox "Ré-ajouter n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour ré-ajouter Véhicule", vbExclamation
            Exit Sub
        End If
    End If
    VCode = txt_Matricule.Text
    If MsgBox("Confirmez vous la ré-ajouter de cette " & vbNewLine & "véhicule", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
        Call LObj_Find.Delete_Add_Vehicule(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "N", LInt_UserId, VCode)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
        MsgBox "Vehicule ré-ajouter avec succes!...", vbInformation, App.ProductName
        Call AfficheRow(VCode)
        Call InitialiseDate
    On Error Resume Next
        Img_Vehicule.Picture = LoadPicture("\\srv-files\Centrano\Image Parcano\Vehicule\car.jpg")
    On Error GoTo Err
        Call EnbDisb(True)
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'# FindView...
Private Sub CmdFind_Click()
On Error Resume Next
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "VehiculeBase"
        .Show vbModal
    End With
End Sub

Private Sub cmdFindenergie_Click()
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "Energie"
        .Show vbModal
    End With
End Sub
Private Sub cmdFindMatricule_Click()
    On Error Resume Next
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "VehiculeActif"
        .Show vbModal
    End With
End Sub
Public Sub AfficheRow(ByVal VCode As String)
    Dim LOBJ_BonVidange     As New BonVidange
    Dim Lobj_Vehicule       As New VEHICULE
    Dim rs                  As New Recordset
    Dim rs1                 As New Recordset
    Dim AA                  As Long
    Dim Pic_Vehicule        As String
On Error GoTo Err
    Call ViderZone(Frm_Vehicule)
    grid.ListItems.Clear
    If (CHECK_ACCES("Maj_vehicule", LInt_UserId) = False) Then
        MsgBox "Ré-ajouter n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour ré-ajouter Véhicule", vbExclamation
        Exit Sub
    End If
    AA = 0
    Set rs = Lobj_Vehicule.GetVehicule(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
        ErrNumber = 0
        Exit Sub
    End If
    Set Lobj_Vehicule = Nothing
    If Not rs.EOF Then
        txt_Matricule.Text = rs("Code")
        If Not IsNull(rs("Matricule")) Then
            txt_libelle.Text = rs("Matricule")
            Txt_NvCpt.Text = MaxCompteurVehicule(rs("Matricule")) 'dernier compt saisi dans la fiche trafic
        End If
        If Not IsNull(rs("Abreviation")) Then Txt_Abr.Text = rs("Abreviation")
        If Not IsNull(rs("Type")) Then txt_Type.Text = rs("TYPE")
        If Not IsNull(rs("marque")) Then txt_Marque.Text = rs("marque")
        If Not IsNull(rs("puissance")) Then txt_Puissance.Text = rs("puissance")
        If Not IsNull(rs("Energie")) Then Cbo_Energie.Text = rs("Energie")
        If Not IsNull(rs("NumSerie")) Then txt_Nserie.Text = rs("NumSerie")
        If Not IsNull(rs("DateCircul")) Then cda_DateCircul.Value = rs("DateCircul")
        If Not IsNull(rs("NumAssur")) Then txt_NumAssur.Text = rs("NumAssur")
        If Not IsNull(rs("FournisAssur")) Then txt_FourAssur.Text = rs("FournisAssur")
        If Not IsNull(rs("AgenceAssur")) Then txt_AgenceAssur.Text = rs("AgenceAssur")
        If Not IsNull(rs("DateDebAssur")) And rs("DateDebAssur") <> "01/01/1900" Then cda_DebuAssur.Value = rs("DateDebAssur")
        If Not IsNull(rs("DateFinAssur")) And rs("DateFinAssur") <> "01/01/1900" Then cda_FinAssur.Value = rs("DAteFinAssur")
        If Not IsNull(rs("CompteurCarburant")) Then txt_BC.Text = rs("CompteurCarburant")
        If Not IsNull(rs("CompteurVidange")) Then
            txt_DerCompteurV.Text = rs("CompteurVidange")
        Else
            txt_DerCompteurV.Text = 0
        End If
        If Not IsNull(rs("CompteurFT")) Then
                If Val(Txt_NvCpt.Text) < Val(rs("CompteurFT")) Then
                    Txt_NvCpt.Text = rs("CompteurFT")
                    'le compteur saisi dans la fiche traffic > à celui saissi dans véhicule (Compteur vidange)
                End If
        End If
        Set rs1 = LOBJ_BonVidange.Get_DerBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("Code"))
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
            ErrNumber = 0
            Exit Sub
        End If
        Set LOBJ_BonVidange = Nothing
        If Not rs1.EOF Then
            If Not IsNull(rs1("CompteurVidange")) Then
                txt_DerCompteurV.Text = rs1("CompteurVidange")
            Else
                 txt_DerCompteurV.Text = rs("CompteurVidange")
            End If
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
        End If
        Set rs1 = Nothing
        If Not IsNull(rs("genre")) Then txt_Genre.Text = rs("genre")
        If Not IsNull(rs("Carrosserie")) Then txt_Carrosserie.Text = rs("Carrosserie")
        If Not IsNull(rs("PlaceAssis")) Then txt_PlasseAssise.Text = rs("PlaceAssis")
        If Not IsNull(rs("PlaceDebout")) Then txt_Plassedebout.Text = rs("PlaceDebout")
        If Not IsNull(rs("Cylindre")) Then txt_Cylindre.Text = rs("Cylindre")
        If Not IsNull(rs("NbrEssieux")) Then txt_NbrEssieux.Text = rs("NbrEssieux")
        If Not IsNull(rs("PTAC")) Then txt_PTAC.Text = rs("PTAC")
        If Not IsNull(rs("PTRA")) Then txt_PTRA.Text = rs("PTRA")
        If Not IsNull(rs("Charge")) Then txt_Charge.Text = rs("Charge")
        If Not IsNull(rs("NbKlmVidange")) Then txt_klmVidange.Text = rs("NbKlmVidange")
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
        If Not IsNull(rs("DateDebVisite")) And rs("DateDebVisite") <> "01/01/1900" Then cda_DebutVesite.Value = rs("DateDebVisite")
        If Not IsNull(rs("DAteFinVisite")) And rs("DAteFinVisite") <> "01/01/1900" Then cda_FinVisite.Value = rs("DAteFinVisite")
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
        If Not IsNull(rs("DateDebTax")) And rs("DateDebTax") <> "01/01/1900" Then cda_debuttax.Value = rs("DateDebTax")
        If Not IsNull(rs("DateFinTax")) And rs("DateFinTax") <> "01/01/1900" Then cda_fintax.Value = rs("DateFinTax")
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
        If Not IsNull(rs("DateSortie")) Then
            cda_DateSortie.Text = rs("DateSortie")
        Else
            cda_DateSortie.Text = ""
        End If
        If Not IsNull(rs("Obs")) Then txt_Obs.Text = rs("Obs")
        If Not IsNull(rs("PicBox")) Then
        On Error Resume Next
            Img_Vehicule = LoadPicture("\\srv-files\Centrano\Image Parcano\Vehicule\" & rs("PicBox"))
            On Error GoTo Err
            Txt_Photo.Text = rs("PicBox")
        Else
        On Error Resume Next
            Img_Vehicule = LoadPicture("\\srv-files\Centrano\Image Parcano\Vehicule\car.jpg")
        On Error GoTo Err
            Txt_Photo.Text = "Vehicule"
        End If
        If rs("Supp") = "O" Then
            EnbDisb (False)
            Cmd_Supp.Visible = True
            Lbl_Supp.Visible = True
        Else
            EnbDisb (True)
            Cmd_Supp.Visible = False
            Lbl_Supp.Visible = False
        End If
    Call AfficheRow_Lubr(VCode)
    Else
        MsgBox "Code introuvable", vbInformation, App.ProductName
        txt_Matricule.SetFocus
    End If
    Set rs = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Public Sub AfficheRow_Lubr(ByVal VCode As String)
    Dim LOBJ_Lub        As New Produit_Lubrifiant
    Dim Lobj_Vehicule   As New VEHICULE
    Dim rs              As New Recordset
    Dim rs1             As New Recordset
    Dim itmX
On Error Resume Next

    Set rs = Lobj_Vehicule.Get_VehVdg(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set Lobj_Vehicule = Nothing
    If Not rs.EOF Then
        While Not rs.EOF
            Set LOBJ_Lub = New Produit_Lubrifiant
            Set rs1 = LOBJ_Lub.Get_ProdLubBycode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("Lubrifiant"))
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
            Set LOBJ_Lub = Nothing
            If Not rs1.EOF Then
                Set itmX = grid.ListItems.Add(, , CStr(rs("Lubrifiant")))
                itmX.SubItems(1) = CStr(rs1("Libelle"))
                itmX.SubItems(2) = CStr(Format(rs1("prixht"), "#,##0.000"))
                itmX.SubItems(3) = CStr(Format(rs1("tva"), "#,##0.000"))
            End If
            Set rs1 = Nothing
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Private Sub Cmd_Annul_Click()
    Dim i       As Integer
On Error GoTo Err
'Si pas de veh sélectionné ou pas en cours de saisie
    If Len(Trim(txt_Matricule.Text)) = 0 Then  'Trim : Renvoie une copie d'une chaîne sans espaces à gauche ni à droite
        MsgBox "Véhicule obligatoire      ", vbInformation, App.ProductName
        txt_Matricule.SetFocus
        Exit Sub
    End If
'Liste de details du bon est vide
    If grid.ListItems.Count <= 0 Then Exit Sub
    If MsgBox("Confirmez vous la suppression de la ligne en cours.?", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
        i = grid.SelectedItem.Index  ' indice de la ligne de detail sélectionné
        grid.ListItems.Remove i     'Supprimer la ligne de la liste
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Private Sub Delete_VehVdg(ByVal Veh As String)
    Dim LOBJ_Veh    As New VEHICULE
On Error GoTo Err
    Call LOBJ_Veh.Delete_VehVdg(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Veh)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'# Enregistre
Private Sub CmdSave_Click()
    Dim Lrs_Veh             As New Recordset
    Dim LOBJ_Veh            As New VEHICULE
    Dim LObj_Find           As New Energie
    Dim rs                  As New Recordset
    Dim LInt_NumCompteur    As Long
On Error GoTo Err
    If Len(Txt_Abr.Text) < 2 Or Trim(Txt_Abr.Text) = "" Then
        MsgBox "Abréviation Obligatoir!...           ", vbExclamation + vbOKOnly + vbDefaultButton2, App.ProductName
        Txt_Abr.SetFocus
        Exit Sub
    End If
    If Trim(txt_Nserie.Text) = "" Then
        MsgBox "N° serie du type : Obligatoire!...           ", vbExclamation + vbOKOnly + vbDefaultButton2, App.ProductName
        txt_Nserie.SetFocus
        Exit Sub
    End If
    
    If Trim(Cbo_Energie.Text) = "" Then
        MsgBox "Energie Obligatoire!...           ", vbExclamation + vbOKOnly + vbDefaultButton2, App.ProductName
        Cbo_Energie.SetFocus
        Exit Sub
    End If
    
    If Trim(txt_NumAssur.Text) = "" Then
        MsgBox "N° Contrat Obligatoire!...           ", vbExclamation + vbOKOnly + vbDefaultButton2, App.ProductName
        SSTab1.Tab = 1
        txt_NumAssur.SetFocus
        Exit Sub
    End If
    
    If Trim(txt_FourAssur.Text) = "" Then
        MsgBox "Assureur Obligatoire!...           ", vbExclamation + vbOKOnly + vbDefaultButton2, App.ProductName
        SSTab1.Tab = 1
        txt_FourAssur.SetFocus
        Exit Sub
    End If
    
    If Trim(txt_AgenceAssur.Text) = "" Then
        MsgBox "Agence assurance Obligatoire!...           ", vbExclamation + vbOKOnly + vbDefaultButton2, App.ProductName
        SSTab1.Tab = 1
        txt_AgenceAssur.SetFocus
        Exit Sub
    End If
    
    If (cda_FinAssur.Value <= cda_DebuAssur.Value) Or (cda_FinAssur.Value <= Date) Then
        MsgBox "Date fin invalide", vbInformation
        SSTab1.Tab = 1
        cda_FinAssur.SetFocus
        Exit Sub
    End If
    If (cda_FinVisite.Value <= cda_DebuAssur.Value) Or (cda_FinVisite.Value <= Date) Then
        MsgBox "Date fin invalide", vbInformation
        SSTab1.Tab = 1
        cda_FinVisite.SetFocus
        Exit Sub
    End If
    If (cda_fintax.Value <= cda_debuttax.Value) Or (cda_fintax.Value <= Date) Then
        MsgBox "Date fin invalide", vbInformation
        SSTab1.Tab = 1
        cda_fintax.SetFocus
        Exit Sub
    End If
    
    '# Vérifier saisie de l'énergie
    Set rs = LObj_Find.Get_EnergByLiborCod(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Cbo_Energie.Text)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    
    If rs.EOF Then
        MsgBox "Energie inexistante, vérifier votre saisie. ", vbInformation, App.ProductName
        rs.Close
        Cbo_Energie.SetFocus
        Exit Sub
    End If
    '# Confirmer Abreviation existe ou non.
    Set Lrs_Veh = LOBJ_Veh.GetAll_Abreviation(ErrNumber, ErrDescription, ErrSourceDetail, Trim(Txt_Abr.Text), Trim(txt_Matricule.Text), CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LOBJ_Veh = Nothing
    If Not Lrs_Veh.EOF Then
        MsgBox "Abreviation existe déja!... Choisir un autre...", vbExclamation, App.ProductName
        Exit Sub
    End If
    Set Lrs_Veh = Nothing
    If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
        If (CHECK_ACCES("Maj_vehicule", LInt_UserId) = False) Then
            MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier Véhicule", vbExclamation
            Exit Sub
        End If
        If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
        Call Modif_Veh
        'Supression de tout les vidanges associé à ce véhicule
        Call Delete_VehVdg(txt_Matricule.Text)
        'Réajout des vidanges
        If grid.ListItems.Count <> 0 Then Call Ajout_VehVdg(txt_Matricule.Text)
    ElseIf txt_Matricule.Text = "Auto" Then
        If (CHECK_ACCES("Maj_vehicule", LInt_UserId) = False) Then
            MsgBox "Insertion n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Ajouter Véhicule", vbExclamation
            Exit Sub
        End If
        If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
        LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_Vehicule")
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        'Insertion enregistrement assiette
        txt_Matricule.Text = Format(LInt_NumCompteur, "00000")
        Call Ajout_Vehicule
        If grid.ListItems.Count <> 0 Then Call Ajout_VehVdg(txt_Matricule.Text)
    End If
    
    Call ViderZone(Frm_Vehicule)
    grid.ListItems.Clear
    Img_Vehicule.Picture = LoadPicture("\\srv-files\Centrano\Image Parcano\Vehicule\car.jpg")
    txt_Matricule.SetFocus
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Private Sub Ajout_Vehicule()
    Dim Lobj_Vehicule   As New VEHICULE
    Dim LRs_NewRecord   As New Recordset
On Error GoTo Err
    StrPicture = txt_libelle.Text & ".Bmp"
    Set LRs_NewRecord = CreateEmptyRS_Vehicule
    With LRs_NewRecord
        .AddNew
        .Fields("Code") = txt_Matricule.Text
        If Trim(txt_libelle.Text) <> "" Then .Fields("Matricule") = txt_libelle.Text
        If Trim(Txt_Abr.Text) <> "" Then .Fields("ABr") = UCase(Txt_Abr.Text)
        If Trim(txt_Type.Text) <> "" Then .Fields("TYPE") = txt_Type.Text
        If Trim(txt_Marque.Text) <> "" Then .Fields("Marque") = txt_Marque.Text
        If Trim(txt_Puissance.Text) <> "" Then .Fields("Puissance") = txt_Puissance.Text
        If Trim(Cbo_Energie.Text) <> "" Then .Fields("Energie") = Cbo_Energie.Text
        If Trim(txt_Genre.Text) <> "" Then .Fields("genre") = txt_Genre.Text
        If Trim(txt_Carrosserie.Text) <> "" Then .Fields("Carrosserie") = txt_Carrosserie.Text
        If Trim(txt_PlasseAssise.Text) <> "" Then .Fields("PlaceAssis") = txt_PlasseAssise.Text
        If Trim(txt_Plassedebout.Text) <> "" Then .Fields("PlaceDebout") = txt_Plassedebout.Text
        If Trim(txt_Cylindre.Text) <> "" Then .Fields("Cylindre") = txt_Cylindre.Text
        If Trim(txt_NbrEssieux.Text) <> "" Then .Fields("NbrEssieux") = txt_NbrEssieux.Text
        If Trim(txt_PTAC.Text) <> "" Then .Fields("PTAC") = txt_PTAC.Text
        If Trim(txt_PTRA.Text) <> "" Then .Fields("PTRA") = txt_PTRA.Text
        If Trim(txt_Charge.Text) <> "" Then .Fields("Charge") = txt_Charge.Text
        If Trim(txt_klmVidange.Text) <> "" Then .Fields("NbKlmVidange") = txt_klmVidange.Text
        If Trim(txt_PoidVide.Text) <> "" Then .Fields("PoidsVide") = txt_PoidVide.Text
        If Trim(txt_TYPCOM.Text) <> "" Then .Fields("TypeComm") = txt_TYPCOM.Text
        If Trim(txt_Nserie.Text) <> "" Then .Fields("NumSerie") = txt_Nserie.Text
        If cda_DateCircul.Value <> "" And cda_DateCircul.Value <> "__/__/____" Then .Fields("DateCircul") = CDate(cda_DateCircul.Value)
        If Trim(txt_NumAssur.Text) <> "" Then .Fields("NumAssur") = txt_NumAssur.Text
        If Trim(txt_FourAssur.Text) <> "" Then .Fields("FournisAssur") = txt_FourAssur.Text
        If Trim(txt_AgenceAssur.Text) <> "" Then .Fields("AgenceAssur") = txt_AgenceAssur.Text
        
        If Trim(cda_DebuAssur.Value) <> "" And cda_DebuAssur.Value <> "__/__/____" Then .Fields("DateDebAssur") = CDate(cda_DebuAssur.Value)
        If Trim(cda_FinAssur.Value) <> "" And cda_FinAssur.Value <> "__/__/____" Then .Fields("DAteFinAssur") = CDate(cda_FinAssur.Value)
        If Trim(cda_DebutVesite.Value) <> "" And cda_DebutVesite.Value <> "__/__/____" Then .Fields("DateDebVisite") = CDate(cda_DebutVesite.Value)
        If Trim(cda_FinVisite.Value) <> "" And cda_FinVisite.Value <> "__/__/____" Then .Fields("DAteFinVisite") = CDate(cda_FinVisite.Value)
        If Trim(cda_debuttax.Value) <> "" And cda_debuttax.Value <> "__/__/____" Then .Fields("DateDebTax") = CDate(cda_debuttax.Value)
        If Trim(cda_fintax.Value) <> "" And cda_fintax.Value <> "__/__/____" Then .Fields("DateFinTax") = CDate(cda_fintax.Value)
        If cda_DateSortie.Text <> "" And cda_DateSortie.Text <> "__/__/____" Then .Fields("DateSortie") = CDate(cda_DateSortie.Text)
        If txt_Obs.Text <> "" Then .Fields("Obs") = txt_Obs.Text
        If txt_BC.Text <> "" Then .Fields("CompteurCarburant") = txt_BC.Text
        .Fields("actif") = 1
        .Fields("Disponible") = "O"
        .Fields("UserInsert") = LInt_UserId
        If ChangPic = True Then
            .Fields("PicBox") = StrPicture
        Else
            .Fields("PicBox") = "Car.Jpg"
        End If
    End With
    Call Lobj_Vehicule.Insert_Veh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion
    Set Lobj_Vehicule = Nothing
    Set LRs_NewRecord = Nothing
On Error Resume Next
    If ChangPic = True Then SavePicture Img_Vehicule.Picture, "\\srv-files\Centrano\Image Parcano\Vehicule\" & StrPicture
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Private Sub Modif_Veh()
    Dim Lobj_Vehicule   As New VEHICULE
    Dim LRs_NewRecord   As New Recordset
On Error GoTo Err
    StrPicture = txt_libelle.Text & ".Bmp"
    Set LRs_NewRecord = CreateEmptyRS_Vehicule
    With LRs_NewRecord
        .AddNew
        .Fields("Code") = txt_Matricule.Text
        If Trim(txt_libelle.Text) <> "" Then .Fields("Matricule") = txt_libelle.Text
        If Trim(Txt_Abr.Text) <> "" Then .Fields("ABr") = UCase(Txt_Abr.Text)
        If Trim(txt_Type.Text) <> "" Then .Fields("TYPE") = txt_Type.Text
        If Trim(txt_Marque.Text) <> "" Then .Fields("Marque") = txt_Marque.Text
        If Trim(txt_Puissance.Text) <> "" Then .Fields("Puissance") = txt_Puissance.Text
        If Trim(Cbo_Energie.Text) <> "" Then .Fields("Energie") = Cbo_Energie.Text
        If Trim(txt_Genre.Text) <> "" Then .Fields("genre") = txt_Genre.Text
        If Trim(txt_Carrosserie.Text) <> "" Then .Fields("Carrosserie") = txt_Carrosserie.Text
        If Trim(txt_PlasseAssise.Text) <> "" Then .Fields("PlaceAssis") = txt_PlasseAssise.Text
        If Trim(txt_Plassedebout.Text) <> "" Then .Fields("PlaceDebout") = txt_Plassedebout.Text
        If Trim(txt_Cylindre.Text) <> "" Then .Fields("Cylindre") = txt_Cylindre.Text
        If Trim(txt_NbrEssieux.Text) <> "" Then .Fields("NbrEssieux") = txt_NbrEssieux.Text
        If Trim(txt_PTAC.Text) <> "" Then .Fields("PTAC") = txt_PTAC.Text
        If Trim(txt_PTRA.Text) <> "" Then .Fields("PTRA") = txt_PTRA.Text
        If Trim(txt_Charge.Text) <> "" Then .Fields("Charge") = txt_Charge.Text
        If Trim(txt_klmVidange.Text) <> "" Then .Fields("NbKlmVidange") = txt_klmVidange.Text
        If Trim(txt_PoidVide.Text) <> "" Then .Fields("PoidsVide") = txt_PoidVide.Text
        If Trim(txt_TYPCOM.Text) <> "" Then .Fields("TypeComm") = txt_TYPCOM.Text
        If Trim(txt_Nserie.Text) <> "" Then .Fields("NumSerie") = txt_Nserie.Text
        If cda_DateCircul.Value <> "" And cda_DateCircul.Value <> "__/__/____" Then .Fields("DateCircul") = CDate(cda_DateCircul.Value)
        If Trim(txt_NumAssur.Text) <> "" Then .Fields("NumAssur") = txt_NumAssur.Text
        If Trim(txt_FourAssur.Text) <> "" Then .Fields("FournisAssur") = txt_FourAssur.Text
        If Trim(txt_AgenceAssur.Text) <> "" Then .Fields("AgenceAssur") = txt_AgenceAssur.Text
        
        If Trim(cda_DebuAssur.Value) <> "" And cda_DebuAssur.Value <> "__/__/____" Then .Fields("DateDebAssur") = CDate(cda_DebuAssur.Value)
        If Trim(cda_FinAssur.Value) <> "" And cda_FinAssur.Value <> "__/__/____" Then .Fields("DAteFinAssur") = CDate(cda_FinAssur.Value)
        If Trim(cda_DebutVesite.Value) <> "" And cda_DebutVesite.Value <> "__/__/____" Then .Fields("DateDebVisite") = CDate(cda_DebutVesite.Value)
        If Trim(cda_FinVisite.Value) <> "" And cda_FinVisite.Value <> "__/__/____" Then .Fields("DAteFinVisite") = CDate(cda_FinVisite.Value)
        If Trim(cda_debuttax.Value) <> "" And cda_debuttax.Value <> "__/__/____" Then .Fields("DateDebTax") = CDate(cda_debuttax.Value)
        If Trim(cda_fintax.Value) <> "" And cda_fintax.Value <> "__/__/____" Then .Fields("DateFinTax") = CDate(cda_fintax.Value)
        If cda_DateSortie.Text <> "" And cda_DateSortie.Text <> "__/__/____" Then .Fields("DateSortie") = CDate(cda_DateSortie.Text)
        If txt_Obs.Text <> "" Then .Fields("Obs") = txt_Obs.Text
        If txt_BC.Text <> "" Then .Fields("CompteurCarburant") = Val(txt_BC.Text)

        .Fields("CompteurVidange") = Val(txt_DerCompteurV.Text)
        .Fields("CompteurFT") = Val(Txt_NvCpt.Text)
        .Fields("actif") = 1
        .Fields("Disponible") = "O"
        .Fields("UserUpdate") = LInt_UserId
        If ChangPic = True Then
            .Fields("PicBox") = StrPicture
            Txt_Photo = StrPicture
        End If
    End With
    Call Lobj_Vehicule.Update_Veh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion
    Set Lobj_Vehicule = Nothing
    Set LRs_NewRecord = Nothing
On Error Resume Next
    If ChangPic = True Then SavePicture Img_Vehicule.Picture, "\\srv-files\Centrano\Image Parcano\Vehicule\" & StrPicture
On Error GoTo Err
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Private Sub Ajout_VehVdg(ByVal Veh As String)
    Dim LOBJ_Veh        As New VEHICULE
    Dim LRs_NewRecord   As New Recordset
    Dim i               As Integer
On Error GoTo Err
    Set LRs_NewRecord = CreateEmptyRS_VehVdg
    For i = 1 To grid.ListItems.Count
        With LRs_NewRecord
            .AddNew
            .Fields("Vehicule") = txt_Matricule.Text
            .Fields("Lubrifiant") = GetCode_Lubrif(grid.ListItems(i).SubItems(1))
            .Fields("UserInsert") = LInt_UserId
        End With
    Next
    Call LOBJ_Veh.Insert_VehVdg(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LOBJ_Veh = Nothing
    Set LRs_NewRecord = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'# Afficher liste des Lubrifiants
Public Sub Affiche_Lubrif_Combo(cbo As ComboBox)
    Dim LOBJ_Lubrifiant As New Produit_Lubrifiant
    Dim rs              As New Recordset
On Error GoTo Err
    Cbo_Lub.Clear
    Set rs = LOBJ_Lubrifiant.Get_LibLubActif(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LOBJ_Lubrifiant = Nothing
    If Not rs.EOF Then
        While Not rs.EOF
            With cbo
                .AddItem rs("Libelle")
            End With
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Public Sub AfficheRow_Lubrif(ByVal VCode As String)
    Dim LOBJ_Lub        As New Produit_Lubrifiant
    Dim rs              As New Recordset
    Dim itmX
On Error GoTo Err
    Set rs = LOBJ_Lub.Get_ProdLubByLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LOBJ_Lub = Nothing
    If Not rs.EOF Then
        While Not rs.EOF
                Set itmX = grid.ListItems.Add(, , CStr(rs("Numero")))
                itmX.SubItems(1) = CStr(rs("Libelle"))
                itmX.SubItems(2) = CStr(Format(rs("prixht"), "#,##0.000"))
                itmX.SubItems(3) = CStr(Format(rs("TVA"), "#,##0.000"))
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Private Sub Cmd_ok_Click()
    Dim Hiem    As Boolean
    Dim itmX    As ListItem
    Dim i       As Integer
    Hiem = False
    For i = 1 To grid.ListItems.Count
        If grid.ListItems(i).SubItems(1) = Cbo_Lub.Text Then
           Hiem = True
           Exit For
        End If
    Next
    If Hiem = True Then
        MsgBox "Lubrifiant existe dans la liste ", vbInformation
        Exit Sub
    Else
        Call AfficheRow_Lubrif(Cbo_Lub.Text)
    End If
End Sub
