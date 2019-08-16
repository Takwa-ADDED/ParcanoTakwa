VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Supervision 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gestion de Traffic Auto"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
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
   Icon            =   "Frm_Supervision.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView Lsv_Depot 
      Height          =   3855
      Left            =   13440
      TabIndex        =   37
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   6800
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vehicule a l'interieure"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   240
   End
   Begin TabDlg.SSTab Tab_Spervision 
      Height          =   9135
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   16113
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Supervision"
      TabPicture(0)   =   "Frm_Supervision.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Anomalie"
      TabPicture(1)   =   "Frm_Supervision.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ComBox_Destination"
      Tab(1).Control(1)=   "ComBox_Vehicule"
      Tab(1).Control(2)=   "ComBox_Conducteur"
      Tab(1).Control(3)=   "Grid_An"
      Tab(1).Control(4)=   "Cda_DateFin"
      Tab(1).Control(5)=   "Cda_DateDebut"
      Tab(1).Control(6)=   "Cmd_LisConducteur"
      Tab(1).Control(7)=   "Cmd_LisVehicule"
      Tab(1).Control(8)=   "Cmd_ListDestination"
      Tab(1).Control(9)=   "Cmd_Vider"
      Tab(1).Control(10)=   "Cmd_Recherche"
      Tab(1).Control(11)=   "Cmd_Print"
      Tab(1).Control(12)=   "Lbl_AnomalieTotal"
      Tab(1).Control(13)=   "Lbl_AnomalieDuree"
      Tab(1).Control(14)=   "Lbl_AnomalieKm"
      Tab(1).Control(15)=   "Image4"
      Tab(1).Control(16)=   "Label3"
      Tab(1).Control(17)=   "Label4"
      Tab(1).Control(18)=   "Label5"
      Tab(1).Control(19)=   "Label6"
      Tab(1).Control(20)=   "Label7"
      Tab(1).Control(21)=   "Image3"
      Tab(1).ControlCount=   22
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8655
         Left            =   120
         ScaleHeight     =   8655
         ScaleWidth      =   13215
         TabIndex        =   11
         Top             =   360
         Width           =   13215
         Begin MSComctlLib.ListView LSV_Exterieur 
            Height          =   6135
            Left            =   120
            TabIndex        =   38
            Top             =   480
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   10821
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
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "N"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Numero"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Conducteur"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Matricule"
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
               Text            =   "KM"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Durée"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Op. Sortie"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Op. Entre"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   1215
            Left            =   9600
            ScaleHeight     =   1215
            ScaleWidth      =   2175
            TabIndex        =   32
            Top             =   6720
            Width           =   2175
            Begin VB.Shape Shape1 
               Height          =   1215
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label Lbl_HorsService 
               BackStyle       =   0  'Transparent
               Caption         =   "Hors-Servicé"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   840
               TabIndex        =   8
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Lbl_EnService 
               BackStyle       =   0  'Transparent
               Caption         =   "En-Servicé"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   840
               TabIndex        =   12
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Lbl_Occupe 
               BackStyle       =   0  'Transparent
               Caption         =   "Occupé"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   840
               TabIndex        =   36
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Lbl_ColHorsService 
               BackColor       =   &H0000FFFF&
               Height          =   255
               Left            =   240
               TabIndex        =   35
               Top             =   840
               Width           =   495
            End
            Begin VB.Label Lbl_ColEnServise 
               BackColor       =   &H0000FF00&
               Height          =   255
               Left            =   240
               TabIndex        =   34
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Lbl_ColOccupe 
               BackColor       =   &H000000FF&
               Height          =   255
               Left            =   240
               TabIndex        =   33
               Top             =   480
               Width           =   495
            End
         End
         Begin SToolBox.SGrid grid_vehicule 
            Height          =   6135
            Left            =   10200
            TabIndex        =   13
            Top             =   480
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   10821
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
            Header          =   0   'False
            HeaderButtons   =   0   'False
            DisableIcons    =   -1  'True
            MaxVisibleRows  =   0
         End
         Begin SToolBox.SGrid grid_Conducteur 
            Height          =   6135
            Left            =   7320
            TabIndex        =   14
            Top             =   480
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   10821
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
            Header          =   0   'False
            HeaderButtons   =   0   'False
            DisableIcons    =   -1  'True
            MaxVisibleRows  =   0
         End
         Begin SToolBox.SCommand Cmd_Vehicule 
            Height          =   375
            Left            =   10320
            TabIndex        =   15
            Top             =   120
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
            Picture         =   "Frm_Supervision.frx":0044
            BackColor       =   16777215
         End
         Begin SToolBox.SCommand Cmd_Conducteur 
            Height          =   375
            Left            =   7440
            TabIndex        =   16
            Top             =   120
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
            Picture         =   "Frm_Supervision.frx":0776
            BackColor       =   16777215
         End
         Begin VB.Image CmdPrint 
            Height          =   505
            Left            =   5280
            Picture         =   "Frm_Supervision.frx":0AF0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1815
         End
         Begin VB.Image cmd_r 
            Height          =   480
            Left            =   3240
            Picture         =   "Frm_Supervision.frx":1215A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1935
         End
         Begin VB.Line Line1 
            X1              =   7200
            X2              =   7200
            Y1              =   0
            Y2              =   6600
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Liste des Missions"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   2895
         End
         Begin VB.Image Image2 
            Height          =   495
            Left            =   7200
            Picture         =   "Frm_Supervision.frx":21AF0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   6015
         End
         Begin VB.Image Img_Conducteur 
            Height          =   615
            Left            =   120
            Stretch         =   -1  'True
            Top             =   7440
            Width           =   735
         End
         Begin VB.Image Img_Vehicule 
            Height          =   615
            Left            =   2760
            Stretch         =   -1  'True
            Top             =   7440
            Width           =   735
         End
         Begin VB.Label Lbl_Conducteur 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   840
            TabIndex        =   19
            Top             =   7440
            Width           =   1935
         End
         Begin VB.Label Lbl_Vehicule 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3480
            TabIndex        =   18
            Top             =   7440
            Width           =   1935
         End
         Begin VB.Label Lbl_Destination 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6120
            TabIndex        =   17
            Top             =   7440
            Width           =   3135
         End
         Begin VB.Image Img_Destination 
            Height          =   615
            Left            =   5400
            Top             =   7440
            Width           =   735
         End
      End
      Begin SToolBox.SBiCombo ComBox_Destination 
         Height          =   315
         Left            =   -73680
         TabIndex        =   2
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
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
      Begin SToolBox.SBiCombo ComBox_Vehicule 
         Height          =   315
         Left            =   -67200
         TabIndex        =   1
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
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
      Begin SToolBox.SBiCombo ComBox_Conducteur 
         Height          =   315
         Left            =   -73680
         TabIndex        =   0
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
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
      Begin SToolBox.SGrid Grid_An 
         Height          =   6495
         Left            =   -75000
         TabIndex        =   10
         Top             =   1560
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   11456
         RowMode         =   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisableIcons    =   -1  'True
         MaxVisibleRows  =   0
      End
      Begin MSComCtl2.DTPicker Cda_DateFin 
         Height          =   375
         Left            =   -67440
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12632256
         Format          =   149618689
         CurrentDate     =   42850
      End
      Begin MSComCtl2.DTPicker Cda_DateDebut 
         Height          =   375
         Left            =   -69360
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12632256
         Format          =   149618689
         CurrentDate     =   42850
      End
      Begin SToolBox.SCommand Cmd_LisConducteur 
         Height          =   375
         Left            =   -69360
         TabIndex        =   20
         Top             =   480
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
         Picture         =   "Frm_Supervision.frx":55B52
         ButtonType      =   1
      End
      Begin SToolBox.SCommand Cmd_LisVehicule 
         Height          =   375
         Left            =   -63000
         TabIndex        =   21
         Top             =   480
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
         Picture         =   "Frm_Supervision.frx":55ED4
         ButtonType      =   1
      End
      Begin SToolBox.SCommand Cmd_ListDestination 
         Height          =   495
         Left            =   -70800
         TabIndex        =   22
         Top             =   960
         Width           =   495
         _ExtentX        =   873
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
         Picture         =   "Frm_Supervision.frx":56256
         ButtonType      =   1
      End
      Begin VB.Image Cmd_Vider 
         Height          =   480
         Left            =   -66960
         Picture         =   "Frm_Supervision.frx":565D8
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Image Cmd_Recherche 
         Height          =   510
         Left            =   -65280
         Picture         =   "Frm_Supervision.frx":688F6
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Image Cmd_Print 
         Height          =   510
         Left            =   -63600
         Picture         =   "Frm_Supervision.frx":794F8
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Lbl_AnomalieTotal 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -74760
         TabIndex        =   30
         Top             =   8160
         Width           =   2295
      End
      Begin VB.Label Lbl_AnomalieDuree 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   -72600
         TabIndex        =   29
         Top             =   8160
         Width           =   2055
      End
      Begin VB.Label Lbl_AnomalieKm 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   -70560
         TabIndex        =   28
         Top             =   8160
         Width           =   2055
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   -75000
         Picture         =   "Frm_Supervision.frx":8A0FA
         Stretch         =   -1  'True
         Top             =   8160
         Width           =   12855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Au"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67800
         TabIndex        =   27
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "De"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69720
         TabIndex        =   26
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Conducteur"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Véhicule"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68400
         TabIndex        =   24
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   1215
         Left            =   -75000
         Picture         =   "Frm_Supervision.frx":BE15C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   12975
      End
   End
   Begin VB.Label Lbl_heure 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9960
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Lbl_date 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contrôle Vehicule / Conducteur"
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
      TabIndex        =   5
      Top             =   360
      Width           =   5295
   End
   Begin VB.Image PicBox_Header 
      Height          =   1000
      Left            =   0
      Picture         =   "Frm_Supervision.frx":F14F6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "Frm_Supervision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim VCodeVehicle As String, VCodeDrive  As String, VCodeDestination  As String
    Dim NAnomalieTotal As Integer, NAnomalieKm As Integer, NAnomalieDuree As Integer
    
    
    Dim Operation() As String
    Dim itmX As ListItem
    
    


'~~~~~~~~~~~~~~~~~~~~
    'Mise en Forme~~~
'~~~~~~~~~~~~~~~~~~~~
Private Sub Form_Load()
    Tab_Spervision.Tab = 0
    Call AfficheExterieur
    Call AfficheDepot
    Call Initgrid_Conducteur
    Call Initgrid_Vehicule
    Call Initgrid_Anomali
    Call Affiche_Vehicule
    Call Affiche_Conducteur
    Call Initialise_SBICombo_Cond(ComBox_Conducteur)
    Call Initialise_SBICombo_AllDest(ComBox_Destination)
    Call Initialise_SBICombo_Vehic(ComBox_Vehicule)
    Call grid_vehicule_ColumnClick(1)
    Call grid_Conducteur_ColumnClick(1)
End Sub
Private Sub Form_Resize()
    Dim WidthForm As Integer, HeightForm As Integer
    WidthForm = Me.Width
    HeightForm = Me.Height
        Lbl_heure.Left = WidthForm - 4500
        Lbl_date.Left = WidthForm - 7300
        PicBox_Header.Width = WidthForm
        Tab_Spervision.Width = WidthForm - 400
        Picture1.Width = Tab_Spervision.Width - 280
        grid_vehicule.Left = WidthForm - 3500
        Cmd_Vehicule.Left = WidthForm - 3500
        grid_Conducteur.Left = WidthForm - 6300
        Cmd_Conducteur.Left = WidthForm - 6300
        Picture2.Left = WidthForm - 3500
        Image2.Left = WidthForm - 6400
        Line1.X1 = WidthForm - 6400
        Line1.X2 = WidthForm - 6400
        LSV_Exterieur.Width = Tab_Spervision.Width - 6200
        CmdPrint.Left = WidthForm - 8500
        cmd_r.Left = WidthForm - 10500
        Grid_An.Width = Tab_Spervision.Width - 50
        Image4.Width = Tab_Spervision.Width - 50
        Image3.Width = Tab_Spervision.Width - 50
        Cmd_Print.Left = WidthForm - 2100
        Cmd_Recherche.Left = WidthForm - 3700
        Cmd_Vider.Left = WidthForm - 5330
        Image4.Top = Tab_Spervision.Height - 350
        Grid_An.Height = Tab_Spervision.Height - 2000
End Sub
Private Sub Form_Initialize()
    Dim dat As Date
    Lbl_heure.Caption = Format(Time, "hh:mm:ss")
    Timer1.Enabled = True
    Timer1.Interval = 1000
    dat = Date
    Lbl_date.Caption = UCase(Format(Now, "dddd-dd-mm-yyyy"))
    dat = Time
    Lbl_heure.Caption = dat
    Cda_DateDebut.Value = "01/" & Month(Date) & "/" & Year(Date)
    Cda_DateFin.Value = Date
    Img_Conducteur.Visible = False
    Lbl_Conducteur.Visible = False
    Img_Vehicule.Visible = False
    Lbl_Vehicule.Visible = False
    Lbl_Destination.Visible = False
    Img_Destination.Visible = False
    VCodeVehicle = "0000"
    VCodeDrive = "0000"
    VCodeDestination = "0000"
End Sub
'~~~~~~~~~~~~~~~~~~~~~~
    'Initialise Grid~~~
'~~~~~~~~~~~~~~~~~~~~~~
Public Sub Initgrid_Anomali()
    With Grid_An
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
        .AddColumn "TypeAn", "Type An", , , 80
        .AddColumn "Conducteur", "Conducteur", , , 100
        .AddColumn "Vehicule", "Vehicule", , , 100
        .AddColumn "Destination", "Destination", , , 100
        .AddColumn "DS", "D.Sortie", , , 80
        .AddColumn "HS", "H.Sortie", , , 80
        .AddColumn "HE", "H.Eentre", , , 80
        .AddColumn "Duree", "Durée", , , 80
        .AddColumn "MaxD", "Max-Durée", , , 80
        .AddColumn "DifD", "Dif-Durée", , , 80
        .AddColumn "CPTS", "CPT.S.", , , 60
        .AddColumn "CPTE", "CPT.E.", , , 60
        .AddColumn "Km", "Km", , , 50
        .AddColumn "MaxKm", "Max-Km", , , 80
        .AddColumn "DifKm", "Dif-Km", , , 80
        .AddColumn "OS", "Op.S.", , , 80
        .AddColumn "OE", "Op.E.", , , 80
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
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
        .AddColumn "Matricule", "Vehicule", , , 140
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
        .AddColumn "Libelle", "Conducteur", , , 140
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Afficher Liste (FindView)~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Cmd_LisConducteur_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "ConducteurSuperv"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub Cmd_LisVehicule_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "VehiculeSuperv"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub Cmd_ListDestination_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "DestinationSuperv"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
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
'~~~~~~~~~~~~~~~~~~~~
    'Afficher Tous~~~
'~~~~~~~~~~~~~~~~~~~~
Public Sub Affiche_Vehicule()
    Dim LObj_Find As New VEHICULE, Lrs_Find As New Recordset, Couleur As String
On Error GoTo Err
    grid_vehicule.ClearRows
    Set Lrs_Find = LObj_Find.GetAllActifVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
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
                    .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Find("Matricule"), , , &H8080FF
                Else
                If (Lrs_Find("disponible") = "O") Then
                    .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Find("Matricule"), , , &HC0FFC0
                    Else
                     .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Find("Matricule"), , , &H80FFFF
                     End If
                End If
            End With
                Lrs_Find.MoveNext
        Wend
        grid_vehicule.Redraw = True
     End If
    grid_vehicule.SelectedRow = 1
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Public Sub Affiche_Conducteur()
    Dim LObj_Find As New Conducteur, Lrs_Find As New Recordset, i As Integer, Couleur As String
On Error GoTo Err
    grid_Conducteur.ClearRows
    Set Lrs_Find = LObj_Find.GetAll_ConducteursActif(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        grid_Conducteur.Redraw = False
        While Not Lrs_Find.EOF
            'Définir couleur***
            Couleur = "vbGreen"
            For i = 1 To LSV_Exterieur.ListItems.Count
                If (Lrs_Find("Libelle") = LSV_Exterieur.ListItems(i).SubItems(2) And LSV_Exterieur.ListItems(i).SubItems(6) = "" And LSV_Exterieur.ListItems(i).SubItems(4) <> "REPARATION") Then Couleur = "vbRed"
            Next i
            
            With grid_Conducteur
                .AddRow
                If Couleur = "vbRed" Then
                    .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find("Libelle"), , , &H8080FF
                Else
                    If (Lrs_Find("disponible") = "O") Then
                    .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find("Libelle"), , , &HC0FFC0
                    Else
                     .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find("Libelle"), , , &H80FFFF
                     End If
                End If
            End With
            Lrs_Find.MoveNext
        Wend
        grid_Conducteur.Redraw = True
     End If
    grid_Conducteur.SelectedRow = 1
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Public Sub AfficheDepot()
    Dim LObj_Find As New VEHICULE, Lrs_Find As New Recordset
    Dim DateSys As Date, i As Integer, J As Integer
On Error GoTo Err
    DateSys = Date
    Lsv_Depot.ListItems.Clear
    Set Lrs_Find = LObj_Find.GetMatricule_Vehicules(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
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
            If (Lsv_Depot.ListItems(i).SubItems(1) = LSV_Exterieur.ListItems(J).SubItems(2)) And (Len(LSV_Exterieur.ListItems(J).SubItems(6)) = 0) Then Lsv_Depot.ListItems.Remove (i)
        Next
    Next
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheExterieur()
    Dim LObj_Find As New Traffic, Lrs_Find As New Recordset
    Dim DateSys As Date
    Dim min As Long
    Dim heur As Long
    Dim Dur As Long
    Dim temp As String
    Dim Name_Tab As String
On Error GoTo Err
    DateSys = Date
    LSV_Exterieur.ListItems.Clear
    Name_Tab = "FicheTraffic"
    'Voitures en exterieure de plus d'un jours
    Set Lrs_Find = LObj_Find.GetAll_TrafficVehiculeExterieur(ErrNumber, ErrDescription, ErrSourceDetail, Name_Tab, DateSys, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
            Set itmX = LSV_Exterieur.ListItems.Add(, , "")
            itmX.SubItems(1) = CStr(Lrs_Find("Numero"))
            itmX.SubItems(2) = Lrs_Find("LibelleCond")
            itmX.SubItems(3) = Lrs_Find("Matriculevehic")
            itmX.SubItems(4) = Lrs_Find("LibelleDest")
            itmX.SubItems(5) = Format(Lrs_Find("heureSortie"), "hh:mm")
            If Not IsNull(Lrs_Find("HeureENtre")) Then itmX.SubItems(6) = Format(Lrs_Find("HeureENtre"), "hh:mm")
            If Not IsNull(Lrs_Find("CompteurSortie")) Then itmX.SubItems(7) = Lrs_Find("CompteurSortie")
            If Not IsNull(Lrs_Find("CompteurEntre")) Then itmX.SubItems(8) = Lrs_Find("CompteurEntre")
            If Not IsNull(Lrs_Find("CompteurEntre")) Then itmX.SubItems(9) = Val(Lrs_Find("CompteurEntre")) - Val(Lrs_Find("CompteurSortie")) & " KM"
            If IsNull(Lrs_Find("HeureENtre")) Then itmX.SubItems(10) = Format(Lrs_Find("heureSortie"), "dd/mm/yyyy hh:mm")
            itmX.SubItems(11) = Lrs_Find("OperateurSortie")
            If Not IsNull(Lrs_Find("OperateurEntre")) Then itmX.SubItems(12) = Lrs_Find("OperateurEntre")
            Lrs_Find.MoveNext
        Wend
    End If
    Set Lrs_Find = Nothing
    'Detail des fiches d'aujourdhui
    Set LObj_Find = New Traffic
    Set Lrs_Find = LObj_Find.GetAll_TrafficByDateSys(ErrNumber, ErrDescription, ErrSourceDetail, Name_Tab, DateSys, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
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
            itmX.SubItems(2) = Lrs_Find("libelleCond")
            itmX.SubItems(3) = Lrs_Find("Matriculevehic")
            itmX.SubItems(4) = Lrs_Find("libelleDest")
            itmX.SubItems(5) = Format(Lrs_Find("heureSortie"), "hh:mm")
            If Not IsNull(Lrs_Find("HeureENtre")) Then itmX.SubItems(6) = Format(Lrs_Find("HeureENtre"), "hh:mm")
            If Not IsNull(Lrs_Find("CompteurSortie")) Then itmX.SubItems(7) = Lrs_Find("CompteurSortie")
            If Not IsNull(Lrs_Find("CompteurEntre")) Then itmX.SubItems(8) = Lrs_Find("CompteurEntre")
            If Not IsNull(Lrs_Find("CompteurEntre")) Then itmX.SubItems(9) = Val(Lrs_Find("CompteurEntre")) - Val(Lrs_Find("CompteurSortie")) & " KM"
            If Not IsNull(Lrs_Find("HeureENtre")) Then
                'Calcule de durée
                Dur = DateDiff("n", Lrs_Find("HeureSortie"), Lrs_Find("HeureEntre"))
                heur = Dur \ 60
                min = Dur - (heur * 60)
                temp = CStr(heur) & ":" & CStr(min)
                itmX.SubItems(10) = temp
            Else
                'Calcule de durée
                Dur = DateDiff("n", Lrs_Find("HeureSortie"), Now)
                heur = Dur \ 60
                min = Dur - (heur * 60)
                temp = CStr(heur) & ":" & CStr(min)
                itmX.SubItems(10) = temp
            End If
            itmX.SubItems(11) = Lrs_Find("OperateurSortie")
            If Not IsNull(Lrs_Find("OperateurEntre")) Then itmX.SubItems(12) = Lrs_Find("OperateurEntre")
            Lrs_Find.MoveNext
        Wend
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Afficher Ligne Séléctionnée~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub AfficheRowVehiculeSup(ByVal VCode As String)
    Dim LObj_Find As New VEHICULE
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        If Not IsNull(Lrs_Find("Matricule")) Then
            ComBox_Vehicule.Text = Lrs_Find("code") & "  -  " & Lrs_Find("Matricule")
            VCodeVehicle = Lrs_Find("code")
        Else
            MsgBox "Code introuvable", vbInformation
        End If
    Else
        MsgBox "Code introuvable", vbInformation
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheRowconducteurSup(ByVal VCode As String)
    Dim LObj_Find As New Conducteur
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetRow_Conducteur_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        If Not IsNull(Lrs_Find("Libelle")) Then
            ComBox_Conducteur.Text = Lrs_Find("code") & "  -  " & Lrs_Find("Libelle")
            VCodeDrive = Lrs_Find("code")
        Else
            MsgBox "Code introuvable", vbInformation
        End If
    Else
        MsgBox "Code introuvable", vbInformation
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheRowDestinationSup(ByVal VCode As String)
    Dim LObj_Find As New DESTINATION
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        If Not IsNull(Lrs_Find("Libelle")) Then
            ComBox_Destination.Text = Lrs_Find("Numero") & "  -  " & Lrs_Find("Libelle")
            VCodeDestination = Lrs_Find("Numero")
        Else
            MsgBox "Code introuvable", vbInformation
        End If
    Else
        MsgBox "Code introuvable", vbInformation
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Séléctionnée ByComboBox~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub ComBox_Conducteur_Click()
    VCodeDrive = ComBox_Conducteur.FirstValue
End Sub
Private Sub ComBox_Vehicule_Click()
    VCodeVehicle = ComBox_Vehicule.FirstValue
End Sub
Private Sub ComBox_Destination_Click()
    VCodeDestination = ComBox_Destination.FirstValue
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Afficher les Anomalies~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Cmd_Recherche_Click()
    Call AfficheAnomali
End Sub
Private Sub AfficheAnomali()
    Dim LObj_Find As New Traffic, Lrs_Find As New Recordset
    Dim DateFin As Date, DATEDEBUT As Date, YearTrafic As Integer, Name_Table As String
    Dim TypeAn As String
On Error GoTo Err
    DATEDEBUT = Cda_DateDebut.Value
    DateFin = Cda_DateFin.Value
    NAnomalieDuree = 0
    NAnomalieKm = 0
    NAnomalieTotal = 0
    Grid_An.ClearRows
    If DATEDEBUT > DateFin Then
        MsgBox "Période de recherche invalide!..." & vbCr & "Vérifier période et date", vbExclamation, App.ProductName
        Exit Sub
    End If
    For YearTrafic = Year(DATEDEBUT) To Year(DateFin)
        Name_Table = "FicheTraffic"
        If YearTrafic < Year(Date) Then Name_Table = "FicheTraffic_" & YearTrafic
        Set Lrs_Find = LObj_Find.GETALL_SUPERVISIONTRAFFICBYDATE(ErrNumber, ErrDescription, ErrSourceDetail, Name_Table, DATEDEBUT, DateFin, VCodeDrive, VCodeVehicle, VCodeDestination, "Anomalie", YearTrafic, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_Find = Nothing
        If Not Lrs_Find.EOF Then
            Grid_An.Redraw = False
            While Not Lrs_Find.EOF
                With Grid_An
                    .AddRow
                    .CellDetails .Rows, .ColumnIndex("Conducteur"), Lrs_Find.Fields("LibelleCond")
                    .CellDetails .Rows, .ColumnIndex("Vehicule"), Lrs_Find.Fields("MatriculeVehic")
                    .CellDetails .Rows, .ColumnIndex("Destination"), Lrs_Find.Fields("LibelleDest")
                    .CellDetails .Rows, .ColumnIndex("DS"), Lrs_Find.Fields("DateSortie")
                    .CellDetails .Rows, .ColumnIndex("HS"), Lrs_Find.Fields("HeureSortie")
                    .CellDetails .Rows, .ColumnIndex("HE"), Lrs_Find.Fields("HeureEntre"), , , &H80FF80, &HFF0000
                    If (Val(Lrs_Find.Fields("Kmt")) < Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") >= Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) >= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") >= Lrs_Find.Fields("MaxDuree")) Then
                        .CellDetails .Rows, .ColumnIndex("Duree"), Lrs_Find.Fields("Duree"), , , &H8080FF
                    Else
                        .CellDetails .Rows, .ColumnIndex("Duree"), Lrs_Find.Fields("Duree")
                    End If
                    If (Val(Lrs_Find.Fields("Kmt")) <= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") >= Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) >= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") >= Lrs_Find.Fields("MaxDuree")) Then .CellDetails .Rows, .ColumnIndex("MaxD"), Lrs_Find.Fields("MaxDuree"), , , &H80FFFF Else .CellDetails .Rows, .ColumnIndex("MaxD"), Lrs_Find.Fields("MaxDuree")
                    .CellDetails .Rows, .ColumnIndex("CPTS"), Lrs_Find.Fields("CompteurSortie")
                    .CellDetails .Rows, .ColumnIndex("CPTE"), Lrs_Find.Fields("CompteurEntre")
                    If (Val(Lrs_Find.Fields("Kmt")) >= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") <= Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) >= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") >= Lrs_Find.Fields("MaxDuree")) Then .CellDetails .Rows, .ColumnIndex("Km"), Lrs_Find.Fields("Kmt"), , , &H8080FF Else .CellDetails .Rows, .ColumnIndex("Km"), Lrs_Find.Fields("Kmt")
                    If (Val(Lrs_Find.Fields("Kmt")) >= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") <= Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) >= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") >= Lrs_Find.Fields("MaxDuree")) Then .CellDetails .Rows, .ColumnIndex("MaxKm"), Lrs_Find.Fields("MaxCompteur"), , , &H80FFFF Else .CellDetails .Rows, .ColumnIndex("MaxKm"), Lrs_Find.Fields("MaxCompteur")
                    .CellDetails .Rows, .ColumnIndex("OS"), Lrs_Find.Fields("OperateurSortie")
                    .CellDetails .Rows, .ColumnIndex("OE"), Lrs_Find.Fields("OperateurEntre")
                    If Val(Lrs_Find.Fields("Kmt")) >= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") < Lrs_Find.Fields("MaxDuree") Then
                        .CellDetails .Rows, .ColumnIndex("TypeAn"), "Km", , , &H8080FF
                        NAnomalieKm = NAnomalieKm + 1
                    End If
                    If Val(Lrs_Find.Fields("Kmt")) < Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") >= Lrs_Find.Fields("MaxDuree") Then
                        .CellDetails .Rows, .ColumnIndex("TypeAn"), "Durée", , , &H8080FF
                        NAnomalieDuree = NAnomalieDuree + 1
                    End If
                    If Val(Lrs_Find.Fields("Kmt")) >= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") >= Lrs_Find.Fields("MaxDuree") Then
                        .CellDetails .Rows, .ColumnIndex("TypeAn"), "Km & Durée", , , &H8080FF
                        NAnomalieDuree = NAnomalieDuree + 1
                        NAnomalieKm = NAnomalieKm + 1
                    End If
                    .CellDetails .Rows, .ColumnIndex("DifKm"), Lrs_Find.Fields("DifK"), , , &H80C0FF, &HFF0000
                    If Lrs_Find.Fields("Duree") >= Lrs_Find.Fields("MaxDuree") Then
                        .CellDetails .Rows, .ColumnIndex("DifD"), Lrs_Find.Fields("DifD"), , , &H80C0FF, &HFF0000
                    Else
                        .CellDetails .Rows, .ColumnIndex("DifD"), "- " & Lrs_Find.Fields("DifDm"), , , &H80C0FF, &HFF0000
                    End If
                End With
                Lrs_Find.MoveNext
            Wend
            NAnomalieTotal = Grid_An.Rows
            Grid_An.Redraw = True
            Grid_An.SelectedRow = 1
            Set Lrs_Find = Nothing
            Tab_Spervision.TabCaption(1) = "Anomalie   (" & NAnomalieTotal & ")"
            Lbl_AnomalieTotal.Caption = "Anomalie Total  (" & Grid_An.Rows & ")"
            Lbl_AnomalieKm = "Anomalie Km(" & NAnomalieKm & ")"
            Lbl_AnomalieDuree = "Anomalie Duree(" & NAnomalieDuree & ")"
        End If
        Set Lrs_Find = Nothing
    Next
    If Grid_An.Rows = 0 Then
        MsgBox "Aucune Anomalie!... dans cette période", vbInformation, App.ProductName
        Grid_An.ClearRows
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Imprimer les Anomalies~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Cmd_Print_Click()
    Dim DateFin As Date, DATEDEBUT As Date
On Error GoTo Err
    Call AfficheAnomali
    DATEDEBUT = Cda_DateDebut.Value
    DateFin = Cda_DateFin.Value
    If DATEDEBUT > DateFin Then
        MsgBox "Période de recherche invalide!..." & vbCr & "Vérifier période et longure de date", vbExclamation, App.ProductName
        Exit Sub
    End If
    If Grid_An.Rows <> 0 Then
        If MsgBox("Imprission de la liste de(s) anomalie(s) en cours...?", vbYesNo + vbDefaultButton1 + vbInformation, App.ProductName) = vbYes Then
            Call Frm_Rpt_Apercus.PrintOutAndApercu_AnomalieTrafic(0, DATEDEBUT, DateFin, VCodeDrive, VCodeVehicle, VCodeDestination, LStr_NameUser, CStr(NAnomalieKm), CStr(NAnomalieDuree), CStr(NAnomalieTotal), True)
            Frm_Rpt_Apercus.Show
        End If
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Initialise liste de Rechercher~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Cmd_Vider_Click()
    ComBox_Conducteur.ListIndex = 0
    ComBox_Vehicule.ListIndex = 0
    ComBox_Destination.ListIndex = 0
    Grid_An.ClearRows
    Cda_DateDebut.Value = Date
    Cda_DateFin.Value = Date
    Tab_Spervision.TabCaption(1) = "Anomalie   ( )"
    Lbl_AnomalieTotal.Caption = ""
    Lbl_AnomalieKm.Caption = ""
    Lbl_AnomalieDuree.Caption = ""
End Sub
'~~~~~~~~~~~~~~~~~~
    'Control Box~~~
'~~~~~~~~~~~~~~~~~~
Private Sub ComBox_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call AfficheAnomali
End Sub
Private Sub ComBox_Destination_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call AfficheAnomali
End Sub
Private Sub ComBox_Vehicule_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call AfficheAnomali
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Msg As String
On Error GoTo Err
    Msg = "Voulez-vous vraiment quitter?"
    If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Cancel = True Else Unload Me
Exit Sub
Err:
   MsgBox Err.Description, 48, App.ProductName
End Sub

Private Sub grid_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub grid_vehicule_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub



Private Sub Timer1_Timer()
    Lbl_heure = Time
End Sub
Private Sub grid_Conducteur_ColumnClick(ByVal lCol As Long)
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
      Case "Libelle"
         .SortType(1) = CCLSortStringNoCase
         .SortType(1) = CCLSortBackColor
      End Select
   End With
   Screen.MousePointer = vbHourglass
   grid_Conducteur.Sort
   Screen.MousePointer = vbDefault
End Sub
Private Sub grid_vehicule_ColumnClick(ByVal lCol As Long)
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
      Case "Matricule"
         .SortType(1) = CCLSortBackColor
      End Select
   End With
   Screen.MousePointer = vbHourglass
   grid_vehicule.Sort
   Screen.MousePointer = vbDefault
End Sub
Private Sub LSV_Exterieur_ItemClick(ByVal Item As ListItem)
    Dim Lobj_Conducteur As Conducteur
    Dim Lrs_Conducteur As Recordset
    Dim Lobj_Vehicule As VEHICULE
    Dim Lrs_Vehicule As Recordset
    Dim Pic_Conducteur As String
    Dim Pic_Vehicule As String
    Dim Conducteur As String
    Dim VEHICULE As String
    Dim DESTINATION As String
    Dim i As Integer
    i = LSV_Exterieur.SelectedItem.Index
On Error GoTo Err
    If LSV_Exterieur.SelectedItem.Index <> 0 Then
        Conducteur = LSV_Exterieur.ListItems(i).SubItems(2)
            '-- Photo Conducteur***
                Set Lobj_Conducteur = New Conducteur
                Set Lrs_Conducteur = Lobj_Conducteur.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, Conducteur, CNB)
                If ErrNumber <> 0 Then
                    ErrNumber = 0
                    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                    Exit Sub
                End If
                Set Lobj_Conducteur = Nothing
                If Not Lrs_Conducteur.EOF Then
                    If IsNull(Lrs_Conducteur("PicBox")) Then Pic_Conducteur = "Null" Else Pic_Conducteur = Lrs_Conducteur("PicBox")
                End If
                Set Lrs_Conducteur = Nothing
        
        VEHICULE = LSV_Exterieur.ListItems(i).SubItems(3)
            '-- Code Vehicule***
                Set Lobj_Vehicule = New VEHICULE
                Set Lrs_Vehicule = Lobj_Vehicule.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VEHICULE)
                If ErrNumber <> 0 Then
                    ErrNumber = 0
                    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                    Exit Sub
                End If
                Set Lobj_Vehicule = Nothing
                If Not Lrs_Vehicule.EOF Then
                    If IsNull(Lrs_Vehicule("PicBox")) Then Pic_Vehicule = "Null" Else Pic_Vehicule = Lrs_Vehicule("PicBox")
                End If
                Set Lrs_Vehicule = Nothing
                
        DESTINATION = LSV_Exterieur.ListItems(i).SubItems(4)
        Lbl_Destination.Visible = True
        Img_Destination.Visible = True
        Img_Conducteur.Visible = True
        Lbl_Conducteur.Visible = True
        Img_Vehicule.Visible = True
        Lbl_Vehicule.Visible = True
        Lbl_Destination.Caption = DESTINATION
        Lbl_Vehicule.Caption = VEHICULE
        Lbl_Conducteur.Caption = Conducteur
    On Error Resume Next
        Img_Conducteur.Picture = LoadPicture("\\srv-files\Centrano\Image Parcano\Personnel\" & Pic_Conducteur)
        Img_Vehicule.Picture = LoadPicture("\\srv-files\Centrano\Image Parcano\Vehicule\" & Pic_Vehicule)
        Img_Destination.Picture = LoadPicture("\\srv-files\Centrano\Image Parcano\Dest\Dest.jpg")
    On Error GoTo Err
    Else
        Lbl_Destination.Visible = False
        Img_Destination.Visible = False
        Img_Conducteur.Visible = False
        Lbl_Conducteur.Visible = False
        Img_Vehicule.Visible = False
        Lbl_Vehicule.Visible = False
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub
'~~~~~~~~~~~~~~~~
    'Actualise~~~
'~~~~~~~~~~~~~~~~
Private Sub cmd_r_Click()
    Call Affiche_Vehicule
    Call Affiche_Conducteur
    Call AfficheDepot
    Call AfficheExterieur
    Call grid_vehicule_ColumnClick(1)
    Call grid_Conducteur_ColumnClick(1)
    Img_Conducteur.Visible = False
    Lbl_Conducteur.Visible = False
    Img_Vehicule.Visible = False
    Lbl_Vehicule.Visible = False
    Lbl_Destination.Visible = False
    Img_Destination.Visible = False
End Sub
Private Sub Cmd_Conducteur_Click()
    Call Affiche_Conducteur
    Call grid_Conducteur_ColumnClick(1)
End Sub
Private Sub Cmd_Vehicule_Click()
    Call Affiche_Vehicule
    Call grid_vehicule_ColumnClick(1)
End Sub
