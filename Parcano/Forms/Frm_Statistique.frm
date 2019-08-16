VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Statistiques 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Statistiques Carburant"
   ClientHeight    =   10425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16320
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   16320
   Begin VB.CommandButton Stat_Dest 
      Caption         =   "Statistiques Destinations"
      Height          =   495
      Left            =   9360
      TabIndex        =   46
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin TabDlg.SSTab Tab_Satistiques 
      Height          =   8295
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Statistiques Carburant"
      TabPicture(0)   =   "Frm_Statistique.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Pic_ControlStatCarburant"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Satistiques Reparation"
      TabPicture(1)   =   "Frm_Statistique.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Pic_ControlStatR"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Satistiques Trafic"
      TabPicture(2)   =   "Frm_Statistique.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Pic_ControlStatFT"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Satistiques En/Hors Service"
      TabPicture(3)   =   "Frm_Statistique.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Pic_ControlStatPersonnel"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Statistiques Destinations"
      TabPicture(4)   =   "Frm_Statistique.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Picture1"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.PictureBox Pic_ControlStatR 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   -74880
         ScaleHeight     =   7815
         ScaleWidth      =   14415
         TabIndex        =   32
         Top             =   360
         Width           =   14415
         Begin MSComctlLib.ListView List_detailRp 
            Height          =   4455
            Left            =   120
            TabIndex        =   42
            Top             =   2280
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   7858
            SortKey         =   1
            View            =   3
            LabelEdit       =   1
            SortOrder       =   -1  'True
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Pièce"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Vehicule"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Désignation"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Qte"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "TOT.TTC.NET"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView List_DetailsRp 
            Height          =   1335
            Left            =   120
            TabIndex        =   44
            Top             =   840
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   2355
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Vehicule"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Nombre réparation"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Key             =   "valeur"
               Text            =   "valeur Reparation"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.ComboBox Cbo_Vehicule 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "Frm_Statistique.frx":008C
            Left            =   1560
            List            =   "Frm_Statistique.frx":008E
            TabIndex        =   33
            Top             =   240
            Width           =   4095
         End
         Begin SToolBox.SCommand cmd_FindMatricule 
            Height          =   345
            Left            =   5760
            TabIndex        =   34
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
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
            Picture         =   "Frm_Statistique.frx":0090
            ButtonType      =   1
         End
         Begin MSComCtl2.DTPicker Dta_Fin 
            Height          =   375
            Left            =   9600
            TabIndex        =   35
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   142868481
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker Dta_Debut 
            Height          =   375
            Left            =   7080
            TabIndex        =   63
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   142868481
            CurrentDate     =   42860
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   9000
            TabIndex        =   39
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   6480
            TabIndex        =   38
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Véhicule:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   1350
         End
         Begin VB.Image Cmd_Find 
            Height          =   495
            Left            =   11760
            Picture         =   "Frm_Statistique.frx":03CA
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.PictureBox Pic_ControlStatFT 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   -74880
         ScaleHeight     =   7815
         ScaleWidth      =   14415
         TabIndex        =   17
         Top             =   360
         Width           =   14415
         Begin MSComctlLib.ListView Lsv_DetailsFT 
            Height          =   1095
            Left            =   0
            TabIndex        =   45
            Top             =   1080
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   1931
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nombre des voyages"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Tot.Durée"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Moy.duré"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Tot.Distance"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Moy.Distance"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.ComboBox cbo_VehiculeFT 
            Height          =   315
            ItemData        =   "Frm_Statistique.frx":10FCC
            Left            =   1080
            List            =   "Frm_Statistique.frx":10FCE
            TabIndex        =   20
            Top             =   120
            Width           =   2055
         End
         Begin VB.ComboBox cbo_ConducteurFT 
            Height          =   315
            ItemData        =   "Frm_Statistique.frx":10FD0
            Left            =   5160
            List            =   "Frm_Statistique.frx":10FD2
            TabIndex        =   19
            Top             =   120
            Width           =   2055
         End
         Begin VB.ComboBox cbo_DestinationFT 
            Height          =   315
            Left            =   9360
            TabIndex        =   18
            Top             =   120
            Width           =   2055
         End
         Begin SToolBox.SCommand Cmd_FindVehiculeFT 
            Height          =   315
            Left            =   3240
            TabIndex        =   24
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
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
            Picture         =   "Frm_Statistique.frx":10FD4
            ButtonType      =   1
         End
         Begin SToolBox.SCommand Cmd_FindConducteurFT 
            Height          =   315
            Left            =   7320
            TabIndex        =   25
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
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
            Picture         =   "Frm_Statistique.frx":1130E
            ButtonType      =   1
         End
         Begin SToolBox.SCommand Cmd_FindDestinationFT 
            Height          =   315
            Left            =   11520
            TabIndex        =   26
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
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
            Picture         =   "Frm_Statistique.frx":11648
            ButtonType      =   1
         End
         Begin MSComCtl2.DTPicker cda_FinFT 
            Height          =   375
            Left            =   9000
            TabIndex        =   27
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   142934017
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker cda_Debutft 
            Height          =   375
            Left            =   6480
            TabIndex        =   28
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   142934017
            CurrentDate     =   42860
         End
         Begin SToolBox.SGrid grid_Ft 
            Height          =   5655
            Left            =   0
            TabIndex        =   31
            Top             =   2280
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   9975
            RowMode         =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   2
            DisableIcons    =   -1  'True
            MaxVisibleRows  =   0
         End
         Begin VB.Image Cmd_SearchFT 
            Height          =   495
            Left            =   11280
            Picture         =   "Frm_Statistique.frx":11982
            Stretch         =   -1  'True
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   8400
            TabIndex        =   30
            Top             =   600
            Width           =   600
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   5880
            TabIndex        =   29
            Top             =   600
            Width           =   600
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Vehicule"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Conducteur"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3720
            TabIndex        =   22
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Destination"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7800
            TabIndex        =   21
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.PictureBox Pic_ControlStatPersonnel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   -74880
         ScaleHeight     =   7815
         ScaleWidth      =   14415
         TabIndex        =   8
         Top             =   360
         Width           =   14415
         Begin VB.ComboBox cbo_Conducteur 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "Frm_Statistique.frx":22584
            Left            =   1920
            List            =   "Frm_Statistique.frx":22586
            TabIndex        =   13
            Top             =   240
            Width           =   3735
         End
         Begin SToolBox.SGrid grid_Service 
            Height          =   6735
            Left            =   0
            TabIndex        =   10
            Top             =   960
            Width           =   12015
            _ExtentX        =   21193
            _ExtentY        =   11880
            RowMode         =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   2
            DisableIcons    =   -1  'True
            MaxVisibleRows  =   0
         End
         Begin MSComCtl2.DTPicker cda_FinService 
            Height          =   375
            Left            =   9600
            TabIndex        =   11
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   142868481
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker cda_DebutService 
            Height          =   375
            Left            =   7080
            TabIndex        =   12
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   142868481
            CurrentDate     =   42860
         End
         Begin SToolBox.SCommand cmdFindConducteur 
            Height          =   345
            Left            =   5760
            TabIndex        =   14
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
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
            Picture         =   "Frm_Statistique.frx":22588
            ButtonType      =   1
         End
         Begin VB.Image Cmd_SearchService 
            Height          =   495
            Left            =   11760
            Picture         =   "Frm_Statistique.frx":228C2
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conducteur:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   6480
            TabIndex        =   16
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   9000
            TabIndex        =   15
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.PictureBox Pic_ControlStatCarburant 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   -74880
         ScaleHeight     =   7815
         ScaleWidth      =   14415
         TabIndex        =   2
         Top             =   360
         Width           =   14415
         Begin MSComctlLib.ListView Lsv_Details 
            Height          =   1455
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   12375
            _ExtentX        =   21828
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Vehicule"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "valeur Carburant"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Prix Litre"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "NB.Litre Carburant"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Km.Parcouru"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Consommation par cent km"
               Object.Width           =   2540
            EndProperty
         End
         Begin SToolBox.SDateBox cda_fin 
            Height          =   285
            Left            =   9600
            TabIndex        =   41
            Top             =   240
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_debut 
            Height          =   285
            Left            =   7080
            TabIndex        =   40
            Top             =   240
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SGrid Grid_Carb 
            Height          =   5415
            Left            =   120
            TabIndex        =   36
            Top             =   2280
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   9551
            RowMode         =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            GroupRowForeColor=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Editable        =   -1  'True
            DisableIcons    =   -1  'True
            MaxVisibleRows  =   0
         End
         Begin VB.ComboBox cbo_Matricule 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "Frm_Statistique.frx":334C4
            Left            =   1560
            List            =   "Frm_Statistique.frx":334C6
            TabIndex        =   3
            Top             =   240
            Width           =   4095
         End
         Begin SToolBox.SCommand cmdFindMatricule 
            Height          =   345
            Left            =   5760
            TabIndex        =   4
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
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
            Picture         =   "Frm_Statistique.frx":334C8
            ButtonType      =   1
         End
         Begin VB.Image Cmd_Search 
            Height          =   495
            Left            =   11400
            Picture         =   "Frm_Statistique.frx":33802
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   9000
            TabIndex        =   7
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   6480
            TabIndex        =   6
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Véhicule:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   1350
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   120
         ScaleHeight     =   7815
         ScaleWidth      =   14415
         TabIndex        =   47
         Top             =   360
         Width           =   14415
         Begin VB.PictureBox Pic_Details_Dest 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   6735
            Left            =   1200
            ScaleHeight     =   6705
            ScaleWidth      =   11865
            TabIndex        =   49
            Top             =   960
            Width           =   11895
            Begin SToolBox.SGrid SGrid_Details 
               Height          =   5295
               Left            =   120
               TabIndex        =   50
               Top             =   1080
               Width           =   11655
               _ExtentX        =   20558
               _ExtentY        =   9340
               RowMode         =   -1  'True
               BackgroundPictureHeight=   0
               BackgroundPictureWidth=   0
               BackColor       =   16777152
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   2
               DisableIcons    =   -1  'True
               SelectionAlphaBlend=   -1  'True
               SelectionOutline=   -1  'True
               MaxVisibleRows  =   0
            End
            Begin VB.Label Label15 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Retour"
               BeginProperty Font 
                  Name            =   "Sitka Text"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   10920
               TabIndex        =   53
               Top             =   6360
               Width           =   735
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Height          =   375
               Left            =   2280
               TabIndex        =   52
               Top             =   480
               Width           =   4335
            End
            Begin VB.Label Label17 
               Caption         =   "Destination:"
               BeginProperty Font 
                  Name            =   "Sitka Small"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   360
               TabIndex        =   51
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.ComboBox Cbo_Destination 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "Frm_Statistique.frx":44404
            Left            =   1800
            List            =   "Frm_Statistique.frx":44406
            TabIndex        =   48
            Top             =   240
            Width           =   3735
         End
         Begin MSComctlLib.ListView List_Details1 
            Height          =   6735
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   11880
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Numero"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Destination"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Nombre Tournées"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Key             =   "valeur"
               Text            =   "Total Distance/KM"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Total Durée/H:M"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Moyenne Distance/KM"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Moyenne Durée/H"
               Object.Width           =   2822
            EndProperty
         End
         Begin SToolBox.SCommand cmd_FindDest 
            Height          =   345
            Left            =   5760
            TabIndex        =   55
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
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
            Picture         =   "Frm_Statistique.frx":44408
            ButtonType      =   1
         End
         Begin MSComCtl2.DTPicker Dta_Fin_Dest 
            Height          =   375
            Left            =   9600
            TabIndex        =   56
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   302972929
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker Dta_Debut_Dest 
            Height          =   375
            Left            =   7080
            TabIndex        =   60
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   302972929
            CurrentDate     =   42860
         End
         Begin VB.Image Cmd_Search_Dest 
            Height          =   495
            Left            =   11880
            Picture         =   "Frm_Statistique.frx":44742
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destination"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   6480
            TabIndex        =   58
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   9000
            TabIndex        =   57
            Top             =   240
            Width           =   600
         End
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7320
      TabIndex        =   61
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   14737632
      Format          =   302972929
      CurrentDate     =   42860
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   7320
      TabIndex        =   62
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   14737632
      Format          =   302972929
      CurrentDate     =   42860
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "Frm_Statistique.frx":55344
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image CmdPrint 
      Height          =   495
      Left            =   11880
      Picture         =   "Frm_Statistique.frx":72A9E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label m 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statistiques"
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
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Image PicBox_Header 
      Height          =   1005
      Left            =   0
      Picture         =   "Frm_Statistique.frx":836A0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14535
   End
End
Attribute VB_Name = "Frm_Statistiques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim thekey As Integer
    Dim theshift As Integer
    Dim itmX
    Dim VCodeVehicle As String              'Code Vehicule
    Dim VCodeDrive  As String               'Code Conducteur
    Dim VCodeDestination  As String         'Code Destination
    Dim NAnomalieTotal As Integer           'Nombre Anomalie Total***
    Dim NAnomalieKm As Integer              'Nombre Anomalie Km***
    Dim NAnomalieDuree As Integer           'Nombre Anomalie Durée***




Private Sub Cbo_Destination_Click()
List_Details1.ListItems.Clear
Pic_Details_Dest.Visible = False
Dim LOBJ_Dest As New DESTINATION
    Dim Lrs_Dest As Recordset
    If Cbo_Destination.ListIndex = 0 Then
        VCodeDestination = "  -  Tous"
    Else
        Set Lrs_Dest = LOBJ_Dest.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, Cbo_Destination.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LOBJ_Dest = Nothing
        
        If Not Lrs_Dest.EOF Then VCodeDestination = Lrs_Dest("Numero")
    End If
End Sub

Private Sub cmd_FindDest_Click()
Pic_Details_Dest.Visible = False
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "STATDestinations"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub

Private Sub Cmd_Search_Dest_Click()
Pic_Details_Dest.Visible = False

  Dim VCode As String, LObj_V As New DESTINATION, Lrs_V As New Recordset
On Error GoTo Err
    If Dta_Debut_Dest.Value > Dta_Fin_Dest.Value Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    List_Details1.ListItems.Clear
  
'    VCode = Cbo_Destination.Text
'     If VCode <> "Tous" And VCode <> "" Then
'        Set Lrs_V = LObj_V.GetDestByLibDest(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
'        If ErrNumber <> 0 Then
'            ErrNumber = 0
'            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
'            Exit Sub
'        End If
'        Set LObj_V = Nothing
'        If Not Lrs_V.EOF Then
'            VCode = Lrs_V("Numero")
'            Set Lrs_V = Nothing
'        Else
'            MsgBox "Destination introuvable!..."
'            Set Lrs_V = Nothing
'            Exit Sub
'        End If
'    End If
    If Cbo_Destination.Text = "Tous" Then
        Call AfficheDetails_Tous_Dest(Dta_Debut_Dest.Value, Dta_Fin_Dest.Value)
    Else
        Call AfficheDetails_ParDestination(Cbo_Destination.Text, CDate(Dta_Debut_Dest.Value), CDate(Dta_Fin_Dest.Value))
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheDetails_Tous_Dest(ByVal VdateD As Date, ByVal vDateF As Date)
    'variables globales
    Dim LOBJ_StatRp As DESTINATION
    Dim rs As New Recordset
    Dim i
    'variables Details
    Dim nbVoy As Double
    Dim Valeur As Double
    Dim duree As Double
    Dim dureeM As Long
    Dim heurmin As Long
    Dim heur As Long
    Dim min As Long
    Dim Moy_Duree As Long
    Dim Moy_DureeM As Long
    Dim Calcul As Long
    Dim HeurMoy As Long
    Dim MinMoy As Long
    Dim Moy_Dist As Double
On Error GoTo Err

    Set itmX = List_Details1.ListItems.Add()
    
    Set LOBJ_StatRp = New DESTINATION
    'nombre des Voyages
    Set rs = LOBJ_StatRp.Get_SumNbrVoyage(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
    
    itmX.SubItems(1) = "Tous"
        nbVoy = 0
        nbVoy = nbVoy + rs("nbVoy")
        
        itmX.SubItems(2) = CStr(rs("nbVoy"))
        'Total Distance
        If Not IsNull(rs("valeur")) Then
            Valeur = 0
            Valeur = Valeur + rs("valeur")
            Moy_Dist = Valeur / nbVoy
            itmX.SubItems(3) = CStr(Format(rs("valeur"), "#,##0.0"))
            itmX.SubItems(5) = Format(Moy_Dist, "#,##0.0")
        Else
            itmX.SubItems(3) = "Valeur Null"
            itmX.SubItems(5) = "Valeur Null"
        End If
        'Total Durée
        If Not IsNull(rs("totduree")) And Not IsNull(rs("totdureeM")) Then
            duree = 0
            dureeM = 0
'            duree = duree + rs("totduree")
            dureeM = dureeM + rs("totdureeM")
            heur = rs("totdureeM") \ 60
            heurmin = 60 * heur
            min = dureeM - heurmin
'            Moy_Duree = duree / nbVoy
            Moy_DureeM = dureeM \ nbVoy
            HeurMoy = Moy_DureeM \ 60
            Calcul = HeurMoy * 60
            MinMoy = Moy_DureeM - Calcul
'            itmX.SubItems(4) = CStr(Format(rs("totduree"), "#,##0.0"))
            itmX.SubItems(4) = CStr(heur) & " : " & CStr(Format(min, "##00"))
            itmX.SubItems(6) = CStr(HeurMoy) & " : " & CStr(Format(MinMoy, "##00"))
        Else
            itmX.SubItems(4) = "Duree Null"
            itmX.SubItems(6) = "Moyenne Duree Null"
        End If
    End If
    rs.Close
'    'Destination + nbr voyage par destination
    Set rs = LOBJ_StatRp.Get_NbrRepStatistGrpDest(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
'    Numero DESTINATION
        
    If Not rs.EOF Then
        While Not rs.EOF
        
            
            If Not IsNull(rs("nbVoy")) Then
            nbVoy = 0
            nbVoy = nbVoy + rs("nbVoy")
            Set itmX = List_Details1.ListItems.Add(, , rs("Numero"))
            itmX.SubItems(1) = CStr(rs("Libelle"))
            itmX.SubItems(2) = CStr(rs("nbVoy"))
            End If
      'Total Distance
            If Not IsNull(rs("valeur")) Then
                Valeur = 0
                Valeur = Valeur + rs("valeur")
                Moy_Dist = Valeur / nbVoy
                itmX.SubItems(3) = CStr(Format(rs("valeur"), "#,##0.0"))
                itmX.SubItems(5) = Format(Moy_Dist, "#,##0.0")
            Else
                itmX.SubItems(3) = "Valeur Null"
                itmX.SubItems(5) = "Valeur Null"
            End If
      'Total Durée
        If Not IsNull(rs("totduree")) Then
           duree = 0
            dureeM = 0
'            duree = duree + rs("totduree")
            dureeM = dureeM + rs("totdureeM")
            heur = rs("totdureeM") \ 60
            heurmin = 60 * heur
            min = dureeM - heurmin
'            Moy_Duree = duree / nbVoy
           Moy_DureeM = dureeM \ nbVoy
            HeurMoy = Moy_DureeM \ 60
            Calcul = HeurMoy * 60
            MinMoy = Moy_DureeM - Calcul
           itmX.SubItems(4) = CStr(heur) & " : " & CStr(Format(min, "##00"))
             itmX.SubItems(6) = CStr(HeurMoy) & " : " & CStr(Format(MinMoy, "##00"))
        Else
            itmX.SubItems(4) = "Duree Null"
            itmX.SubItems(6) = "Moyenne Duree Null"
        End If
            
            
        rs.MoveNext
        Wend
    End If
    rs.Close
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheDetails_ParDestination(ByVal Libelle As String, ByVal VdateD As Date, ByVal vDateF As Date)
    Dim LOBJ_StatRp As New DESTINATION
    Dim rs As New Recordset
   
    'variables Details
   Dim nbVoy As Double
    Dim Valeur As Double
'    Dim duree As Double
    Dim dureeM As Long
    Dim heurmin As Long
    Dim heur As Long
    Dim min As Long
'    Dim Moy_Duree As Long
    Dim Moy_DureeM As Long
    Dim Calcul As Long
    Dim HeurMoy As Long
    Dim MinMoy As Long
    Dim Moy_Dist As Double
On Error GoTo Err

    'nombre des Voyages
    Set rs = LOBJ_StatRp.Get_ValRepStatistDest(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Libelle, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        Set itmX = List_Details1.ListItems.Add(, , rs("Numero"))
        itmX.SubItems(1) = CStr(rs("Libelle"))
    
        If Not IsNull(rs("nbVoy")) Then
            nbVoy = 0
            nbVoy = nbVoy + rs("nbVoy")
            
            itmX.SubItems(2) = CStr(rs("nbVoy"))
            Else
            itmX.SubItems(2) = "Valeur Null"
        End If
        'Total Distance
        If Not IsNull(rs("valeur")) Then
            Valeur = 0
            Valeur = Valeur + rs("valeur")
            Moy_Dist = Valeur / nbVoy
            itmX.SubItems(3) = CStr(Format(rs("valeur"), "#,##0.0"))
            itmX.SubItems(5) = Format(Moy_Dist, "#,##0.0")
        Else
            itmX.SubItems(3) = "Valeur Null"
            itmX.SubItems(5) = "Valeur Null"
        End If
        'Total Durée
        If Not IsNull(rs("totdureeM")) Then
            
            dureeM = 0
            dureeM = dureeM + rs("totdureeM")
            heur = rs("totdureeM") \ 60
            heurmin = 60 * heur
            min = dureeM - heurmin
            Moy_DureeM = dureeM \ nbVoy
            HeurMoy = Moy_DureeM \ 60
            Calcul = HeurMoy * 60
            MinMoy = Moy_DureeM - Calcul
            itmX.SubItems(4) = CStr(heur) & " : " & CStr(Format(min, "##00"))
            itmX.SubItems(6) = CStr(HeurMoy) & " : " & CStr(Format(MinMoy, "##00"))
        Else
            itmX.SubItems(4) = "Duree Null"
            itmX.SubItems(6) = "Moyenne Duree Null"
        End If
    End If
    rs.Close
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub







Private Sub Dta_Debut_Dest_Click()
Pic_Details_Dest.Visible = False
End Sub



Private Sub Dta_Fin_Dest_Click()
Pic_Details_Dest.Visible = False
End Sub

'~~~~~~~~~~~~~~~~~~~~
    'Mise en Forme~~~
'~~~~~~~~~~~~~~~~~~~~
Private Sub Form_Load()
    Me.WindowState = 2
    Tab_Satistiques.Tab = 0
    cda_debut.Text = "01/" & Month(Date) & "/" & Year(Date)
    cda_fin.Text = Date
    cda_DebutService.Value = "01/" & Month(Date) & "/" & Year(Date)
    cda_FinService.Value = Date
    cda_Debutft.Value = "01/" & Month(Date) & "/" & Year(Date)
    cda_FinFT.Value = Date
    Dta_Debut.Value = "01/" & Month(Date) & "/" & Year(Date)
    Dta_Fin.Value = Date
    Dta_Debut_Dest.Value = "01/" & Month(Date) & "/" & Year(Date)
    Dta_Fin_Dest.Value = Date
    
    
    Call Initgrid_Services
    Call Initgrid_FT
    Call Initgrid_Carb
    Call Initgrid_Details_Dest
    
    
    Cbo_Vehicule.AddItem "Tous", 0
    cbo_Matricule.AddItem "Tous", 0
    cbo_VehiculeFT.AddItem "Tous", 0
    cbo_ConducteurFT.AddItem ("Tous"), 0
    cbo_DestinationFT.AddItem ("Tous"), 0
    Cbo_Destination.AddItem ("Tous"), 0
    Call Affiche_Matricule_Combo(Cbo_Vehicule)
    Call Affiche_Matricule_Combo(cbo_Matricule)
    Call Affiche_Matricule_Combo(cbo_VehiculeFT)
    Call Affiche_Personnel_Combo(cbo_ConducteurFT)
    Call Affiche_Personnel_Combo(cbo_Conducteur)
    Call Affiche_Destination_Combo(cbo_DestinationFT)
    Call Affiche_Destination_Combo(Cbo_Destination)
    
    Cbo_Vehicule.ListIndex = 0
    cbo_Matricule.ListIndex = 0
    cbo_VehiculeFT.ListIndex = 0
    cbo_ConducteurFT.ListIndex = 0
    cbo_DestinationFT.ListIndex = 0
    Cbo_Destination.ListIndex = 0
    VCodeVehicle = "  -  Tous"
    VCodeDrive = "  -  Tous"
    VCodeDestination = "  -  Tous"
    
End Sub
Private Sub Form_Resize()
    Dim WidthForm As Integer, HeightForm As Integer
    WidthForm = Me.Width
    HeightForm = Me.Height
        PicBox_Header.Width = WidthForm
        Tab_Satistiques.Width = WidthForm - 400
        CmdPrint.Left = WidthForm - 2000
        Stat_Dest.Left = WidthForm - 5000
        Pic_ControlStatCarburant.Width = Tab_Satistiques.Width - 200
        Pic_ControlStatPersonnel.Width = Tab_Satistiques.Width - 200
        Pic_ControlStatFT.Width = Tab_Satistiques.Width - 200
        Pic_ControlStatR.Width = Tab_Satistiques.Width - 200
        Picture1.Width = Tab_Satistiques.Width - 200
        Cmd_SearchService.Left = WidthForm - 3000
        Cmd_SearchService.Top = 200
        Cmd_SearchFT.Left = WidthForm - 3000
        Cmd_SearchFT.Top = 200
        Cmd_Find.Left = WidthForm - 3000
        Cmd_Find.Top = 200
        Cmd_Search.Left = WidthForm - 3000
        Cmd_Search.Top = 200
        Cmd_Search_Dest.Left = WidthForm - 3000
        Cmd_Search_Dest.Top = 200
        grid_Service.Width = Tab_Satistiques.Width - 200
        grid_Ft.Width = Tab_Satistiques.Width - 200
        Lsv_DetailsFT.Width = Tab_Satistiques.Width - 200
        List_DetailsRp.Width = Tab_Satistiques.Width - 200
        List_detailRp.Width = Tab_Satistiques.Width - 200
        Lsv_Details.Width = Tab_Satistiques.Width - 200
        Grid_Carb.Width = Tab_Satistiques.Width - 200
        List_Details1.Width = Tab_Satistiques.Width - 200
        SGrid_Details.Width = Tab_Satistiques.Width - 200
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
'~~~~~~~~~~~~~~~~~~~~~~~
    'Initialise SGrid~~~
'~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Initgrid_Services()
    With grid_Service
        .Redraw = False
        .ClearSelection
        .ClearRows
        .Clear
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
        .AddColumn "Numero", "", , , 500, , , , , , , CCLSortNumeric
        .AddColumn "Date", "Date", , , 100
        .AddColumn "Etat", "Etat", , , 0
        .AddColumn "HDebut", "Heure Sortie", , , 100
        .AddColumn "HFin", "Heure Entre", , , 100
        .AddColumn "DureTrafic", "Durée Trafic", , , 100
        .AddColumn "Activités", "Activités", , , 500
        .AddColumn "Null", ""
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Private Sub Initgrid_FT()
    With grid_Ft
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
        .AddColumn "DateFT", "Date", , , 80
        .AddColumn "Numero", "Numero", , , 60, False, , , , , , CCLSortNumeric
        .AddColumn "Matricule", "Matricule", , , 100
        .AddColumn "Conducteur", "Conducteur", , , 100
        .AddColumn "Destination", "Destination", , , 140
        .AddColumn "HeureS", "H.Sortie", , , 60
        .AddColumn "HeureE", "H.Entrée", , , 60
        .AddColumn "CPTS", "CPT.S", , , 60
        .AddColumn "CPTE", "CPT.E", , , 60
        .AddColumn "Distance", "Distance(KM)", , , 40
        .AddColumn "MaxK", "Max-Km", , , 60
        .AddColumn "DifK", "Dif Km", , , 60
        .AddColumn "Dure", "Durée(Heure)", , , 60
        .AddColumn "MaxD", "Max-Durée", , , 60
        .AddColumn "DifD", "Dif-Durée", , , 60 ', False
        .AddColumn "OS", "Op.Sortie)", , , 100
        .AddColumn "OE", "Op.Entre", , , 100
        .AddColumn "Null", ""
        .StretchLastColumnToFit = True
    End With
End Sub
Private Sub Initgrid_Carb()
    With Grid_Carb
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
        .AddColumn "Numero", "Pièce", , , 80, , , , , , , CCLSortNumeric
        .AddColumn "Vehicule", "Véhicule", , , 120
        .AddColumn "Date", "Date", , , 90
        .AddColumn "NbrL", "Nbr.Litres", , , 70
        .AddColumn "Montant", "Montant", , , 90
        .AddColumn "Compteur", "Compteur", , , 90
        .AddColumn "KmParc", "Km.Parcouru", , , 80
        .AddColumn "Consom", "Consomation/100Km", , , 80
        .AddColumn "Anomalie", "Anomalie.consomation", , , 80
        .AddColumn "NULL", ""
        .StretchLastColumnToFit = True
    .Redraw = True
    End With
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Afficher Liste (FindView)~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub cmdFindConducteur_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "ConducteurE/H"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub Cmd_FindConducteurFT_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "ConducteurStque"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub Cmd_FindDestinationFT_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "DestinationE/H"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub Cmd_FindVehiculeFT_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "VehiculeStqueTF"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub cmd_FindMatricule_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "VehiculeStqueRp"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub cmdFindMatricule_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "VehiculeStqueCBr"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub

Private Sub Grid_Carb_ColumnClick(ByVal lCol As Long)
Dim sTag As String
    Dim i As Long
    With Grid_Carb.SortObject
        .Clear
        .SortColumn(1) = lCol
        sTag = Grid_Carb.ColumnTag(lCol)
        If (sTag = "") Then
            sTag = "DESC"
            .SortOrder(1) = CCLOrderAscending
        Else
            sTag = ""
            .SortOrder(1) = CCLOrderDescending
        End If
        grid_Ft.ColumnTag(lCol) = sTag
        Select Case Grid_Carb.ColumnKey(lCol)
            Case "Numero"
                 .SortType(1) = CCLSortNumeric
            Case "Vehicule"
                 .SortType(1) = CCLSortString
            Case "Date"
                 .SortType(1) = CCLSortDate
            Case "NbrL"
                 .SortType(1) = CCLSortNumeric
            Case "Montant"
                 .SortType(1) = CCLSortNumeric
            Case "Compteur"
                 .SortType(1) = CCLSortNumeric
            Case "KmParc"
                 .SortType(1) = CCLSortNumeric
            Case "Consom"
                 .SortType(1) = CCLSortNumeric
            Case "Anomalie"
                 .SortType(1) = CCLSortNumeric
        End Select
    End With
    Screen.MousePointer = vbHourglass
    Grid_Carb.Sort
    Screen.MousePointer = vbDefault
End Sub

'~~~~~~~~~~~~~~~~~
    'ControlBox~~~
'~~~~~~~~~~~~~~~~~
Private Sub grid_Service_ColumnClick(ByVal lCol As Long)
    Dim sTag As String, i As Long
    With grid_Service.SortObject
        .Clear
        .SortColumn(1) = lCol
        sTag = grid_Service.ColumnTag(lCol)
        If (sTag = "") Then
            sTag = "DESC"
            .SortOrder(1) = CCLOrderAscending
        Else
            sTag = ""
            .SortOrder(1) = CCLOrderDescending
        End If
        grid_Service.ColumnTag(lCol) = sTag
        Select Case grid_Service.ColumnKey(lCol)
            Case "Conducteur"
                 .SortType(1) = CCLSortString
            Case "Etat"
                 .SortType(1) = CCLSortString
            Case "HDebut"
                 .SortType(1) = CCLSortDateHourAccuracy
            Case "HFin"
                 .SortType(1) = CCLSortDateHourAccuracy
            Case "Dure"
                 .SortType(1) = CCLSortDateHourAccuracy
        End Select
    End With
    Screen.MousePointer = vbHourglass
    grid_Service.Sort
    Screen.MousePointer = vbDefault
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'Afficher Row***
Public Sub AfficheRow_Vehicule(ByVal VCode As String)
    Dim LObj_Find As New VEHICULE, Lrs_Find As New Recordset, cbo As ComboBox
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        If Frm_FindView.StrSource = "VehiculeStqueCBr" Then Set cbo = cbo_Matricule
        If Frm_FindView.StrSource = "VehiculeStqueRp" Then Set cbo = Cbo_Vehicule
        If Frm_FindView.StrSource = "VehiculeStqueTF" Then Set cbo = cbo_VehiculeFT
        
        If Not IsNull(Lrs_Find("Matricule")) Then
            cbo.Text = Lrs_Find("Matricule")
            VCode = Lrs_Find("code")
        End If
    Else
        MsgBox "Code introuvable", vbInformation
        cbo.SetFocus
        Exit Sub
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
    'Satistiques En/Hors Service***
Private Sub Cmd_SearchService_Click()
    Call Affiche_FP
End Sub
Public Sub Affiche_FP()
    Dim LObj_Find As New Traffic, Lrs_Find As Recordset
    Dim VdateD As Date, vDateF As Date
    Dim Conducteur As String, DESTINATION As String
    Dim YearTrafic As Integer, Name_Table As String
On Error GoTo Err
    Call Initgrid_Services
    Conducteur = cbo_Conducteur.Text
    VdateD = cda_DebutService.Value
    vDateF = cda_FinService.Value
    If cda_DebutService.Value > cda_FinService.Value Then
        MsgBox "Vérifier dates de recherche!...", vbExclamation, App.ProductName
        Exit Sub
    End If
    If Conducteur = "" Then
        MsgBox "Choisir un conducteur !...", vbExclamation, App.ProductName
        Exit Sub
    End If
    For YearTrafic = Year(VdateD) To Year(vDateF)
        Name_Table = "FicheTraffic"
        If YearTrafic < Year(Date) Then Name_Table = "FicheTraffic_" & YearTrafic
        Set Lrs_Find = LObj_Find.GETALL_STATISTIQUESSERVICES(ErrNumber, ErrDescription, ErrSourceDetail, Name_Table, VdateD, vDateF, Conducteur, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_Find = Nothing
    Next
    Dim Ass_Statistique
    If Not Lrs_Find.EOF Then
        grid_Service.Redraw = False
        While Not Lrs_Find.EOF
            If Not ((IsNull(Lrs_Find("HDebut"))) And (IsNull(Lrs_Find("HFin")))) Then Ass_Statistique = UCase(Format(Lrs_Find("DATEdebut"), "dddd-dd-mm-yyyy")) & "     |      Du: " & Lrs_Find("HDebut") & "   |=> Au: " & Lrs_Find("HFin") & "     ||      Durée : " & Lrs_Find("DUREE")
            With grid_Service
                If (Lrs_Find("Etat") = "En-Service") Then
                    .AddRow
                    If Not (IsNull(Lrs_Find("Ndisp"))) Then .CellDetails .Rows, .ColumnIndex("Numero"), Ass_Statistique, , , &HC0FFC0
                    If Not (IsNull(Lrs_Find("Etat"))) Then .CellDetails .Rows, .ColumnIndex("Etat"), Lrs_Find("ETAT"), , , &HC0FFC0
                    If Not (IsNull(Lrs_Find("HDebut"))) Then .CellDetails .Rows, .ColumnIndex("Date"), Lrs_Find("dateDebut"), , , &HC0FFC0
                    If Not (IsNull(Lrs_Find("HDebut"))) Then .CellDetails .Rows, .ColumnIndex("HDebut"), Lrs_Find("Heuresortie"), , , &HC0FFC0
                    If Not (IsNull(Lrs_Find("HFin"))) Then .CellDetails .Rows, .ColumnIndex("HFin"), Lrs_Find("Heureentre"), , , &HC0FFC0
                    If Not (IsNull(Lrs_Find("Heureentre"))) Then .CellDetails .Rows, .ColumnIndex("DureTrafic"), Lrs_Find("DUREETRAFIC"), , , &HC0FFC0
                    DESTINATION = ""
                    If Not (IsNull(Lrs_Find("Destination"))) Then DESTINATION = DESTINATION & " | " & Format(Lrs_Find("HeureSortie"), "hh:mm") & " Aller à " & Lrs_Find("Destination") & " par " & Lrs_Find("vehicule")
                    .CellDetails .Rows, .ColumnIndex("Activités"), DESTINATION, , , &HC0FFC0
                    .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HC0FFC0
                End If
            End With
            Lrs_Find.MoveNext
            Ass_Statistique = ""
        Wend
        grid_Service.Redraw = True
        Set Lrs_Find = Nothing
        With grid_Service
            .GroupRowBackColor = RGB(251, 246, 206)
            .GroupRowForeColor = QBColor(12)
            .ColumnIsGrouped(1) = True
            .GroupRowForeColor = QBColor(10)
            .HideGroupingBox = True
            .AllowGrouping = True
        End With
    End If
    Set Lrs_Find = Nothing
    If grid_Service.Rows > 0 Then grid_Service.SelectedRow = 1
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
    'Satistiques FicheTrafic***
Public Sub Cmd_Searchft_Click()
    Dim VCode As String, CCode As String, DCode As String
    Dim LObj_V As New VEHICULE, LObj_C As New Conducteur, LObj_D As New DESTINATION
    Dim Lrs_V As New Recordset, Lrs_C As New Recordset, Lrs_D As New Recordset
On Error GoTo Err
        VCode = cbo_VehiculeFT.Text
        CCode = cbo_ConducteurFT.Text
        DCode = cbo_DestinationFT.Text
    If VCode <> "Tous" And VCode <> "" Then
        Set Lrs_V = LObj_V.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, cbo_VehiculeFT.Text)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_V = Nothing
        If Not Lrs_V.EOF Then
            VCode = Lrs_V("code")
            Set Lrs_V = Nothing
        Else
            MsgBox "Vehicule invalide!..."
            Set Lrs_V = Nothing
            Exit Sub
        End If
    End If
    If CCode <> "Tous" And CCode <> "" Then
        Set Lrs_C = LObj_C.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, cbo_ConducteurFT.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_C = Nothing
        If Not Lrs_C.EOF Then
            CCode = Lrs_C("code")
            Set Lrs_C = Nothing
        Else
            MsgBox "Conducteur invalide!..."
            Set Lrs_C = Nothing
            Exit Sub
        End If
    End If
    If DCode <> "Tous" And DCode <> "" Then
        Set Lrs_D = LObj_D.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, cbo_DestinationFT.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_D = Nothing
        If Not Lrs_D.EOF Then
            DCode = Lrs_D("numero")
            Set Lrs_D = Nothing
        Else
            MsgBox "Déstination invalide!..."
            Set Lrs_D = Nothing
            Exit Sub
        End If
    End If
    If cda_Debutft.Value > cda_FinFT.Value Then
        MsgBox "Période de recherche invalide,..."
        Exit Sub
    End If
    If (cbo_ConducteurFT.Text = "") Or (cbo_VehiculeFT.Text = "") Or (cbo_DestinationFT.Text = "") Then
        MsgBox "Vérifier les informations de recherche,..."
        Exit Sub
    End If
    Call Affiche_FT(VCode, CCode, DCode, cda_Debutft.Value, cda_FinFT.Value)
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub Affiche_FT(ByVal VEHICULE As String, _
                    ByVal Conducteur As String, _
                    ByVal DESTINATION As String, _
                    ByVal VdateD As String, _
                    ByVal vDateF As String)
                    
    Dim LObj_Find As New Traffic, Lrs_Find As New Recordset
    Dim YearTrafic As Integer, Name_Table As String, itmX As ListItem
    Dim Voyage As Long, NSecond As Long, Distance As Long
    Dim S As Long, m As Long, H As Long, x As Long, y As Long, C As Long, Time As String
    Dim w As Long, V As Long, Q As Long, A As Long, T As String
    Dim MoyDur As Long, MoyDis As Long
        Voyage = 0
        NSecond = 0
        Distance = 0
        NAnomalieDuree = 0
        NAnomalieKm = 0
        NAnomalieTotal = 0
    grid_Ft.ClearRows
    For YearTrafic = Year(VdateD) To Year(vDateF)
        Name_Table = "FicheTraffic"
        If YearTrafic < Year(Date) Then Name_Table = "FicheTraffic_" & YearTrafic
        Set Lrs_Find = LObj_Find.GETALL_SUPERVISIONTRAFFICBYDATE(ErrNumber, ErrDescription, ErrSourceDetail, Name_Table, VdateD, vDateF, Conducteur, VEHICULE, DESTINATION, "Statistique", YearTrafic, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        grid_Ft.Redraw = False
        While Not Lrs_Find.EOF
            With grid_Ft
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Numero"), Lrs_Find("Numero")
                .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Find("MatriculeVehic")
                .CellDetails .Rows, .ColumnIndex("Conducteur"), Lrs_Find("LibelleCond")
                .CellDetails .Rows, .ColumnIndex("Destination"), Lrs_Find("LibelleDest")
                .CellDetails .Rows, .ColumnIndex("DateFT"), Lrs_Find("DateSortie"), , , &H80FF80, &HFF0000
                .CellDetails .Rows, .ColumnIndex("HeureS"), Lrs_Find("HeureSortie")
                .CellDetails .Rows, .ColumnIndex("OS"), Lrs_Find("OperateurSortie")
                If Not IsNull(Lrs_Find("OperateurEntre")) Then .CellDetails .Rows, .ColumnIndex("OE"), Lrs_Find("OperateurEntre")
                If Not IsNull(Lrs_Find("DifK")) Then .CellDetails .Rows, .ColumnIndex("DifK"), Lrs_Find.Fields("DifK"), , , &H80C0FF, &HFF0000
                If Not IsNull(Lrs_Find("MaxDuree")) Then
                    If Lrs_Find.Fields("Duree") >= Lrs_Find.Fields("MaxDuree") Then
                        .CellDetails .Rows, .ColumnIndex("DifD"), Lrs_Find.Fields("DifD"), , , &H80C0FF, &HFF0000
                    Else
                        .CellDetails .Rows, .ColumnIndex("DifD"), "- " & Lrs_Find.Fields("DifDm"), , , &H80C0FF, &HFF0000
                    End If
                End If
                If Not IsNull(Lrs_Find("MaxCompteur")) And Not IsNull(Lrs_Find.Fields("Duree")) Then
                    If (Val(Lrs_Find.Fields("Kmt")) <= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Then .CellDetails .Rows, .ColumnIndex("Dure"), Lrs_Find.Fields("Duree"), , , &H8080FF Else .CellDetails .Rows, .ColumnIndex("Dure"), Lrs_Find.Fields("Duree")
                    If (Val(Lrs_Find.Fields("Kmt")) <= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Then .CellDetails .Rows, .ColumnIndex("MaxD"), Lrs_Find.Fields("MaxDuree"), , , &H80FFFF Else .CellDetails .Rows, .ColumnIndex("MaxD"), Lrs_Find.Fields("MaxDuree")
                    If (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") <= Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Then .CellDetails .Rows, .ColumnIndex("Distance"), Lrs_Find.Fields("Kmt"), , , &H8080FF Else .CellDetails .Rows, .ColumnIndex("Distance"), Lrs_Find.Fields("Kmt")
                    If (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") <= Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Then .CellDetails .Rows, .ColumnIndex("MaxK"), Lrs_Find.Fields("MaxCompteur"), , , &H80FFFF Else .CellDetails .Rows, .ColumnIndex("MaxK"), Lrs_Find.Fields("MaxCompteur")
                End If
                If Not (IsNull(Lrs_Find("HeureENtre"))) Then .CellDetails .Rows, .ColumnIndex("HeureE"), Lrs_Find("HeureENtre")
                .CellDetails .Rows, .ColumnIndex("CPTS"), Lrs_Find("CompteurSortie")
                If Not (IsNull(Lrs_Find("HeureENtre"))) Then .CellDetails .Rows, .ColumnIndex("CPTE"), Lrs_Find("CompteurEntre")
                If Not IsNull(Lrs_Find("MaxCompteur")) And Not IsNull(Lrs_Find.Fields("Duree")) Then
                    If Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") <= Lrs_Find.Fields("MaxDuree") Then
                        NAnomalieKm = NAnomalieKm + 1
                    End If
                    If Val(Lrs_Find.Fields("Kmt")) <= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree") Then
                        NAnomalieDuree = NAnomalieDuree + 1
                    End If
                    If Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree") Then
                        NAnomalieDuree = NAnomalieDuree + 1
                        NAnomalieKm = NAnomalieKm + 1
                    End If
                End If
            End With
            ' Affiche LSv_details
            If Not (IsNull(Lrs_Find("HeureENtre"))) And Not IsNull(Lrs_Find.Fields("Kmt")) Then
                If Val(Lrs_Find.Fields("Kmt")) > 0 Then
                    Voyage = Voyage + 1
                    NSecond = NSecond + Lrs_Find("NSecond")
    '                Dure = Minute(Dure) + Minute(Lrs_Find("duree"))
                    Distance = Distance + Lrs_Find("kmt")
                End If
            End If
            Lrs_Find.MoveNext
        Wend
        grid_Ft.Redraw = True
    End If
    x = NSecond \ 60
    y = x * 60
    S = NSecond - y     'N° second "Z"
    H = x \ 60         'N° Heure
    m = x - (H * 60)   'N° Minute
    Time = CStr(H) & ":" & CStr(m) & ":" & CStr(S)
    'Lsv_toto
    Lsv_DetailsFT.ListItems.Clear
    Set itmX = Lsv_DetailsFT.ListItems.Add(, , CStr(Voyage))
    If Voyage > 0 Then
        MoyDur = NSecond \ Voyage
        w = MoyDur \ 60
        V = w * 60
        Q = MoyDur - V     'N° second "Z"
        C = w \ 60          'N° Heure
        A = w - (C * 60)   'N° Minute
        T = CStr(C) & ":" & CStr(A) & ":" & CStr(Q)
    End If
    itmX.SubItems(1) = CStr(Time)
    If Distance > 0 Then
        itmX.SubItems(2) = CStr(T)
    Else
        itmX.SubItems(2) = "Voyages égale à zéro"
    End If
    itmX.SubItems(3) = CStr(Distance)
    If Voyage > 0 Then
        MoyDis = Distance \ Voyage
    End If
    If Voyage > 0 Then
        itmX.SubItems(4) = CStr(MoyDis)
    Else
        itmX.SubItems(4) = "Voyages égale à zéro"
    End If
    Lrs_Find.Close
    Set Lrs_Find = Nothing
    Next
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'Satistiques Reparation***
Private Sub Cmd_Find_Click()
    Dim VCode As String, LObj_V As New VEHICULE, Lrs_V As New Recordset
On Error GoTo Err
    If Dta_Debut.Value > Dta_Fin.Value Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    List_DetailsRp.ListItems.Clear
    List_detailRp.ListItems.Clear
    VCode = Cbo_Vehicule.Text
     If VCode <> "Tous" And VCode <> "" Then
        Set Lrs_V = LObj_V.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_V = Nothing
        If Not Lrs_V.EOF Then
            VCode = Lrs_V("code")
            Set Lrs_V = Nothing
        Else
            MsgBox "Vehicule introuvable!..."
            Set Lrs_V = Nothing
            Exit Sub
        End If
    End If
    If Cbo_Vehicule.Text = "Tous" Then
        Call AfficheDetailsRp_Tous(Dta_Debut.Value, Dta_Fin.Value)
    Else
        Call AfficheDetailsRp_ParVehicule(Cbo_Vehicule.Text, Dta_Debut.Value, Dta_Fin.Value)
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheDetailsRp_Tous(ByVal VdateD As Date, ByVal vDateF As Date)
    'variables globales
    Dim LOBJ_StatRp As PieceReparation
    Dim rs As New Recordset
    'variables DetailsP
    Dim TotHTBrut As Double
    Dim TotTTC As Double
    Dim Fcode As String
    Dim Qte As Double
    Dim PUHT As Double
    Dim Remise As Double
    Dim tva As Double
    Dim RP As Double
    Dim TotalG As Double
    Dim i
    'variables Details
    Dim nbRep As Double
    Dim Valeur As Double
    Dim MOeuvre As Double
On Error GoTo Err
    Set itmX = List_DetailsRp.ListItems.Add(, , "Tous")
    Set LOBJ_StatRp = New PieceReparation
    'nombre des reparations
    Set rs = LOBJ_StatRp.Get_SumNbrRepStatist(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        nbRep = 0
        nbRep = nbRep + rs("nbrRep")
        itmX.SubItems(1) = CStr(rs("nbrRep"))
        'TTC
        If Not IsNull(rs("Valeur")) Then
            Valeur = 0
            Valeur = Valeur + rs("valeur")
            itmX.SubItems(2) = CStr(Format(rs("valeur"), "#,##0.000"))
        Else
            itmX.SubItems(2) = "Valeur Null"
        End If
    End If
    rs.Close
    'Vehicule + nbr reparations par vehicule
    Set rs = LOBJ_StatRp.Get_NbrRepStatistGrpVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        While Not rs.EOF
            nbRep = 0
            nbRep = nbRep + rs("nbrRep")
            Set itmX = List_DetailsRp.ListItems.Add(, , CStr(rs("vehicule")))
            itmX.SubItems(1) = CStr(rs("nbrRep"))
        rs.MoveNext
        Wend
    End If
    rs.Close
    'Valeur Reparation par vehicule
    For i = 2 To List_DetailsRp.ListItems.Count
        Set rs = LOBJ_StatRp.Get_ValRepStatistVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, List_DetailsRp.ListItems(i), VdateD, vDateF)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        TotalG = 0
        If Not rs.EOF Then
            While Not rs.EOF
                TotTTC = 0
                TotHTBrut = 0
                Qte = 0
                PUHT = 0
                Remise = 0
                tva = 0
                RP = 0
                Qte = rs("Qte")
                PUHT = rs("PUHT")
                Remise = rs("Remise")
                tva = rs("tva")
                RP = rs("remisePiece")
                TotHTBrut = FrmSaisiePieceReparation.Return_TotHT(Qte, PUHT, Remise)
                TotTTC = TotTTC + (TotHTBrut + (TotHTBrut * (tva / 100)))
                TotTTC = TotTTC - (TotTTC * RP / 100)
                TotalG = TotalG + TotTTC
                List_DetailsRp.ListItems(i).SubItems(2) = Format(TotalG, "#,##0.000")
                rs.MoveNext
            Wend
        End If
        rs.Close
    Next
    Valeur = 0
    For i = 2 To List_DetailsRp.ListItems.Count
        Valeur = Valeur + List_DetailsRp.ListItems(i).SubItems(2)
    Next
    Set itmX = List_DetailsRp.ListItems.Item(1)
    itmX.SubItems(2) = CStr(Format(Valeur, "#,##0.000"))
    'detailP
    Set rs = LOBJ_StatRp.Get_DetRepStatist(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
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
            RP = rs("remisePiece")
            TotHTBrut = FrmSaisiePieceReparation.Return_TotHT(Qte, PUHT, Remise)
            TotTTC = TotHTBrut + (TotHTBrut * (tva / 100))
            TotTTC = TotTTC - (TotTTC * RP / 100)
               Set itmX = List_detailRp.ListItems.Add(, , rs("Numero"))
                itmX.SubItems(1) = rs("datePiece")
                itmX.SubItems(2) = rs("Vehicule")
                itmX.SubItems(3) = rs("Designation")
                itmX.SubItems(4) = rs("Qte")
                itmX.SubItems(5) = Format(TotTTC, "#,##0.000")
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set LOBJ_StatRp = Nothing
    List_detailRp.ColumnHeaders(3).Width = 1640
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheDetailsRp_ParVehicule(ByVal Matricule As String, ByVal VdateD As Date, ByVal vDateF As Date)
    'variables globales
    Dim LOBJ_StatRp As PieceReparation
    Dim rs As New Recordset
    Dim i As Integer
    'variables DetailsP
    Dim TotHTBrut As Double
    Dim TotTTC As Double
    Dim Fcode As String
    Dim Qte As Double
    Dim PUHT As Double
    Dim Remise As Double
    Dim tva As Double
    Dim RemiseP As Double
    'variables Details
    Dim nbRep As Double
    Dim Valeur As Double
On Error GoTo Err
    Set itmX = List_DetailsRp.ListItems.Add(, , Cbo_Vehicule.Text)
    Set LOBJ_StatRp = New PieceReparation
    'nombre des reparations
    Set rs = LOBJ_StatRp.Get_NbrRepStatistVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Matricule, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        nbRep = 0
        nbRep = nbRep + rs("nbrRep")
        itmX.SubItems(1) = CStr(rs("nbrRep"))
    End If
    rs.Close
    'detail P
    Set rs = LOBJ_StatRp.Get_PieceRepStatistVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Matricule, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
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
        RemiseP = 0
        Qte = rs("Qte")
        PUHT = rs("PUHT")
        Remise = rs("Remise")
        tva = rs("tva")
        RemiseP = rs("RemisePiece")
        TotHTBrut = FrmSaisiePieceReparation.Return_TotHT(Qte, PUHT, Remise)
        TotHTBrut = TotHTBrut + (TotHTBrut * (tva / 100))
        TotTTC = TotHTBrut - (RemiseP * TotHTBrut / 100)
                Set itmX = List_detailRp.ListItems.Add(, , rs("Numero"))
                itmX.SubItems(1) = rs("datePiece")
                itmX.SubItems(2) = rs("Vehicule")
                itmX.SubItems(3) = rs("Designation")
                itmX.SubItems(4) = rs("Qte")
                itmX.SubItems(5) = Format(TotTTC, "#,##0.000")
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set LOBJ_StatRp = Nothing
     List_detailRp.ColumnHeaders(3).Width = 0
    'Totale réparation
    Valeur = 0
    If List_detailRp.ListItems.Count > 0 Then
        For i = 1 To List_detailRp.ListItems.Count
            Valeur = Valeur + List_detailRp.ListItems(i).SubItems(5)
        Next
    End If
    Set itmX = List_DetailsRp.ListItems.Item(1)
    itmX.SubItems(2) = CStr(Valeur)
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'Satistiques Carburant***
Private Sub Cmd_Search_Click()
    Dim VCode As String, LObj_V As New VEHICULE, Lrs_V As New Recordset
On Error GoTo Err
    If CDate(cda_debut.Text) > CDate(cda_fin.Text) Then
        MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
        Exit Sub
    End If
    Lsv_Details.ListItems.Clear
    Grid_Carb.ClearRows
    VCode = cbo_Matricule.Text
     If VCode <> "Tous" And VCode <> "" Then
        Set Lrs_V = LObj_V.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_V = Nothing
        If Not Lrs_V.EOF Then
            VCode = Lrs_V("code")
            Set Lrs_V = Nothing
        Else
            MsgBox "Vehicule introuvable!..."
            Set Lrs_V = Nothing
            Exit Sub
        End If
    End If
    If cbo_Matricule.Text = "Tous" Then
        Call AfficheDetails_Tous(cda_debut.Text, cda_fin.Text)
    Else
        Call AfficheDetails_ParVehicule(cbo_Matricule.Text, cda_debut.Text, cda_fin.Text)
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheDetails_Tous(ByVal VdateD As Date, ByVal vDateF As Date)
    Dim LOBJ_BC As BonCarburant
    Dim rs As New Recordset
    Dim rD As New Recordset
    Dim i
    Dim TLitre As Double
    Dim Valeur As Double
    Dim MaxC As Long
    Dim MinC As Long
    Dim NbKM As Long
    Dim KmCarburant As Double
On Error GoTo Err
    'detail P
    Set LOBJ_BC = New BonCarburant
    Set rs = LOBJ_BC.Get_StatistDetBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        Grid_Carb.Redraw = False
        While Not rs.EOF
            With Grid_Carb
                .AddRow
                .CellDetails .Rows, 1, rs("Numero")
                .CellDetails .Rows, .ColumnIndex("Vehicule"), rs("Matricule")
                .CellDetails .Rows, .ColumnIndex("Date"), rs("DateDoc")
                .CellDetails .Rows, .ColumnIndex("NbrL"), Format(rs("Litre"), "#,##0.00")
                .CellDetails .Rows, .ColumnIndex("Montant"), Format(rs("Litre") * rs("prixLitre"), "#,##0.000")
                .CellDetails .Rows, .ColumnIndex("Compteur"), rs("Compteur")
                'Get ancien compteur pour chaque voiture et chaque boncarb
                Set rD = LOBJ_BC.Get_AnComptCar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("Numero"), rs("Vehicule"))
                If ErrNumber <> 0 Then
                    ErrNumber = 0
                    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                    Exit Sub
                End If
                If Not rD.EOF Then
                    .CellDetails .Rows, .ColumnIndex("KmParc"), (Val(rs("Compteur")) - Val(rD("maxCpt")))
                    If (Val(rs("Compteur")) - Val(rD("maxCpt"))) <> 0 Then .CellDetails .Rows, .ColumnIndex("Consom"), Format((rs("Litre") * 100) / (Val(rs("Compteur")) - Val(rD("maxCpt"))), "#,##0.000")
                End If
                rD.Close
                If Not IsNull(rs("AnomalieConsom")) Then
                    If CDbl(rs("AnomalieConsom")) >= 2 Then
                        .CellDetails .Rows, .ColumnIndex("Anomalie"), Format(rs("AnomalieConsom"), "#,##0.00"), , , &H8080FF
                    Else
                        .CellDetails .Rows, .ColumnIndex("Anomalie"), Format(rs("AnomalieConsom"), "#,##0.00")
                    End If
                End If
            End With
        rs.MoveNext
        Wend
    Grid_Carb.ColumnWidth("Vehicule") = 120
    Grid_Carb.Redraw = True
    Else
        MsgBox "Pas de donnèes à visualiser !", vbInformation
        Exit Sub
    End If
    rs.Close
    'totaux Lsv Details
    'Parcourir Grid_carb
    Dim Valc As Double
    Dim NBL As Double
    Dim itm
    Dim Veh As String
    Dim ii
    Dim KmParcVeh As Double
    Dim TotLitVeh As Double
    Dim TotKm As Long
    Valc = 0
    NBL = 0
    TotKm = 0
    For i = 1 To Grid_Carb.Rows
        Valc = Valc + Grid_Carb.CellText(i, 5)
        NBL = NBL + Grid_Carb.CellText(i, 4)
        TotKm = TotKm + Grid_Carb.CellText(i, 7)
    Next
    Set itm = Lsv_Details.ListItems.Add(, , "Tous")
        itm.SubItems(1) = CStr(Format(Valc, "#,##0.000"))
        itm.SubItems(3) = CStr(Format(NBL, "#,##0.00"))
        itm.SubItems(4) = CStr(TotKm)
        If TotKm <> 0 Then itm.SubItems(5) = CStr(Format(NBL * 100 / TotKm, "#,##0.00"))
        
    'Details Lsv_details
    Set rs = LOBJ_BC.Get_StatistBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        While Not rs.EOF
        'Lire et prix Litre
            TLitre = 0
            Valeur = 0
            KmParcVeh = 0
            TotLitVeh = 0
            TLitre = TLitre + rs("Litre")
            Valeur = Valeur + rs("Litre") * rs("Prix")
            For ii = 1 To Grid_Carb.Rows
                If Grid_Carb.CellText(ii, 2) = rs("Vehicule") Then
                    KmParcVeh = KmParcVeh + Grid_Carb.CellText(ii, 7)
                    TotLitVeh = TotLitVeh + Grid_Carb.CellText(ii, 4)
                End If
            Next
            'Consommation par 100 KM
            If Not (KmParcVeh = 0) Then
                KmCarburant = ((TotLitVeh * 100) / KmParcVeh)
            Else
                KmCarburant = 0
            End If
            Set itmX = Lsv_Details.ListItems.Add(, , CStr(rs("Vehicule")))
                itmX.SubItems(1) = CStr(Format(Valeur, "#,##0.000"))
                itmX.SubItems(2) = CStr(Format(rs("Prix"), "#,##0.000"))
                itmX.SubItems(3) = CStr(Format(TLitre, "#,##0.00"))
                itmX.SubItems(4) = CStr(KmParcVeh)
                itmX.SubItems(5) = CStr(Format(KmCarburant, "#,##0.000"))
            
            rs.MoveNext
        Wend
    End If
    rs.Close
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheDetails_ParVehicule(ByVal Matricule As String, ByVal VdateD As Date, ByVal vDateF As Date)
    Dim LOBJ_BC As BonCarburant
    Dim rs As New Recordset
    Dim rD As New Recordset
    Dim i As Integer
    Dim itm
    Dim TLitre As Double
    Dim Valeur As Double
    Dim MaxC As Long
    Dim MinC As Long
    Dim NbKM As Long
    ''selection de code de Vehicule
    Dim CodV As String
On Error GoTo Err
    CodV = Return_CodVehicule(Matricule)
    'Remplissage de Lsv_detailP
    Set LOBJ_BC = New BonCarburant
    Set rs = LOBJ_BC.Get_StatistBCVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF, CodV)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
    Grid_Carb.Redraw = False
        While Not rs.EOF
            With Grid_Carb
                .AddRow
                .CellDetails .Rows, 1, rs("Numero")
                .CellDetails .Rows, .ColumnIndex("Vehicule"), rs("Matricule")
                .CellDetails .Rows, .ColumnIndex("Date"), rs("DateDoc")
                .CellDetails .Rows, .ColumnIndex("NbrL"), Format(rs("Litre"), "#,##0.00")
                .CellDetails .Rows, .ColumnIndex("Montant"), Format(rs("Litre") * rs("prixLitre"), "#,##0.000")
                .CellDetails .Rows, .ColumnIndex("Compteur"), rs("Compteur")
                'Get ancien compteur pour chaque voiture et chaque boncarb
                Set rD = LOBJ_BC.Get_AnComptCar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("Numero"), CodV)
                If ErrNumber <> 0 Then
                    ErrNumber = 0
                    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                    Exit Sub
                End If
                If Not rD.EOF Then
                    .CellDetails .Rows, .ColumnIndex("KmParc"), (Val(rs("Compteur")) - Val(rD("maxCpt")))
                    If (Val(rs("Compteur")) - Val(rD("maxCpt"))) <> 0 Then .CellDetails .Rows, .ColumnIndex("Consom"), Format((rs("Litre") * 100) / (Val(rs("Compteur")) - Val(rD("maxCpt"))), "#,##0.000")
                End If
                rD.Close
                If Not IsNull(rs("AnomalieConsom")) Then
                    If CDbl(rs("AnomalieConsom")) >= 2 Then
                        .CellDetails .Rows, .ColumnIndex("Anomalie"), Format(rs("AnomalieConsom"), "#,##0.00"), , , &H8080FF
                    Else
                        .CellDetails .Rows, .ColumnIndex("Anomalie"), Format(rs("AnomalieConsom"), "#,##0.00")
                    End If
                End If
            End With
            rs.MoveNext
        Wend
        Grid_Carb.ColumnWidth("Vehicule") = 0
        Grid_Carb.Redraw = True
    Else
        MsgBox "Pas de donnèes à visualiser !", vbInformation
        Exit Sub
    End If
    rs.Close
    'Parcourir Grid_Carb
    Dim Valc As Double
    Dim NBL As Double
    Dim KmParcVeh As Long
    Dim KmCarburant As Double
    Valc = 0
    NBL = 0
    KmParcVeh = 0
    KmCarburant = 0
    For i = 1 To Grid_Carb.Rows
        Valc = Valc + Grid_Carb.CellText(i, 5)
        NBL = NBL + Grid_Carb.CellText(i, 4)
        KmParcVeh = KmParcVeh + Grid_Carb.CellText(i, 7)
    Next
    Set itm = Lsv_Details.ListItems.Add(, , CStr(Matricule))
        itm.SubItems(1) = CStr(Format(Valc, "#,##0.000"))
        itm.SubItems(3) = CStr(Format(NBL, "#,##0.00"))
        itm.SubItems(4) = KmParcVeh
        'consommation par 100 km
    If Not (itm.SubItems(4) = 0) Then
        KmCarburant = ((itm.SubItems(3) * 100) / itm.SubItems(4))
        itm.SubItems(5) = CStr(Format(KmCarburant, "#,##0.000"))
    Else
        itm.SubItems(5) = "zéro km!!"
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
    'Print***
Private Sub CmdPrint_Click()
    Dim VCode As String
    Dim DATEDEBUT As Date
    Dim DateFin As Date
    Dim TotLitre As Double
    Dim nbrRep As Long
    Dim Total As Double
    Dim J
On Error GoTo Err
'Imprimer statistique carburant
If Tab_Satistiques.Tab = 0 Then
    If Grid_Carb.Rows = 0 Then
        MsgBox "Pas de données à imprimer .", vbInformation
        Exit Sub
    End If
    
    DATEDEBUT = cda_debut.Text
    DateFin = cda_fin.Text
    If DATEDEBUT > DateFin Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    VCode = cbo_Matricule.Text
    TotLitre = 0
    Total = 0
    If MsgBox("Imprimer statistiques carburant   ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
        If VCode = "Tous" Then
            For J = 1 To Lsv_Details.ListItems.Count
                If (Lsv_Details.ListItems(J)) = "Tous" Then
                    TotLitre = Lsv_Details.ListItems(J).ListSubItems(3)
                    Total = Lsv_Details.ListItems(J).ListSubItems(1)
                End If
            Next
        Else
            TotLitre = Lsv_Details.ListItems(1).ListSubItems(3)
            Total = Lsv_Details.ListItems(1).ListSubItems(1)
        End If
        Call Frm_Rpt_Apercus.PrintOutAndApercu_StatCarb(0, DATEDEBUT, DateFin, VCode, LStr_NameUser, TotLitre, Total)
        Frm_Rpt_Apercus.Show
    End If
 'Imprimer statistique réparation
ElseIf Tab_Satistiques.Tab = 1 Then
    If List_detailRp.ListItems.Count = 0 Then
        MsgBox "Pas de données à imprimer .", vbInformation
        Exit Sub
    End If
    DATEDEBUT = Dta_Debut.Value
    DateFin = Dta_Fin.Value
     If DATEDEBUT > DateFin Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    VCode = Cbo_Vehicule.Text
    nbrRep = 0
    Total = 0
    If MsgBox("Imprimer statistiques réparation   ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
        If VCode = "Tous" Then
            For J = 1 To List_DetailsRp.ListItems.Count
                If (List_DetailsRp.ListItems(J)) = "Tous" Then
                    nbrRep = List_DetailsRp.ListItems(J).ListSubItems(1)
                    Total = List_DetailsRp.ListItems(1).ListSubItems(2)
                End If
            Next
        Else
            nbrRep = List_DetailsRp.ListItems(1).ListSubItems(1)
            Total = List_DetailsRp.ListItems(1).ListSubItems(2)
        End If
        Call Frm_Rpt_Apercus.PrintOutAndApercu_StatRep(0, DATEDEBUT, DateFin, VCode, LStr_NameUser, nbrRep, Total)
        Frm_Rpt_Apercus.Show
    End If
 'Imprimer statistique traffic
ElseIf Tab_Satistiques.Tab = 2 Then
    If grid_Ft.Rows = 0 Then
        MsgBox "Pas de données à imprimer.", vbInformation
        Exit Sub
    End If
    DATEDEBUT = cda_Debutft.Value
    DateFin = cda_FinFT.Value
     If DATEDEBUT > DateFin Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    If MsgBox("Imprimer statistique traffic ?", vbYesNo + vbDefaultButton1 + vbInformation, App.ProductName) = vbYes Then
        Call Frm_Rpt_Apercus.PrintOutAndApercu_AnomalieTrafic(0, DATEDEBUT, DateFin, VCodeDrive, VCodeVehicle, VCodeDestination, LStr_NameUser, CStr(NAnomalieKm), CStr(NAnomalieDuree), CStr(NAnomalieTotal), False)
        Frm_Rpt_Apercus.Show
    End If
'Imprimer statistique Services
ElseIf Tab_Satistiques.Tab = 3 Then
    If grid_Service.Rows = 0 Then
        MsgBox "Pas de données à imprimer.", vbInformation
        Exit Sub
    End If
    DATEDEBUT = cda_DebutService.Value
    DateFin = cda_FinService.Value
     If DATEDEBUT > DateFin Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    If MsgBox("Imprimer statistique service ?", vbYesNo + vbDefaultButton1 + vbInformation, App.ProductName) = vbYes Then
        Call Frm_Rpt_Apercus.PrintOutAndApercu_StatService(0, DATEDEBUT, DateFin, cbo_Conducteur.Text, LStr_NameUser)
        Frm_Rpt_Apercus.Show
    End If
End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'=================================================
Public Sub AfficheRow_Conducteur(ByVal VCode As String)
    Dim LOBJ_Cond As New Conducteur
    Dim Lrs_Cond As New Recordset
    Dim cboPers As ComboBox
On Error GoTo Err
    Set Lrs_Cond = LOBJ_Cond.GetRow_Conducteur_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LOBJ_Cond = Nothing
    If Not Lrs_Cond.EOF Then
        If Frm_FindView.StrSource = "ConducteurStque" Then Set cboPers = cbo_ConducteurFT
        If Frm_FindView.StrSource = "ConducteurE/H" Then Set cboPers = cbo_Conducteur
        If Not IsNull(Lrs_Cond("Libelle")) Then
            cboPers.Text = Lrs_Cond("Libelle")
            VCodeDrive = Lrs_Cond("code")
        End If
    End If
    Lrs_Cond.Close
    Set Lrs_Cond = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Public Sub AfficheRow_Destination(ByVal VCode As String)
    Dim LObj_Find As New DESTINATION
    Dim Lrs_Find As New Recordset
    Dim cboDest As ComboBox
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        'Charge
        If Frm_FindView.StrSource = "DestinationE/H" Then Set cboDest = cbo_DestinationFT
        If Not IsNull(Lrs_Find("Libelle")) Then
            cboDest.Text = Lrs_Find("Libelle")
            VCodeDestination = Lrs_Find("Numero")
        End If
    End If
    Lrs_Find.Close
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub cbo_ConducteurFT_Click()
    Dim LOBJ_Cond As New Conducteur
    Dim Lrs_Cond As Recordset
    If cbo_ConducteurFT.ListIndex = 0 Then
        VCodeDrive = "  -  Tous"
    Else
        Set Lrs_Cond = LOBJ_Cond.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, cbo_ConducteurFT.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LOBJ_Cond = Nothing
        
        If Not Lrs_Cond.EOF Then VCodeDrive = Lrs_Cond("Code")
    End If
End Sub
Private Sub cbo_VehiculeFT_Click()
    Dim Lobj_Vehicule As New VEHICULE
    Dim Lrs_Vehicule As Recordset
    If cbo_VehiculeFT.ListIndex = 0 Then
        VCodeVehicle = "  -  Tous"
    Else
         Set Lrs_Vehicule = Lobj_Vehicule.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, cbo_VehiculeFT.Text)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set Lobj_Vehicule = Nothing
        
        If Not Lrs_Vehicule.EOF Then VCodeVehicle = Lrs_Vehicule("Code")
        Set Lrs_Vehicule = Nothing
    End If
End Sub
Private Sub cbo_DestinationFT_Click()
    Dim LOBJ_Dest As New DESTINATION
    Dim Lrs_Dest As Recordset
    If cbo_DestinationFT.ListIndex = 0 Then
        VCodeDestination = "  -  Tous"
    Else
        Set Lrs_Dest = LOBJ_Dest.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, cbo_DestinationFT.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LOBJ_Dest = Nothing
        
        If Not Lrs_Dest.EOF Then VCodeDestination = Lrs_Dest("Numero")
    End If
End Sub
Private Sub Cbo_Conducteur_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_Conducteur.Text)
        For i = 0 To cbo_Conducteur.ListCount - 1
            If Left(cbo_Conducteur.List(i), start) = cbo_Conducteur.Text Then
                cbo_Conducteur.Text = cbo_Conducteur.List(i)
            End If
        Next
        cbo_Conducteur.SelStart = start
        cbo_Conducteur.SelLength = Len(cbo_Conducteur.Text)
    End If
End Sub
Private Sub Cbo_Conducteur_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub cbo_vehiculeft_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_VehiculeFT.Text)
        For i = 0 To cbo_VehiculeFT.ListCount - 1
            If Left(cbo_VehiculeFT.List(i), start) = cbo_VehiculeFT.Text Then
                cbo_VehiculeFT.Text = cbo_VehiculeFT.List(i)
            End If
        Next
        cbo_VehiculeFT.SelStart = start
        cbo_VehiculeFT.SelLength = Len(cbo_VehiculeFT.Text)
    End If
End Sub
Private Sub cbo_Vehiculeft_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub cbo_ConducteurFT_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_ConducteurFT.Text)
        For i = 0 To cbo_ConducteurFT.ListCount - 1
            If Left(cbo_ConducteurFT.List(i), start) = cbo_ConducteurFT.Text Then
                cbo_ConducteurFT.Text = cbo_ConducteurFT.List(i)
            End If
        Next
        cbo_ConducteurFT.SelStart = start
        cbo_ConducteurFT.SelLength = Len(cbo_ConducteurFT.Text)
    End If
End Sub
Private Sub Cbo_Conducteurft_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub cbo_destinationft_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_DestinationFT.Text)
        For i = 0 To cbo_DestinationFT.ListCount - 1
            If Left(cbo_DestinationFT.List(i), start) = cbo_DestinationFT.Text Then
                cbo_DestinationFT.Text = cbo_DestinationFT.List(i)
            End If
        Next
        cbo_DestinationFT.SelStart = start
        cbo_DestinationFT.SelLength = Len(cbo_DestinationFT.Text)
    End If
End Sub
Private Sub cbo_destinationft_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub grid_ft_ColumnClick(ByVal lCol As Long)
    Dim sTag As String
    Dim i As Long
    With grid_Ft.SortObject
        .Clear
        .SortColumn(1) = lCol
        sTag = grid_Ft.ColumnTag(lCol)
        If (sTag = "") Then
            sTag = "DESC"
            .SortOrder(1) = CCLOrderAscending
        Else
            sTag = ""
            .SortOrder(1) = CCLOrderDescending
        End If
        grid_Ft.ColumnTag(lCol) = sTag
        Select Case grid_Ft.ColumnKey(lCol)
            Case "Matricule"
                 .SortType(1) = CCLSortString
            Case "Conducteur"
                 .SortType(1) = CCLSortString
            Case "Destination"
                 .SortType(1) = CCLSortString
            Case "DateFT"
                 .SortType(1) = CCLSortDate
            Case "HeureS"
                 .SortType(1) = CCLSortDateHourAccuracy
            Case "HeureE"
                 .SortType(1) = CCLSortDateHourAccuracy
            Case "CPTS"
                 .SortType(1) = CCLSortNumeric
            Case "CPTE"
                 .SortType(1) = CCLSortNumeric
            Case "Distance"
                 .SortType(1) = CCLSortNumeric
            Case "Dure"
                 .SortType(1) = CCLSortDateHourAccuracy
        End Select
    End With
    Screen.MousePointer = vbHourglass
    grid_Ft.Sort
    Screen.MousePointer = vbDefault
End Sub
Private Sub grid_ft_DblClick(ByVal lRow As Long, ByVal lCol As Long)
On Error GoTo Err
    With Frm_MajStatFT
        .selectFT (grid_Ft.CellText(grid_Ft.SelectedRow, grid_Ft.ColumnIndex("Numero")))
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub Cbo_Vehicule_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(Cbo_Vehicule.Text)
        For i = 0 To Cbo_Vehicule.ListCount - 1
            If Left(Cbo_Vehicule.List(i), start) = Cbo_Vehicule.Text Then
                Cbo_Vehicule.Text = Cbo_Vehicule.List(i)
            End If
        Next
        Cbo_Vehicule.SelStart = start
        Cbo_Vehicule.SelLength = Len(Cbo_Vehicule.Text)
    End If
End Sub
Public Sub LV_ColumnSort(ListViewControl As ListView, _
  Column As ColumnHeader)
 With ListViewControl
  If .SortKey <> Column.Index - 1 Then
   .SortKey = Column.Index - 1
   .SortOrder = lvwAscending
  Else
   If .SortOrder = lvwAscending Then
    .SortOrder = lvwDescending
   Else
    .SortOrder = lvwAscending
   End If
  End If
  .Sorted = -1
 End With
End Sub
Public Sub ListView_Header_Click(LView As ListView, ColumnHeader As MSComctlLib.ColumnHeader)
            LView.Sorted = True
            LView.SortOrder = Abs(Not -LView.SortOrder)
            LView.SortKey = ColumnHeader.Index - 1
            LView.Sorted = False
    End Sub

Private Sub Label22_Click()

End Sub



Private Sub Label15_Click()
Pic_Details_Dest.Visible = False
End Sub

Private Sub List_detailRp_Click()


'Call LV_ColumnSort(List_detailRp, List_detailRp.ColumnHeaders(1))

End Sub

Private Sub List_detailRp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ListView_Header_Click(List_detailRp, ColumnHeader)
    
End Sub

Private Sub List_detailRp_DblClick()
    Dim VCode
    Dim i As Integer
On Error GoTo Err
    i = List_detailRp.SelectedItem.Index
    VCode = List_detailRp.ListItems.Item(i)
    'ViderZone (frm)
    With FrmConsultPieceReception       'FrmPieceReparation
        .AfficheRow (VCode)
        .Show
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub Cbo_Vehicule_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub Cbo_Vehicule_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub cbo_Matricule_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_Matricule.Text)
        For i = 0 To cbo_Matricule.ListCount - 1
            If Left(cbo_Matricule.List(i), start) = cbo_Matricule.Text Then
                cbo_Matricule.Text = cbo_Matricule.List(i)
            End If
        Next
        cbo_Matricule.SelStart = start
        cbo_Matricule.SelLength = Len(cbo_Matricule.Text)
    End If
End Sub
Private Sub Cbo_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cbo_Matricule_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub Grid_Carb_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    Dim VCode
On Error GoTo Err
    VCode = Grid_Carb.CellText(lRow, 1)
    With frmConsultBC
        .AfficheRow (VCode)
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Public Function Return_CodVehicule(ByVal Matricule As String) As String
    Dim Lobj_Vehicule As New VEHICULE
    Dim Lrs_Vehicule As Recordset
On Error GoTo Err
    Set Lrs_Vehicule = Lobj_Vehicule.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Matricule)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Function
    End If
    Set Lobj_Vehicule = Nothing
    If Not Lrs_Vehicule.EOF Then
        'Charge
        If Not IsNull(Lrs_Vehicule("Code")) Then
            Return_CodVehicule = CStr(Lrs_Vehicule("Code"))
        End If
    End If
    Lrs_Vehicule.Close
    Set Lrs_Vehicule = Nothing
Exit Function
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Function


Public Sub ListView_Header(LView As ListView, ColumnHeader As MSComctlLib.ColumnHeader)
            LView.Sorted = True
            LView.SortOrder = Abs(Not -LView.SortOrder)
            LView.SortKey = ColumnHeader.Index - 1
            LView.Sorted = False
    End Sub

Private Sub List_Details1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Pic_Details_Dest.Visible = False
 Call ListView_Header(List_Details1, ColumnHeader)
End Sub

Private Sub List_Details1_DblClick()
Dim VCode
Dim DateD As Date
Dim DateF As Date
Dim i As Integer
On Error GoTo Err
    i = List_Details1.SelectedItem.Index
    VCode = List_Details1.ListItems.Item(i)
    DateD = Dta_Debut_Dest.Value
    DateF = Dta_Fin_Dest.Value
    Call AfficheDetailsDes(VCode, DateD, DateF)
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Public Sub AfficheDetailsDes(ByVal VCode As String, ByVal DateD As Date, ByVal DateF As Date)
    Dim LOBJ_Dest   As New DESTINATION
    Dim Lrs_Dest    As New Recordset
On Error GoTo Err
    
    SGrid_Details.ClearRows
    
    Set Lrs_Dest = LOBJ_Dest.GetRow_Details_Dest(ErrNumber, ErrDescription, ErrSourceDetail, VCode, DateD, DateF, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    
    Set LOBJ_Dest = Nothing

If Not Lrs_Dest.EOF Then
Label16.Caption = Lrs_Dest("Destination")
        While Not Lrs_Dest.EOF
        SGrid_Details.Redraw = False
            With SGrid_Details
            .AddRow
            .CellDetails .Rows, .ColumnIndex("Vehicule"), Lrs_Dest.Fields("Vehicule"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("Conducteur"), Lrs_Dest.Fields("Conducteur"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("CompteurSortie"), Lrs_Dest.Fields("CompteurSortie"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("CompteurEntre"), Lrs_Dest.Fields("CompteurEntre"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("HeureSortie"), Lrs_Dest.Fields("HeureSortie"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("HeureEntre"), Lrs_Dest.Fields("HeureEntre"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("OperateurSortie"), Lrs_Dest.Fields("OperateurSortie"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("OperateurEntre"), Lrs_Dest.Fields("OperateurEntre"), DT_LEFT
                    
            End With
        Lrs_Dest.MoveNext
        Wend
        
        SGrid_Details.Redraw = True
        Pic_Details_Dest.Visible = True
        cmd_FindDest.SetFocus
        
        If SGrid_Details.Rows > 0 Then
            SGrid_Details.SelectedRow = 1
            SGrid_Details.SetFocus
        End If

    
    Else
        MsgBox "Veuillez Choisir une Destination!", vbInformation
     
    End If
On Error Resume Next
    Set Lrs_Dest = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub



Private Sub SGrid_Details_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
If KeyCode = vbKeyEscape Then Pic_Details_Dest.Visible = False
End Sub

Private Sub Stat_Dest_Click()
Frm_Stat_Dest.Show
End Sub

Private Sub SGrid_Details_ColumnClick(ByVal lCol As Long)

Dim sTag As String
    Dim i As Long
    With SGrid_Details.SortObject
        .Clear
        .SortColumn(1) = lCol
        sTag = SGrid_Details.ColumnTag(lCol)
        If (sTag = "") Then
            sTag = "DESC"
            .SortOrder(1) = CCLOrderAscending
        Else
            sTag = ""
            .SortOrder(1) = CCLOrderDescending
        End If
        SGrid_Details.ColumnTag(lCol) = sTag
        Select Case SGrid_Details.ColumnKey(lCol)
            Case "Numero"
                 .SortType(1) = CCLSortNumeric
            Case "Vehicule"
                 .SortType(1) = CCLSortString
            Case "Conducteur"
                 .SortType(1) = CCLSortString
            Case "CompteurSortie"
                 .SortType(1) = CCLSortNumeric
            Case "CompteurEntre"
                 .SortType(1) = CCLSortNumeric
            Case "HeureSortie"
                 .SortType(1) = CCLSortDate
            Case "HeureEntre"
                 .SortType(1) = CCLSortDate
            Case "OperateurSortie"
                 .SortType(1) = CCLSortString
            Case "OperateurEntre"
                 .SortType(1) = CCLSortString
        End Select
    End With
    Screen.MousePointer = vbHourglass
    SGrid_Details.Sort
    Screen.MousePointer = vbDefault
End Sub
Private Sub Initgrid_Details_Dest()
With SGrid_Details
.Redraw = False

    

    .HideGroupingBox = True
    .AllowGrouping = True
 
    .GroupRowBackColor = vbWindowBackground
    .GroupRowForeColor = vbWindowText
    
    .GridLineColor = &HE0E0E0
    .GridFillLineColor = vbBlue
    .GridLines = True
    
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .DrawFocusRectangle = False
    
 
    .AddColumn "Numero", "Numero", , , , False
    .AddColumn "Vehicule", "Vehicule", , , 100, , , , , , , CCLSortString
    .AddColumn "Conducteur", "Conducteur", , , 100
    .AddColumn "CompteurSortie", "CPT Sortie", , , 60
    .AddColumn "CompteurEntre", "CPT Entrée", , , 60
    .AddColumn "HeureSortie", "HeureSortie", , , 100
    .AddColumn "HeureEntre", "HeureEntre", , , 100
    .AddColumn "OperateurSortie", "OperateurSortie", , , 100
    .AddColumn "OperateurEntre", "OperateurEntre", , , 100
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
.Redraw = True
End With

End Sub
