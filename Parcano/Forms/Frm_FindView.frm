VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form Frm_FindView 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FindView"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Tab_FindView2 
      Height          =   2535
      Left            =   360
      TabIndex        =   28
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Consultation"
      TabPicture(0)   =   "Frm_FindView.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Pic_ConsltConge"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox Pic_ConsltConge 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   7815
         TabIndex        =   38
         Top             =   360
         Width           =   7815
         Begin MSComCtl2.DTPicker Cda_FinConge 
            Height          =   375
            Left            =   6000
            TabIndex        =   39
            Top             =   120
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
            Format          =   147980289
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker Cda_DebutConge 
            Height          =   375
            Left            =   3600
            TabIndex        =   40
            Top             =   120
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
            Format          =   147980289
            CurrentDate     =   42860
         End
         Begin SToolBox.SBiCombo Cbo_CondConge 
            Height          =   405
            Left            =   240
            TabIndex        =   41
            Top             =   840
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   714
            FontBold        =   -1  'True
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Image Pic_PrintConge 
            Height          =   495
            Left            =   5760
            Picture         =   "Frm_FindView.frx":001C
            Stretch         =   -1  'True
            Top             =   1440
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   5520
            TabIndex        =   44
            Top             =   0
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   3105
            TabIndex        =   43
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Conducteur ..."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   2175
         End
         Begin VB.Image Pic_FindConge 
            Height          =   495
            Left            =   5760
            Picture         =   "Frm_FindView.frx":10C1E
            Stretch         =   -1  'True
            Top             =   840
            Width           =   1935
         End
      End
   End
   Begin VB.TextBox txt_Libelle 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   0
      TabIndex        =   27
      Text            =   "   Rechercher..."
      Top             =   7800
      Width           =   9015
   End
   Begin TabDlg.SSTab Tab_FindView1 
      Height          =   2535
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "A Faire..."
      TabPicture(0)   =   "Frm_FindView.frx":21820
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tous..."
      TabPicture(1)   =   "Frm_FindView.frx":2183C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   -74880
         ScaleHeight     =   2055
         ScaleWidth      =   7815
         TabIndex        =   16
         Top             =   360
         Width           =   7815
         Begin SToolBox.SCheckBox ChBox_Supprimer 
            Height          =   375
            Left            =   3720
            TabIndex        =   25
            Top             =   840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            ForeColor       =   255
         End
         Begin MSComCtl2.DTPicker Cda_FinPgHT 
            Height          =   375
            Left            =   6000
            TabIndex        =   17
            Top             =   120
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
            Format          =   92864513
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker Cda_DebutPgHT 
            Height          =   375
            Left            =   3600
            TabIndex        =   18
            Top             =   120
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
            Format          =   92864513
            CurrentDate     =   42860
         End
         Begin SToolBox.SBiCombo Cbo_CondPgHT 
            Height          =   405
            Left            =   240
            TabIndex        =   19
            Top             =   1440
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   714
            FontBold        =   -1  'True
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Supprime..."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   4080
            TabIndex        =   26
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "|> Rechercher par date et Conducteur..."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   6015
         End
         Begin VB.Image Pic_FindPgHT 
            Height          =   495
            Left            =   5760
            Picture         =   "Frm_FindView.frx":21858
            Stretch         =   -1  'True
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Conducteur ..."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   3105
            TabIndex        =   21
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   5520
            TabIndex        =   20
            Top             =   0
            Width           =   480
         End
         Begin VB.Image Image3 
            Height          =   495
            Left            =   5760
            Picture         =   "Frm_FindView.frx":3245A
            Stretch         =   -1  'True
            Top             =   1440
            Visible         =   0   'False
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   7815
         TabIndex        =   9
         Top             =   360
         Width           =   7815
         Begin MSComCtl2.DTPicker Cda_FinPgHF 
            Height          =   375
            Left            =   6000
            TabIndex        =   10
            Top             =   0
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
            Format          =   92536833
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker Cda_DebutPgHF 
            Height          =   375
            Left            =   3600
            TabIndex        =   11
            Top             =   0
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   92536833
            CurrentDate     =   42860
         End
         Begin SToolBox.SBiCombo Cbo_CondPgHF 
            Height          =   405
            Left            =   240
            TabIndex        =   12
            Top             =   1440
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   714
            FontBold        =   -1  'True
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "|> Tous les Programmes à faire et non Supprimée..."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   6015
         End
         Begin VB.Image Pic_FindPgHF 
            Height          =   495
            Left            =   5760
            Picture         =   "Frm_FindView.frx":4305C
            Stretch         =   -1  'True
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Conducteur ..."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   3105
            TabIndex        =   14
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   5520
            TabIndex        =   13
            Top             =   0
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   495
            Left            =   5760
            Picture         =   "Frm_FindView.frx":53C5E
            Stretch         =   -1  'True
            Top             =   1440
            Visible         =   0   'False
            Width           =   1935
         End
      End
   End
   Begin TabDlg.SSTab Tab_FindView 
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Consultation"
      TabPicture(0)   =   "Frm_FindView.frx":64860
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Pic_ConsltSPLNG"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tournée de Garde"
      TabPicture(1)   =   "Frm_FindView.frx":6487C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Pic_ConsltSPLNG_TG"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Pic_ConsltSPLNG_TG 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   -74880
         ScaleHeight     =   2055
         ScaleWidth      =   7815
         TabIndex        =   45
         Top             =   360
         Width           =   7815
         Begin MSComCtl2.DTPicker Cda_FinSPLNG_TG 
            Height          =   375
            Left            =   6000
            TabIndex        =   46
            Top             =   120
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
            Format          =   149946369
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker Cda_DebutSPLNG_TG 
            Height          =   375
            Left            =   3600
            TabIndex        =   47
            Top             =   120
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
            Format          =   149946369
            CurrentDate     =   42860
         End
         Begin SToolBox.SBiCombo Cbo_CondSPLNG_TG 
            Height          =   405
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   714
            FontBold        =   -1  'True
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SToolBox.SBiCombo Cbo_DestSPLNG_TG 
            Height          =   405
            Left            =   120
            TabIndex        =   52
            Top             =   1560
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   714
            FontBold        =   -1  'True
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label22 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Destination ..."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Image Image2 
            Height          =   495
            Left            =   5760
            Picture         =   "Frm_FindView.frx":64898
            Stretch         =   -1  'True
            Top             =   1440
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   5520
            TabIndex        =   51
            Top             =   0
            Width           =   480
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   3105
            TabIndex        =   50
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label21 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Conducteur ..."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   2175
         End
         Begin VB.Image Pic_FindSPLNG_TG 
            Height          =   495
            Left            =   5760
            Picture         =   "Frm_FindView.frx":7549A
            Stretch         =   -1  'True
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.PictureBox Pic_ConsltSPLNG 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   7815
         TabIndex        =   29
         Top             =   360
         Width           =   7815
         Begin MSComCtl2.DTPicker Cda_FinSPLNG 
            Height          =   375
            Left            =   6000
            TabIndex        =   30
            Top             =   120
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
            Format          =   149946369
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker Cda_DebutSPLNG 
            Height          =   375
            Left            =   3600
            TabIndex        =   31
            Top             =   120
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
            Format          =   149946369
            CurrentDate     =   42860
         End
         Begin SToolBox.SBiCombo Cbo_CondSPLNG 
            Height          =   405
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   714
            FontBold        =   -1  'True
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SToolBox.SBiCombo Cbo_DestSPLNG 
            Height          =   405
            Left            =   120
            TabIndex        =   33
            Top             =   1560
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   714
            FontBold        =   -1  'True
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Destination ..."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Image Pic_PrintSPLNG 
            Height          =   495
            Left            =   5760
            Picture         =   "Frm_FindView.frx":8609C
            Stretch         =   -1  'True
            Top             =   1440
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   5520
            TabIndex        =   36
            Top             =   0
            Width           =   480
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   3105
            TabIndex        =   35
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Conducteur ..."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   2175
         End
         Begin VB.Image Pic_FindSPLNG 
            Height          =   495
            Left            =   5760
            Picture         =   "Frm_FindView.frx":96C9E
            Stretch         =   -1  'True
            Top             =   840
            Width           =   1935
         End
      End
   End
   Begin VB.PictureBox Pic_Date 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1800
      ScaleHeight     =   375
      ScaleWidth      =   6255
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label Lbl_DateAuStiquePNG 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Top             =   0
         Width           =   2610
      End
      Begin VB.Label Lbl_DateDuStiquePNG 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   600
         TabIndex        =   6
         Top             =   0
         Width           =   1770
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Au :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2415
         TabIndex        =   5
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Du :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   495
      End
   End
   Begin SToolBox.SGrid Grid_FindView 
      Height          =   6495
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   8950
      _ExtentX        =   15796
      _ExtentY        =   11456
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin VB.Image Pic_PrintCompteur 
      Height          =   495
      Left            =   6960
      Picture         =   "Frm_FindView.frx":A78A0
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label LBL_Titre 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   8415
   End
   Begin VB.Image Pic_MaskMenu 
      Height          =   375
      Left            =   8400
      Picture         =   "Frm_FindView.frx":B84A2
      Stretch         =   -1  'True
      Top             =   840
      Width           =   495
   End
   Begin VB.Image Pic_ShowMenu 
      Height          =   375
      Left            =   8400
      Picture         =   "Frm_FindView.frx":B91C8
      Stretch         =   -1  'True
      Top             =   840
      Width           =   495
   End
   Begin VB.Image Pic_Header 
      Height          =   1095
      Left            =   0
      Picture         =   "Frm_FindView.frx":B9B1E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "Frm_FindView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public StrSource As String
    Dim CodeCondtr As String
    Dim ViewSupp As String

Private Sub Form_Load()
    Call Initialiser
    Select Case StrSource
        Case "ConsultConge"
            LBL_Titre.Caption = "Consultation des 'Congés'"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Pic_ShowMenu.Visible = True
            Pic_Date.Visible = True
            Call Initialise_SBICombo_Cond(Cbo_CondConge)
            Call InitGrid_FindView_ConsultConge
            Call Affiche_ConsultConge
            Lbl_DateDuStiquePNG.Caption = Cda_DebutConge.Value
            Lbl_DateAuStiquePNG.Caption = Cda_FinConge.Value
            
        Case "ConducteurPing", "ConducteurPH", "ConducteurSuperv", "ConducteurE/H", "ConducteurStque", "Personnel", "Personnel_actif", "Personnel_base"
            LBL_Titre.Caption = "Liste des Conducteurs"
            If StrSource = "Personnel" Or StrSource = "Personnel_actif" Or StrSource = "Personnel_base" Then LBL_Titre.Caption = "Liste de Personnel"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Call InitGrid_FindView_Personnel
            Call Affiche_Personnel
        
        Case "VehiculePing", "VehiculePH", "VehiculeSuperv", "VehiculeStqueTF", _
                "VehiculeStqueRp", "VehiculeStqueCBr", "VehiculeBase", "VehiculeActif"
            LBL_Titre.Caption = "Liste des Vehicules"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Call InitGrid_FindView_Vehicule
            Call Affiche_Vehicule
            
        Case "StiquePLNG"
            LBL_Titre.Caption = "Statistiques PLANNING"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Pic_ShowMenu.Visible = True
            Pic_Date.Visible = True
            Call Initialise_SBICombo_Cond(Cbo_CondSPLNG)
            Call Initialise_SBICombo_PngDest(Cbo_DestSPLNG)
            Call Initialise_SBICombo_Cond(Cbo_CondSPLNG_TG)
            Call Initialise_SBICombo_PngDest(Cbo_DestSPLNG_TG)
            Call InitGrid_FindView_StatistiquePLNG
            Call Affiche_StatistiquePLNG(Cda_DebutSPLNG.Value, Cda_FinSPLNG.Value, Cbo_DestSPLNG.FirstValue, Cbo_CondSPLNG.FirstValue)
            Lbl_DateDuStiquePNG.Caption = Cda_DebutSPLNG.Value
            Lbl_DateAuStiquePNG.Caption = Cda_FinSPLNG.Value
            
        Case "ProgChauf"
            LBL_Titre.Caption = "Liste des Programmes"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Pic_ShowMenu.Visible = True
            Call Initialise_SBICombo_Cond(Cbo_CondPgHF)
            Call Initialise_SBICombo_Cond(Cbo_CondPgHT)
            Call InitGrid_FindView_ProgChauf
            Call Affiche_ProgChauffeursAvecDetails(Cda_DebutPgHF.Value, Cda_FinPgHF.Value, 4, ViewSupp, Cbo_CondPgHF.FirstValue)
            
        Case "FournisseurPH"
            LBL_Titre.Caption = "Liste des Fournisseurs"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Call InitGrid_FindView_FournisseurPH
            Call Affiche_FournisseurPH
            
        Case "Compteurs"
            LBL_Titre.Caption = "Liste des Compteurs"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Pic_PrintCompteur.Visible = True
            Call InitGrid_FindView_Compteur
            Call Affiche_Compteur
            
        Case "DestinationSuperv", "DestinationE/H", "STATDestinations"
            LBL_Titre.Caption = "Liste des Destinations"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Call InitGrid_FindView_Destination
            Call Affiche_Destination
            
        Case "ConducteurPLNG"
            LBL_Titre.Caption = "Liste des conducteurs n'ont pas de repos "
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Call InitGrid_FindView_Personnel
            Call Affiche_condRepos(Frm_PLANNING.CondRepos)
            
        Case "Utilisateur"
            LBL_Titre.Caption = "Liste des Utilisateurs"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Call InitGrid_FindView_Personnel
            Call Affiche_Utilisateur
            
            Case "Utilisateur_actif"
            LBL_Titre.Caption = "Liste des Utilisateurs actifs"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Call InitGrid_FindView_Personnel
            Call Affiche_Utilisateur_actif
            
        Case "Lubrifiant"
            LBL_Titre.Caption = "Lubrifiant"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Call InitGrid_FindView_Lubrifiant
            Call Affiche_Lubrifiant
     
        Case "Energie"
            LBL_Titre.Caption = "Energie"
            Me.Caption = LBL_Titre.Caption & " | " & App.ProductName
            Call InitGrid_FindView_Energie
            Call Affiche_Energie
    End Select
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
Private Sub Grid_FindView_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    Dim VCode
On Error GoTo Err
    VCode = Grid_FindView.CellText(lRow, 1)
    Select Case StrSource
        Case "VehiculePing"
            Unload Me
            Frm_PLANNING.AfficheRowVehiculePing (VCode)
        Case "ConducteurPing"
            Unload Me
            Frm_PLANNING.AfficheRowConducteurPing (VCode)
        Case "ProgChauf"
            VCode = Grid_FindView.CellText(lRow, 2)
            If VCode <> "" Then
                Unload Me
                Frm_PrgChauf.AfficheRowProgrammeCH (VCode)
            End If
        Case "VehiculePH"
'            Unload Me
'            Frm_PrgChauf.AfficheRowVehiculePH (VCode)
        Case "FournisseurPH"
'            Unload Me
'            Frm_PrgChauf.AfficheRowFournisseurPH (VCode)
        Case "ConducteurPH"
'            Unload Me
'            Frm_PrgChauf.AfficheRowconducteurPH (VCode)
        Case "VehiculeSuperv"
            Unload Me
            Frm_Supervision.AfficheRowVehiculeSup (VCode)
        Case "ConducteurSuperv"
            Unload Me
            Frm_Supervision.AfficheRowconducteurSup (VCode)
        Case "DestinationSuperv"
            Unload Me
            Frm_Supervision.AfficheRowDestinationSup (VCode)
        Case "ConducteurE/H"
            Unload Me
            Frm_Statistiques.AfficheRow_Conducteur (VCode)
        Case "ConducteurStque"
            Unload Me
            Frm_Statistiques.AfficheRow_Conducteur (VCode)
            
            
        Case "STATDestinations"
            Unload Me
            Frm_Stat_Dest.AfficheRow_Destination (VCode)
        Case "DestinationE/H"
            Unload Me
            Frm_Statistiques.AfficheRow_Destination (VCode)
        Case "VehiculeStqueTF"
            Unload Me
            Frm_Statistiques.AfficheRow_Vehicule (VCode)
        Case "VehiculeStqueRp"
            Unload Me
            Frm_Statistiques.AfficheRow_Vehicule (VCode)
        Case "VehiculeStqueCBr"
            Unload Me
            Frm_Statistiques.AfficheRow_Vehicule (VCode)
            
        Case "Utilisateur"
            Unload Me
'            Frm_Utilisateur.cmdFindMatricule.SetFocus
            Frm_Utilisateur.AfficheRow (VCode)
            
        Case "Utilisateur_actif"
            Unload Me
'            Frm_Utilisateur.cmdFindMatricule.SetFocus
            Frm_Utilisateur.AfficheRow (VCode)
            
        Case "Personnel", "Personnel_actif", "Personnel_base"
            Unload Me
            Frm_Personnel.AfficheRow (VCode)
        Case "Lubrifiant"
            Unload Me
            Frm_Vehicule.AfficheRow_Lubr (VCode)
        Case "VehiculeBase", "VehiculeActif"
            Unload Me
            Frm_Vehicule.AfficheRow (VCode)
        Case "Energie"
            Unload Me
            Frm_Vehicule.Cbo_Energie.Text = Grid_FindView.CellText(lRow, 2)
    End Select
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub Grid_FindView_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    Dim VCode
    Dim lRow As Integer
On Error GoTo Err
    If Grid_FindView.Rows > 0 Then
        lRow = Grid_FindView.SelectedRow
        VCode = Grid_FindView.CellText(lRow, 1)
        Select Case KeyCode
            Case vbKeyReturn
                Select Case StrSource
                    Case "VehiculePing"
                        Unload Me
                        Frm_PLANNING.AfficheRowVehiculePing (VCode)
                    Case "ConducteurPing"
                        Unload Me
                        Frm_PLANNING.AfficheRowConducteurPing (VCode)
                    Case "ProgChauf"
                        VCode = Grid_FindView.CellText(lRow, 2)
                        If VCode <> "" Then
                            Unload Me
                            Frm_PrgChauf.AfficheRowProgrammeCH (VCode)
                        End If
                    Case "VehiculePH"
'                        Unload Me
'                        Frm_PrgChauf.AfficheRowVehiculePH (VCode)
                    Case "FournisseurPH"
'                        Unload Me
'                        Frm_PrgChauf.AfficheRowFournisseurPH (VCode)
                    Case "ConducteurPH"
'                        Unload Me
'                        Frm_PrgChauf.AfficheRowconducteurPH (VCode)
                    Case "VehiculeSuperv"
                        Unload Me
                        Frm_Supervision.AfficheRowVehiculeSup (VCode)
                    Case "ConducteurSuperv"
                        Unload Me
                        Frm_Supervision.AfficheRowconducteurSup (VCode)
                    Case "DestinationSuperv"
                        Unload Me
                        Frm_Supervision.AfficheRowDestinationSup (VCode)
                    Case "ConducteurE/H"
                        Unload Me
                        Frm_Statistiques.AfficheRow_Conducteur (VCode)
                    Case "ConducteurStque"
                        Unload Me
                        Frm_Statistiques.AfficheRow_Conducteur (VCode)
                    Case "DestinationE/H"
                        Unload Me
                        Frm_Statistiques.AfficheRow_Destination (VCode)
                    Case "VehiculeStqueTF"
                        Unload Me
                        Frm_Statistiques.AfficheRow_Vehicule (VCode)
                    Case "VehiculeStqueRp"
                        Unload Me
                        Frm_Statistiques.AfficheRow_Vehicule (VCode)
                    Case "VehiculeStqueCBr"
                        Unload Me
                        Frm_Statistiques.AfficheRow_Vehicule (VCode)
                 
                    Case "Utilisateur", "Utilisateur_actif"
                        Unload Me
                        Frm_Utilisateur.AfficheRow (VCode)
                        
                    Case "Personnel", "Personnel_actif", "Personnel_base"
                        Unload Me
                        Frm_Personnel.AfficheRow (VCode)
                    Case "Lubrifiant"
                        Unload Me
                        Frm_Vehicule.AfficheRow_Lubr (VCode)
                        
                    Case "VehiculeBase", "VehiculeActif"
                        Unload Me
                        Frm_Vehicule.AfficheRow (VCode)
                        
                    
                End Select
                Case vbKeyEscape
                        Unload Me
        End Select
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'Initialized ControlBox***
Private Sub Initialiser()
    Pic_MaskMenu.Visible = False
    Tab_FindView.Visible = False
    Tab_FindView1.Visible = False
    Tab_FindView2.Visible = False
    Pic_ShowMenu.Visible = False
    Pic_Date.Visible = False
    Pic_PrintCompteur.Visible = False
    Cda_DebutConge.Value = "01/" & Month(Date) & "/" & Year(Date)
    Cda_FinConge.Value = Date
    Cda_DebutSPLNG.Value = "01/" & Month(Date) & "/" & Year(Date)
    Cda_FinSPLNG.Value = Date
    Cda_DebutSPLNG_TG.Value = "01/" & Month(Date) & "/" & Year(Date)
    Cda_FinSPLNG_TG.Value = Date
    Cda_DebutPgHF.Value = Date
    Cda_FinPgHF.Value = Date
    Cda_DebutPgHT.Value = "01/" & Month(Date) & "/" & Year(Date)
    Cda_FinPgHT.Value = Date
    ViewSupp = "N"
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'Initialized Grid_FindView***
Public Sub InitGrid_FindView_ConsultConge()
    With Grid_FindView
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
        .AddColumn "Code", "Code", , , 20, False, , , , , , CCLSortNumeric
        .AddColumn "Conducteur", "Conducteur", , , 200
        .AddColumn "TypeConge", "Type de Conge", , , 140
        .AddColumn "DateDu", "Date Du", , , 120
        .AddColumn "DateAu", "Date Au", , , 120
        .AddColumn "Null", "", , , , False
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Public Sub InitGrid_FindView_StatistiquePLNG()
    With Grid_FindView
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
        .AddColumn "Code", "Code", , , 20, False, , , , , , CCLSortNumeric
        .AddColumn "Conducteur", "Conducteur", , , 200, , , , , , , CCLSortString
        .AddColumn "Destination", "Destination", , , 160
        .AddColumn "Date", "Date", , , 150, , , , , , , CCLSortDate
        .AddColumn "Null", "", , , , False
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Public Sub InitGrid_FindView_StatistiquePLNG_TG()
    With Grid_FindView
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
        .AddColumn "Code", "Code", , , 20, False, , , , , , CCLSortNumeric
        .AddColumn "Date", "Date", , , 150, , , , , , , CCLSortDate
        .AddColumn "Conducteur", "Conducteur", , , 200, , , , , , , CCLSortString
        .AddColumn "Vehicule", "Véhicule", , , 160
        .AddColumn "Destination", "Destination", , , 160
        .AddColumn "Null", "", , , , False
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Private Sub InitGrid_FindView_Personnel()
    With Grid_FindView
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
        .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
        .AddColumn "Libelle", "Nom et prénom", , , 200
        If StrSource <> "ConducteurPH" And StrSource <> "ConducteurPing" And StrSource <> "ConducteurPLNG" And StrSource <> "ConducteurSuperv" Then .AddColumn "Actif", "Actif", , , 40
        .AddColumn "Null", ""
        .StretchLastColumnToFit = True
    End With
End Sub
Public Sub InitGrid_FindView_ProgChauf()
    With Grid_FindView
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
        .AddColumn "Code", ""
        .AddColumn "CodeProg", "Code", , , , False
        .AddColumn "Order", "Order", , , 60
        .AddColumn "Fournisseur", "Fournisseur", , , 150
        .AddColumn "TxtCommande", "Commande", , , 150
        .AddColumn "TxtPaiement", "Paiement", , , 150
        .AddColumn "TxtObservation", "Observation", , , 150
        .AddColumn "Null", ""
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Public Sub InitGrid_FindView_Vehicule()
    With Grid_FindView
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
        .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
        .AddColumn "Libelle", "Matricule", , , 160
        .AddColumn "Marque", "Marque", , , 150
        .AddColumn "Type", "Type", eSortType:=CCLSortStringNoCase
        .AddColumn "Energie", "Energie", eSortType:=CCLSortStringNoCase
        .AddColumn "Puissance", "Puissance", sFmtString:="short date", eSortType:=CCLSortDateDayAccuracy
        If StrSource <> "VehiculePH" And StrSource <> "VehiculePing" Then .AddColumn "Actif", "Acif", , , 40
        .AddColumn "Null", ""
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Private Sub InitGrid_FindView_FournisseurPH()
    With Grid_FindView
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
        .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
        .AddColumn "Libelle", "Libelle", , , 140
        .AddColumn "Type", "Type", , , 100
        If StrSource <> "FournisseurPH" Then .AddColumn "Activité", "Activité", , , 140
        If StrSource = "FournisseurPH" Then
            .AddColumn "Adresse", "Adresse", , , 280
        Else
            .AddColumn "Adresse", "Adresse", , , , 180
        End If
        .AddColumn "Actif", "Actif", , , , 40
        .AddColumn "Null", ""
        .StretchLastColumnToFit = True
    .Redraw = True
    End With
End Sub
Public Sub InitGrid_FindView_Compteur()
    With Grid_FindView
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
        .AddColumn "Code", "Code", , , 40, False, , , , , , CCLSortNumeric
        .AddColumn "Matricule", "Matricule", , , 140
        .AddColumn "CPTFT", "CPT.FT", , , 120
        .AddColumn "CPTBC", "CPT.BC", , , 120
        .AddColumn "CPTBV", "CPT.BV", , , 120
        .AddColumn "Null", "", , , 40, False
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Private Sub InitGrid_FindView_Destination()
    With Grid_FindView
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
        .AddColumn "Numero", "Numero", , , 60, False, , , , , , CCLSortNumeric
        .AddColumn "Libelle", "Libelle", , , 200
        .AddColumn "Type", "Type", , , 140
        .AddColumn "Actif", "Actif", , , 60
        .AddColumn "Null", ""
        .StretchLastColumnToFit = True
    End With
End Sub
Private Sub InitGrid_FindView_Lubrifiant()
    With Grid_FindView
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
        .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
        .AddColumn "A", "", , , 60, False
        .AddColumn "Libelle", "Libelle", , , 180
        .AddColumn "Prix", "Prix.TTC", , , 140
        .AddColumn "Null", ""
        .StretchLastColumnToFit = True
    End With
End Sub
Private Sub InitGrid_FindView_Energie()
    With Grid_FindView
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
        .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
        .AddColumn "Libelle", "Libelle", , , 140
        .AddColumn "Prix", "Prix.TTC", , , 140
        .AddColumn "Null", ""
        .StretchLastColumnToFit = True
    End With
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
Public Sub Affiche_ConsultConge()
    Dim Lrs_Find As New Recordset
    Dim LObj_Find As New Conducteur
On Error GoTo Err
    Set Lrs_Find = LObj_Find.Get_AllCongeByCond_Date(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Cda_DebutConge.Value, Cda_FinConge.Value, Cbo_CondConge.FirstValue)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not Lrs_Find.EOF Then
        Grid_FindView.Redraw = False
        While Not Lrs_Find.EOF
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find("Numero")
                .CellDetails .Rows, .ColumnIndex("Conducteur"), Lrs_Find("Conducteur")
                If Lrs_Find("Type") = "Repos" Then
                    .CellDetails .Rows, .ColumnIndex("TypeConge"), Lrs_Find("Type"), , , &HE0E0E0
                    .CellDetails .Rows, .ColumnIndex("DateDu"), Lrs_Find("datedu"), , , &HE0E0E0
                    .CellDetails .Rows, .ColumnIndex("DateAu"), Lrs_Find("Observation"), , , &HE0E0E0
                Else
                    .CellDetails .Rows, .ColumnIndex("TypeConge"), Lrs_Find("Type"), , , &HC0C0C0
                    .CellDetails .Rows, .ColumnIndex("DateDu"), Lrs_Find("datedu"), , , &HC0C0C0
                    .CellDetails .Rows, .ColumnIndex("DateAu"), Lrs_Find("dateau"), , , &HC0C0C0
                End If
            End With
            Lrs_Find.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
    Set LObj_Find = Nothing
    Set Lrs_Find = Nothing
    If Grid_FindView.Rows <> 0 Then
        With Grid_FindView
             .GroupRowBackColor = RGB(251, 246, 206)
             .GroupRowForeColor = QBColor(12)
             .ColumnIsGrouped(2) = True
             .GroupRowForeColor = QBColor(10)
             .HideGroupingBox = True
             .AllowGrouping = True
        End With
    End If
    If Grid_FindView.Rows > 0 Then Grid_FindView.SelectedRow = 1
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Affiche_Personnel()
    Dim LOBJ_Pers As New Conducteur
    Dim rs As New Recordset
    Dim Actif As String
    Dim Supp As String
On Error GoTo Err
    Actif = "O"
    Supp = "N"
    If StrSource = "Personnel_base" Or StrSource = "personnel" Then
        Actif = "A"
        Supp = "A"
    End If
    
    Set rs = LOBJ_Pers.GetAll_ConducteursActifNonSupp(ErrNumber, ErrDescription, ErrSourceDetail, Actif, Supp, CNB)
'    Set rs = LOBJ_Pers.GetAll_ConducteursActifNonSupp(ErrNumber, ErrDescription, ErrSourceDetail, StrSource, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs.EOF Then
        Grid_FindView.Redraw = False
        While Not rs.EOF
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, 1, rs("Code")
                .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle"), , , &H808080, &HFFFFFF
                If StrSource <> "ConducteurPH" And StrSource <> "ConducteurPing" And StrSource <> "ConducteurSuperv" Then .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HE0E0E0
            End With
            rs.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
    If Grid_FindView.Rows > 0 Then Grid_FindView.SelectedRow = 1
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub

Public Sub Affiche_Vehicule()
    Dim LObj_Find As New VEHICULE
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetAll_VehiculeActifNonSupp(ErrNumber, ErrDescription, ErrSourceDetail, CNB, StrSource)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        Grid_FindView.Redraw = False
        While Not Lrs_Find.EOF
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, 1, Lrs_Find("Code")
                .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find("Matricule"), , , &H808080, &HFFFFFF
                .CellDetails .Rows, .ColumnIndex("Marque"), Lrs_Find("Marque"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Type"), Lrs_Find("Type"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Energie"), Lrs_Find("Energie"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Puissance"), Lrs_Find("Puissance"), , , &HE0E0E0
                If StrSource <> "VehiculePH" And StrSource <> "VehiculePing" Then .CellDetails .Rows, .ColumnIndex("Actif"), Lrs_Find("Actif"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HE0E0E0
            End With
            Lrs_Find.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
    Set Lrs_Find = Nothing
    If Grid_FindView.Rows > 0 Then Grid_FindView.SelectedRow = 1
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
'Afficher tout les tournée effectuées par les conducteurs choisit et selon les destinations choisit pour la période précise
Private Sub Affiche_StatistiquePLNG(ByVal Date_db As Date, ByVal Date_f As Date, ByVal Desto As String, ByVal condt As String)
    Dim LObj_Find As New PLANNING
    Dim Lrs_Find As New Recordset
    Dim Lrs As New Recordset
    Dim JourDate As Integer
    Dim NbrPlng As Integer
On Error GoTo Err
    Set Lrs_Find = LObj_Find.Get_DetailPLNG(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Date_db, Date_f, Desto, condt)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
            Set Lrs = LObj_Find.Get_CountPLNG(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Date_db, Date_f, Lrs_Find("TOURNEE"), Lrs_Find("CONDUCTEUR"))
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            If Not Lrs.EOF Then
                NbrPlng = Lrs("nbrPLNG")
            End If
            Set LObj_Find = Nothing
            Set Lrs = Nothing
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find("Numero")
                .CellDetails .Rows, .ColumnIndex("Conducteur"), Lrs_Find("CONDUCTEUR") & " --    Tournée:  " & Lrs_Find("TOURNEE") & " --            " & NbrPlng
                .CellDetails .Rows, .ColumnIndex("Destination"), "", , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Date"), Format(Lrs_Find("DATEjour"), " dddd - dd/mm/yyyy"), , , &HE0E0E0
            End With
            Lrs_Find.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
    Set Lrs_Find = Nothing
    If Grid_FindView.Rows <> 0 Then
        With Grid_FindView
            .GroupRowBackColor = RGB(251, 246, 206)
            .GroupRowForeColor = QBColor(12)
            .ColumnIsGrouped(2) = True
            .GroupRowForeColor = QBColor(10)
            .HideGroupingBox = True
            .AllowGrouping = True
        End With
    End If
    If Grid_FindView.Rows > 0 Then Grid_FindView.SelectedRow = 1
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
'Afficher tout les tournée de garde effectuées par les conducteurs choisit et selon les destinations choisit pour la période précise
Private Sub Affiche_StatistiquePLNG_TG(ByVal Date_db As Date, ByVal Date_f As Date, ByVal Desto As String, ByVal condt As String)
    Dim LObj_Find As New PLANNING
    Dim Lrs_Find As New Recordset
    Dim Lrs As New Recordset
    Dim JourDate As Integer
On Error GoTo Err
    Set Lrs_Find = LObj_Find.Get_DetailPLNG_TG(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Date_db, Date_f, Desto, condt)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find("Numero")
                .CellDetails .Rows, .ColumnIndex("Date"), Lrs_Find("Datejour") & "  (" & Lrs_Find("jour") & ")"
                .CellDetails .Rows, .ColumnIndex("Conducteur"), Lrs_Find("CONDUCTEUR"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Vehicule"), Lrs_Find("vehicule"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Destination"), Lrs_Find("tournee"), , , &HE0E0E0
            End With
            Lrs_Find.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
    Set Lrs_Find = Nothing
    If Grid_FindView.Rows <> 0 Then
        With Grid_FindView
            .GroupRowBackColor = RGB(251, 246, 206)
            .GroupRowForeColor = QBColor(12)
            .ColumnIsGrouped(2) = True
            .GroupRowForeColor = QBColor(10)
            .HideGroupingBox = True
            .AllowGrouping = True
        End With
    End If
    If Grid_FindView.Rows > 0 Then Grid_FindView.SelectedRow = 1
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Affiche_ProgChauffeursAvecDetails(ByVal Ddebut As String, ByVal Dfin As String, ByVal Param As Integer, ByVal ViewSupp As String, ByVal cond As String)
    Dim LObj_Find As New ProgChauf
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetRow_ProgramChauffeur(ErrNumber, ErrDescription, ErrSourceDetail, Ddebut, Dfin, Param, ViewSupp, cond, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        Grid_FindView.Redraw = False
        While Not Lrs_Find.EOF
        Dim AssProg As String
        AssProg = Lrs_Find.Fields("DateProgramme") & "    Conducteur :  " & Lrs_Find.Fields("conducteur") & "    Véhicule :  " & Lrs_Find.Fields("Matricule")
        If Lrs_Find.Fields("Supp") = "O" Then AssProg = Lrs_Find.Fields("DateProgramme") & "    Conducteur :  " & Lrs_Find.Fields("conducteur") & "    Véhicule :  " & Lrs_Find.Fields("Matricule") & "            || ** Programme Supprimer **"
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Code"), AssProg, , , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("CodeProg"), Lrs_Find("Code")
                .CellDetails .Rows, .ColumnIndex("Order"), Lrs_Find.Fields("ProgOrder"), DT_CENTER, , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("Fournisseur"), Lrs_Find.Fields("Fournisseur"), , , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("TxtCommande"), Lrs_Find.Fields("TxtCommande"), , , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("TxtPaiement"), Lrs_Find.Fields("TxtPaiement"), , , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("TxtObservation"), Lrs_Find.Fields("TxtObservation"), , , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("Null"), "", , , RGB(225, 237, 226)
            End With
            Lrs_Find.MoveNext
        Wend
        Grid_FindView.Redraw = True
        Set Lrs_Find = Nothing
        With Grid_FindView
            .GroupRowBackColor = RGB(251, 246, 206)
            .GroupRowForeColor = QBColor(12)
            .ColumnIsGrouped(1) = True
            .GroupRowForeColor = QBColor(10)
            .HideGroupingBox = True
            .AllowGrouping = True
       End With
    End If
    If Grid_FindView.Rows > 0 Then Grid_FindView.SelectedRow = 1
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Affiche_FournisseurPH()
    Dim LObj_Find As New Fournisseur
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.Get_FournisAchat(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        Grid_FindView.Redraw = False
        While Not Lrs_Find.EOF
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, 1, Lrs_Find("Code")
                .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find("libelle"), , , &H808080, &HFFFFFF
                .CellDetails .Rows, .ColumnIndex("Type"), Lrs_Find("Type"), , , &HE0E0E0
                If StrSource <> "FournisseurPH" Then .CellDetails .Rows, .ColumnIndex("Activité"), Lrs_Find("Activite"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Adresse"), Lrs_Find("Adresse"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Actif"), Lrs_Find("Actif"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HE0E0E0
            End With
            Lrs_Find.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
    Set Lrs_Find = Nothing
    If Grid_FindView.Rows > 0 Then Grid_FindView.SelectedRow = 1
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Affiche_Compteur()
    Dim LObj_Find As New VEHICULE
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Grid_FindView.ClearRows
    Set Lrs_Find = LObj_Find.GetAllActifVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        Grid_FindView.Redraw = False
        While Not Lrs_Find.EOF
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Code"), "", , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Find("Matricule"), , , &H808080, &HFFFFFF
                .CellDetails .Rows, .ColumnIndex("CPTFT"), Lrs_Find("CompteurFT"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("CPTBC"), Lrs_Find("CompteurCarburant"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("CPTBV"), Lrs_Find("CompteurVidange"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HE0E0E0
            End With
            Lrs_Find.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
    Grid_FindView.SelectedRow = 1
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Affiche_Destination()
    Dim LObj_Find As New DESTINATION
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.Get_DestTrafic(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        Grid_FindView.Redraw = False
        While Not Lrs_Find.EOF
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, 1, Lrs_Find("Numero")
                .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find("Libelle"), , , &H808080, &HFFFFFF
                .CellDetails .Rows, .ColumnIndex("Type"), Lrs_Find("Type"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Actif"), Lrs_Find("Actif"), , , &HE0E0E0
                .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HE0E0E0
            End With
            Lrs_Find.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Affiche_condRepos(ByVal listCond As String)
    Dim XChp() As String
    Dim XCount As Integer, J As Integer
On Error GoTo Err
    Grid_FindView.Redraw = False
    XChp = Split(listCond, "||")
    XCount = UBound(XChp)
    For J = 0 To XCount
        With Grid_FindView
            .AddRow
            .CellDetails .Rows, .ColumnIndex("Code"), "", , , &HE0E0E0
            .CellDetails .Rows, .ColumnIndex("Libelle"), XChp(J), , , &H808080, &HFFFFFF
            .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HE0E0E0
        End With
    Next J
    Grid_FindView.Redraw = True
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Affiche_Utilisateur()
    Dim LOBJ_Personnel As personnel
    Dim rs As New Recordset
On Error GoTo Err
    Set LOBJ_Personnel = New personnel
    Set rs = LOBJ_Personnel.Get_AllUsers(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs.EOF Then
        Grid_FindView.Redraw = False
        While Not rs.EOF
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, 1, rs("Code"), , , &HC0C0C0
                .CellDetails .Rows, .ColumnIndex("Libelle"), rs("NOMPRN"), , , &H808080, &HFFFFFF
                .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif"), , , &HC0C0C0
                .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HE0E0E0
            End With
            rs.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
    If Grid_FindView.Rows > 0 Then Grid_FindView.SelectedRow = 1
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub

Public Sub Affiche_Utilisateur_actif()
    Dim LOBJ_Personnel As personnel
    Dim rs As New Recordset
On Error GoTo Err
    Set LOBJ_Personnel = New personnel
    Set rs = LOBJ_Personnel.Get_AllActifUsers(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs.EOF Then
        Grid_FindView.Redraw = False
        While Not rs.EOF
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, 1, rs("Code"), , , &HC0C0C0
                .CellDetails .Rows, .ColumnIndex("Libelle"), rs("NOMPRN"), , , &H808080, &HFFFFFF
                .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif"), , , &HC0C0C0
                .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HE0E0E0
            End With
            rs.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
    If Grid_FindView.Rows > 0 Then Grid_FindView.SelectedRow = 1
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Affiche_Lubrifiant()
    Dim LObj_Find   As New Produit_Lubrifiant
    Dim Lrs_Find    As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.Get_Lubrifiant(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        Grid_FindView.Redraw = False
        While Not Lrs_Find.EOF
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, 1, Lrs_Find("Numero"), , , &HC0C0C0
                .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find("libelle"), , , &H808080, &HFFFFFF
                .CellDetails .Rows, .ColumnIndex("Prix"), Format(Lrs_Find("prixht"), "#,##0.000"), , , &HC0C0C0
                .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HE0E0E0
            End With
            Lrs_Find.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Affiche_Energie()
    Dim LOBJ_Energie    As New Energie
    Dim rs              As New Recordset
On Error GoTo Err
    Set rs = LOBJ_Energie.Get_Energ(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs.EOF Then
        Grid_FindView.Redraw = False
        While Not rs.EOF
            With Grid_FindView
                .AddRow
                .CellDetails .Rows, 1, rs("Code"), , , &HC0C0C0
                .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle"), , , &H808080, &HFFFFFF
                .CellDetails .Rows, .ColumnIndex("Prix"), Format(rs("Prix"), "#,##0.000"), , , &HC0C0C0
                .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HE0E0E0
            End With
            rs.MoveNext
        Wend
        Grid_FindView.Redraw = True
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub

'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
Private Sub Pic_FindConge_Click()
On Error GoTo Err
    If Cda_DebutConge.Value > Cda_FinConge.Value Then
        MsgBox "Vérifier date de recherche!...", vbExclamation, App.ProductName
        Exit Sub
    End If
    Lbl_DateDuStiquePNG.Caption = Cda_DebutConge.Value
    Lbl_DateAuStiquePNG.Caption = Cda_FinConge.Value
    Call InitGrid_FindView_ConsultConge
    Call Affiche_ConsultConge
    If Grid_FindView.Rows = 0 Then MsgBox "Aucun congé pour " & Cbo_CondConge.SecondValue, vbInformation, App.ProductName
    Call Pic_MaskMenu_Click
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub Pic_FindSPLNG_Click()
    Dim condt As String
    Dim Desto As String
    Dim Text As String
On Error GoTo Err
    If Cda_DebutSPLNG.Value > Cda_FinSPLNG.Value Then
        MsgBox "Vérifier date de recherche!...", vbExclamation, App.ProductName
        Exit Sub
    End If
    condt = Cbo_CondSPLNG.FirstValue
    Desto = Cbo_DestSPLNG.FirstValue
    Lbl_DateDuStiquePNG.Caption = Cda_DebutSPLNG.Value
    Lbl_DateAuStiquePNG.Caption = Cda_FinSPLNG.Value
    Call InitGrid_FindView_StatistiquePLNG
    Call Affiche_StatistiquePLNG(Cda_DebutSPLNG.Value, Cda_FinSPLNG.Value, Desto, condt)
    If Grid_FindView.Rows = 0 Then MsgBox "Aucun Planning pour " & Cbo_CondSPLNG.SecondValue, vbInformation, App.ProductName
    Call Pic_MaskMenu_Click
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub Pic_FindSPLNG_TG_Click()
    Dim condt As String
    Dim Desto As String
On Error GoTo Err
    If Cda_DebutSPLNG_TG.Value > Cda_FinSPLNG_TG.Value Then
        MsgBox "Vérifier date de recherche!...", vbExclamation, App.ProductName
        Exit Sub
    End If
    condt = Cbo_CondSPLNG_TG.FirstValue
    Desto = Cbo_DestSPLNG_TG.FirstValue
    Lbl_DateDuStiquePNG.Caption = Cda_DebutSPLNG_TG.Value
    Lbl_DateAuStiquePNG.Caption = Cda_FinSPLNG_TG.Value
    Call InitGrid_FindView_StatistiquePLNG_TG
    Call Affiche_StatistiquePLNG_TG(Cda_DebutSPLNG_TG.Value, Cda_FinSPLNG_TG.Value, Desto, condt)
    If Grid_FindView.Rows = 0 Then MsgBox "Aucun Planning pour " & Cbo_CondSPLNG.SecondValue, vbInformation, App.ProductName
    Call Pic_MaskMenu_Click
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub Pic_FindPgHF_Click()
    Dim condt As String
    Dim Text As String
On Error GoTo Err
    If Cda_DebutPgHF.Value > Cda_FinPgHF.Value Then
        MsgBox "Date de recherche invalide!...     " & vbCr & "Verifier 'Date Au' ...", vbExclamation, App.ProductName
        Exit Sub
    End If
    condt = Cbo_CondPgHF.FirstValue
    Call InitGrid_FindView_ProgChauf
    Call Affiche_ProgChauffeursAvecDetails(Cda_DebutPgHF.Value, Cda_FinPgHF.Value, 4, ViewSupp, Cbo_CondPgHF.FirstValue)
    If Grid_FindView.Rows = 0 Then
        Dim Msg As VbMsgBoxResult
        Msg = MsgBox("Aucun programme en attend!..." & vbCr & " Voulez-vous afficher Tous", vbOKCancel + vbInformation, App.ProductName)
        If Msg = vbCancel Then Exit Sub
        Call Affiche_ProgChauffeursAvecDetails(Cda_DebutPgHT.Value, Cda_FinPgHT.Value, 0, ViewSupp, Cbo_CondPgHF.FirstValue)
        Tab_FindView1.Tab = 1
    End If
    Call Pic_MaskMenu_Click
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub Pic_FindPgHT_Click()
    Dim condt As String
    Dim Text As String
On Error GoTo Err
    condt = Cbo_CondPgHT.FirstValue
    If Cda_DebutPgHT.Value > Cda_FinPgHT.Value Then
        MsgBox "Date de recherche invalide!...     " & vbCr & "Verifier 'Date Au' ...", vbExclamation, App.ProductName
        Exit Sub
    Else
        If ChBox_Supprimer.Value = vbChecked Then ViewSupp = "O" Else ViewSupp = "N"
        Call InitGrid_FindView_ProgChauf
        Call Affiche_ProgChauffeursAvecDetails(Cda_DebutPgHT.Value, Cda_FinPgHT.Value, 0, ViewSupp, Cbo_CondPgHT.FirstValue)
    End If
    Call Pic_MaskMenu_Click
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub



'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'Recherche Par TextBox***
Private Sub Txt_Libelle_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err
    If Len(Trim(txt_libelle.Text)) <> 0 Then
        Call FindArticleKeydown(txt_libelle.Text, Grid_FindView)
    End If
Exit Sub
Err:
 MsgBox Err.Description & vbNewLine & Err.Source, vbQuestion
End Sub
Private Sub txt_Libelle_KeyPress(KeyAscii As Integer)
    Dim rech As String
    If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 26 And KeyAscii <> 127 Then
        rech = txt_libelle.Text & Chr(KeyAscii)
        If Len(Trim(rech)) <> 0 Then Call FindArticleKeydown(rech, Grid_FindView)
    Else
        If txt_libelle <> "" Then
            rech = Left(txt_libelle.Text, Len(txt_libelle.Text) - 1)
            Call FindArticleKeydown(rech, Grid_FindView)
        End If
    End If
End Sub
Private Sub txt_Libelle_GotFocus()
    If txt_libelle = "   Rechercher..." Then txt_libelle = ""
End Sub
Private Sub txt_Libelle_LostFocus()
    If txt_libelle = "" Then txt_libelle = "   Rechercher..."
End Sub
Private Sub FindArticleKeydown(ByVal vString As String, vGrid As SGrid)
    Dim i As Long
    For i = 1 To vGrid.Rows
        If Len(vString) = 0 Then
            vGrid.RowVisible(i) = True
            vGrid.Redraw = True
        Else
            If UCase(Mid(vGrid.CellText(i, 2), 1, Len(vString))) = UCase(vString) Then
                vGrid.RowVisible(i) = True
                vGrid.Redraw = True
            Else
                vGrid.RowVisible(i) = False
                vGrid.Redraw = True
            End If
        End If
    Next
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'Print Compteur***
Private Sub Pic_PrintCompteur_Click()
    Dim F As Form
    If MsgBox("Imprimer la liste!...", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
        Set F = New Frm_Rpt_Apercus
        With F
            Call .PrintOutAndApercu_Compteurs(0)
            .Show vbModal
        End With
    End If
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'Show & Mask SearchMenu***
Private Sub Pic_ShowMenu_Click()
    If Tab_FindView.Visible = False And Tab_FindView1.Visible = False And Tab_FindView2.Visible = False Then
        Pic_MaskMenu.Visible = True
        Pic_ShowMenu.Visible = False
        Select Case StrSource
            Case "ConsultConge"
                Tab_FindView2.Visible = True
            Case "StiquePLNG"
                Tab_FindView.Tab = 0
                Tab_FindView.Visible = True
                
            Case "ProgChauf"
                Tab_FindView1.Tab = 0
                Tab_FindView1.Visible = True
        End Select
    Else
        Pic_MaskMenu.Visible = False
        Pic_ShowMenu.Visible = True
        Tab_FindView.Visible = False
        Tab_FindView1.Visible = False
        Tab_FindView2.Visible = False
    End If
End Sub
Private Sub Pic_ConsltConge_DblClick()
    Pic_MaskMenu.Visible = False
    Pic_ShowMenu.Visible = True
    Tab_FindView.Visible = False
    Tab_FindView1.Visible = False
    Tab_FindView2.Visible = False
End Sub
Private Sub Pic_Header_DblClick()
    Call Pic_ShowMenu_Click
End Sub
Private Sub Pic_MaskMenu_Click()
    Pic_MaskMenu.Visible = False
    Pic_ShowMenu.Visible = True
    Tab_FindView.Visible = False
    Tab_FindView1.Visible = False
    Tab_FindView2.Visible = False
End Sub
Private Function DateCell(ByVal DateSearch As Date, ByVal jour As Integer) As Date
    Dim LObj_Find As New PLANNING, Lrs_Date As New Recordset
On Error GoTo Err
    Set Lrs_Date = LObj_Find.GetDate_NewPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, DateSearch, jour, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Function
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Date.EOF Then DateCell = Lrs_Date.Fields("datedebut")
    Set Lrs_Date = Nothing
Exit Function
Err:
    MsgBox Err.Description, vbExclamation
End Function




Private Sub Cbo_CondConge_LostFocus()
    Call ExistDonnee(Cbo_CondConge)
End Sub
Private Sub Cbo_CondPgHF_LostFocus()
    Call ExistDonnee(Cbo_CondPgHF)
End Sub
Private Sub Cbo_CondPgHT_LostFocus()
    Call ExistDonnee(Cbo_CondPgHT)
End Sub
Private Sub Cbo_CondSPLNG_LostFocus()
    Call ExistDonnee(Cbo_CondSPLNG)
End Sub
Private Sub Cbo_DestSPLNG_LostFocus()
    Call ExistDonnee(Cbo_DestSPLNG)
End Sub







