VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{AC8BD728-BB17-45B6-9ABA-3482EB8738F7}#1.2#0"; "HLBButton6.ocx"
Begin VB.Form Frm_PrgChauf 
   BackColor       =   &H80000000&
   Caption         =   "Programme chauffeurs"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   765
   ClientWidth     =   18210
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   18210
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Tab_Supp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9360
      ScaleHeight     =   375
      ScaleWidth      =   5415
      TabIndex        =   50
      Top             =   2180
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton cmd_Oui 
         BackColor       =   &H8000000C&
         Caption         =   "Ré-ajouter?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Lbl_Msg 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "=> Programme Supprimer, Voulez-Vous "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   52
         Top             =   60
         Width           =   3975
      End
   End
   Begin VB.PictureBox Tab_Button 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   16200
      ScaleHeight     =   4335
      ScaleWidth      =   615
      TabIndex        =   47
      Top             =   4250
      Width           =   615
   End
   Begin VB.PictureBox TAB_Find 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   4440
      Left            =   840
      ScaleHeight     =   4410
      ScaleWidth      =   1125
      TabIndex        =   36
      Top             =   2660
      Visible         =   0   'False
      Width           =   1155
      Begin SToolBox.SGrid Grid_Recherche 
         Height          =   3975
         Left            =   60
         TabIndex        =   37
         ToolTipText     =   "Sélectionner Tournée"
         Top             =   420
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7011
         RowMode         =   -1  'True
         GridLines       =   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   16777215
         GridLineColor   =   16777215
         GridFillLineColor=   16777215
         HighlightBackColor=   0
         HighlightForeColor=   16777215
         GroupRowBackColor=   -2147483624
         GroupRowForeColor=   192
         AlternateRowBackColor=   14737632
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
      Begin HLBButton6.HLBBttn Lbl_MaskTab 
         Height          =   300
         Left            =   6360
         TabIndex        =   38
         Top             =   60
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Frm_PrgChauf.frx":0000
         CaptionAlign    =   5
         PictureAlign    =   1
         BackColor       =   16777215
         Caption         =   "Fermer"
         ForeColor       =   0
         Appearance      =   3
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientBackColor=   -1  'True
         GradientColorStart=   16777215
         GradientColorEnd=   -2147483636
      End
      Begin VB.Label lbl_TabFind 
         BackColor       =   &H00808080&
         Caption         =   "  Liste des recherchés"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   300
         Left            =   60
         TabIndex        =   39
         Top             =   60
         Width           =   7935
      End
   End
   Begin VB.PictureBox Tab_Demo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   60
      ScaleHeight     =   4455
      ScaleWidth      =   14295
      TabIndex        =   45
      Top             =   4250
      Width           =   14295
      Begin SToolBox.SGrid Grid_Programme 
         Height          =   4215
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   7435
         RowMode         =   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         HighlightBackColor=   0
         HighlightForeColor=   16777215
         AlternateRowBackColor=   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderButtons   =   0   'False
         HeaderHeight    =   30
         BorderStyle     =   0
         DisableIcons    =   -1  'True
         MaxVisibleRows  =   0
      End
   End
   Begin VB.PictureBox Tab_Four 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   460
      Left            =   60
      ScaleHeight     =   465
      ScaleWidth      =   17055
      TabIndex        =   40
      Top             =   2180
      Width           =   17055
      Begin VB.TextBox Txt_Libelle_F 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   60
         Width           =   4815
      End
      Begin VB.TextBox Txt_Code_F 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1680
         TabIndex        =   42
         Top             =   60
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fournisseur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   41
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Tab_Assiet 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   800
      Left            =   60
      ScaleHeight     =   795
      ScaleWidth      =   17055
      TabIndex        =   26
      Top             =   1360
      Width           =   17055
      Begin VB.PictureBox PicBox_DProg 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   13200
         ScaleHeight     =   375
         ScaleWidth      =   1335
         TabIndex        =   34
         Top             =   120
         Width           =   1335
         Begin MSComCtl2.DTPicker DBox_DateProgramme 
            Height          =   375
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   149487617
            CurrentDate     =   42859
         End
      End
      Begin VB.TextBox Txt_Libelle_V 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   420
         Width           =   4815
      End
      Begin VB.TextBox Txt_Code_V 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1680
         TabIndex        =   31
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox Txt_Libelle_C 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   60
         Width           =   4815
      End
      Begin VB.TextBox Txt_Code_C 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1680
         TabIndex        =   28
         Top             =   60
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date de Programme"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   11280
         TabIndex        =   33
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicule"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Conducteur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Tab_Lign 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   80
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   7725
      TabIndex        =   25
      Top             =   1260
      Width           =   7725
   End
   Begin VB.PictureBox Tab_Code 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   60
      ScaleHeight     =   615
      ScaleWidth      =   17055
      TabIndex        =   18
      Top             =   680
      Width           =   17055
      Begin VB.PictureBox PicBox_DateCeartion 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Left            =   13320
         ScaleHeight     =   495
         ScaleWidth      =   3735
         TabIndex        =   22
         Top             =   0
         Width           =   3735
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Date de Création"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   120
            Width           =   1455
         End
         Begin VB.Image Image3 
            Height          =   300
            Left            =   0
            Picture         =   "Frm_PrgChauf.frx":0182
            Stretch         =   -1  'True
            Top             =   60
            Width           =   300
         End
         Begin VB.Label Lbl_DateCreation 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   1800
            TabIndex        =   23
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.TextBox TxtBox_Code 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   405
         Left            =   1680
         TabIndex        =   19
         Top             =   90
         Width           =   2055
      End
      Begin SToolBox.SCommand Cmd_LisCode 
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   120
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
         Picture         =   "Frm_PrgChauf.frx":0CAC
         ButtonType      =   1
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Left            =   360
         TabIndex        =   20
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   18180
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   18210
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programme Chauffeurs"
         BeginProperty Font 
            Name            =   "Footlight MT Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   360
         TabIndex        =   17
         Top             =   120
         Width           =   3345
      End
      Begin VB.Image Img_CN 
         Height          =   375
         Left            =   15240
         Picture         =   "Frm_PrgChauf.frx":0FE6
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   2760
         X2              =   15120
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   12200
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   12200
         X2              =   0
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label LBL_titre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   45
      End
   End
   Begin VB.PictureBox barMain 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   18210
      TabIndex        =   9
      Top             =   9105
      Width           =   18210
      Begin VB.PictureBox picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   12375
         TabIndex        =   10
         Top             =   60
         Width           =   12375
         Begin HLBButton6.HLBBttn Lbl_New 
            Height          =   420
            Left            =   1560
            TabIndex        =   11
            Top             =   30
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   741
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "Frm_PrgChauf.frx":7C04
            CaptionAlign    =   5
            PictureAlign    =   1
            BackColor       =   16777215
            Caption         =   "Nouveau"
            ForeColor       =   0
            Appearance      =   3
            BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientBackColor=   -1  'True
            GradientColorStart=   16777215
            GradientColorEnd=   -2147483636
         End
         Begin HLBButton6.HLBBttn Lbl_Print 
            Height          =   420
            Left            =   8760
            TabIndex        =   12
            Top             =   30
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Picture         =   "Frm_PrgChauf.frx":7F57
            CaptionAlign    =   5
            PictureAlign    =   1
            BackColor       =   16777215
            Caption         =   "Imprimer"
            ForeColor       =   -2147483642
            Appearance      =   3
            BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientBackColor=   -1  'True
            GradientColorStart=   16777215
            GradientColorEnd=   -2147483636
         End
         Begin HLBButton6.HLBBttn LBL_Sortir 
            Height          =   420
            Left            =   60
            TabIndex        =   13
            Top             =   30
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   741
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "Frm_PrgChauf.frx":80D9
            CaptionAlign    =   5
            PictureAlign    =   1
            BackColor       =   16777215
            Caption         =   "Quitter"
            ForeColor       =   -2147483642
            Appearance      =   3
            BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientBackColor=   -1  'True
            GradientColorStart=   16777215
            GradientColorEnd=   -2147483636
         End
         Begin HLBButton6.HLBBttn lbl_recherche 
            Height          =   420
            Left            =   3240
            TabIndex        =   14
            Top             =   30
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "Frm_PrgChauf.frx":825B
            CaptionAlign    =   5
            PictureAlign    =   1
            BackColor       =   16777215
            Caption         =   "Rechercher"
            ForeColor       =   0
            Appearance      =   3
            BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientBackColor=   -1  'True
            GradientColorStart=   16777215
            GradientColorEnd=   -2147483636
         End
         Begin HLBButton6.HLBBttn Lbl_Supp 
            Height          =   420
            Left            =   5040
            TabIndex        =   48
            Top             =   30
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   741
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
            Picture         =   "Frm_PrgChauf.frx":85AE
            CaptionAlign    =   5
            PictureAlign    =   1
            BackColor       =   16777215
            Caption         =   "F4 : Supprimer"
            ForeColor       =   0
            Appearance      =   3
            BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientBackColor=   -1  'True
            GradientColorStart=   16777215
            GradientColorEnd=   -2147483636
         End
         Begin HLBButton6.HLBBttn Lbl_Save 
            Height          =   420
            Left            =   6960
            TabIndex        =   49
            Top             =   30
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   741
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
            Picture         =   "Frm_PrgChauf.frx":8730
            CaptionAlign    =   5
            PictureAlign    =   1
            BackColor       =   16777215
            Caption         =   "Enregistre"
            ForeColor       =   0
            Appearance      =   3
            BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientBackColor=   -1  'True
            GradientColorStart=   16777215
            GradientColorEnd=   -2147483636
         End
      End
   End
   Begin VB.PictureBox Tab_Cmd 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   60
      ScaleHeight     =   1575
      ScaleWidth      =   17055
      TabIndex        =   3
      Top             =   2660
      Width           =   17055
      Begin VB.CommandButton Cmd_Valide 
         BackColor       =   &H8000000C&
         Caption         =   "Ajouter"
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1120
         Width           =   1455
      End
      Begin VB.TextBox TxtBox_Observation 
         BackColor       =   &H00FFFFFF&
         Height          =   1035
         Left            =   8400
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   60
         Width           =   6375
      End
      Begin VB.ComboBox TxtBox_Paiment 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   5295
      End
      Begin SToolBox.SMTextBox TxtBox_Commande 
         Height          =   975
         Left            =   1680
         TabIndex        =   0
         Top             =   60
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Lbl_NObservation 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(0/200)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7440
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "*: Moins de 200 caractéres"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8400
         TabIndex        =   7
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Observation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Paiement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Commande"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Menu sMenu 
      Caption         =   ""
      Begin VB.Menu sMenuAdd 
         Caption         =   "Ajouter"
      End
      Begin VB.Menu sMenuEdit 
         Caption         =   "Modifier"
      End
      Begin VB.Menu sMenuDelete 
         Caption         =   "Supprimer"
      End
   End
End
Attribute VB_Name = "Frm_PrgChauf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TypeFind                                As String
    Dim SaveNew                                 As Boolean
    Dim SaveEdit                                As Boolean
    Dim VCodeP                                  As String
    Public NProg As Integer
    Dim VCodeDrive As String, VCodeVehicle As String, VCodeFournisseur, VLibFournisseur As String


Private Sub Form_Load()
    TypeFind = ""
    DBox_DateProgramme.Value = Date
    SaveNew = False
    SaveEdit = False
    Call InitialGrid_Rech
    Call InitComBox_Payement
    Call EnabledControlBox(False)
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
Private Sub Form_Resize()
On Error Resume Next
    picBanner.Move 60, 0, Me.ScaleWidth - (Me.ScaleX(8, vbPixels, Me.ScaleMode))
    Img_CN.Left = Me.ScaleWidth - (Me.ScaleX(8, vbPixels, Me.ScaleMode)) - Img_CN.Width - 60
    Tab_Code.Move 60, Tab_Code.Top, Me.ScaleWidth - (Me.ScaleX(8, vbPixels, Me.ScaleMode))
    PicBox_DateCeartion.Left = Me.ScaleWidth - (Me.ScaleX(8, vbPixels, Me.ScaleMode)) - PicBox_DateCeartion.Width - 60
    Tab_Lign.Move 0, Tab_Lign.Top, Me.ScaleWidth
    Tab_Assiet.Move 60, Tab_Assiet.Top, Me.ScaleWidth - (Me.ScaleX(8, vbPixels, Me.ScaleMode))
    Tab_Four.Move 60, Tab_Four.Top, Me.ScaleWidth - (Me.ScaleX(8, vbPixels, Me.ScaleMode))
    Tab_Cmd.Move 60, Tab_Cmd.Top, Me.ScaleWidth - (Me.ScaleX(8, vbPixels, Me.ScaleMode))
    Tab_Demo.Move 60, Tab_Demo.Top, Me.ScaleWidth - Tab_Button.Width - 160, Me.ScaleHeight - Tab_Demo.Top - barMain.Height - 60
    Grid_Programme.Move 0, 0, Tab_Demo.Width, Tab_Demo.Height
    Tab_Button.Move 80 + Tab_Demo.Width, Tab_Demo.Top, Tab_Button.Width, Tab_Demo.Height
    Tab_Button.Left = 80 + Tab_Demo.Width
    TAB_Find.Move TAB_Find.Left, TAB_Find.Top, ScaleWidth - TAB_Find.Left * 2, Me.ScaleHeight - TAB_Find.Top - barMain.Height - 60
    Grid_Recherche.Move 60, 420, TAB_Find.Width - 120, TAB_Find.Height - 480
    lbl_TabFind.Move 60, 60, TAB_Find.Width - 120, lbl_TabFind.Height
    Lbl_MaskTab.Left = TAB_Find.Width - Lbl_MaskTab.Width - 60
End Sub
Private Sub barMain_Resize()
    Picture1.Move 60, 60, barMain.Width - 120, Picture1.Height
End Sub
Private Sub InitialGrid_Rech()
    With Grid_Recherche
        .Redraw = False
        .AllowGrouping = False
        .GroupRowBackColor = vbWindowBackground
        .GroupRowForeColor = vbWindowText
        .GridLineColor = vbWindowBackground
        .GridFillLineColor = vbWindowBackground
        .GridLines = True
        .SelectionAlphaBlend = True
        .SelectionOutline = True
        .DrawFocusRectangle = False
        .AddColumn "N", "N°", ecgHdrTextALignRight, , 30, eSortType:=CCLSortNumeric
        .AddColumn "Code", "Code", ecgHdrTextALignRight, , eSortType:=CCLSortNumeric
        .AddColumn "Libelle", "Libelle", eSortType:=CCLSortString
        .BackColor = RGB(225, 237, 226)
        .Redraw = True
        .StretchLastColumnToFit = True
    End With
    With Grid_Programme
        .Redraw = False
        .AllowGrouping = False
        .GroupRowBackColor = vbWindowBackground
        .GroupRowForeColor = vbWindowText
        .GridLineColor = vbWindowBackground
        .GridFillLineColor = vbWindowBackground
        .GridLines = True
        .SelectionAlphaBlend = True
        .SelectionOutline = True
        .DrawFocusRectangle = False
        .AddColumn "Numero", "N°", , , 50, , , , , , , CCLSortNumeric
        .AddColumn "Order", "Ordre", , , 50
        .AddColumn "CodeFR", "CodeFr", , , , False
        .AddColumn "Fournisseur", "Fournisseur", , , 200
        .AddColumn "Commande", "Commande", , , 90
        .AddColumn "Paiment", "Paiment", , , 90
        .AddColumn "Observation", "Observation", , , 250
        .AddColumn "DateProg", "Date Programme", , , 90
        .AddColumn "Null", ""
        .BackColor = RGB(225, 237, 226)
        .StretchLastColumnToFit = True
    End With
End Sub
Private Sub InitialBlanc()
    Grid_Programme.ClearRows
    Grid_Recherche.ClearRows
    Txt_Code_C.Text = ""
    Txt_Code_V.Text = ""
    Txt_Code_F.Text = ""
    Txt_Libelle_C.Text = ""
    Txt_Libelle_V.Text = ""
    Txt_Libelle_F.Text = ""
    TxtBox_Commande.Text = ""
    TxtBox_Observation.Text = ""
    TxtBox_Paiment.Text = ""
    DBox_DateProgramme.Value = Date
    Lbl_DateCreation.Caption = Date
End Sub
Private Sub InitComBox_Payement()
    TxtBox_Paiment.AddItem "Chèque"
    TxtBox_Paiment.AddItem "Traite"
    TxtBox_Paiment.AddItem "Espèce"
    TxtBox_Paiment.ListIndex = 0
End Sub
Private Sub EnabledControlBox(ByVal TYP As Boolean)
    Tab_Assiet.Enabled = TYP
    Tab_Four.Enabled = TYP
    Tab_Cmd.Enabled = TYP
    Tab_Demo.Enabled = TYP
End Sub
Private Sub MASK_GRID()
    TAB_Find.Visible = False
    Grid_Recherche.ClearRows
End Sub
Private Sub InitializeAdd()
    Txt_Code_F.Text = ""
    Txt_Libelle_F.Text = ""
    TxtBox_Commande.Text = ""
    TxtBox_Paiment.Text = ""
    TxtBox_Observation.Text = ""
    Cmd_Valide.Caption = "Ajouter"
    NProg = 0
End Sub






'========== ControlBox***
Private Sub TxtBox_Code_GotFocus()
    TxtBox_Code.BackColor = &HC0FFFF
    Call MASK_GRID
End Sub
Private Sub TxtBox_Code_LostFocus()
    TxtBox_Code.BackColor = &HFFFFFF
End Sub
Private Sub Txt_Code_C_GotFocus()
    Txt_Code_C.BackColor = &HC0FFFF
    Call MASK_GRID
End Sub
Private Sub Txt_Code_C_LostFocus()
    Txt_Code_C.BackColor = &HFFFFFF
End Sub
Private Sub Txt_Code_C_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TypeFind = "COND"
            Call GetByCode(TypeFind, Txt_Code_C.Text)
        Case vbKeyRight
            TypeFind = "COND"
            Call FindByKeyDown(TypeFind, Txt_Code_C.Text)
        Case vbKeyEscape
        
    End Select
End Sub
Private Sub Txt_Code_V_GotFocus()
    If Len(Trim(Txt_Code_C.Text)) = 0 Then
        Txt_Code_C.SetFocus
        Exit Sub
    End If
    Txt_Code_V.BackColor = &HC0FFFF
    Call MASK_GRID
End Sub
Private Sub Txt_Code_V_LostFocus()
    Txt_Code_V.BackColor = &HFFFFFF
End Sub
Private Sub Txt_Code_V_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TypeFind = "VEH"
            Call GetByCode(TypeFind, Txt_Code_V.Text)
        Case vbKeyRight
            TypeFind = "VEH"
            Call FindByKeyDown(TypeFind, Txt_Code_V.Text)
        Case vbKeyEscape
        
    End Select
End Sub
Private Sub Txt_Code_f_GotFocus()
    If Len(Trim(Txt_Code_C.Text)) = 0 Then
        Txt_Code_C.SetFocus
        Exit Sub
    End If
    If Len(Trim(Txt_Code_V.Text)) = 0 Then
        Txt_Code_V.SetFocus
        Exit Sub
    End If
    Txt_Code_F.BackColor = &HC0FFFF
    Call MASK_GRID
End Sub
Private Sub Txt_Code_f_LostFocus()
    Txt_Code_F.BackColor = &HFFFFFF
End Sub
Private Sub Txt_Code_f_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TypeFind = "FOUR"
            Call GetByCode(TypeFind, Txt_Code_F.Text)
        Case vbKeyRight
            TypeFind = "FOUR"
            Call FindByKeyDown(TypeFind, Txt_Code_F.Text)
        Case vbKeyEscape
        
    End Select
End Sub
Private Sub GetByCode(ByVal xTypeFind As String, ByVal xCode As String)
    Dim Lrs_Find                    As New Recordset
On Error GoTo Err
    If xTypeFind = "COND" Then
        Dim LObj_Find                   As New Conducteur
        Set Lrs_Find = LObj_Find.GetRow_Conducteur_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, xCode, CNB)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
        If Not Lrs_Find.EOF Then
            Txt_Code_C.Text = Lrs_Find("Code")
            Txt_Libelle_C.Text = Lrs_Find("Libelle")
            Txt_Code_V.SetFocus
        Else
            MsgBox "Code Conducteur Invalid!!...       ", vbExclamation
            Txt_Code_C.SetFocus
            Txt_Code_C.SelStart = 0
            Txt_Code_C.SelLength = Len(Txt_Code_C.Text)
        End If
    ElseIf xTypeFind = "VEH" Then
        Dim LObj_FindV                   As New VEHICULE
        Set Lrs_Find = LObj_FindV.GetRow_Vehicule_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, xCode)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_FindV = Nothing
        If Not Lrs_Find.EOF Then
            Txt_Code_V.Text = Lrs_Find("Code")
            Txt_Libelle_V.Text = Lrs_Find("Matricule")
            Txt_Code_F.SetFocus
        Else
            MsgBox "Code Véhicule Invalid!!...       ", vbExclamation
            Txt_Code_V.SetFocus
            Txt_Code_V.SelStart = 0
            Txt_Code_V.SelLength = Len(Txt_Code_V.Text)
        End If
    ElseIf xTypeFind = "FOUR" Then
        Dim LObj_FindF                   As New Fournisseur
        Set Lrs_Find = LObj_FindF.GetRow_Fournisseur_ACHAT_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, xCode, CNB)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_FindV = Nothing
        If Not Lrs_Find.EOF Then
            Txt_Code_F.Text = Lrs_Find("Code")
            Txt_Libelle_F.Text = Lrs_Find("Libelle")
            TxtBox_Commande.SetFocus
        Else
            MsgBox "Code Fournisseur Invalid!!...       ", vbExclamation
            Txt_Code_F.SetFocus
            Txt_Code_F.SelStart = 0
            Txt_Code_F.SelLength = Len(Txt_Code_F.Text)
        End If
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description
End Sub
Private Sub FindByKeyDown(ByVal xTypeFind As String, ByVal xLibelle As String)
    Dim Lrs_Find                    As New Recordset
On Error GoTo Err
    If xTypeFind = "COND" Then
        Dim LObj_Find                   As New Conducteur
        If Len(xLibelle) > 0 Then
            Set Lrs_Find = LObj_Find.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, xLibelle, CNB)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
        ElseIf Len(xLibelle) = 0 Then
            Set Lrs_Find = LObj_Find.GetAll_ConducteursActif(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
        End If
        Set LObj_Find = Nothing
        If Not Lrs_Find.EOF Then
            If Lrs_Find.RecordCount = 1 Then
                Txt_Code_C.Text = Lrs_Find("Code")
                Txt_Libelle_C.Text = Lrs_Find("Libelle")
                Txt_Code_V.SetFocus
            Else
                Grid_Recherche.Redraw = False
                With Grid_Recherche
                    While Not Lrs_Find.EOF
                        .AddRow
                        .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find.Fields("Code"), DT_RIGHT
                        .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find.Fields("Libelle")
                        Lrs_Find.MoveNext
                    Wend
                End With
                Call N_Ligne(Grid_Recherche)
                Grid_Recherche.Redraw = True
                TAB_Find.Visible = True
                Grid_Recherche.SetFocus
                Grid_Recherche.SelectedRow = 1
            End If
        Else
            MsgBox "Conducteur Invalid!!...       ", vbExclamation
            Txt_Code_C.SetFocus
            Txt_Code_C.SelStart = 0
            Txt_Code_C.SelLength = Len(Txt_Code_C.Text)
        End If
    ElseIf xTypeFind = "VEH" Then
        Dim LObj_FindV                   As New VEHICULE
        If Len(xLibelle) > 0 Then
            Set Lrs_Find = LObj_FindV.GetRow_Vehicule_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, CNB, xLibelle)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
        ElseIf Len(xLibelle) = 0 Then
            Set Lrs_Find = LObj_FindV.GetAll_Vehicule_Actif(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
        End If
        Set LObj_FindV = Nothing
        If Not Lrs_Find.EOF Then
            If Lrs_Find.RecordCount = 1 Then
                Txt_Code_V.Text = Lrs_Find("Code")
                Txt_Libelle_V.Text = Lrs_Find("Matricule")
                Txt_Code_F.SetFocus
            Else
                Grid_Recherche.Redraw = False
                With Grid_Recherche
                    While Not Lrs_Find.EOF
                        .AddRow
                        .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find.Fields("Code"), DT_RIGHT
                        .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find.Fields("Matricule")
                        Lrs_Find.MoveNext
                    Wend
                End With
                Call N_Ligne(Grid_Recherche)
                Grid_Recherche.Redraw = True
                TAB_Find.Visible = True
                Grid_Recherche.SetFocus
                Grid_Recherche.SelectedRow = 1
            End If
        Else
            MsgBox "Véhicule Invalid!!...       ", vbExclamation
            Txt_Code_V.SetFocus
            Txt_Code_V.SelStart = 0
            Txt_Code_V.SelLength = Len(Txt_Code_V.Text)
        End If
    ElseIf xTypeFind = "FOUR" Then
        Dim LObj_FindF                   As New Fournisseur
        If Len(xLibelle) > 0 Then
            Set Lrs_Find = LObj_FindF.GetRow_Fournisseur_ACHAT_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, CNB, xLibelle)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
        ElseIf Len(xLibelle) = 0 Then
            Set Lrs_Find = LObj_FindF.GetAll_Fournisseur_ACHAT(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
        End If
        Set LObj_FindV = Nothing
        If Not Lrs_Find.EOF Then
            If Lrs_Find.RecordCount = 1 Then
                Txt_Code_F.Text = Lrs_Find("Code")
                Txt_Libelle_F.Text = Lrs_Find("Libelle")
                TxtBox_Commande.SetFocus
            Else
                Grid_Recherche.Redraw = False
                With Grid_Recherche
                    While Not Lrs_Find.EOF
                        .AddRow
                        .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find.Fields("Code"), DT_RIGHT
                        .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Find.Fields("Libelle")
                        Lrs_Find.MoveNext
                    Wend
                End With
                Call N_Ligne(Grid_Recherche)
                Grid_Recherche.Redraw = True
                TAB_Find.Visible = True
                Grid_Recherche.SetFocus
                Grid_Recherche.SelectedRow = 1
            End If
        Else
            MsgBox "Fournisseur Invalid!!...       ", vbExclamation
            Txt_Code_F.SetFocus
            Txt_Code_F.SelStart = 0
            Txt_Code_F.SelLength = Len(Txt_Code_F.Text)
        End If
        
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description
End Sub
Private Sub Grid_Recherche_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    If TypeFind = "COND" Then
        Txt_Code_C.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Code"))
        Txt_Libelle_C.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Libelle"))
        Txt_Code_V.SetFocus
    ElseIf TypeFind = "VEH" Then
        Txt_Code_V.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Code"))
        Txt_Libelle_V.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Libelle"))
        Txt_Code_F.SetFocus
    ElseIf TypeFind = "FOUR" Then
        Txt_Code_F.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Code"))
        Txt_Libelle_F.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Libelle"))
        TxtBox_Commande.SetFocus
    End If
    Call MASK_GRID
End Sub
Private Sub Grid_Recherche_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyReturn Then
        If TypeFind = "COND" Then
            Txt_Code_C.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Code"))
            Txt_Libelle_C.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Libelle"))
            Txt_Code_V.SetFocus
        ElseIf TypeFind = "VEH" Then
            Txt_Code_V.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Code"))
            Txt_Libelle_V.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Libelle"))
            Txt_Code_F.SetFocus
        ElseIf TypeFind = "FOUR" Then
            Txt_Code_F.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Code"))
            Txt_Libelle_F.Text = Grid_Recherche.CellText(Grid_Recherche.SelectedRow, Grid_Recherche.ColumnIndex("Libelle"))
            TxtBox_Commande.SetFocus
        End If
        Call MASK_GRID
    ElseIf KeyCode = vbKeyEscape Then
        If TypeFind = "COND" Then
            Txt_Code_C.SetFocus
        ElseIf TypeFind = "VEH" Then
            Txt_Code_V.SetFocus
        ElseIf TypeFind = "FOUR" Then
            Txt_Code_F.SetFocus
        End If
        Call MASK_GRID
    End If
End Sub
Private Sub Lbl_MaskTab_Click()
    If TypeFind = "COND" Then
        Txt_Code_C.SetFocus
    ElseIf TypeFind = "VEH" Then
        Txt_Code_V.SetFocus
    ElseIf TypeFind = "FOUR" Then
        Txt_Code_F.SetFocus
    End If
    Call MASK_GRID
End Sub
'====Liste(Find View)~~~
Private Sub Cmd_LisCode_Click()
    If TxtBox_Code.Text = "Auto" Then If MsgBox("Voulez-vous annuler la création en cours?...", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "ProgChauf"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub DBox_DateProgramme_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
        Txt_Code_C.SelStart = 0
        Txt_Code_C.SelLength = Len(Txt_Code_C.Text)
        Txt_Code_C.SetFocus
    End If
End Sub
Private Sub TxtBox_Commande_GotFocus()
    If Len(Trim(Txt_Code_C.Text)) = 0 Then
        Txt_Code_C.SetFocus
        Exit Sub
    End If
    If Len(Trim(Txt_Code_V.Text)) = 0 Then
        Txt_Code_V.SetFocus
        Exit Sub
    End If
    If Len(Trim(Txt_Code_F.Text)) = 0 Then
        Txt_Code_F.SetFocus
        Exit Sub
    End If
    TxtBox_Observation.BackColor = &HC0FFFF
    Call MASK_GRID
End Sub
Private Sub TxtBox_Commande_LostFocus()
    TxtBox_Commande.BackColor = &HFFFFFF
End Sub
Private Sub TxtBox_Observation_GotFocus()
    If Len(Trim(Txt_Code_C.Text)) = 0 Then
        Txt_Code_C.SetFocus
        Exit Sub
    End If
    If Len(Trim(Txt_Code_V.Text)) = 0 Then
        Txt_Code_V.SetFocus
        Exit Sub
    End If
    If Len(Trim(Txt_Code_F.Text)) = 0 Then
        Txt_Code_F.SetFocus
        Exit Sub
    End If
    TxtBox_Observation.BackColor = &HC0FFFF
    Call MASK_GRID
End Sub
Private Sub TxtBox_Observation_LostFocus()
    TxtBox_Observation.BackColor = &HFFFFFF
End Sub
Private Sub TxtBox_Observation_KeyPress(KeyAscii As Integer)
    Lbl_NObservation.Caption = "(" & Len(TxtBox_Observation.Text) & "/200)"
End Sub
Private Sub TxtBox_Paiment_GotFocus()
    If Len(Trim(Txt_Code_C.Text)) = 0 Then
        Txt_Code_C.SetFocus
        Exit Sub
    End If
    If Len(Trim(Txt_Code_V.Text)) = 0 Then
        Txt_Code_V.SetFocus
        Exit Sub
    End If
    If Len(Trim(Txt_Code_F.Text)) = 0 Then
        Txt_Code_F.SetFocus
        Exit Sub
    End If
    TxtBox_Paiment.BackColor = &HC0FFFF
    Call MASK_GRID
End Sub
Private Sub TxtBox_Paiment_LostFocus()
    TxtBox_Paiment.BackColor = &HFFFFFF
End Sub
Private Sub TxtBox_Paiment_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Cmd_Valide.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        TxtBox_Observation.SetFocus
    End If
End Sub
Private Sub Cmd_Valide_Click()
    If Len(Trim(Txt_Code_C.Text)) = 0 Then
        Txt_Code_C.SetFocus
        Exit Sub
    End If
    If Len(Trim(Txt_Code_V.Text)) = 0 Then
        Txt_Code_V.SetFocus
        Exit Sub
    End If
    If Len(Trim(Txt_Code_F.Text)) = 0 Then
        Txt_Code_F.SetFocus
        Exit Sub
    End If
    If Cmd_Valide.Caption = "Ajouter" Then
        If (CHECK_ACCES("Ins_PCH", LInt_UserId) = True) Then
            Call Add
            Lbl_Save.Enabled = True
        Else
            MsgBox "Insertion, n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-etre pas des autorisations nécessaires pour ajouter un programme", vbInformation, App.ProductName
            Exit Sub
        End If
    ElseIf Cmd_Valide.Caption = "Modifier" Then
        If (CHECK_ACCES("Maj_PCH", LInt_UserId) = True) Then
            Call Edit
            Lbl_Save.Enabled = True
        Else
            MsgBox "Modification, n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-etre pas des autorisations nécessaires pour modifier un programme", vbInformation, App.ProductName
            Exit Sub
        End If
    End If
End Sub
Private Sub Add()
    Dim DateProgram                         As Date
    Dim TOrder                              As Integer
    Dim i                                   As Integer
    Dim Msg                                 As VbMsgBoxResult
    Dim TCommand                            As String
    Dim TPayment                            As String
    Dim TObservation                        As String
On Error GoTo Err
    DateProgram = DBox_DateProgramme.Value
    If Grid_Programme.Rows = 0 Then
        TOrder = 1
    Else
        TOrder = Grid_Programme.CellText(Grid_Programme.Rows, 2) + 1
    End If
    If Len(TxtBox_Observation.Text) > 200 Then
        MsgBox "Text d'observation plus long..." & vbCr & "Nombre de caractéres plus long", vbInformation, App.ProductName
        TxtBox_Observation.SetFocus
        Exit Sub
    End If
    If Len(TxtBox_Commande.Text) > 200 Then
        MsgBox "Text de commande plus long..." & vbCr & "Nombre de caractéres plus long", vbInformation, App.ProductName
        TxtBox_Commande.SetFocus
        Exit Sub
    End If
    If Len(TxtBox_Paiment.Text) > 200 Then
        MsgBox "Text de Paiment plus long..." & vbCr & "Nombre de caractéres plus long", vbInformation, App.ProductName
        TxtBox_Paiment.SetFocus
        Exit Sub
    End If
    'Verification de contenue de programme
    If TxtBox_Commande.Text <> "" Then TCommand = TxtBox_Commande.Text Else TCommand = "----"
    If TxtBox_Paiment.Text <> "" Then TPayment = TxtBox_Paiment.Text Else TPayment = "----"
    If TxtBox_Observation.Text <> "" Then TObservation = TxtBox_Observation.Text Else TObservation = "----"
    If TCommand = "----" And TPayment = "----" And TObservation = "----" Then
        MsgBox "Aucun détail pour l'enregistrer!...", vbInformation
        TxtBox_Commande.SetFocus
        Exit Sub
    End If
    If DateProgram < Date Then
        MsgBox "Date de 'Programme' invalide" & vbCr & "Cette Date " & DBox_DateProgramme.Value & " est passée", vbExclamation, App.ProductName
        Exit Sub
    Else
        Grid_Programme.Redraw = False
        With Grid_Programme
            .AddRow
            .CellDetails .Rows, .ColumnIndex("Numero"), Grid_Programme.Rows
            .CellDetails .Rows, .ColumnIndex("Order"), TOrder
            .CellDetails .Rows, .ColumnIndex("CodeFR"), Txt_Code_F.Text
            .CellDetails .Rows, .ColumnIndex("Fournisseur"), Txt_Libelle_F.Text
            .CellDetails .Rows, .ColumnIndex("Commande"), TCommand
            .CellDetails .Rows, .ColumnIndex("Paiment"), TPayment
            .CellDetails .Rows, .ColumnIndex("Observation"), TObservation
            .CellDetails .Rows, .ColumnIndex("DateProg"), DateProgram
        End With
        Grid_Programme.Redraw = True
        Grid_Programme.SelectedRow = 1
    End If
    Call InitializeAdd
    Lbl_NObservation.Caption = "(0/200)"
    Txt_Code_F.SetFocus
Exit Sub
Err:
   MsgBox Err.Description, 48, App.ProductName
End Sub
Private Sub Edit()
    Dim i As Integer, Msg As VbMsgBoxResult, DateProgram As String, TCommand As String
    Dim TPayment As String, TObservation As String
On Error GoTo Err
    If DBox_DateProgramme.Value < Date Then
        MsgBox "Ce programme est passé!...          " & vbCr & "Vérifier sa date !", vbInformation, App.ProductName
        DBox_DateProgramme.SetFocus
        Exit Sub
    End If
    DateProgram = DBox_DateProgramme.Value
    If Len(TxtBox_Observation.Text) > 200 Then
        MsgBox "Text d'observation est long...", vbInformation, App.ProductName
        TxtBox_Observation.SetFocus
        Exit Sub
    End If
    If Len(TxtBox_Commande.Text) > 200 Then
        MsgBox "Text de commande est long...", vbInformation, App.ProductName
        TxtBox_Commande.SetFocus
        Exit Sub
    End If
    'Verification de contenue de programme
    If TxtBox_Commande.Text <> "" Then TCommand = TxtBox_Commande.Text Else TCommand = "----"
    If TxtBox_Paiment.Text <> "" Then TPayment = TxtBox_Paiment.Text Else TPayment = "----"
    If TxtBox_Observation.Text <> "" Then TObservation = TxtBox_Observation.Text Else TObservation = "----"
    If TCommand = "----" And TPayment = "----" And TObservation = "----" Then
        MsgBox "Aucun détail pour l'enregistrer!...", vbInformation
        Exit Sub
    End If
    Msg = MsgBox("Voulez-vous modifier ce programme", vbOKCancel + vbExclamation, App.ProductName)
        If Msg = vbCancel Then Exit Sub
        With Grid_Programme
            .CellText(NProg, 2) = NProg
            .CellText(NProg, 3) = Txt_Code_F.Text
            .CellText(NProg, 4) = Txt_Libelle_F.Text
            .CellText(NProg, 5) = TCommand
            .CellText(NProg, 6) = TPayment
            .CellText(NProg, 7) = TObservation
            .CellText(NProg, 8) = DateProgram
        End With
        MsgBox "Programme Modifier avec succès", vbInformation, App.ProductName
    Call InitializeAdd
    Lbl_NObservation.Caption = "(0/200)"
    Txt_Code_F.SetFocus
Exit Sub
Err:
   MsgBox Err.Description, 48, App.ProductName
End Sub
Private Sub Grid_Programme_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    Call EditProg(Grid_Programme.SelectedRow)
End Sub
Private Sub Grid_Programme_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    NProg = lRow
End Sub
Private Sub Grid_Programme_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With Grid_Programme
        sMenuEdit.Enabled = False
        sMenuAdd.Enabled = False
        sMenuDelete.Enabled = False
        If DBox_DateProgramme.Value >= Date Then
            sMenuAdd.Enabled = True
            If .Rows > 0 Then
                sMenuEdit.Enabled = True
                sMenuDelete.Enabled = True
            End If
        End If
        If (Button = vbRightButton) Then
            Me.PopupMenu sMenu
        End If
    End With
End Sub
Private Sub EditProg(ByVal NProg As Integer)
    Dim i As Integer, CountFournisseur As Integer, TFournisseur
On Error GoTo Err:
    If DBox_DateProgramme.Value < Date Then
        MsgBox "Ce programme est passé!...          " & vbCr & "Vérifier la date du programme!", vbInformation, App.ProductName
        Exit Sub
    End If
    TxtBox_Commande.Text = ""
    TxtBox_Paiment.Text = ""
    TxtBox_Observation.Text = ""
    If NProg > 0 Then
        Txt_Code_F.Text = Grid_Programme.CellText(NProg, 3)
        Txt_Libelle_F.Text = Grid_Programme.CellText(NProg, 4)
    Else
        MsgBox "Code Invalide!...", vbExclamation
        Exit Sub
    End If
    DBox_DateProgramme.Value = Grid_Programme.CellText(NProg, 8)
'    TxtBox_Order.text = Grid_Programme.CellText(NProg, 2)
    If Grid_Programme.CellText(NProg, 5) <> "" And Grid_Programme.CellText(NProg, 5) <> "----" Then TxtBox_Commande.Text = Grid_Programme.CellText(NProg, 5) Else TxtBox_Commande.Text = ""
    If Grid_Programme.CellText(NProg, 6) <> "" And Grid_Programme.CellText(NProg, 6) <> "----" Then TxtBox_Paiment.Text = Grid_Programme.CellText(NProg, 6) Else TxtBox_Paiment.Text = ""
    If Grid_Programme.CellText(NProg, 7) <> "" And Grid_Programme.CellText(NProg, 7) <> "----" Then
        TxtBox_Observation.Text = Grid_Programme.CellText(NProg, 7)
        Lbl_NObservation.Caption = "(" & Len(Grid_Programme.CellText(NProg, 7)) & "/200)"
    Else
        TxtBox_Observation.Text = ""
        Lbl_NObservation.Caption = "(0/200)"
    End If
    Cmd_Valide.Caption = "Modifier"
    Txt_Code_F.SetFocus
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub sMenuEdit_click()
    Call EditProg(Grid_Programme.SelectedRow)
End Sub
Private Sub sMenuAdd_click()
    Dim Msg As VbMsgBoxResult
    If TxtBox_Commande.Text <> "" Or TxtBox_Observation.Text <> "" Or TxtBox_Paiment.Text <> "" Then
        Msg = MsgBox("Voulez-vous rénitialise l'ajout", vbExclamation + vbYesNo, App.ProductName)
        If Msg = vbNo Then Exit Sub
    Else
        Call InitializeAdd
        Lbl_NObservation.Caption = "(0/200)"
        Txt_Code_F.SetFocus
    End If
End Sub
Private Sub sMenuDelete_click()
    Call SuppRow(Grid_Programme.SelectedRow)
End Sub
Private Sub SuppRow(ByVal NProg As Integer)
        Dim Msg As VbMsgBoxResult
On Error GoTo Err
    If NProg > 0 Then
        Msg = MsgBox("Voulez-Vous Supprimer cet Programme", vbExclamation + vbOKCancel, App.ProductName)
        If Msg = vbOK Then
            With Grid_Programme
                .RemoveRow (NProg)
            End With
            NProg = 0
        Else
            Exit Sub
        End If
    Else
        MsgBox "Selectionner un Programme puis supprimer", vbExclamation, App.ProductName
        Exit Sub
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub LBL_Sortir_Click()
    Unload Me
End Sub
'====Ajouter nouveau Programme~~~
Private Sub Lbl_New_Click()
    If Lbl_New.Caption = "Nouveau" Then
        If (CHECK_ACCES("Ins_PCH", LInt_UserId) = True) Then
            Tab_Code.Enabled = False
            Call EnabledControlBox(True)
            Lbl_Print.Enabled = False
            Lbl_Supp.Enabled = False
            lbl_recherche.Enabled = False
            Lbl_Save.Enabled = True
            Tab_Supp.Visible = False
            Lbl_New.Caption = "Annuler"
            Call InitialBlanc
            TxtBox_Code.Text = "Auto"
            Txt_Code_C.SetFocus
            SaveNew = True
        Else
            MsgBox "Insertion n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-être pas des autorisations nécessaires pour ajouter un programme", vbInformation, App.ProductName
            Exit Sub
        End If
    ElseIf Lbl_New.Caption = "Annuler" Then
        If MsgBox("Voulez-vous annuler la création en cours?...", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
        Tab_Code.Enabled = True
        Lbl_Print.Enabled = False
        Lbl_Supp.Enabled = False
        lbl_recherche.Enabled = True
        Lbl_Save.Enabled = False
        Tab_Supp.Visible = False
        Lbl_New.Caption = "Nouveau"
        Call InitialBlanc
        TxtBox_Code.Text = ""
        Call EnabledControlBox(False)
        SaveNew = False
    End If
End Sub
Private Sub lbl_recherche_Click()
    Tab_Code.Enabled = True
    Call InitialBlanc
    TxtBox_Code.Text = ""
    Lbl_Print.Enabled = False
    Lbl_Supp.Enabled = False
    Lbl_Save.Enabled = False
    Lbl_New.Enabled = True
    Tab_Supp.Visible = False
    Call Cmd_LisCode_Click
End Sub
Private Sub Lbl_Save_Click()
On Error GoTo Err
    If DBox_DateProgramme.Value >= Date And IsDate(DBox_DateProgramme.Value) Then
        If SaveNew = True Then
            If (CHECK_ACCES("Ins_PCH", LInt_UserId) = True) Then
                Call SaveProgram
            Else
                MsgBox "Insertion n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-être pas des autorisations nécessaires pour ajouter un programme", vbInformation, "Parcano..."
                Exit Sub
            End If
        ElseIf SaveEdit = True Then
            If (CHECK_ACCES("Maj_PCH", LInt_UserId) = True) Then
                Call EditProgram
            Else
                MsgBox "Modification n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-être pas des autorisations nécessaires pour modifier un programme", vbInformation, "Parcano..."
                Exit Sub
            End If
        End If
    Else
        MsgBox "Date de 'Programme' invalide" & vbCr & "Cette Date " & DBox_DateProgramme.Value & " est passée", vbExclamation, "Parcano..."
    End If
Exit Sub
Err:
   MsgBox Err.Description, 48, App.ProductName
End Sub
Private Sub SaveProgram()
    Dim VCodeProg                       As String
    Dim LInt_NumCompteur                As Long
    Dim i                               As Integer
    Dim Msg                             As VbMsgBoxResult
    Dim Lobj_SaveProg                   As New ProgChauf
    Dim Lrs_Ass_ProgCH                  As New Recordset
    Dim Lrs_Det_ProgCH                  As New Recordset
On Error GoTo Err
    VCodeProg = TxtBox_Code.Text
    If TxtBox_Code.Text = "Auto" And Grid_Programme.Rows > 0 Then
        Msg = MsgBox("Voulez-vous enregistre cet Programme", vbYesNo + vbExclamation, "Parcano...")
        If Msg = vbNo Then Exit Sub
        If VCodeProg = "Auto" Then
            LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "ProgramChauffeur")
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
            TxtBox_Code.Text = Format(LInt_NumCompteur, "00000")
            VCodeProg = Format(LInt_NumCompteur, "00000")
        End If
        Set Lrs_Ass_ProgCH = CreateEmptyRS_Ass_ProgChauf()
        With Lrs_Ass_ProgCH
            .AddNew
            .Fields("Code") = VCodeProg
            .Fields("CodeConducteur") = Txt_Code_C.Text
            .Fields("CodeVehicule") = Txt_Code_V.Text
            .Fields("DateCreation") = Lbl_DateCreation.Caption
            .Fields("DateProgramme") = Format(DBox_DateProgramme.Value, "dd/mm/yyyy")
            .Fields("UserInsert") = LInt_UserId
        End With
        Set Lrs_Det_ProgCH = CreateEmptyRS_Det_ProgChauf()
        For i = 1 To Grid_Programme.Rows
            With Lrs_Det_ProgCH
                .AddNew
                .Fields("CodeProgChauf") = VCodeProg
                .Fields("ProgOrder") = Grid_Programme.CellText(i, 2)
                .Fields("CodeFournisseur") = Grid_Programme.CellText(i, 3)
                .Fields("TxtCommande") = Grid_Programme.CellText(i, 5)
                .Fields("TxtPaiement") = Grid_Programme.CellText(i, 6)
                .Fields("TxtObservation") = Grid_Programme.CellText(i, 7)
            End With
        Next i
        Set Lobj_SaveProg = New ProgChauf
        Call Lobj_SaveProg.Save_AssProgCH(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Ass_ProgCH)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set Lobj_SaveProg = Nothing
        Set Lobj_SaveProg = New ProgChauf
        Call Lobj_SaveProg.Save_DetProgCH(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Det_ProgCH)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set Lobj_SaveProg = Nothing
        MsgBox "Programme Enregistre avec succès!...", vbInformation, "Parcano..."
        Call Lbl_Print_Click
        Call InitialBlanc
        TxtBox_Code.Text = ""
        Call EnabledControlBox(False)
        Lbl_New.Caption = "Nouveau"
        Tab_Code.Enabled = True
        Lbl_Print.Enabled = False
        Lbl_Supp.Enabled = False
        Lbl_Save.Enabled = False
        lbl_recherche.Enabled = True
        SaveNew = False
        SaveEdit = False
    Else
        MsgBox " Aucun Programme pour enregistrer", vbExclamation, "Parcano..."
    End If
    Set Lrs_Det_ProgCH = Nothing
    Set Lrs_Ass_ProgCH = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub EditProgram()
    Dim Lrs_Ass_ProgCH                      As New Recordset
    Dim Lrs_Det_ProgCH                      As New Recordset
    Dim LObj_Delete                         As New ProgChauf
    Dim LObj_Update                         As New ProgChauf
    Dim Msg                                 As VbMsgBoxResult
    Dim VCodeProgram                        As String
    Dim DateProg                            As String
    Dim i                                   As Integer
On Error GoTo Err
    VCodeProgram = TxtBox_Code.Text
    DateProg = DBox_DateProgramme.Value
    If DateProg >= Date And IsDate(DBox_DateProgramme.Value) Then
        If TxtBox_Code.Text <> "Auto" And Grid_Programme.Rows > 0 Then
            Msg = MsgBox("Voulez-vous modifier le programme N°: " & VCodeProgram, vbExclamation + vbOKCancel, "Parcano...")
            If Msg = vbCancel Then Exit Sub
            Set Lrs_Ass_ProgCH = CreateEmptyRS_Ass_ProgChauf()
            With Lrs_Ass_ProgCH
                .AddNew
                .Fields("Code") = VCodeProgram
                .Fields("CodeConducteur") = Txt_Code_C.Text
                .Fields("CodeVehicule") = Txt_Code_V.Text
                .Fields("DateProgramme") = Format(DBox_DateProgramme.Value, "dd/mm/yyyy")
                .Fields("UserUpdate") = LInt_UserId
            End With
            Set Lrs_Det_ProgCH = CreateEmptyRS_Det_ProgChauf()
            For i = 1 To Grid_Programme.Rows
                With Lrs_Det_ProgCH
                    .AddNew
                    .Fields("CodeProgChauf") = VCodeProgram
                    .Fields("ProgOrder") = Grid_Programme.CellText(i, 2)
                    .Fields("CodeFournisseur") = Grid_Programme.CellText(i, 3)
                    .Fields("TxtCommande") = Grid_Programme.CellText(i, 5)
                    .Fields("TxtPaiement") = Grid_Programme.CellText(i, 6)
                    .Fields("TxtObservation") = Grid_Programme.CellText(i, 7)
                End With
            Next i
            Call LObj_Delete.Delete_Det_Chauf(ErrNumber, ErrDescription, ErrSourceDetail, VCodeProgram, CNB)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
            Set LObj_Delete = Nothing
            Call LObj_Update.Update_Ass_ProgChauf(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Ass_ProgCH)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
            Set LObj_Update = Nothing
            Call LObj_Update.Save_DetProgCH(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Det_ProgCH)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
            Set LObj_Update = Nothing
            MsgBox "Programme Modifier avec succès!...", vbInformation, "Parcano..."
            Call Lbl_Print_Click
            Call InitialBlanc
            TxtBox_Code.Text = ""
            Call EnabledControlBox(False)
            Lbl_New.Caption = "Nouveau"
            Tab_Code.Enabled = True
            Lbl_Print.Enabled = False
            Lbl_Supp.Enabled = False
            Lbl_Save.Enabled = False
            lbl_recherche.Enabled = True
            SaveNew = False
            SaveEdit = False
        Else
            MsgBox "Aucun Programme pour enregistrer", vbExclamation, "Parcano..."
        End If
    Else
        MsgBox "Date de 'Programme' invalide" & vbCr & "Cet Date " & DBox_DateProgramme.Value & " il passe", vbExclamation, "Parcano..."
        Exit Sub
    End If
    Set Lrs_Ass_ProgCH = Nothing
    Set Lrs_Det_ProgCH = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub Lbl_Print_Click()
    On Error GoTo Err
    If TxtBox_Code.Text = "Auto" Then
        MsgBox "Séléctionner un programme puis imprimer...", vbExclamation, App.ProductName
        Exit Sub
    End If
    If Grid_Programme.Rows = 0 Then
        MsgBox "Aucun Programme", vbExclamation, "Parcano..."
        Exit Sub
    End If
    If MsgBox("Imprimer la Programme en cours...?        ", vbYesNo + vbDefaultButton1 + vbInformation, App.ProductName) = vbYes Then
        Frm_Rpt_Apercus.Numero = TxtBox_Code.Text
        Call Frm_Rpt_Apercus.PrintOutAndApercu_ProgChauf(0)
        Frm_Rpt_Apercus.Show
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub Lbl_Supp_Click()
    Dim CodeProgChauf                   As String
    CodeProgChauf = TxtBox_Code.Text
    If (CHECK_ACCES("Supp_PCH", LInt_UserId) = True) Then
        Call DeleteAddProgram(CodeProgChauf, "O", LInt_UserId)
        Call InitialBlanc
        TxtBox_Code.Text = ""
        Call EnabledControlBox(False)
        Lbl_New.Caption = "Nouveau"
        Tab_Code.Enabled = True
        Lbl_Print.Enabled = False
        Lbl_Supp.Enabled = False
        Lbl_Save.Enabled = False
    Else
        MsgBox "Suppression n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-etre pas des autorisations nécessaires pour Supprimer un programme", vbInformation, App.ProductName
        Exit Sub
    End If
End Sub
Private Sub cmd_Oui_Click()
    Dim CodeProgChauf As String
    CodeProgChauf = TxtBox_Code.Text
    If (CHECK_ACCES("Supp_PCH", LInt_UserId) = True) Then
        Call DeleteAddProgram(CodeProgChauf, "N", LInt_UserId)
        Call InitialBlanc
        TxtBox_Code.Text = ""
        Call EnabledControlBox(False)
        Lbl_New.Caption = "Nouveau"
        Tab_Code.Enabled = True
        Lbl_Print.Enabled = False
        Lbl_Supp.Enabled = False
        Lbl_Save.Enabled = False
        Tab_Supp.Visible = False
    Else
        MsgBox "Ajoutation n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-etre pas des autorisations nécessaires pour Ré-ajouter un programme", vbInformation, App.ProductName
        Exit Sub
    End If
End Sub
Private Sub DeleteAddProgram(ByVal CodeP As String, ByVal LettreSupp As String, ByVal CodeUser As String)
    Dim LObj_Find                       As New ProgChauf
    Dim Msg                             As VbMsgBoxResult
On Error GoTo Err
    If CodeP <> "Auto" Then
        If LettreSupp = "O" Then
            Msg = MsgBox("Voulez-vous supprimer cet programme " & CodeP & " !...", vbExclamation + vbOKCancel, App.ProductName)
        ElseIf LettreSupp = "N" Then
            Msg = MsgBox("Voulez-vous re-ajouter cet programme " & CodeP & " !...", vbExclamation + vbOKCancel, App.ProductName)
        End If
        If Msg = vbCancel Then Exit Sub
        Call LObj_Find.Delete_Restaurer(ErrNumber, ErrDescription, ErrSourceDetail, CNB, CodeP, LettreSupp, CodeUser)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
        If LettreSupp = "O" Then
            MsgBox "Programme Supprimer avec succes!...", vbInformation, App.ProductName
        ElseIf LettreSupp = "N" Then
            MsgBox "Programme Ré-ajouter avec succes!...", vbInformation, App.ProductName
        End If
    Else
        MsgBox "Séléctionner un Programme", App.ProductName
        Exit Sub
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub
Public Sub AfficheRowProgrammeCH(ByVal VCode As String)
    Dim LObj_Find                       As New ProgChauf
    Dim Lrs_Ass_Prog                    As New Recordset
    Dim Lrs_Det_Prog                    As New Recordset
    Dim i                               As Integer
On Error GoTo Err
    Call InitialBlanc
    Set Lrs_Ass_Prog = LObj_Find.GetRow_Ass_ProgramCH(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Ass_Prog.EOF Then
        TxtBox_Code.Text = Lrs_Ass_Prog("Code")
        Txt_Code_C.Text = Lrs_Ass_Prog("codepersonne")
        Txt_Libelle_C.Text = Lrs_Ass_Prog("Libelle")
        Txt_Code_V.Text = Lrs_Ass_Prog("codevehicule")
        Txt_Libelle_V.Text = Lrs_Ass_Prog("Matricule")
        If Not IsNull(Lrs_Ass_Prog("DateCreation")) Then Lbl_DateCreation.Caption = Lrs_Ass_Prog("DateCreation")
        If Not IsNull(Lrs_Ass_Prog("DateProgramme")) Then DBox_DateProgramme.Value = Lrs_Ass_Prog("DateProgramme")
        Set LObj_Find = New ProgChauf
        Set Lrs_Det_Prog = LObj_Find.GetRow_DetailsProgramCH(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
        Grid_Programme.Redraw = False
        While Not Lrs_Det_Prog.EOF
            With Grid_Programme
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Numero"), Lrs_Det_Prog.Fields("code")
                .CellDetails .Rows, .ColumnIndex("Order"), Lrs_Det_Prog.Fields("ProgOrder")
                .CellDetails .Rows, .ColumnIndex("CodeFR"), Lrs_Det_Prog.Fields("codeFr")
                .CellDetails .Rows, .ColumnIndex("Fournisseur"), Lrs_Det_Prog.Fields("Libelle")
                .CellDetails .Rows, .ColumnIndex("Commande"), Lrs_Det_Prog.Fields("TxtCommande")
                .CellDetails .Rows, .ColumnIndex("Paiment"), Lrs_Det_Prog.Fields("TxtPaiement")
                .CellDetails .Rows, .ColumnIndex("Observation"), Lrs_Det_Prog.Fields("TxtObservation")
                .CellDetails .Rows, .ColumnIndex("DateProg"), Lrs_Det_Prog.Fields("DateProgramme")
            End With
            Lrs_Det_Prog.MoveNext
        Wend
        Grid_Programme.Redraw = True
        Grid_Programme.SelectedRow = 1
        Lbl_Print.Enabled = True
        SaveNew = False
        If DBox_DateProgramme.Value < Date Then
            Call EnabledControlBox(False)
            SaveEdit = False
            Lbl_Save.Enabled = False
            Lbl_Supp.Enabled = False
        Else
            Call EnabledControlBox(True)
            SaveEdit = True
            Lbl_Save.Enabled = True
            Lbl_Print.Enabled = True
            Lbl_Supp.Enabled = True
        End If
        If Lrs_Ass_Prog.Fields("Supp") = "O" Then
            Call EnabledControlBox(False)
            SaveEdit = False
            Tab_Supp.Visible = True
            Lbl_Supp.Enabled = False
            Lbl_Print.Enabled = False
            Lbl_Save.Enabled = False
        Else
            Tab_Supp.Visible = False
        End If
    Else
        MsgBox "Code introuvable", vbInformation
    End If
    Set Lrs_Ass_Prog = Nothing
    Set Lrs_Det_Prog = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
