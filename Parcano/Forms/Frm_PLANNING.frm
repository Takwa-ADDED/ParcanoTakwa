VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form Frm_PLANNING 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Planning semaine"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17355
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   17355
   WindowState     =   2  'Maximized
   Begin SToolBox.SCommand Cmd_suivt 
      Height          =   495
      Left            =   11160
      TabIndex        =   32
      Top             =   960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_PLANNING.frx":0000
      BackColor       =   16777215
   End
   Begin VB.PictureBox PicBox_Planning 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   2760
      ScaleHeight     =   3735
      ScaleWidth      =   9975
      TabIndex        =   19
      Top             =   2160
      Width           =   9975
      Begin VB.PictureBox Pic_HeureEntre 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6720
         ScaleHeight     =   255
         ScaleWidth      =   2535
         TabIndex        =   40
         Top             =   600
         Width           =   2535
         Begin VB.CheckBox Chk_HEntre 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Heure d'entre"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   0
            Width           =   255
         End
         Begin SToolBox.STimeBox Txt_Heurentre 
            Height          =   285
            Left            =   1560
            TabIndex        =   41
            Top             =   0
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            Enabled         =   0   'False
            BackColor       =   14737632
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Heure d'entre"
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
            Height          =   255
            Left            =   360
            TabIndex        =   42
            Top             =   0
            Width           =   1215
         End
      End
      Begin SToolBox.SOptionButton OptSoir 
         Height          =   315
         Left            =   5280
         TabIndex        =   35
         Top             =   1200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BackStyle       =   0
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
         ForeColor       =   255
      End
      Begin SToolBox.SOptionButton Opt_Matin 
         Height          =   315
         Left            =   5280
         TabIndex        =   34
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BackStyle       =   0
         Value           =   1
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
         ForeColor       =   255
      End
      Begin SToolBox.SCommand Cmd_Cancel 
         Height          =   495
         Left            =   8520
         TabIndex        =   33
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "Annuler"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   8421504
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   9975
         TabIndex        =   28
         Top             =   3360
         Width           =   9975
      End
      Begin VB.PictureBox PicBox_Tournee 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   9975
         TabIndex        =   20
         Top             =   0
         Width           =   9975
         Begin VB.Label Cbo_Tournee 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   495
            Left            =   4680
            TabIndex        =   30
            Top             =   0
            Width           =   4815
         End
         Begin VB.Label Lbl_Journee 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   495
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   3615
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "TOURNEE"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Index           =   0
            Left            =   3600
            TabIndex        =   21
            Top             =   0
            Width           =   1095
         End
      End
      Begin SToolBox.SGrid Grid_DetPlanning 
         Height          =   1695
         Left            =   0
         TabIndex        =   7
         Top             =   1560
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2990
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
      Begin SToolBox.SCommand Cmd_Vehicule 
         Height          =   375
         Left            =   9240
         TabIndex        =   9
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
         Picture         =   "Frm_PLANNING.frx":09A2
         ButtonType      =   1
      End
      Begin SToolBox.SCommand Cmd_Conducteur 
         Height          =   375
         Left            =   4320
         TabIndex        =   8
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
         Picture         =   "Frm_PLANNING.frx":0D24
         ButtonType      =   1
      End
      Begin SToolBox.SCommand Cmd_AddNew 
         Height          =   495
         Left            =   8520
         TabIndex        =   3
         Top             =   1560
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
         Picture         =   "Frm_PLANNING.frx":10A6
      End
      Begin SToolBox.SCommand Cmd_AddPlanning 
         Height          =   495
         Left            =   8520
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
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
         Picture         =   "Frm_PLANNING.frx":1228
         BackColor       =   14737632
      End
      Begin SToolBox.SBiCombo cbo_conducteur 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SToolBox.SBiCombo cbo_vehicule 
         Height          =   360
         Left            =   4800
         TabIndex        =   2
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SToolBox.SCommand Cmd_Edit 
         Height          =   495
         Left            =   9000
         TabIndex        =   4
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
         Picture         =   "Frm_PLANNING.frx":13AA
      End
      Begin SToolBox.SCommand Cmd_Ok 
         Height          =   495
         Left            =   9000
         TabIndex        =   5
         Top             =   1560
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
         Picture         =   "Frm_PLANNING.frx":1604
         BackColor       =   14737632
      End
      Begin SToolBox.SCommand CmdDelete 
         Height          =   495
         Left            =   9480
         TabIndex        =   27
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
         Picture         =   "Frm_PLANNING.frx":1786
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Après midi"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5520
         TabIndex        =   38
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Matin"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5520
         TabIndex        =   36
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line2 
         X1              =   9960
         X2              =   0
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "VEHICULE"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   23
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CONDUCTEUR"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.PictureBox Pic_Menu 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   8775
      Left            =   0
      ScaleHeight     =   8775
      ScaleWidth      =   2655
      TabIndex        =   17
      Top             =   1560
      Width           =   2655
      Begin SToolBox.SCommand Cmd_Cog 
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   4560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Congés"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SToolBox.SCommand Cmd_Rep 
         Height          =   495
         Left            =   120
         TabIndex        =   45
         Top             =   4560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Repos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
      End
      Begin SToolBox.SCommand Cmd_Repos 
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   3960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Sans Repos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SToolBox.SCommand Cmd_Conge 
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   3360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Conge"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   8421504
         ForeColor       =   16777215
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   2775
         TabIndex        =   24
         Top             =   0
         Width           =   2775
         Begin VB.Label Lbl_Search 
            BackStyle       =   0  'Transparent
            Caption         =   "Rechercher"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   600
            TabIndex        =   26
            Top             =   0
            Width           =   1455
         End
         Begin VB.Image Pic_Show 
            Height          =   375
            Left            =   0
            Picture         =   "Frm_PLANNING.frx":1AD9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Pic_Mask 
            Height          =   375
            Left            =   2160
            Picture         =   "Frm_PLANNING.frx":242F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
      End
      Begin MSComCtl2.MonthView MonthView 
         Height          =   2370
         Left            =   0
         TabIndex        =   18
         Top             =   840
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   14737632
         BackColor       =   8421504
         Appearance      =   1
         MonthBackColor  =   4210752
         StartOfWeek     =   149028866
         TitleBackColor  =   12632256
         TitleForeColor  =   16711680
         TrailingForeColor=   8421504
         CurrentDate     =   42868
      End
      Begin VB.Image Img_Cog 
         Height          =   1455
         Left            =   0
         Picture         =   "Frm_PLANNING.frx":4511
         Stretch         =   -1  'True
         Top             =   6720
         Width           =   405
      End
      Begin VB.Image Img_Rep 
         Height          =   1455
         Left            =   0
         Picture         =   "Frm_PLANNING.frx":984B
         Stretch         =   -1  'True
         Top             =   6720
         Width           =   405
      End
      Begin VB.Image Pic_Conge 
         Height          =   1455
         Left            =   0
         Picture         =   "Frm_PLANNING.frx":E7D9
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   405
      End
      Begin VB.Image Pic_Repos 
         Height          =   1455
         Left            =   0
         Picture         =   "Frm_PLANNING.frx":13A9B
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   405
      End
      Begin VB.Image Pic_Consulter 
         Height          =   1575
         Left            =   0
         Picture         =   "Frm_PLANNING.frx":18D5D
         Stretch         =   -1  'True
         Top             =   720
         Width           =   405
      End
      Begin VB.Line Line1 
         X1              =   1440
         X2              =   2640
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Lbl_Title 
         BackStyle       =   0  'Transparent
         Caption         =   "PLANNING par semaine"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.PictureBox Pic_Planning 
      BackColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   2640
      ScaleHeight     =   6435
      ScaleWidth      =   14595
      TabIndex        =   16
      Top             =   1560
      Width           =   14655
      Begin SToolBox.SGrid Grid_Conge 
         Height          =   1215
         Left            =   120
         TabIndex        =   44
         Top             =   5160
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   2143
         GridLineMode    =   1
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   12648384
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
         BorderStyle     =   0
         DisableIcons    =   -1  'True
         DefaultRowHeight=   60
         MaxVisibleRows  =   0
      End
      Begin SToolBox.SGrid Grid_Repos 
         Height          =   1215
         Left            =   120
         TabIndex        =   39
         Top             =   5160
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   2143
         GridLineMode    =   1
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   12640511
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
         BorderStyle     =   0
         DisableIcons    =   -1  'True
         DefaultRowHeight=   60
         MaxVisibleRows  =   0
      End
      Begin SToolBox.SGrid Grid_Planning 
         Height          =   4935
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   8705
         GridLines       =   -1  'True
         GridLineMode    =   1
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   12632256
         GridLineColor   =   0
         GridFillLineColor=   0
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
         BorderStyle     =   0
         Editable        =   -1  'True
         DisableIcons    =   -1  'True
         DefaultRowHeight=   55
         MaxVisibleRows  =   0
      End
   End
   Begin VB.PictureBox PicBox_Date 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      ScaleHeight     =   375
      ScaleWidth      =   4935
      TabIndex        =   11
      Top             =   0
      Width           =   4935
      Begin MSComCtl2.DTPicker Cda_DebutPlg 
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   149094401
         CurrentDate     =   42865
      End
      Begin MSComCtl2.DTPicker Cda_FinPlg 
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   149094401
         CurrentDate     =   42865
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Du:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Au:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   0
         Picture         =   "Frm_PLANNING.frx":1DD5F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
   Begin SToolBox.SCommand Cmd_precd 
      Height          =   495
      Left            =   6240
      TabIndex        =   31
      Top             =   960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_PLANNING.frx":1E889
      BackColor       =   16777215
   End
   Begin VB.Image Pic_PlannigSemaine 
      Height          =   530
      Left            =   6960
      Picture         =   "Frm_PLANNING.frx":1F22B
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image CmdPrint 
      Height          =   495
      Left            =   13920
      Picture         =   "Frm_PLANNING.frx":2EC45
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Image Pic_Prochain 
      Height          =   495
      Left            =   9120
      Picture         =   "Frm_PLANNING.frx":3F847
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Image Pic_Ref 
      Height          =   470
      Left            =   4320
      Picture         =   "Frm_PLANNING.frx":50599
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Image Pic_Find 
      Height          =   495
      Left            =   11880
      Picture         =   "Frm_PLANNING.frx":5FF2F
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLANNING Semaine"
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
      TabIndex        =   10
      Top             =   360
      Width           =   3210
   End
   Begin VB.Image PicBox_Header 
      Height          =   1215
      Left            =   -120
      Picture         =   "Frm_PLANNING.frx":70B31
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   15735
   End
End
Attribute VB_Name = "Frm_PLANNING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim DCode           As String
    Dim CCode           As String
    Dim VeCode          As String
    Dim Date_Repos      As Date
    Dim DLibelle        As String
    Dim CLibelle        As String
    Dim VMatricule      As String
    Dim GridValid       As Boolean
    Dim Row             As Long
    Dim Col             As Long
    Dim ZRow            As Long
    Dim Edit            As Boolean
    Dim ColR            As Integer
    Dim RowR            As Integer
    Public ErreurABr    As Boolean
Private Sub Form_Load()
    PicBox_Planning.Visible = False
    Cmd_Ok.Visible = False
    Grid_Conge.Visible = False
    Img_Rep.Visible = False
    Call Initialise_SBICombo_Cond(Cbo_Conducteur)
    Call Initialise_SBICombo_Vehic(cbo_vehicule)
    Call Initgrid_Planning
    Call Initgrid_Repos
    Call Initgrid_Conge
    Call Initgrid_DetPlanning
    Call FindDest
    Call Pic_Mask_Click
    Cda_DebutPlg.Value = DatePlanning(Date)
    Cda_FinPlg.Value = DateWEnd(DatePlanning(Date))
    Call SearchPLANNING(DatePlanning(Date))
    MonthView.Value = Date
End Sub
Private Sub Form_Resize()
    Dim WidthForm           As Integer
    Dim ScreenWidth         As Integer
    Dim ScreenHeight        As Integer
    Dim ScreenResolution    As String
On Error Resume Next
    WidthForm = Me.Width
    ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    ScreenResolution = ScreenWidth & "x" & ScreenHeight
    PicBox_Header.Width = WidthForm
    CmdPrint.Left = WidthForm - 2300
    Pic_Find.Left = WidthForm - 4300
    Cmd_suivt.Left = WidthForm - 5100
    Pic_Prochain.Left = WidthForm - 7200
    Pic_PlannigSemaine.Left = WidthForm - 9350
    Cmd_precd.Left = WidthForm - 10100
    Pic_Ref.Left = WidthForm - 12100
    Pic_Planning.Width = WidthForm - 850
    Grid_Planning.Width = Pic_Planning.Width - 180
    Grid_Repos.Width = Pic_Planning.Width - 160
    Grid_Conge.Width = Pic_Planning.Width - 160
    Pic_Menu.Height = 15000
    If ScreenResolution = "1600x900" Then
        PicBox_Date.Left = WidthForm - 18000
        PicBox_Date.Top = 1100
    ElseIf ScreenResolution = "800x600" Then
        PicBox_Date.Left = WidthForm - 6000
        PicBox_Date.Top = 1
    ElseIf ScreenResolution = "1024x768" Then
        PicBox_Date.Left = WidthForm - 6000
        PicBox_Date.Top = 1
        Pic_Planning.Height = Me.Height - 2100
        Grid_Planning.Height = Pic_Planning.Height - 1800
        Grid_Repos.Top = Pic_Planning.Height - 1500
        Grid_Conge.Top = Pic_Planning.Height - 1500
    Else
        PicBox_Date.Left = 100
        PicBox_Date.Top = 1100
    End If
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

Private Sub Grid_Repos_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyF3 Then Call CmdPrint_Click
End Sub
Private Sub Img_Cog_Click()
    Img_Rep.Visible = True
    Img_Cog.Visible = False
    Grid_Conge.Visible = True
End Sub
Private Sub Img_Rep_Click()
    Img_Cog.Visible = True
    Img_Rep.Visible = False
    Grid_Conge.Visible = False
End Sub
Private Sub Cmd_Cog_Click()
    Cmd_Rep.Visible = True
    Cmd_Cog.Visible = False
    Grid_Conge.Visible = True
End Sub
Private Sub Cmd_Rep_Click()
    Cmd_Cog.Visible = True
    Cmd_Rep.Visible = False
    Grid_Conge.Visible = False
End Sub
Private Sub Pic_Consulter_Click()
    Call Pic_Show_Click
End Sub
Private Sub Pic_Conge_Click()
    Call Cmd_Conge_Click
End Sub
Private Sub Pic_Repos_Click()
    Call Cmd_repos_Click
End Sub
'Afficher et masquer menu date***
Private Sub Pic_Menu_DblClick()
    If Pic_Menu.Width = 490 Then Call Pic_Show_Click Else If Pic_Menu.Width = 2655 Then Call Pic_Mask_Click
End Sub
Private Sub Pic_Mask_Click()
    Pic_Menu.Width = 490
    MonthView.Visible = False
    Pic_Mask.Visible = False
    Pic_Show.Visible = True
    Cmd_Conge.Visible = False
    Cmd_Repos.Visible = False
    Lbl_Search.Visible = False
    Lbl_Title.Visible = False
    Pic_Planning.Left = 490
    Pic_Planning.Width = Me.Width - 850
    Grid_Planning.Width = Pic_Planning.Width
    Grid_Repos.Width = Pic_Planning.Width
    Pic_Conge.Visible = True
    Pic_Repos.Visible = True
    Pic_Consulter.Visible = True
    Img_Rep.Visible = True
    Img_Cog.Visible = True
    Cmd_Rep.Visible = False
    Cmd_Cog.Visible = False
End Sub
Private Sub Pic_Show_Click()
    Pic_Menu.Width = 2655
    MonthView.Visible = True
    Pic_Mask.Visible = True
    Pic_Show.Visible = False
    Cmd_Conge.Visible = True
    Cmd_Repos.Visible = True
    Lbl_Search.Visible = True
    Lbl_Title.Visible = True
    Pic_Planning.Left = 2655
    Pic_Planning.Width = Me.Width - 3000
    Grid_Planning.Width = Pic_Planning.Width
    Grid_Repos.Width = Pic_Planning.Width
    Pic_Conge.Visible = False
    Pic_Repos.Visible = False
    Pic_Consulter.Visible = False
    Img_Rep.Visible = False
    Img_Cog.Visible = False
    Cmd_Rep.Visible = True
    Cmd_Cog.Visible = True
End Sub
'Initialise SGrid***
Public Sub Initgrid_Planning()
    With Grid_Planning
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
        .AddColumn "Destination", "TOURNEE", , , 140, , True
        .AddColumn "Lundi", "LUNDI", , , 100
        .AddColumn "Mardi", "MARDI", , , 100
        .AddColumn "Mercredi", "MERCREDI", , , 100
        .AddColumn "Jeudi", "JEUDI", , , 100
        .AddColumn "Vendredi", "VENDREDI", , , 100
        .AddColumn "Samdi", "SAMEDI", , , 100
        .AddColumn "Dimanche", "DIMANCHE", , , 100
        .AddColumn "Temps", "Temps", , , 50, False
        .AddColumn "NULL", "", , , 0, False
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Public Sub Initgrid_Repos()
    With Grid_Repos
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
        .AddColumn "Destination", "Repos", , , 140, , True
        .AddColumn "Lundi", "LUNDI", , , 100
        .AddColumn "Mardi", "MARDI", , , 100
        .AddColumn "Mercredi", "MERCREDI", , , 100
        .AddColumn "Jeudi", "JEUDI", , , 100
        .AddColumn "Vendredi", "VENDREDI", , , 100
        .AddColumn "Samdi", "SAMEDI", , , 100
        .AddColumn "Dimanche", "DIMANCHE", , , 100
        .AddColumn "Temps", "Temps", , , 50, False
        .AddColumn "NULL", "", , , 0, False
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Public Sub Initgrid_DetPlanning()
    With Grid_DetPlanning
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
        .AddColumn "CodeDestination", "", , , 40, False
        .AddColumn "CodeConducteur", "", , , 40, False
        .AddColumn "Conducteur", "Conducteur", , , 280
        .AddColumn "CodeVehicule", "", , , 40, False
        .AddColumn "Vehicule", "Vehicule", , , 280
        .AddColumn "TypeRep", "Temps de Repos", , , 100
        .AddColumn "HeureEntre", "Heure Entre", , , 80
        .AddColumn "Code", "", , , 40, False
        .AddColumn "NULL", "", , , 0, True
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Public Sub Initgrid_Conge()
    With Grid_Conge
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
        .AddColumn "Destination", "Conge", , , 140, , True
        .AddColumn "Lundi", "LUNDI", , , 100
        .AddColumn "Mardi", "MARDI", , , 100
        .AddColumn "Mercredi", "MERCREDI", , , 100
        .AddColumn "Jeudi", "JEUDI", , , 100
        .AddColumn "Vendredi", "VENDREDI", , , 100
        .AddColumn "Samdi", "SAMEDI", , , 100
        .AddColumn "Dimanche", "DIMANCHE", , , 100
        .AddColumn "NULL", "", , , 0, False
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
'Tous les destinations***
Private Sub FindDest()
    Dim LObj_Find       As New DESTINATION
    Dim Lrs_Dest        As New Recordset
On Error GoTo Err
    Set Lrs_Dest = LObj_Find.GetAll_DestinationActifTourneeDisponibleExist(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    Grid_Planning.ClearRows
    Grid_Planning.Redraw = False
    While Not Lrs_Dest.EOF
        With Grid_Planning
            .AddRow
            .CellDetails .Rows, .ColumnIndex("Destination"), Lrs_Dest.Fields("Libelle"), , , &HE0E0E0
            .CellDetails .Rows, .ColumnIndex("Temps"), Lrs_Dest.Fields("Temps"), , , &H404040, &HFFFFFF
        End With
        Lrs_Dest.MoveNext
    Wend
    Grid_Planning.Redraw = True
    Grid_Planning.SelectedRow = 1
    Set Lrs_Dest = Nothing
    If Grid_Repos.Rows = 0 Then
        Grid_Repos.Redraw = False
        With Grid_Repos
            .AddRow
            .CellDetails .Rows, .ColumnIndex("Destination"), "Repos", , , &H808080, &HFFFFFF
            .CellText(.Rows, 10) = "Repot"
        End With
        Grid_Repos.Redraw = True
    End If
    If Grid_Conge.Rows = 0 Then
        Grid_Conge.Redraw = False
        With Grid_Conge
            .AddRow
            .CellDetails .Rows, .ColumnIndex("Destination"), "Conge", , , &H808080, &HFFFFFF
        End With
        Grid_Conge.Redraw = True
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'Vehicule***
Public Sub AfficheRowVehiculePing(ByVal ECode As String)
    Dim LObj_Find       As VEHICULE
    Dim Lrs_Vehicule    As Recordset
On Error GoTo Err
    Set LObj_Find = New VEHICULE
    Set Lrs_Vehicule = LObj_Find.GetVehiculeByCode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, ECode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Vehicule.EOF Then
        If Not IsNull(Lrs_Vehicule("Matricule")) Then
            cbo_vehicule.Text = Lrs_Vehicule("Code") & "  -  " & Lrs_Vehicule("Matricule")
            VeCode = Lrs_Vehicule("Code")
            VMatricule = Lrs_Vehicule("Matricule")
        Else
            MsgBox "Code introuvable", vbInformation
        End If
    Else
        MsgBox "Code introuvable", vbInformation
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Conducteur***
Public Sub AfficheRowConducteurPing(ByVal VeCode As String)
    Dim LObj_Find       As Conducteur
    Dim Lrs_Conducteur  As Recordset
On Error GoTo Err
    Set LObj_Find = New Conducteur
    Set Lrs_Conducteur = LObj_Find.GetRow_Conducteur_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VeCode, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Conducteur.EOF Then
        If Not IsNull(Lrs_Conducteur("Libelle")) Then
            Cbo_Conducteur.Text = Lrs_Conducteur("code") & "  -  " & Lrs_Conducteur("Libelle")
            CCode = Lrs_Conducteur("code")
            CLibelle = Lrs_Conducteur("Libelle")
        Else
            MsgBox "Code introuvable", vbInformation
        End If
    Else
        MsgBox "Code introuvable", vbInformation
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub cbo_Conducteur_Click()
    CCode = Cbo_Conducteur.FirstValue
    CLibelle = Cbo_Conducteur.SecondValue
End Sub
Private Sub cbo_vehicule_Click()
    VeCode = cbo_vehicule.FirstValue
    VMatricule = cbo_vehicule.SecondValue
End Sub
'Afficher Frm de statistiques de planning
Private Sub Pic_Find_Click()
On Error GoTo Err
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload FrmFind
    Unload FrmFind_BCarb
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "StiquePLNG"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub Cmd_Conducteur_Click()
On Error GoTo Err
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload FrmFind
    Unload FrmFind_BCarb
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "ConducteurPing"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub Cmd_Vehicule_Click()
On Error GoTo Err
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload FrmFind
    Unload FrmFind_BCarb
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "VehiculePing"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Consult les congés***
Private Sub Cmd_Conge_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "ConsultConge"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub Cmd_repos_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "ConducteurPLNG"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub cbo_conducteur_LostFocus()
    Call ExistDonnee(Cbo_Conducteur)
End Sub
Private Sub cbo_vehicule_LostFocus()
    If cbo_vehicule.Enabled = True Then Call ExistDonnee(cbo_vehicule)
End Sub
'Enabled | Disabled |-> ControlBox***
Private Sub EnbDisb(ByVal TYP As Boolean)
    Grid_Planning.Enabled = TYP
    Pic_Menu.Enabled = TYP
    Cmd_Edit.Enabled = TYP
    CmdDelete.Enabled = TYP
    Cmd_AddPlanning.Enabled = TYP
    CmdPrint.Enabled = TYP
    Pic_Prochain.Enabled = TYP
    Grid_Repos.Enabled = TYP
    Pic_Ref.Enabled = TYP
    Pic_PlannigSemaine.Enabled = TYP
    Pic_Prochain.Enabled = TYP
    Pic_Find.Enabled = TYP
    CmdPrint.Enabled = TYP
    Cmd_precd.Enabled = TYP
    Cmd_suivt.Enabled = TYP
    Cmd_Conge.Enabled = TYP
    Cmd_Repos.Enabled = TYP
    If Cda_DebutPlg.Value = DateNewPlanning(Date) Then
        Pic_Prochain.Enabled = False
        Cmd_suivt.Enabled = False
    End If
End Sub
Private Sub Chk_HEntre_Click()
    If Chk_HEntre.Value = 0 Then
        Txt_Heurentre.Text = ""
        Txt_Heurentre.Enabled = False
    ElseIf Chk_HEntre.Value = 1 Then
        Txt_Heurentre.Enabled = True
    End If
End Sub
'# Grid PLANNING***
Private Sub Grid_Planning_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    Dim LObj_Find       As New DESTINATION
    Dim Lrs_Destination As New Recordset
    Dim DESTINATION     As String
On Error GoTo Err
    DESTINATION = Grid_Planning.CellText(lRow, 1)
    Set Lrs_Destination = LObj_Find.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, DESTINATION, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Destination.EOF Then
        Cbo_Tournee.Caption = Lrs_Destination.Fields("Numero") & "  -  " & Lrs_Destination.Fields("Libelle")
        DCode = Lrs_Destination.Fields("Numero")
        DLibelle = Lrs_Destination.Fields("Libelle")
    Else
        Cbo_Tournee.Caption = "Repos"
    End If
    Set Lrs_Destination = Nothing
    With Grid_Planning
        Col = .SelectedCol
        Row = .SelectedRow
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub Grid_Planning_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyReturn Then
        With Grid_Planning
            Col = .SelectedCol
            Row = .SelectedRow
        End With
        Call Grid_Planning_DblClick(Row, Col)
    End If
    If KeyCode = vbKeyF3 Then Call CmdPrint_Click
End Sub
Private Sub Grid_Planning_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    Dim DateAu      As Date
On Error GoTo Err
    If (CHECK_ACCES("MAJ_PLING", LInt_UserId) = False) Then
        MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier PLANNING", vbExclamation
        Exit Sub
    End If
    DateAu = Cda_FinPlg.Value
    If DateAu < Date Then
        MsgBox " Planning déjà fait" & vbCr & " Modification refusée...", vbExclamation, App.ProductName
        Exit Sub
    End If
    With Grid_Planning
        Col = .SelectedCol
        Row = .SelectedRow
        Date_Repos = Format(DateCell(Cda_DebutPlg.Value, Col), "dd/mm/yyyy")
        Lbl_Journee.Caption = Format(DateCell(Cda_DebutPlg.Value, Col), " dddd - dd/mm/yyyy")
        If Col <> 1 And Col <> 9 Then
            If Row <> .Rows And Row <> .Rows - 1 Then
                Label5.Caption = "Véhicule"
                Grid_DetPlanning.ColumnWidth("Vehicule") = 230
                Grid_DetPlanning.ColumnWidth("Conducteur") = 230
                Grid_DetPlanning.ColumnWidth("TypeRep") = 0
                Grid_DetPlanning.ColumnWidth("HeureEntre") = 100
                Opt_Matin.Visible = False
                OptSoir.Visible = False
                cbo_vehicule.Visible = True
                Cmd_Vehicule.Visible = True
                Label6.Visible = False
                Label8.Visible = False
                Pic_HeureEntre.Visible = True
            Else
                Label5.Caption = ""
                Opt_Matin.Visible = False
                OptSoir.Visible = False
                cbo_vehicule.Visible = False
                Cmd_Vehicule.Visible = False
                Label6.Visible = False
                Label8.Visible = False
                Pic_HeureEntre.Visible = True
                Grid_DetPlanning.ColumnWidth("Vehicule") = 0
                Grid_DetPlanning.ColumnWidth("TypeRep") = 0
                Grid_DetPlanning.ColumnWidth("HeureEntre") = 100
                Grid_DetPlanning.ColumnWidth("Conducteur") = 460
            End If
            GridValid = True 'Tournee (false Repos)
            Grid_DetPlanning.ClearRows
            If .CellText(Row, Col) <> "" Then
                Call SelectCellule(Row, Col)
                Edit = True
            Else
                Edit = False
            End If
            Call EnbDisb(False)
            PicBox_Planning.Visible = True
            If Grid_DetPlanning.Rows > 0 Then Cmd_AddPlanning.Enabled = True
        End If
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub SelectCellule(ByVal NRow As Long, ByVal NCol As Long)
    Dim Lobj_PLANNING   As New PLANNING
    Dim Lobj_Conducteur As New Conducteur
    Dim Lobj_Vehicule   As New VEHICULE
    Dim Lrs_Find        As New Recordset
    Dim DateDu          As Date
    Dim jour            As String
    Dim XChp()          As String
    Dim XChps()         As String
    Dim CodeConducteur  As String
    Dim CodeVehicule    As String
    Dim XCount          As Integer
    Dim XCountS         As Integer
    Dim i               As Integer
On Error GoTo Err
    'Tournee***
    If GridValid = True Then
        With Grid_Planning
            XChp = Split(.CellText(NRow, NCol), vbCr)
            XCount = UBound(XChp)
            DateDu = Cda_DebutPlg.Value
            jour = Format(DateCell(Cda_DebutPlg.Value, NCol), "dddd")
            For i = 0 To XCount
                XChps = Split(XChp(i), "||")
                XCountS = UBound(XChps) + 1
                With Grid_DetPlanning
                    .AddRow
                    .CellDetails .Rows, .ColumnIndex("Conducteur"), RTrim(XChps(0)), , , &HC0C0C0
                    .CellDetails .Rows, .ColumnIndex("CodeDestination"), DCode, , , &HC0C0C0
                    If XCountS = 3 Then .CellDetails .Rows, .ColumnIndex("HeureEntre"), Trim(XChps(2)), , , &HC0C0C0 Else .CellDetails .Rows, .ColumnIndex("HeureEntre"), "", , , &HC0C0C0
                    Set Lrs_Find = Lobj_Conducteur.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, RTrim(XChps(0)), CNB)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                        Exit Sub
                    End If
                    Set Lobj_Conducteur = Nothing
                    If Not Lrs_Find.EOF Then
                        .CellDetails .Rows, .ColumnIndex("CodeConducteur"), Lrs_Find("Code"), , , &HC0C0C0
                        CodeConducteur = Lrs_Find("Code")
                    End If
                    Set Lrs_Find = Nothing
                    If Row = Grid_Planning.Rows Or Row = Grid_Planning.Rows - 1 Then
                        If XCountS = 2 Then .CellDetails .Rows, .ColumnIndex("HeureEntre"), Trim(XChps(1)), , , &HC0C0C0 Else .CellDetails .Rows, .ColumnIndex("HeureEntre"), "", , , &HC0C0C0
                    End If
                    If Row <> Grid_Planning.Rows And Row <> Grid_Planning.Rows - 1 Then
                        If XCountS = 3 Then .CellDetails .Rows, .ColumnIndex("HeureEntre"), Trim(XChps(2)), , , &HC0C0C0 Else .CellDetails .Rows, .ColumnIndex("HeureEntre"), "", , , &HC0C0C0
                        .CellDetails .Rows, .ColumnIndex("Vehicule"), Trim(XChps(1)), , , &HC0C0C0
                        Set Lrs_Find = Lobj_Vehicule.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Trim(XChps(1)))
                        If ErrNumber <> 0 Then
                            ErrNumber = 0
                            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                            Exit Sub
                        End If
                        Set Lobj_Vehicule = Nothing
                        If Not Lrs_Find.EOF Then
                            .CellDetails .Rows, .ColumnIndex("CodeVehicule"), Lrs_Find("Code"), , , &HC0C0C0
                            CodeVehicule = Lrs_Find("Code")
                        End If
                        Set Lrs_Find = Nothing
                    End If
                    Set Lrs_Find = New Recordset
                    Set Lrs_Find = Lobj_PLANNING.GetCode_PLANNING(ErrNumber, ErrDescription, ErrSourceDetail, DateDu, jour, DCode, CodeConducteur, CNB)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                        Exit Sub
                    End If
                    Set Lobj_PLANNING = Nothing
                    If Not Lrs_Find.EOF Then .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Find("Code"), , , &HC0C0C0 Else .CellBackColor(.Rows, 8) = &HC0C0C0
                    Set Lrs_Find = Nothing
                End With
            Next i
        End With
    'Repos***
    ElseIf GridValid = False Then
        Dim Lrs_Rep As New Recordset
        Dim LObj_Rep As New Conducteur

        XChp = Split(Grid_Repos.CellText(NRow, NCol), vbCr)
        XCount = UBound(XChp)
        DateDu = Cda_DebutPlg.Value
        jour = Format(DateCell(Cda_DebutPlg.Value, NCol), "dddd")
        For i = 0 To XCount
            XChps = Split(XChp(i), "||")
            XCountS = UBound(XChps) + 1
            With Grid_DetPlanning
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Conducteur"), RTrim(XChps(0)), , , &HC0C0C0
                .CellDetails .Rows, .ColumnIndex("TypeRep"), Trim(XChps(1)), , , &HC0C0C0
                Set Lrs_Find = Lobj_Conducteur.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, RTrim(XChps(0)), CNB)
                If ErrNumber <> 0 Then
                    ErrNumber = 0
                    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                    Exit Sub
                End If
                Set Lobj_Conducteur = Nothing
                If Not Lrs_Find.EOF Then
                    .CellDetails .Rows, .ColumnIndex("CodeConducteur"), Lrs_Find("Code"), , , &HC0C0C0
                    CodeConducteur = Lrs_Find("Code")
                End If
                Set Lrs_Find = Nothing

                Set Lrs_Rep = LObj_Rep.Get_ReposByCodeCondAndDate(ErrNumber, ErrDescription, ErrSourceDetail, CodeConducteur, DateCell(Cda_DebutPlg.Value, NCol), CNB)
                If ErrNumber <> 0 Then
                    ErrNumber = 0
                    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                    Exit Sub
                End If
                Set LObj_Rep = Nothing
                If Not Lrs_Rep.EOF Then .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Rep("numero"), , , &HC0C0C0 Else .CellBackColor(.Rows, 7) = &HC0C0C0
                Set Lrs_Rep = Nothing
            End With
        Next i
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'# Grid_Repos
Private Sub Grid_Repos_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    With Grid_Repos
        ColR = .SelectedCol
        RowR = .SelectedRow
    End With
End Sub
Private Sub Grid_Repos_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    Dim DateAu          As Date
On Error GoTo Err
    'Acces USER***
    '--------------
    If (CHECK_ACCES("MAJ_PLING", LInt_UserId) = False) Then
        MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier PLANNING", vbExclamation
        Exit Sub
    End If
    DateAu = Cda_FinPlg.Value
    If DateAu < Date Then
        MsgBox " Planning déjà fait" & vbCr & " Modification refusée...", vbExclamation, App.ProductName
        Exit Sub
    End If
    Cmd_Repos.Visible = True
     With Grid_Repos
        ColR = .SelectedCol
        RowR = .SelectedRow
        Date_Repos = Format(DateCell(Cda_DebutPlg.Value, ColR), "dd/mm/yyyy")
        Lbl_Journee.Caption = Format(DateCell(Cda_DebutPlg.Value, ColR), " dddd - dd/mm/yyyy")
        If ColR <> 1 And ColR <> 9 Then
            Label5.Caption = "Temps - Repot"
            Grid_DetPlanning.ColumnWidth("Vehicule") = 0
            Grid_DetPlanning.ColumnWidth("HeureEntre") = 0
            Grid_DetPlanning.ColumnWidth("TypeRep") = 100
            Grid_DetPlanning.ColumnWidth("Conducteur") = 460
            Opt_Matin.Visible = True
            OptSoir.Visible = True
            cbo_vehicule.Visible = False
            Cmd_Vehicule.Visible = False
            Label6.Visible = True
            Label8.Visible = True
            Pic_HeureEntre.Visible = False
            Grid_DetPlanning.ClearRows
            GridValid = False 'Repos (true Tournee)
            If .CellText(RowR, ColR) <> "" Then
                Call SelectCellule(RowR, ColR)
                Edit = True
            Else
                Edit = False
            End If
            Call EnbDisb(False)
            Call Pic_Mask_Click
            PicBox_Planning.Visible = True
            If Grid_DetPlanning.Rows > 0 Then Cmd_AddPlanning.Enabled = True
        End If
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'# Grid Detail PLANNING***
Private Sub Grid_DetPlanning_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    ZRow = lRow
    Cmd_Edit.Enabled = True
    CmdDelete.Enabled = True
End Sub
'Ajouter Nouveau Detail Planning***
'==========================
Private Sub Cmd_AddNew_Click()
    Dim Msg             As VbMsgBoxResult
    Dim LObj_Find       As New Conducteur
    Dim Lrs_Find        As New Recordset
    Dim i               As Integer
    Dim J               As Integer
    Dim XCount          As Integer
    Dim XCountS         As Integer
    Dim XRow            As Integer
    Dim XChp()          As String
    Dim XChps()         As String
    Dim Conducteur      As String
    Dim VEHICULE        As String
    Dim CHeureEntre     As String
    Dim cond            As String
    Dim Vehic           As String
    Dim DateT           As String
    Dim Text            As String
    Dim TextV           As String
    Dim xDate           As Date
On Error GoTo Err
    With Grid_Planning
    '# Tournee
        If GridValid = True Then
            If Cbo_Conducteur.ListIndex = 0 Or Cbo_Conducteur.Text = "" Then
                MsgBox "Séléctionner... 'Conducteur'...", vbExclamation, App.ProductName
                Exit Sub
            End If
            If Row <> .Rows And Row <> .Rows - 1 Then
                If cbo_vehicule.ListIndex = 0 Or cbo_vehicule.Text = "" Then
                    MsgBox "Séléctionner... 'Véhicule'..", vbExclamation, App.ProductName
                    Exit Sub
                End If
            End If
            If Chk_HEntre.Value = 0 Then
                Txt_Heurentre.Text = ""
            ElseIf Chk_HEntre.Value = 1 Then
                If Trim(Txt_Heurentre.Text) = "" Or Txt_Heurentre.Text = "__:__:__" Or Txt_Heurentre.Text = "00:00:00" Then
                    MsgBox "Saisie Heure d'Entre...      ", vbInformation, App.ProductName
                    Exit Sub
                Else
                    CHeureEntre = Mid(Replace(Txt_Heurentre.Text, "_", "0"), 1, 5)
                End If
            End If
            xDate = DateCell(Cda_DebutPlg.Value, Col)
            cond = Cbo_Conducteur.SecondValue
            Vehic = cbo_vehicule.SecondValue
            If Row = .Rows Or Row = .Rows - 1 Then Vehic = ""
            Text = ""
            TextV = ""
            '# Si tournée choisit pour toute une journée
            If .CellText(Row, 9) = "Journée" Then
                For i = 1 To .Rows
                    If (.CellText(i, Col) <> "" And .CellText(i, Col) <> " ") Then
                        XChp = Split(.CellText(i, Col), vbCr)
                        XCount = UBound(XChp)
                        For J = 0 To XCount
                            XChps = Split(XChp(J), "||")
                            Conducteur = RTrim(XChps(0))
                            '# Msg Pour Conducteur***
                            If Row <> i Then
                                If RTrim(cond) = Conducteur Then
                                    If Text = "" Then
                                        Text = Conducteur & " est affecté pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                    Else
                                        Text = Text & vbCr & " - " & .CellText(i, 1)
                                    End If
                                End If
                            End If
                            '# tournée <> de bizerte donc tester si aussi véhicule est affecté à une autre tournée
                            If i <= .Rows - 2 Then
                                VEHICULE = Trim(XChps(1))
                                '# Msg Pour Vehicule***
                                If Trim(Vehic) = VEHICULE Then
                                    If TextV = "" Then
                                        TextV = Vehic & " est affectée pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                    Else
                                        TextV = TextV & vbCr & " - " & .CellText(i, 1)
                                    End If
                                End If
                            End If
                        Next J
                    End If
                Next i
            Else
                For i = 1 To .Rows - 2
                    If (.CellText(i, Col) <> "" And .CellText(i, Col) <> " ") Then
                        '# Cellule choisit est bizerte soir***
                        If Row = .Rows Then
                            XChp = Split(.CellText(i, Col), vbCr)
                            XCount = UBound(XChp)
                            For J = 0 To XCount
                                XChps = Split(XChp(J), "||")
                                Conducteur = RTrim(XChps(0))
                                '# Msg Pour Conducteur***
                                If RTrim(cond) = Conducteur Then
                                    If (.CellText(i, 9) = "Après midi" Or .CellText(i, 9) = "Journée") Then
                                        If Text = "" Then
                                            Text = Conducteur & " est affecté pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                        Else
                                            Text = Text & vbCr & " - " & .CellText(i, 1)
                                        End If
                                    End If
                                End If
                            Next J
                        '# Cellule choisit bizerte matin***
                        ElseIf Row = .Rows - 1 Then
                            XChp = Split(.CellText(i, Col), vbCr)
                            XCount = UBound(XChp)
                            For J = 0 To XCount
                                XChps = Split(XChp(J), "||")
                                Conducteur = RTrim(XChps(0))
                                '# Msg Pour Conducteur***
                                If RTrim(cond) = Conducteur Then
                                    If (.CellText(i, 9) = "Matin") Or (.CellText(i, 9) = "Journée") Then
                                        If Text = "" Then
                                            Text = Conducteur & " est affecté pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                        Else
                                            Text = Text & vbCr & " - " & .CellText(i, 1)
                                        End If
                                    End If
                                End If
                            Next J
                        '# Cellule choisit tournée <> bizerte***
                        ElseIf .CellText(Row, 9) <> "Journée" And Row <> .Rows And Row <> .Rows - 1 Then
                            XChp = Split(.CellText(i, Col), vbCr)
                            XCount = UBound(XChp)
                            For J = 0 To XCount
                                XChps = Split(XChp(J), "||")
                                Conducteur = RTrim(XChps(0))
                                '# Msg Pour Conducteur***
                                If Row <> i Then
                                    If RTrim(cond) = Conducteur Then
                                        If Text = "" Then
                                            Text = Conducteur & " est affecté pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                        Else
                                            Text = Text & vbCr & " - " & .CellText(i, 1)
                                        End If
                                    End If
                                End If
                                VEHICULE = Trim(XChps(1))
                                '# Msg Pour Vehicule***
                                If Trim(Vehic) = VEHICULE And .CellText(i, 9) = "Journée" Then
                                    If TextV = "" Then
                                        TextV = Vehic & " est affectée pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                    Else
                                        TextV = TextV & vbCr & " - " & .CellText(i, 1)
                                    End If
                                End If
                            Next J
                        End If
                    End If
                Next i
                If .CellText(Row, 9) = "Matin" Then
                    If .CellText(.Rows - 1, Col) <> "" Then
                        XChp = Split(.CellText(.Rows - 1, Col), vbCr)
                        XCount = UBound(XChp)
                        For i = 0 To XCount
                            XChps = Split(XChp(i), "||")
                            XCountS = UBound(XChps) + 1
                            '//Msg Pour Conducteur***
                            If Row <> .Rows - 1 Then
                                If RTrim(CLibelle) = RTrim(XChps(0)) Then
                                    Msg = MsgBox(Text & vbCr & "Verifier Champs 'Bizerte Matin' puis ajouter un programme.", vbExclamation + vbYesNo + vbDefaultButton2, App.ProductName)
                                    If Msg = vbNo Then Exit Sub
                                    If Text = "" Then
                                        Text = Conducteur & " est affecté pour tournée à : " & vbCr & " - " & .CellText(.Rows - 1, 1)
                                    Else
                                        Text = Text & vbCr & " - " & .CellText(.Rows - 1, 1)
                                    End If
                                End If
                            End If
                        Next i
                    End If
                ElseIf .CellText(Row, 9) = "Après midi" Then
                    If .CellText(.Rows, Col) <> "" Then
                        XChp = Split(.CellText(.Rows, Col), vbCr)
                        XCount = UBound(XChp)
                        For i = 0 To XCount
                            XChps = Split(XChp(i), "||")
                            XCountS = UBound(XChps) + 1
                            '# Msg Pour Conducteur***
                            If RTrim(CLibelle) = RTrim(XChps(0)) Then
                                Msg = MsgBox(Text & vbCr & " Verifier Champ 'Bizerte Soir' puis ajouter un programme.", vbExclamation + vbYesNo + vbDefaultButton2, App.ProductName)
                                If Msg = vbNo Then Exit Sub
                                If Text = "" Then
                                    Text = Conducteur & " est affecté pour tourneé à : " & vbCr & " - " & .CellText(.Rows, 1)
                                Else
                                    Text = Text & vbCr & " - " & .CellText(.Rows, 1)
                                End If
                            End If
                        Next i
                    End If
                End If
            End If
            '# Afficher résultat du parcours du grid : les tournée auxquelles le conducteur ou véhicule sont affécté***
            If Text <> "" Or TextV <> "" Then
                If TextV <> "" Then
                    Msg = MsgBox(Text & vbCr & TextV & vbCr & " voulez-vous confirmez ?", vbInformation + vbYesNo, App.ProductName)
                    If Msg = vbNo Then Exit Sub
                Else
                    Msg = MsgBox(Text & vbCr & " voulez-vous confirmez ?", vbInformation + vbYesNo, App.ProductName)
                    If Msg = vbNo Then Exit Sub
                End If
            End If
            '# Parcours du grid_repos***
            If Grid_Repos.CellText(1, Col) <> "" Then
                XChp = Split(Grid_Repos.CellText(1, Col), vbCr)
                XCount = UBound(XChp)
                For i = 0 To XCount
                    XChps = Split(XChp(i), "||")
                    XCountS = UBound(XChps) + 1
                    If RTrim(CLibelle) = RTrim(XChps(0)) And (.CellText(Row, 9) = Trim(XChps(1)) Or .CellText(Row, 9) = "Journée") Then
                        MsgBox "Conducteur en Repos!... Vérifier Champs Repos puis ajouter un programme.", vbExclamation, App.ProductName
                        Exit Sub
                    End If
                Next i
            End If
            Set Lrs_Find = LObj_Find.GetRow_CongeByCodeCondAndDate(ErrNumber, ErrDescription, ErrSourceDetail, CCode, xDate, CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            Set LObj_Find = Nothing
            If Row <> .Rows And Row <> .Rows - 1 Then
                If Lrs_Find.EOF Then
                    With Grid_DetPlanning
                        XRow = .Rows + 1
                        If .Rows = 0 Then
                            .Redraw = False
                            .AddRow
                            .CellDetails XRow, .ColumnIndex("CodeDestination"), DCode
                            .CellDetails XRow, .ColumnIndex("CodeConducteur"), CCode
                            .CellDetails XRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                            .CellDetails XRow, .ColumnIndex("CodeVehicule"), VeCode
                            .CellDetails XRow, .ColumnIndex("Vehicule"), VMatricule, , , &HC0C0C0
                            .CellDetails XRow, .ColumnIndex("HeureEntre"), CHeureEntre, , , &HC0C0C0
                            .Redraw = True
                        ElseIf .Rows > 0 Then
                            For i = 1 To .Rows
                                If RTrim(Cbo_Conducteur.Text) = .CellText(i, 2) & "  -  " & .CellText(i, 3) And RTrim(cbo_vehicule.Text) = .CellText(i, 4) & "  -  " & .CellText(i, 5) Then
                                    MsgBox "Vérifier 'Conducteur' et 'Véhicule' exist en planning!..      ", vbExclamation, App.ProductName
                                    Exit Sub
                                ElseIf RTrim(Cbo_Conducteur.Text) = .CellText(i, 2) & "  -  " & .CellText(i, 3) And RTrim(cbo_vehicule.Text) <> .CellText(i, 4) & "  -  " & .CellText(i, 5) Then
                                    Msg = MsgBox("'Conducteur' exist en planning!..     ", vbExclamation, App.ProductName)
                                    Exit Sub
                                ElseIf RTrim(Cbo_Conducteur.Text) <> .CellText(i, 2) & "  -  " & .CellText(i, 3) And RTrim(cbo_vehicule.Text) = .CellText(i, 4) & "  -  " & .CellText(i, 5) Then
                                    Msg = MsgBox("'Vehicule' exist en planning!.." & vbCr & vbCr & " Voulez-vous confirmer", vbInformation + vbYesNo, App.ProductName)
                                    If Msg = vbNo Then Exit Sub Else Exit For
                                End If
                            Next i
                            .Redraw = False
                            .AddRow
                            .CellDetails XRow, .ColumnIndex("CodeDestination"), DCode
                            .CellDetails XRow, .ColumnIndex("CodeConducteur"), CCode
                            .CellDetails XRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                            .CellDetails XRow, .ColumnIndex("CodeVehicule"), VeCode
                            .CellDetails XRow, .ColumnIndex("Vehicule"), VMatricule, , , &HC0C0C0
                            .CellDetails XRow, .ColumnIndex("HeureEntre"), CHeureEntre, , , &HC0C0C0
                            .Redraw = True
                        End If
                        cbo_vehicule.ListIndex = 0
                        Cbo_Conducteur.ListIndex = 0
                    End With
                    If Grid_DetPlanning.Rows > 0 Then Cmd_AddPlanning.Enabled = True
                Else
                    MsgBox "Conducteur : " & Lrs_Find.Fields("libelle") & ", en congé le : " & xDate, vbExclamation, App.ProductName
                    cbo_vehicule.ListIndex = 0
                    Cbo_Conducteur.ListIndex = 0
                    Exit Sub
                End If
            Else
                If Lrs_Find.EOF Then
                    With Grid_DetPlanning
                        XRow = .Rows + 1
                        If .Rows = 0 Then
                            .Redraw = False
                            .AddRow
                            .CellDetails XRow, .ColumnIndex("CodeDestination"), DCode
                            .CellDetails XRow, .ColumnIndex("CodeConducteur"), CCode
                            .CellDetails XRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                            .CellDetails XRow, .ColumnIndex("HeureEntre"), CHeureEntre, , , &HC0C0C0
                            .Redraw = True
                        ElseIf .Rows > 0 Then
                            For i = 1 To .Rows
                                If Trim(Cbo_Conducteur.Text) = .CellText(i, 2) & "  -  " & .CellText(i, 3) Then
                                    MsgBox "Vérifier 'Conducteur' exist en planning!..", vbExclamation, App.ProductName
                                    Exit Sub
                                End If
                            Next i
                            .Redraw = False
                            .AddRow
                            .CellDetails XRow, .ColumnIndex("CodeDestination"), DCode
                            .CellDetails XRow, .ColumnIndex("CodeConducteur"), CCode
                            .CellDetails XRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                            .CellDetails XRow, .ColumnIndex("HeureEntre"), CHeureEntre, , , &HC0C0C0
                            .Redraw = True
                        End If
                        Cbo_Conducteur.ListIndex = 0
                    End With
                    If Grid_DetPlanning.Rows > 0 Then Cmd_AddPlanning.Enabled = True
                Else
                    MsgBox "Conducteur : " & Lrs_Find.Fields("libelle") & ", en congé le : " & xDate, vbExclamation, App.ProductName
                    cbo_vehicule.ListIndex = 0
                    Cbo_Conducteur.ListIndex = 0
                    Exit Sub
                End If
                Set Lrs_Find = Nothing
            End If
    '# Repos***
        ElseIf GridValid = False Then
            Dim Temps       As String
            cond = Cbo_Conducteur.SecondValue
            xDate = DateCell(Cda_DebutPlg.Value, ColR)
            If Cbo_Conducteur.ListIndex = 0 Or Cbo_Conducteur.Text = "" Then
                MsgBox "Séléctionner... 'Conducteur'.", vbExclamation, App.ProductName
                Exit Sub
            End If
            If Opt_Matin.Value = vbGrayed And OptSoir.Value = vbGrayed Then
                MsgBox "Choisir le temps de Repos Matin ou Après midi!..." & vbCr & vbCr & "Cocher le temps de repos(Matin - Après midi)", vbExclamation, App.ProductName
                Exit Sub
            End If
            If Opt_Matin.Value = vbChecked Then
                Temps = "Matin"
            ElseIf OptSoir.Value = vbChecked Then
                Temps = "Après midi"
            End If
            For i = 1 To .Rows
                If (.CellText(i, ColR) <> "" And .CellText(i, ColR) <> " ") Then
                    XChp = Split(.CellText(i, ColR), vbCr)
                    XCount = UBound(XChp)
                    For J = 0 To XCount
                        XChps = Split(XChp(J), "||")
                        Conducteur = RTrim(XChps(0))
                        If RTrim(cond) = Conducteur Then
                            If (Temps = .CellText(i, 9) Or .CellText(i, 9) = "Journée") Then
                                MsgBox Conducteur & " Déja affectée pour tournée de " & .CellText(i, 1), vbExclamation, App.ProductName
                                Exit Sub
                            End If
                        End If
                    Next J
                End If
            Next i
            Set Lrs_Find = LObj_Find.GetRow_CongeByCodeCondAndDate(ErrNumber, ErrDescription, ErrSourceDetail, CCode, xDate, CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            Set LObj_Find = Nothing
            If Lrs_Find.EOF Then
                With Grid_DetPlanning
                    XRow = .Rows + 1
                    If .Rows = 0 Then
                        .Redraw = False
                        .AddRow
                        .CellDetails XRow, .ColumnIndex("CodeConducteur"), CCode
                        .CellDetails XRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                        If Opt_Matin.Value = vbChecked Then
                            .CellDetails XRow, .ColumnIndex("TypeRep"), "Matin", , , &HC0C0C0
                        ElseIf OptSoir.Value = vbChecked Then
                            .CellDetails XRow, .ColumnIndex("TypeRep"), "Après midi", , , &HC0C0C0
                        End If
                        .Redraw = True
                    ElseIf .Rows > 0 Then
                        For i = 1 To .Rows
                            If RTrim(Cbo_Conducteur.Text) = .CellText(i, 2) & "  -  " & .CellText(i, 3) Then
                                MsgBox "Vérifier " & .CellText(i, 3) & " exist déja en Repos!..", vbExclamation, App.ProductName
                                Exit Sub
                            End If
                        Next i
                        .Redraw = False
                        .AddRow
                        .CellDetails XRow, .ColumnIndex("CodeConducteur"), CCode
                        .CellDetails XRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                        If Opt_Matin.Value = vbChecked Then
                            .CellDetails XRow, .ColumnIndex("TypeRep"), "Matin", , , &HC0C0C0
                        ElseIf OptSoir.Value = vbChecked Then
                            .CellDetails XRow, .ColumnIndex("TypeRep"), "Après midi", , , &HC0C0C0
                        End If
                        .Redraw = True
                    End If
                    cbo_vehicule.ListIndex = 0
                    Cbo_Conducteur.ListIndex = 0
                End With
                If Grid_DetPlanning.Rows > 0 Then Cmd_AddPlanning.Enabled = True
            Else
                MsgBox "Conducteur : " & Lrs_Find.Fields("libelle") & ", en congé le : " & xDate, vbExclamation, App.ProductName
                cbo_vehicule.ListIndex = 0
                Cbo_Conducteur.ListIndex = 0
                Exit Sub
            End If
            Set Lrs_Find = Nothing
        End If
    End With
    Chk_HEntre.Value = 0
    Txt_Heurentre.Text = ""
    Txt_Heurentre.Enabled = False
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Modification Detail PLANNING***
Private Sub Cmd_ok_Click()
    Dim LObj_Find           As New Conducteur
    Dim Lrs_Find            As New Recordset
    Dim Msg                 As VbMsgBoxResult
    Dim cond                As String
    Dim Vehic               As String
    Dim Text                As String
    Dim TextV               As String
    Dim XChp()              As String
    Dim XChps()             As String
    Dim Conducteur          As String
    Dim VEHICULE            As String
    Dim i                   As Integer
    Dim J                   As Integer
    Dim XCount              As Integer
    Dim XCountS             As Integer
    Dim xDate               As Date
    Dim CHeureEntre         As String
On Error GoTo Err
    With Grid_Planning
    '# Tournee
        If GridValid = True Then
            If Cbo_Conducteur.ListIndex = 0 Then
                MsgBox "Séléctionner... 'Conducteur'...", vbExclamation, App.ProductName
                Exit Sub
            End If
            If Row <> .Rows And Row <> .Rows - 1 Then
                If cbo_vehicule.ListIndex = 0 Then
                    MsgBox "Séléctionner... 'Véhicule'..", vbExclamation, App.ProductName
                    Exit Sub
                End If
            End If
            If Chk_HEntre.Value = 0 Then
                Txt_Heurentre.Text = ""
            ElseIf Chk_HEntre.Value = 1 Then
                If Trim(Txt_Heurentre.Text) = "" Or Txt_Heurentre.Text = "__:__:__" Or Txt_Heurentre.Text = "00:00:00" Then
                    MsgBox "Saisie Heure d'Entre...      ", vbInformation, App.ProductName
                    Exit Sub
                Else
                    CHeureEntre = Mid(Replace(Txt_Heurentre.Text, "_", "0"), 1, 5)
                End If
            End If
            cond = Cbo_Conducteur.SecondValue
            Vehic = cbo_vehicule.SecondValue
            xDate = DateCell(Cda_DebutPlg.Value, Col)
            If Row = .Rows Or Row = .Rows - 1 Then Vehic = ""
            Text = ""
            TextV = ""
            '# Si tournée choisit pour toute une journée
            If .CellText(Row, 9) = "Journée" Then
                For i = 1 To .Rows
                    If (.CellText(i, Col) <> "" And .CellText(i, Col) <> " ") Then
                        XChp = Split(.CellText(i, Col), vbCr)
                        XCount = UBound(XChp)
                        For J = 0 To XCount
                            XChps = Split(XChp(J), "||")
                            Conducteur = RTrim(XChps(0))
                            '# Msg Pour Conducteur***
                            If RTrim(cond) = Conducteur Then
                                If Text = "" Then
                                    Text = RTrim(cond) & " est affecté pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                Else
                                    Text = Text & vbCr & " - " & .CellText(i, 1)
                                End If
                            End If
                            '# Si tournée <> de bizerte donc tester si aussi véhicule est affecté à une autre tournée
                            If i <= .Rows - 2 Then
                                VEHICULE = Trim(XChps(1))
                                '//Msg Pour Vehicule***
                                If Trim(Vehic) = VEHICULE Then
                                    If TextV = "" Then
                                        TextV = Vehic & " est affectée pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                    Else
                                        TextV = TextV & vbCr & " - " & .CellText(i, 1)
                                    End If
                                End If
                            End If
                        Next J
                    End If
                Next i
            Else
                For i = 1 To .Rows - 2
                    If (.CellText(i, Col) <> "" And .CellText(i, Col) <> " ") Then
                        '# Cellule choisit est bizerte soir***
                        If Row = .Rows Then
                            XChp = Split(.CellText(i, Col), vbCr)
                            XCount = UBound(XChp)
                            For J = 0 To XCount
                                XChps = Split(XChp(J), "||")
                                Conducteur = RTrim(XChps(0))
                                '# Msg Pour Conducteur***
                                If RTrim(cond) = Conducteur Then
                                    If (.CellText(i, 9) = "Après midi" Or .CellText(i, 9) = "Journée") Then
                                        If Text = "" Then
                                            Text = RTrim(cond) & " est affecté pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                        Else
                                            Text = Text & vbCr & " - " & .CellText(i, 1)
                                        End If
                                    End If
                                End If
                            Next J
                        '# Cellule choisit bizerte matin***
                        ElseIf Row = .Rows - 1 Then
                            XChp = Split(.CellText(i, Col), vbCr)
                            XCount = UBound(XChp)
                            For J = 0 To XCount
                                XChps = Split(XChp(J), "||")
                                Conducteur = RTrim(XChps(0))
                                '# Msg Pour Conducteur***
                                If RTrim(cond) = Conducteur Then
                                    If (.CellText(i, 9) = "Matin") Or (.CellText(i, 9) = "Journée") Then
                                        If Text = "" Then
                                            Text = RTrim(cond) & " est affecté pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                        Else
                                            Text = Text & vbCr & " - " & .CellText(i, 1)
                                        End If
                                    End If
                                End If
                            Next J
                        '# Cellule choisit tournée <> bizerte***
                        ElseIf .CellText(Row, 9) <> "Journée" And Row <> .Rows And Row <> .Rows - 1 Then
                            XChp = Split(.CellText(i, Col), vbCr)
                            XCount = UBound(XChp)
                            For J = 0 To XCount
                                XChps = Split(XChp(J), "||")
                                Conducteur = RTrim(XChps(0))
                                '# Msg Pour Conducteur***
                                If RTrim(cond) = Conducteur Then
                                    If Text = "" Then
                                        Text = Conducteur & " est affecté pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                    Else
                                        Text = Text & vbCr & " - " & .CellText(i, 1)
                                    End If
                                End If
                                VEHICULE = Trim(XChps(1))
                                '# Msg Pour Vehicule***
                                If Trim(Vehic) = VEHICULE And .CellText(i, 9) = "Journée" Then
                                    If TextV = "" Then
                                        TextV = Vehic & " est affectée pour tournée à : " & vbCr & " - " & .CellText(i, 1)
                                    Else
                                        TextV = TextV & vbCr & " - " & .CellText(i, 1)
                                    End If
                                End If
                            Next J
                        End If
                    End If
                Next i
                If .CellText(Row, 9) = "Matin" Then
                    If .CellText(.Rows - 1, Col) <> "" Then
                        XChp = Split(.CellText(.Rows - 1, Col), vbCr)
                        XCount = UBound(XChp)
                        For i = 0 To XCount
                            XChps = Split(XChp(i), "||")
                            XCountS = UBound(XChps) + 1
                            '# Msg Pour Conducteur***
                            If RTrim(CLibelle) = RTrim(XChps(0)) Then
                                Msg = MsgBox(Text & vbCr & "Verifier Champs 'Bizerte Matin' puis ajouter un programme.", vbExclamation + vbYesNo + vbDefaultButton2, App.ProductName)
                                If Msg = vbNo Then Exit Sub
                                If Text = "" Then
                                    Text = RTrim(CLibelle) & " est affecté pour tournée à : " & vbCr & " - " & .CellText(.Rows - 1, 1)
                                Else
                                    Text = Text & vbCr & " - " & .CellText(.Rows - 1, 1)
                                End If
                            End If
                        Next i
                    End If
                ElseIf .CellText(Row, 9) = "Après midi" Then
                    If .CellText(.Rows, Col) <> "" Then
                        XChp = Split(.CellText(.Rows, Col), vbCr)
                        XCount = UBound(XChp)
                        For i = 0 To XCount
                            XChps = Split(XChp(i), "||")
                            XCountS = UBound(XChps) + 1
                            '# Msg Pour Conducteur***
                            If RTrim(CLibelle) = RTrim(XChps(0)) Then
                                Msg = MsgBox(Text & vbCr & " Verifier Champ 'Bizerte Soir' puis ajouter un programme.", vbExclamation + vbYesNo + vbDefaultButton2, App.ProductName)
                                If Msg = vbNo Then Exit Sub
                                If Text = "" Then
                                    Text = RTrim(CLibelle) & " est affecté pour tourneé à : " & vbCr & " - " & .CellText(.Rows, 1)
                                Else
                                    Text = Text & vbCr & " - " & .CellText(.Rows, 1)
                                End If
                            End If
                        Next i
                    End If
                End If
            End If
            '# Afficher résultat du parcours du grid : les tournée auxquelles le conducteur ou véhicule sont affécté***
            If Text <> "" Or TextV <> "" Then
                If TextV <> "" Then
                    Msg = MsgBox(Text & vbCr & TextV & vbCr & " voulez-vous confirmez ?", vbInformation + vbYesNo, App.ProductName)
                    If Msg = vbNo Then Exit Sub
                Else
                    Msg = MsgBox(Text & vbCr & " voulez-vous confirmez ?", vbInformation + vbYesNo, App.ProductName)
                    If Msg = vbNo Then Exit Sub
                End If
            End If
            '# Parcours du grid_repos***
            If Grid_Repos.CellText(1, Col) <> "" Then
                XChp = Split(Grid_Repos.CellText(1, Col), vbCr)
                XCount = UBound(XChp)
                For i = 0 To XCount
                    XChps = Split(XChp(i), "||")
                    XCountS = UBound(XChps) + 1
                    If RTrim(CLibelle) = RTrim(XChps(0)) And (.CellText(Row, 9) = Trim(XChps(1)) Or .CellText(Row, 9) = "Journée") Then
                        MsgBox "Conducteur en Repos!... Vérifier Champs Repos puis ajouter un programme.", vbExclamation, App.ProductName
                        Exit Sub
                    End If
                Next i
            End If
            Set Lrs_Find = LObj_Find.GetRow_CongeByCodeCondAndDate(ErrNumber, ErrDescription, ErrSourceDetail, CCode, xDate, CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            Set LObj_Find = Nothing
            If Row <> .Rows And Row <> .Rows - 1 Then
                If Lrs_Find.EOF Then
                    With Grid_DetPlanning
                        If .Rows = 0 Then
                            .Redraw = False
                            .AddRow
                            .CellDetails ZRow, .ColumnIndex("CodeConducteur"), CCode
                            .CellDetails ZRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                            .CellDetails ZRow, .ColumnIndex("CodeVehicule"), VeCode
                            .CellDetails ZRow, .ColumnIndex("Vehicule"), VMatricule, , , &HC0C0C0
                            .CellDetails ZRow, .ColumnIndex("HeureEntre"), CHeureEntre, , , &HC0C0C0
                            .Redraw = True
                        ElseIf .Rows > 0 Then
                            For i = 1 To .Rows
                                If i <> ZRow Then
                                    If Cbo_Conducteur.Text = .CellText(i, 2) & "  -  " & .CellText(i, 3) And cbo_vehicule.Text = .CellText(i, 4) & "  -  " & .CellText(i, 5) Then
                                        MsgBox "Vérifier 'Conducteur' et 'Véhicule' exist déja en planning!..", vbExclamation, App.ProductName
                                        Exit Sub
                                    ElseIf Cbo_Conducteur.Text = .CellText(i, 2) & "  -  " & .CellText(i, 3) And cbo_vehicule.Text <> .CellText(i, 4) & "  -  " & .CellText(i, 5) Then
                                        Msg = MsgBox("'Conducteur' exist déja en planning!.." & vbCr & vbCr & " Voulez-vous confirmer", vbInformation + vbYesNo, App.ProductName)
                                        If Msg = vbNo Then Exit Sub Else Exit For
                                    End If
                                End If
                            Next i
                            .Redraw = False
                            .CellDetails ZRow, .ColumnIndex("CodeConducteur"), CCode
                            .CellDetails ZRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                            .CellDetails ZRow, .ColumnIndex("CodeVehicule"), VeCode
                            .CellDetails ZRow, .ColumnIndex("Vehicule"), VMatricule, , , &HC0C0C0
                            .CellDetails ZRow, .ColumnIndex("HeureEntre"), CHeureEntre, , , &HC0C0C0
                            .Redraw = True
                        End If
                    End With
                    If Grid_DetPlanning.Rows > 0 Then Cmd_AddPlanning.Enabled = True
                Else
                    MsgBox "Conducteur : " & Lrs_Find.Fields("libelle") & ", en conge le : " & xDate, vbExclamation, App.ProductName
                    Cmd_Ok.Visible = False
                    Cmd_Edit.Visible = True
                    CmdDelete.Enabled = False
                    Cmd_Edit.Enabled = False
                    Cmd_AddNew.Enabled = True
                    Cmd_AddPlanning.Enabled = True
                    Grid_DetPlanning.Enabled = True
                    cbo_vehicule.ListIndex = 0
                    Cbo_Conducteur.ListIndex = 0
                    ZRow = 0
                    Exit Sub
                End If
            Else
                If Lrs_Find.EOF Then
                    With Grid_DetPlanning
                        If .Rows = 0 Then
                            .Redraw = False
                            .AddRow
                            .CellDetails ZRow, .ColumnIndex("CodeConducteur"), CCode
                            .CellDetails ZRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                            .CellDetails ZRow, .ColumnIndex("HeureEntre"), CHeureEntre, , , &HC0C0C0
                            .Redraw = True
                        ElseIf .Rows > 0 Then
                            For i = 1 To .Rows
                                If i <> ZRow Then
                                    If Cbo_Conducteur.Text = .CellText(i, 2) & "  -  " & .CellText(i, 3) Then
                                        MsgBox "Vérifier 'Conducteur' et 'Véhicule' exist déja en planning!..", vbExclamation, App.ProductName
                                        Exit Sub
                                    End If
                                End If
                            Next i
                            .Redraw = False
                            .CellDetails ZRow, .ColumnIndex("CodeConducteur"), CCode
                            .CellDetails ZRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                            .CellDetails ZRow, .ColumnIndex("HeureEntre"), CHeureEntre, , , &HC0C0C0
                            .Redraw = True
                        End If
                    End With
                    If Grid_DetPlanning.Rows > 0 Then Cmd_AddPlanning.Enabled = True
                Else
                    MsgBox "Conducteur : " & Lrs_Find.Fields("libelle") & ", en conge le : " & xDate, vbExclamation, App.ProductName
                    Cmd_Ok.Visible = False
                    Cmd_Edit.Visible = True
                    CmdDelete.Enabled = False
                    Cmd_Edit.Enabled = False
                    Cmd_AddNew.Enabled = True
                    Cmd_AddPlanning.Enabled = True
                    Grid_DetPlanning.Enabled = True
                    cbo_vehicule.ListIndex = 0
                    Cbo_Conducteur.ListIndex = 0
                    ZRow = 0
                    Exit Sub
                End If
            End If
    '# Repos***
        ElseIf GridValid = False Then
            If Cbo_Conducteur.ListIndex = 0 = "" Then
                MsgBox "Séléctionner... 'Conducteur'.", vbExclamation, App.ProductName
                Exit Sub
            End If
            cond = CLibelle
            For i = 1 To .Rows
                If i <> .Rows Then
                    If (.CellText(i, ColR) <> "" And .CellText(i, ColR) <> " ") Then
                        XChp = Split(.CellText(i, ColR), vbCr)
                        XCount = UBound(XChp)
                        For J = 0 To XCount
                            XChps = Split(XChp(J), "||")
                            Conducteur = RTrim(XChps(0))
                            If RTrim(cond) = Conducteur Then
                                MsgBox Conducteur & " Déja affecteé pour tourneé de " & .CellText(i, 1), vbExclamation, App.ProductName
                                Exit Sub
                            End If
                        Next J
                    End If
                End If
            Next i
            xDate = DateCell(Cda_DebutPlg.Value, ColR)
            Set Lrs_Find = LObj_Find.GetRow_CongeByCodeCondAndDate(ErrNumber, ErrDescription, ErrSourceDetail, CCode, xDate, CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            Set LObj_Find = Nothing
            If Lrs_Find.EOF Then
                With Grid_DetPlanning
                    If .Rows = 0 Then
                        .Redraw = False
                        .AddRow
                        .CellDetails ZRow, .ColumnIndex("CodeConducteur"), CCode
                        .CellDetails ZRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                        If Opt_Matin.Value = vbChecked Then
                            .CellDetails ZRow, .ColumnIndex("TypeRep"), "Matin", , , &HC0C0C0
                        ElseIf OptSoir.Value = vbChecked Then
                            .CellDetails ZRow, .ColumnIndex("TypeRep"), "Soir", , , &HC0C0C0
                        End If
                        .Redraw = True
                    ElseIf .Rows > 0 Then
                        For i = 1 To .Rows
                            If i <> ZRow Then
                                If Cbo_Conducteur.Text = .CellText(i, 2) & "  -  " & .CellText(i, 3) Then
                                    MsgBox "Vérifier " & .CellText(i, 2) & " exist déja en Repos!..", vbExclamation, App.ProductName
                                    Exit Sub
                                End If
                            End If
                        Next i
                        .Redraw = False
                        .CellDetails ZRow, .ColumnIndex("CodeConducteur"), CCode
                        .CellDetails ZRow, .ColumnIndex("Conducteur"), CLibelle, , , &HC0C0C0
                        If Opt_Matin.Value = vbChecked Then
                            .CellDetails ZRow, .ColumnIndex("TypeRep"), "Matin", , , &HC0C0C0
                        ElseIf OptSoir.Value = vbChecked Then
                            .CellDetails ZRow, .ColumnIndex("TypeRep"), "Soir", , , &HC0C0C0
                        End If
                        .Redraw = True
                    End If
                    cbo_vehicule.ListIndex = 0
                    Cbo_Conducteur.ListIndex = 0
                End With
                If Grid_DetPlanning.Rows > 0 Then Cmd_AddPlanning.Enabled = True
            Else
                MsgBox "Conducteur : " & Lrs_Find.Fields("libelle") & ", en congé le : " & xDate, vbExclamation, App.ProductName
                cbo_vehicule.ListIndex = 0
                Cbo_Conducteur.ListIndex = 0
                Exit Sub
            End If
            Set Lrs_Find = Nothing
        End If
    End With
    Cmd_Ok.Visible = False
    Cmd_Edit.Visible = True
    CmdDelete.Enabled = False
    Cmd_Edit.Enabled = False
    Cmd_AddNew.Enabled = True
    Cmd_AddPlanning.Enabled = True
    Grid_DetPlanning.Enabled = True
    cbo_vehicule.ListIndex = 0
    Cbo_Conducteur.ListIndex = 0
    ZRow = 0
    Set Lrs_Find = Nothing
    Chk_HEntre.Value = 0
    Txt_Heurentre.Text = ""
    Txt_Heurentre.Enabled = False
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub
'# Edit Detail PLANNING***
Private Sub Grid_DetPlanning_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    ZRow = lRow
    Call Cmd_Edit_Click
End Sub
Private Sub Cmd_Edit_Click()
On Error GoTo Err
    If Grid_DetPlanning.CellText(ZRow, 8) <> "" Then
        If (CHECK_ACCES("MAJ_PLING", LInt_UserId) = False) Then
            MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier PLANNING", vbExclamation
            Exit Sub
        End If
    End If
    If GridValid = True Then
        If (Row <> Grid_Planning.Rows And Row <> Grid_Planning.Rows - 1) Then
            With Grid_DetPlanning
                If .Rows > 0 Then
                    If ZRow > 0 Then
                        Cbo_Conducteur.Text = .CellText(ZRow, 2) & "  -  " & .CellText(ZRow, 3)
                        cbo_vehicule.Text = .CellText(ZRow, 4) & "  -  " & .CellText(ZRow, 5)
                        If Trim(.CellText(ZRow, 7)) <> "" Then
                            Txt_Heurentre.Text = .CellText(ZRow, 7) & ":00"
                            Txt_Heurentre.Enabled = True
                            Chk_HEntre.Value = 1
                        Else
                            Txt_Heurentre.Text = ""
                            Txt_Heurentre.Enabled = False
                            Chk_HEntre.Value = 0
                        End If
                        CCode = .CellText(ZRow, 2)
                        VeCode = .CellText(ZRow, 4)
                        CLibelle = .CellText(ZRow, 3)
                        VMatricule = .CellText(ZRow, 5)
                    End If
                End If
            End With
        ElseIf (Row = Grid_Planning.Rows Or Row = Grid_Planning.Rows - 1) Then
            With Grid_DetPlanning
                If .Rows > 0 Then
                    If ZRow > 0 Then
                        Cbo_Conducteur.Text = .CellText(ZRow, 2) & "  -  " & .CellText(ZRow, 3)
                        If Trim(.CellText(ZRow, 7)) <> "" Then
                            Txt_Heurentre.Text = .CellText(ZRow, 7) & ":00"
                            Txt_Heurentre.Enabled = True
                            Chk_HEntre.Value = 1
                        Else
                            Txt_Heurentre.Text = ""
                            Txt_Heurentre.Enabled = False
                            Chk_HEntre.Value = 0
                        End If
                        CCode = .CellText(ZRow, 2)
                        CLibelle = .CellText(ZRow, 3)
                    End If
                End If
            End With
        End If
    ElseIf GridValid = False Then
        With Grid_DetPlanning
            If .Rows > 0 Then
                If ZRow > 0 Then
                    Cbo_Conducteur.Text = .CellText(ZRow, 2) & "  -  " & .CellText(ZRow, 3)
                    CCode = .CellText(ZRow, 2)
                    CLibelle = .CellText(ZRow, 3)
                    If .CellText(ZRow, 6) = "Matin" Then
                        Opt_Matin.Value = vbChecked
                    ElseIf .CellText(ZRow, 6) = "Après midi" Then
                        OptSoir.Value = vbChecked
                    End If
                End If
            End If
        End With
    End If
    Cmd_Ok.Visible = True
    Cmd_Edit.Visible = False
    Cmd_AddNew.Enabled = False
    Cmd_AddPlanning.Enabled = False
    CmdDelete.Enabled = False
    Grid_DetPlanning.Enabled = False
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'# Supprimer Planning***
Private Sub CmdDelete_Click()
On Error GoTo Err
    With Grid_DetPlanning
        If .CellText(ZRow, 8) <> "" Then
            If (CHECK_ACCES("SUPP_PLING", LInt_UserId) = False) Then
                MsgBox "Suppression n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Supprimer PLANNING", vbExclamation
                Exit Sub
            End If
        End If
        If .Rows > 0 Then
            If ZRow > 0 Then
                .RemoveRow (ZRow)
                ZRow = 0
                Cmd_Edit.Enabled = False
                CmdDelete.Enabled = False
            Else
                MsgBox "Séléctionner un ligne pour supprimer!...", vbExclamation, App.ProductName
                Exit Sub
            End If
        End If
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'# Button Annuler***
Private Sub Cmd_Cancel_Click()
    PicBox_Planning.Visible = False
    Grid_DetPlanning.Enabled = True
    Grid_DetPlanning.ClearRows
    Cmd_Ok.Visible = False
    Cmd_Edit.Visible = True
    Cmd_AddNew.Enabled = True
    Cbo_Conducteur.Text = ""
    cbo_vehicule.Text = ""
    Cbo_Tournee.Caption = ""
    Row = 0
    Col = 0
    ZRow = 0
    Call EnbDisb(True)
    Chk_HEntre.Value = vbUnchecked
    Txt_Heurentre.Enabled = False
    Txt_Heurentre.Text = ""
End Sub
Private Sub Pic_Cancel_Click()
    Call Cmd_Cancel_Click
End Sub
Public Function CondRepos() As String
    Dim LOBJ_Cond       As New Conducteur
    Dim Lrs_Find        As New Recordset
    Dim trv             As Boolean
    Dim i               As Integer
    Dim J               As Integer
    Dim Msg             As String
    Dim XChp()          As String
    Dim XChps()         As String
    Dim XCount          As Integer
    Dim cond            As String
On Error GoTo Err
    Msg = ""
    Set Lrs_Find = LOBJ_Cond.GetAll_ConducteursActifNonSupp(ErrNumber, ErrDescription, ErrSourceDetail, "O", "N", CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbInformation
        ErrNumber = 0
        Exit Function
    End If
    Set LOBJ_Cond = Nothing
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
            trv = False
            For i = 1 To Grid_Repos.Columns
                If Grid_Repos.CellText(1, i) <> "" Then
                    XChp = Split(Grid_Repos.CellText(1, i), vbCr)
                    XCount = UBound(XChp)
                    For J = 0 To XCount
                        XChps = Split(XChp(J), "||")
                        cond = Trim(XChps(0))
                        If Trim(cond) = Trim(Lrs_Find("Libelle")) Then
                            trv = True
                            Exit For
                        End If
                    Next J
                End If
                If trv = True Then Exit For
            Next i
            If trv = False Then
                If Msg = "" Then
                    Msg = Lrs_Find("Libelle")
                Else
                    Msg = Msg & "||" & Lrs_Find("Libelle")
                End If
            End If
            Lrs_Find.MoveNext
        Wend
    End If
    Set Lrs_Find = Nothing
    CondRepos = Msg
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Function
'Chercher PLANNING By MonthDate***
Private Sub MonthView_DateClick(ByVal DateClicked As Date)
    Dim MDate As Date
On Error GoTo Err
    MDate = MonthView.Value
    If MDate >= DateNewPlanning(Date) Then
        Cmd_suivt.Enabled = False
        Pic_Prochain.Enabled = False
        Cda_DebutPlg.Value = DateNewPlanning(Date)
        Cda_FinPlg.Value = DateWEnd(DateNewPlanning(Date))
        Call SearchPLANNING(DateNewPlanning(Date))
    Else
        Cmd_suivt.Enabled = True
        Pic_Prochain.Enabled = True
        Cda_DebutPlg.Value = DatePlanning(MDate)
        Cda_FinPlg.Value = DateWEnd(DatePlanning(MDate))
        SearchPLANNING (DatePlanning(MDate))
    End If
    Call Pic_Show_Click
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Planning Suivant***
Private Sub Cmd_suivt_Click()
    Dim Date_Search     As Date
    Dim LObj_Find       As New PLANNING
    Dim Lrs_Date        As New Recordset
On Error GoTo Err
    Call Pic_Mask_Click
    PicBox_Planning.Visible = False
    Cmd_Ok.Visible = False
    Cmd_AddNew.Enabled = True
    Grid_Planning.ClearRows
    Grid_DetPlanning.ClearRows
    Grid_Planning.Enabled = True
    Grid_DetPlanning.Enabled = True
    CmdPrint.Enabled = True
    Pic_Prochain.Enabled = True
    Call FindDest
    Date_Search = Cda_DebutPlg.Value
    Set Lrs_Date = LObj_Find.GetDate_NewPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, Date_Search, 1, CNB)
    If ErrNumber <> 0 Then
         ErrNumber = 0
         MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
         Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Date.EOF Then
         Date_Search = Lrs_Date.Fields("datedebut")
    End If
    Set Lrs_Date = Nothing
    Cda_DebutPlg.Value = DateNewPlanning(Date_Search)
    Cda_FinPlg.Value = DateWEnd(DateNewPlanning(Date_Search))
    Call SearchPLANNING(DateNewPlanning(Date_Search))
    If Cda_DebutPlg.Value = DateNewPlanning(Date) Then
        Cmd_suivt.Enabled = False
        Pic_Prochain.Enabled = False
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Planning Précedant***
Private Sub Cmd_precd_Click()
    Dim Date_Search     As Date
    Dim LObj_Find       As New PLANNING
    Dim Lrs_Date        As New Recordset
On Error GoTo Err
    Call Pic_Mask_Click
    PicBox_Planning.Visible = False
    Cmd_Ok.Visible = False
    Cmd_AddNew.Enabled = True
    Grid_Planning.ClearRows
    Grid_DetPlanning.ClearRows
    Grid_Planning.Enabled = True
    Grid_DetPlanning.Enabled = True
    CmdPrint.Enabled = True
    Pic_Prochain.Enabled = True
    Cmd_suivt.Enabled = True
    Call FindDest
    Date_Search = Cda_DebutPlg.Value
    Set Lrs_Date = LObj_Find.GetDate_NewPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, Date_Search, -1, CNB)
    If ErrNumber <> 0 Then
         ErrNumber = 0
         MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
         Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Date.EOF Then
         Date_Search = Lrs_Date.Fields("datedebut")
    End If
    Set Lrs_Date = Nothing
    Cda_DebutPlg.Value = DatePlanning(Date_Search)
    Cda_FinPlg.Value = DateWEnd(DatePlanning(Date_Search))
    Call SearchPLANNING(DatePlanning(Date_Search))
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Planning Prochain***
Private Sub Pic_Prochain_Click()
On Error GoTo Err
    Call Pic_Mask_Click
    PicBox_Planning.Visible = False
    Cmd_Ok.Visible = False
    Cmd_AddNew.Enabled = True
    Grid_Planning.ClearRows
    Grid_DetPlanning.ClearRows
    Grid_Planning.Enabled = True
    Grid_DetPlanning.Enabled = True
    CmdPrint.Enabled = True
    Call FindDest
    Cda_DebutPlg.Value = DateNewPlanning(Date)
    Cda_FinPlg.Value = DateWEnd(DateNewPlanning(Date))
    Call SearchPLANNING(DateNewPlanning(Date))
    If Cda_DebutPlg.Value = DateNewPlanning(Date) Then
        Pic_Prochain.Enabled = False
        Cmd_suivt.Enabled = False
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Actualiser Planning***
Private Sub Pic_Ref_Click()
On Error GoTo Err
    Call Pic_Mask_Click
    PicBox_Planning.Visible = False
    Cmd_Ok.Visible = False
    Cmd_AddNew.Enabled = True
    Grid_Planning.ClearRows
    Grid_DetPlanning.ClearRows
    Grid_Planning.Enabled = True
    Grid_DetPlanning.Enabled = True
    CmdPrint.Enabled = True
    Call FindDest
    Call SearchPLANNING(Cda_DebutPlg.Value)
    If Cda_DebutPlg.Value = DateNewPlanning(Date) Then
        Pic_Prochain.Enabled = False
        Cmd_suivt.Enabled = False
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Planning Actuelle***
Private Sub Pic_PlannigSemaine_Click()
On Error GoTo Err
    Call Pic_Mask_Click
    PicBox_Planning.Visible = False
    Cmd_Ok.Visible = False
    Cmd_AddNew.Enabled = True
    Grid_Planning.ClearRows
    Grid_DetPlanning.ClearRows
    Grid_Planning.Enabled = True
    Grid_DetPlanning.Enabled = True
    CmdPrint.Enabled = True
    Call FindDest
    Cda_DebutPlg.Value = DatePlanning(Date)
    Cda_FinPlg.Value = DateWEnd(DatePlanning(Date))
    Call SearchPLANNING(DatePlanning(Date))
    Pic_Prochain.Enabled = True
    Cmd_suivt.Enabled = True
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Rechercher PLANNING***
Private Sub SearchPLANNING(ByVal DateSearch As Date)
    Dim LObj_Find       As New PLANNING
    Dim Lrs_PLANNING    As New Recordset
    Dim TOURNEE         As String
    Dim ZCol            As Integer
    Dim i               As Integer
On Error GoTo Err
    Cmd_Ok.Visible = False
    Cmd_AddNew.Enabled = True
    Grid_Planning.ClearRows
    Grid_Repos.ClearRows
    Grid_Conge.ClearRows
    Grid_DetPlanning.ClearRows
    Call FindDest
'# Planning***
    Set Lrs_PLANNING = LObj_Find.GetRow_PLANNINGByDateDu(ErrNumber, ErrDescription, ErrSourceDetail, DateSearch, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_PLANNING.EOF Then
        With Grid_Planning
            While Not Lrs_PLANNING.EOF
                For i = 1 To .Rows
                    TOURNEE = .CellText(i, 1)
                    If Lrs_PLANNING.Fields("Tournee") = TOURNEE Then
                        If Lrs_PLANNING.Fields("Jour") = "lundi" Then ZCol = 2
                        If Lrs_PLANNING.Fields("Jour") = "mardi" Then ZCol = 3
                        If Lrs_PLANNING.Fields("Jour") = "mercredi" Then ZCol = 4
                        If Lrs_PLANNING.Fields("Jour") = "jeudi" Then ZCol = 5
                        If Lrs_PLANNING.Fields("Jour") = "vendredi" Then ZCol = 6
                        If Lrs_PLANNING.Fields("Jour") = "samedi" Then ZCol = 7
                        If Lrs_PLANNING.Fields("Jour") = "dimanche" Then ZCol = 8
                        If ZCol <> 0 Then .CellTextAlign(i, ZCol) = DT_EXPANDTABS
                        If Not (IsNull(Lrs_PLANNING.Fields("Vehicule"))) Then
                            If .CellText(i, ZCol) = "" Then
                                If Lrs_PLANNING.Fields("HeureEntre") <> "NULL" And Trim(Lrs_PLANNING.Fields("HeureEntre")) <> "" Then
                                    .CellText(i, ZCol) = Lrs_PLANNING.Fields("Conducteur") & " || " & Lrs_PLANNING.Fields("Vehicule") & " || " & Mid(Lrs_PLANNING.Fields("HeureEntre"), 1, 5)
                                Else
                                    .CellText(i, ZCol) = Lrs_PLANNING.Fields("Conducteur") & " || " & Lrs_PLANNING.Fields("Vehicule")
                                End If
                            Else
                                If Lrs_PLANNING.Fields("HeureEntre") <> "NULL" And Trim(Lrs_PLANNING.Fields("HeureEntre")) <> "" Then
                                    .CellText(i, ZCol) = .CellText(i, ZCol) & vbCr & Lrs_PLANNING.Fields("Conducteur") & " || " & Lrs_PLANNING.Fields("Vehicule") & " || " & Mid(Lrs_PLANNING.Fields("HeureEntre"), 1, 5)
                                Else
                                    .CellText(i, ZCol) = .CellText(i, ZCol) & vbCr & Lrs_PLANNING.Fields("Conducteur") & " || " & Lrs_PLANNING.Fields("Vehicule")
                                End If
                            End If
                        Else
                            If .CellText(i, ZCol) = "" Then
                                If Lrs_PLANNING.Fields("HeureEntre") <> "NULL" And Trim(Lrs_PLANNING.Fields("HeureEntre")) <> "" Then
                                    .CellText(i, ZCol) = Lrs_PLANNING.Fields("Conducteur") & " || " & Mid(Lrs_PLANNING.Fields("HeureEntre"), 1, 5)
                                Else
                                    .CellText(i, ZCol) = Lrs_PLANNING.Fields("Conducteur")
                                End If
                            Else
                                If Lrs_PLANNING.Fields("HeureEntre") <> "NULL" And Trim(Lrs_PLANNING.Fields("HeureEntre")) <> "" Then
                                    .CellText(i, ZCol) = .CellText(i, ZCol) & vbCr & Lrs_PLANNING.Fields("Conducteur") & " || " & Mid(Lrs_PLANNING.Fields("HeureEntre"), 1, 5)
                                Else
                                    .CellText(i, ZCol) = .CellText(i, ZCol) & vbCr & Lrs_PLANNING.Fields("Conducteur")
                                End If
                            End If
                        End If
                    End If
                Next i
                Lrs_PLANNING.MoveNext
            Wend
        End With
    End If
    Set Lrs_PLANNING = Nothing
'# Repos***
    Dim Lrs_Rep     As New Recordset
    Dim LObj_Rep    As New Conducteur
    Set Lrs_Rep = LObj_Rep.Get_repos(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Cda_DebutPlg.Value, Cda_FinPlg.Value)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Rep.EOF Then
        With Grid_Repos
            While Not Lrs_Rep.EOF
                    If .CellText(1, 1) = "Repos" Then
                        If Format(Lrs_Rep.Fields("DateDu"), "dddd") = "lundi" Then ZCol = 2
                        If Format(Lrs_Rep.Fields("DateDu"), "dddd") = "mardi" Then ZCol = 3
                        If Format(Lrs_Rep.Fields("DateDu"), "dddd") = "mercredi" Then ZCol = 4
                        If Format(Lrs_Rep.Fields("DateDu"), "dddd") = "jeudi" Then ZCol = 5
                        If Format(Lrs_Rep.Fields("DateDu"), "dddd") = "vendredi" Then ZCol = 6
                        If Format(Lrs_Rep.Fields("DateDu"), "dddd") = "samedi" Then ZCol = 7
                        If Format(Lrs_Rep.Fields("DateDu"), "dddd") = "dimanche" Then ZCol = 8
                        
                        If ZCol <> 0 Then .CellTextAlign(1, ZCol) = DT_EXPANDTABS
                        If .CellText(1, ZCol) = "" Then
                            .CellText(1, ZCol) = Lrs_Rep.Fields("Libelle") & " || " & Lrs_Rep.Fields("Observation")
                        Else
                            .CellText(1, ZCol) = .CellText(1, ZCol) & vbCr & Lrs_Rep.Fields("Libelle") & " || " & Lrs_Rep.Fields("Observation")
                        End If
                    End If
                Lrs_Rep.MoveNext
            Wend
        End With
    End If
    Set Lrs_Rep = Nothing
'# Congé
    Dim Lrs_Cong        As New Recordset
    Dim LObj_Cong       As New Conducteur
    Dim Date_Conge      As Date
    For i = 2 To 7
        Date_Conge = Format(DateCell(Cda_DebutPlg.Value, i), "dd/mm/yyyy")
        Set Lrs_Cong = LObj_Cong.GetRow_CongeByDate(ErrNumber, ErrDescription, ErrSourceDetail, Date_Conge, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
            Exit Sub
        End If
        Set LObj_Cong = Nothing
        If Not Lrs_Cong.EOF Then
            With Grid_Conge
                While Not Lrs_Cong.EOF
                    If .CellText(1, i) = "" Then
                        .CellText(1, i) = Lrs_Cong("Libelle")
                    Else
                        .CellText(1, i) = .CellText(1, i) & vbCr & Lrs_Cong("Libelle")
                    End If
                    Lrs_Cong.MoveNext
                Wend
            End With
        End If
        Set Lrs_Cong = Nothing
    Next i
    Set Lrs_Cong = Nothing
    
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'# Ajouter a PLANNING et enregistre***
Private Sub Cmd_AddPlanning_Click()
    Dim i       As Integer
    Dim Msg     As VbMsgBoxResult
On Error GoTo Err
    If Grid_DetPlanning.Rows = 0 Then
        Msg = MsgBox("Planning est vide!...      " & vbCr & "Aucun Planning pour enregistre!...", vbExclamation + vbOKCancel, App.ProductName)
        If Msg = vbCancel Then Exit Sub
    End If
    If GridValid = True Then
        If Edit = False Then
            Call SaveNewPLANNING(Col)
        Else
            Call SaveNewPLANNING(Col)
            Call SaveEditPLANNING(Col)
            Call RemovePlanning(DCode, Format(DateCell(Cda_DebutPlg.Value, Col), "dddd"), Cda_DebutPlg.Value, Col)
        End If
        Grid_Planning.CellText(Row, Col) = ""
        With Grid_DetPlanning
            Grid_Planning.CellTextAlign(Row, Col) = DT_EXPANDTABS
            For i = 1 To .Rows
                If Row <> Grid_Planning.Rows And Row <> Grid_Planning.Rows - 1 Then
                    If Grid_Planning.CellText(Row, Col) = "" Then
                        If .CellText(i, 3) <> "" And .CellText(i, 5) <> "" And .CellText(i, 7) <> "" Then
                            Grid_Planning.CellText(Row, Col) = .CellText(i, 3) & "  ||  " & .CellText(i, 5) & "  ||  " & .CellText(i, 7)
                        ElseIf .CellText(i, 3) <> "" And .CellText(i, 5) <> "" Then
                            Grid_Planning.CellText(Row, Col) = .CellText(i, 3) & "  ||  " & .CellText(i, 5)
                        End If
                    Else
                        If .CellText(i, 3) <> "" And .CellText(i, 5) <> "" And .CellText(i, 7) <> "" Then
                            Grid_Planning.CellText(Row, Col) = Grid_Planning.CellText(Row, Col) & vbCr & .CellText(i, 3) & "  ||  " & .CellText(i, 5) & "  ||  " & .CellText(i, 7)
                        ElseIf .CellText(i, 3) <> "" And .CellText(i, 5) <> "" Then
                            Grid_Planning.CellText(Row, Col) = Grid_Planning.CellText(Row, Col) & vbCr & .CellText(i, 3) & "  ||  " & .CellText(i, 5)
                        End If
                    End If
                Else
                    If Grid_Planning.CellText(Row, Col) = "" Then
                        If .CellText(i, 3) <> "" And .CellText(i, 7) <> "" Then
                            Grid_Planning.CellText(Row, Col) = .CellText(i, 3) & "  ||  " & .CellText(i, 7)
                        Else
                            Grid_Planning.CellText(Row, Col) = .CellText(i, 3)
                        End If
                    Else
                        If .CellText(i, 3) <> "" And .CellText(i, 7) <> "" Then
                            Grid_Planning.CellText(Row, Col) = Grid_Planning.CellText(Row, Col) & vbCr & .CellText(i, 3) & "  ||  " & .CellText(i, 7)
                        Else
                            Grid_Planning.CellText(Row, Col) = Grid_Planning.CellText(Row, Col) & vbCr & .CellText(i, 3)
                        End If
                    End If
                End If
            Next i
        End With
'# Enregistrer Repos dans Table Congés
    ElseIf GridValid = False Then
        If Edit = False Then
            Call SaveNewPLANNING(ColR)
        Else
            Call SaveNewPLANNING(ColR)
            Call SaveEditPLANNING(ColR)
            Call RemovePlanning(DCode, Format(DateCell(Cda_DebutPlg.Value, ColR), "dddd"), Cda_DebutPlg.Value, ColR)
        End If
        Grid_Repos.CellText(RowR, ColR) = ""
        With Grid_DetPlanning
            Grid_Repos.CellTextAlign(1, ColR) = DT_EXPANDTABS
            For i = 1 To .Rows
                If .CellText(i, 3) <> "" Then
                    If Grid_Repos.CellText(1, ColR) = "" Then
                        Grid_Repos.CellText(1, ColR) = .CellText(i, 3) & "  ||  " & .CellText(i, 6)
                       ' Grid_Repos.CellBackColor(1, ColR) = &H404040
                       ' Grid_Repos.CellForeColor(1, ColR) = &HFFFFFF
                    Else
                        Grid_Repos.CellText(1, ColR) = Grid_Repos.CellText(1, ColR) & vbCr & .CellText(i, 3) & "  ||  " & .CellText(i, 6)
                        'Grid_Repos.CellBackColor(1, ColR) = &H404040
                       ' Grid_Repos.CellForeColor(1, ColR) = &HFFFFFF
                    End If
                End If
            Next i
        End With
    End If
    
    PicBox_Planning.Visible = False
    Grid_DetPlanning.ClearRows
    Cmd_Edit.Enabled = False
    CmdDelete.Enabled = False
    Cmd_Ok.Visible = False
    Cbo_Conducteur.Text = ""
    cbo_vehicule.Text = ""
    EnbDisb (True)
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'# Nouveau Planning***
Private Sub SaveNewPLANNING(ByVal Colm As Integer)
    Dim Lrs_PLANNING    As New Recordset
    Dim Lobj_PLANNING   As New PLANNING
    Dim DateDu          As String
    Dim DateAu          As Date
    Dim jour            As String
    Dim Z               As Integer
    Dim CHE             As String
On Error GoTo Err
    With Grid_DetPlanning
        DateDu = Cda_DebutPlg.Value
        DateAu = Cda_FinPlg.Value
        jour = Format(DateCell(Cda_DebutPlg.Value, Colm), "dddd")
        If DateAu < Date Then
            MsgBox " Planning déja fait", vbExclamation, App.ProductName
            Exit Sub
        End If
        If (CHECK_ACCES("SUPP_PLING", LInt_UserId) = False) Then
            MsgBox "Insertion n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Ajouter PLANNING", vbExclamation
            Exit Sub
        End If
    '# Tournee**
        If GridValid = True Then
            'Save PLANNING***
            Set Lrs_PLANNING = CreateEmptyRS_PLANNING()
            For Z = 1 To .Rows
                If Trim(Grid_DetPlanning.CellText(Z, 7)) <> "" Then CHE = Mid(Grid_DetPlanning.CellText(Z, 7), 1, 5) Else CHE = "NULL"
                'If CodePLANNING IS NULL***
                If .CellText(Z, 8) = "" Then
                    With Lrs_PLANNING
                        .AddNew
                        .Fields("Datedu") = Format(DateDu, "dd/mm/yyyy hh:mm:ss")
                        .Fields("dateau") = Format(DateAu, "dd/mm/yyyy hh:mm:ss")
                        .Fields("DateCreat") = Format(Date, "dd/mm/yyyy hh:mm:ss")
                        .Fields("DateJour") = Format(DateCell(DateDu, Colm), "dd/mm/yyyy hh:mm:ss")
                        .Fields("Tournee") = Grid_DetPlanning.CellText(Z, 1)
                        .Fields("HeureEntre") = CHE
                        .Fields("Jour") = jour
                        .Fields("Conducteur") = Grid_DetPlanning.CellText(Z, 2)
                        If (Grid_DetPlanning.CellText(Z, 4)) <> "" Then .Fields("Vehicule") = Grid_DetPlanning.CellText(Z, 4)
                        .Fields("userinsert") = LInt_UserId
                    End With
                    'Insert***
                    Call Lobj_PLANNING.Save_PLANNING(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_PLANNING)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
                        Exit Sub
                    End If
                    Set Lobj_PLANNING = Nothing
                End If
            Next Z
            Set Lrs_PLANNING = Nothing
    '# Repos***
        ElseIf GridValid = False Then
            Dim Lrs_Repos       As New Recordset
            Dim lobj_Repos      As New Conducteur
            'Save Repos***
            Set Lrs_Repos = CreateEmptyRS_Conge()
            For Z = 1 To .Rows
                'If CodeRepos IS NULL***
                If .CellText(Z, 8) = "" Then
                    With Lrs_Repos
                        .AddNew
                        .Fields("Conducteur") = Grid_DetPlanning.CellText(Z, 2)
                        .Fields("datedu") = DateCell(Cda_DebutPlg.Value, Colm)
                        .Fields("Dateau") = DateCell(Cda_DebutPlg.Value, Colm)
                        .Fields("Type") = "Repos"
                        .Fields("Observation") = Grid_DetPlanning.CellText(Z, 6)
                        .Fields("userinsert") = LInt_UserId
                    End With
                    'Insert***
                    Call lobj_Repos.Insert_Conge(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Repos)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
                        Exit Sub
                    End If
                    Set lobj_Repos = Nothing
                End If
            Next Z
            Set Lrs_Repos = Nothing
        End If
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Edit PLANNING***
Private Sub SaveEditPLANNING(ByVal Colm As Integer)
    Dim DateAu          As Date
    Dim jour            As String
    Dim Z               As Integer
    Dim Lrs_PLANNING    As New Recordset
    Dim Lobj_PLANNING   As New PLANNING
    Dim Lrs_Find        As New Recordset
    Dim CodePLANNING    As Integer
    Dim CHE             As String
    Dim CVEH            As String
On Error GoTo Err
    With Grid_DetPlanning
        DateAu = Cda_FinPlg.Value
        jour = Format(DateCell(Cda_DebutPlg.Value, Colm), "dddd")
        If DateAu < Date Then
            MsgBox " Planning déja fait", vbExclamation, App.ProductName
            Exit Sub
        End If
        If (CHECK_ACCES("SUPP_PLING", LInt_UserId) = False) Then
            MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier PLANNING", vbExclamation
            Exit Sub
        End If
    '# Tournee
        If GridValid = True Then
            'Save PLANNING***
            Set Lrs_PLANNING = CreateEmptyRS_PLANNING()
            For Z = 1 To .Rows
                If Trim(Grid_DetPlanning.CellText(Z, 7)) <> "" Then CHE = Mid(Grid_DetPlanning.CellText(Z, 7), 1, 5) Else CHE = "NULL"
                If Trim(.CellText(Z, 4)) = "" Then CVEH = "Null" Else CVEH = .CellText(Z, 4)
                'If CodePLANNING IS NOT NULL***
                If .CellText(Z, 8) <> "" Then
                    CodePLANNING = .CellText(Z, 8)
                    Set Lrs_Find = New Recordset
                    Set Lrs_Find = Lobj_PLANNING.GetPLANNING_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, CodePLANNING, CNB)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                        Exit Sub
                    End If
                    Set Lobj_PLANNING = Nothing
                    If Not Lrs_Find.EOF Then
                        Dim LrsHE   As String
                        Dim LrsVEH  As String
                        If IsNull(Lrs_Find("HeureEntre")) Then LrsHE = "NULL" Else LrsHE = Lrs_Find("HeureEntre")
                        If IsNull(Lrs_Find("Vehicule")) Then LrsVEH = "NULL" Else LrsVEH = Lrs_Find("Vehicule")
                        If .CellText(Z, 1) <> Lrs_Find("Tournee") Or CVEH <> LrsVEH Or .CellText(Z, 2) <> Lrs_Find("Conducteur") Or CHE <> LrsHE Then
                            With Lrs_PLANNING
                                .AddNew
                                .Fields("DateEdit") = Format(Date, "dd/mm/yyyy hh:mm:ss")
                                .Fields("Tournee") = Grid_DetPlanning.CellText(Z, 1)
                                .Fields("DateJour") = Format(DateCell(Cda_DebutPlg.Value, Colm), "dd/mm/yyyy hh:mm:ss")
                                .Fields("Jour") = jour
                                .Fields("Conducteur") = Grid_DetPlanning.CellText(Z, 2)
                                .Fields("HeureEntre") = CHE
                                If (Grid_DetPlanning.CellText(Z, 4)) <> "" Then .Fields("Vehicule") = Grid_DetPlanning.CellText(Z, 4)
                                .Fields("userupdate") = LInt_UserId
                            End With
                            'Update***
                            Set Lobj_PLANNING = New PLANNING
                            Call Lobj_PLANNING.Update_PLANNING(ErrNumber, ErrDescription, ErrSourceDetail, CodePLANNING, CNB, Lrs_PLANNING)
                            If ErrNumber <> 0 Then
                                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
                                ErrNumber = 0
                                Exit Sub
                            End If
                            Set Lobj_PLANNING = Nothing
                        End If
                    End If
                    Set Lrs_Find = Nothing
                End If
                CodePLANNING = 0
            Next Z
            Set Lrs_PLANNING = Nothing
    '# Repos
        ElseIf GridValid = False Then
            'Save Repos***
            Dim Lrs_Repos   As New Recordset
            Dim lobj_Repos  As New Conducteur
            Dim CodeRepos   As Integer
            Set Lrs_Repos = CreateEmptyRS_Conge()
            For Z = 1 To .Rows
                'If CodeRepos IS NOT NULL***
                If .CellText(Z, 8) <> "" Then
                    CodeRepos = .CellText(Z, 8)
                    Set Lrs_Find = New Recordset
                    Set Lrs_Find = lobj_Repos.Get_CongeCondByCode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, CodeRepos)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                        Exit Sub
                    End If
                    Set lobj_Repos = Nothing
                    If Not Lrs_Find.EOF Then
                        If Trim(.CellText(Z, 2)) <> Trim(Lrs_Find("conducteur")) Or Trim(.CellText(Z, 6)) <> Trim(Lrs_Find("observation")) Then
                            With Lrs_Repos
                                .AddNew
                                .Fields("numero") = CodeRepos
                                .Fields("Conducteur") = Grid_DetPlanning.CellText(Z, 2)
                                .Fields("Observation") = Grid_DetPlanning.CellText(Z, 6)
                                .Fields("userupdate") = LInt_UserId
                            End With
                            'Update***
                            Set lobj_Repos = New Conducteur
                            Call lobj_Repos.Update_Repos(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Repos)
                            If ErrNumber <> 0 Then
                                MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion, App.ProductName
                                ErrNumber = 0
                                Exit Sub
                            End If
                            Set lobj_Repos = Nothing
                        End If
                    End If
                    Set Lrs_Find = Nothing
                End If
                CodeRepos = 0
            Next Z
            Set Lrs_Repos = Nothing
        End If
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Delete***
'=========
Private Sub RemovePlanning(ByVal TOURNEE As String, ByVal jour As String, ByVal DateDu As Date, ByVal Colm As Integer)
    Dim LObj_Find       As New PLANNING
    Dim Lrs_Find        As New Recordset
    Dim Existe          As Boolean
    Dim CodePLANNING    As Long
    Dim i               As Integer
On Error GoTo Err
'# Tournee
    If GridValid = True Then
        Set Lrs_Find = LObj_Find.GetCode_PLANNINGByJourTourneeDate(ErrNumber, ErrDescription, ErrSourceDetail, DateDu, jour, TOURNEE, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
            Exit Sub
        End If
        Set LObj_Find = Nothing
        Existe = False
        If Grid_DetPlanning.Rows = 0 Then
            If Not Lrs_Find.EOF Then
                While Not Lrs_Find.EOF
                    CodePLANNING = Lrs_Find.Fields("code")
                    Set LObj_Find = New PLANNING
                    Call LObj_Find.Delete_RestaurerPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, CodePLANNING, "O", LInt_UserId, Date, CNB)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                        Exit Sub
                    End If
                    Set LObj_Find = Nothing
                    Lrs_Find.MoveNext
                Wend
            End If
        Else
            If Not Lrs_Find.EOF Then
                While Not Lrs_Find.EOF
                    CodePLANNING = Lrs_Find.Fields("code")
                    For i = 1 To Grid_DetPlanning.Rows
                        If Grid_DetPlanning.CellText(i, 8) = CodePLANNING Or Grid_DetPlanning.CellText(i, 8) = "" Then
                            Existe = True
                            Exit For
                        Else
                            Existe = False
                        End If
                    Next i
                    If i = Grid_DetPlanning.Rows + 1 Then
                        If Existe = False Then
                            Set LObj_Find = New PLANNING
                            Call LObj_Find.Delete_RestaurerPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, CodePLANNING, "O", LInt_UserId, Date, CNB)
                            If ErrNumber <> 0 Then
                                ErrNumber = 0
                                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                                Exit Sub
                            End If
                            Set LObj_Find = Nothing
                        End If
                    End If
                    Lrs_Find.MoveNext
                    Existe = False
                Wend
            End If
        End If
'# Repos
    Else
        Dim Lrs_Rep     As New Recordset
        Dim LObj_Rep    As New Conducteur
        Dim CodeRepos   As Integer
        Set Lrs_Rep = LObj_Rep.Get_ReposByDate(ErrNumber, ErrDescription, ErrSourceDetail, DateCell(DateDu, Colm), CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
            Exit Sub
        End If
        Set LObj_Rep = Nothing
        Existe = False
        If Grid_DetPlanning.Rows = 0 Then
            If Not Lrs_Rep.EOF Then
                While Not Lrs_Rep.EOF
                    Call LObj_Rep.Delete_Repos(ErrNumber, ErrDescription, ErrSourceDetail, Lrs_Rep("Conducteur"), DateCell(DateDu, Colm), LInt_UserId, CNB)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                        Exit Sub
                    End If
                    Set LObj_Rep = Nothing
                    Lrs_Rep.MoveNext
                Wend
            End If
        Else
            If Not Lrs_Rep.EOF Then
                While Not Lrs_Rep.EOF
                    CodeRepos = Lrs_Rep.Fields("numero")
                    For i = 1 To Grid_DetPlanning.Rows
                        If Grid_DetPlanning.CellText(i, 8) = CodeRepos Or Grid_DetPlanning.CellText(i, 8) = "" Then
                            Existe = True
                            Exit For
                        Else
                            Existe = False
                        End If
                    Next i
                    If i = Grid_DetPlanning.Rows + 1 Then
                        If Existe = False Then
                            Dim LrsCond     As New Recordset
                            Dim LObjCond    As New Conducteur
                            Set LrsCond = LObjCond.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, Lrs_Rep("conducteur"), CNB)
                            If ErrNumber <> 0 Then
                                ErrNumber = 0
                                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                                Exit Sub
                            End If
                            Set LObjCond = Nothing
                            If Not LrsCond.EOF Then
                                Call LObj_Rep.Delete_Repos(ErrNumber, ErrDescription, ErrSourceDetail, LrsCond.Fields("Code"), DateCell(DateDu, Colm), LInt_UserId, CNB)
                                If ErrNumber <> 0 Then
                                    ErrNumber = 0
                                    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                                    Exit Sub
                                End If
                                Set LObj_Rep = Nothing
                            End If
                            Set LrsCond = Nothing
                        End If
                    End If
                    Lrs_Rep.MoveNext
                    Existe = False
                Wend
            End If
        End If
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'==================================================================================================================================
'Imprimer Programme***
'=====================
Private Sub CmdPrint_Click()
    Dim i           As Integer
    Dim J           As Integer
    Dim MsgTest     As Boolean
    Dim CountRow    As Integer
    Dim CountCol    As Integer
    Dim DateDu      As Date
    Dim DateAu      As Date
    Dim Msg         As VbMsgBoxResult
On Error GoTo Err
    With Grid_Planning
        CountRow = .Rows
        CountCol = 8
        MsgTest = False
        DateDu = Cda_DebutPlg.Value
        DateAu = Cda_FinPlg.Value
        ErreurABr = False
        For i = 1 To CountRow
            For J = 2 To CountCol
                If .CellText(i, J) <> "" And .CellText(i, J) <> " " Then
                    Msg = MsgBox("Imprimer le PLANNING en cours...?        " & vbCr & "En Format d'Abreviation...", vbYesNoCancel + vbDefaultButton1 + vbInformation, "Parcano...")
                    If Msg = vbYes Then
                        Call TmpPLANNING(DateDu)
                        If ErreurABr = False Then
                            Call Frm_Rpt_Apercus.PrintOutAndApercu_PLANNING(0, DateDu, DateAu, "Abreviation")
                            Frm_Rpt_Apercus.Show
                        End If
                    ElseIf Msg = vbNo Then
                        Call TmpPLANNING_Normal(DateDu)
                        Call Frm_Rpt_Apercus.PrintOutAndApercu_PLANNING(0, DateDu, DateAu, "Normal")
                        Frm_Rpt_Apercus.Show
                    End If
                    MsgTest = True
                    Exit For
                End If
            Next J
            If MsgTest = True Then Exit For
        Next i
        If MsgTest = False Then MsgBox "Aucun PLANNING, pour Imprimer!...", vbExclamation, App.ProductName
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'Temp PLANNING***
'================
Private Sub TmpPLANNING_Normal(ByVal DatePrint As Date)
    Dim Lobj_Tmp        As New PLANNING
    Dim MaxCount        As Integer
    Dim RXCount         As Integer
    Dim CXCount         As Integer
    Dim XCount          As Integer
    Dim i               As Integer
    Dim J               As Integer
    Dim DateAu          As Date
    Dim TOURNEE         As String
    Dim SChp()          As String
    Dim RCount As Integer
    Dim LUNDI           As String
    Dim MARDI           As String
    Dim MERCREDI        As String
    Dim JEUDI           As String
    Dim VENDREDI        As String
    Dim SAMDI           As String
    Dim DIMANCHE        As String
On Error GoTo Err
    '========================
    'Deleted ALL DataTemp***
    '========================
    Call Lobj_Tmp.Delete_TmpPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
        Exit Sub
    End If
    Set Lobj_Tmp = Nothing
    '================
    'Selected Data***
    '================
    With Grid_Planning
        RXCount = .Rows
        CXCount = 8
        MaxCount = 1
        DateAu = Cda_FinPlg.Value
        For i = 1 To RXCount
            For RCount = 0 To 10
                For J = 2 To CXCount
                    TOURNEE = .CellText(i, 1)
                    If .CellText(i, J) <> "" And .CellText(i, J) <> " " Then
                        SChp = Split(.CellText(i, J), vbCr)
                        XCount = UBound(SChp) + 1
                        If MaxCount < XCount Then MaxCount = XCount
                        If J = 2 Then If RCount < XCount Then LUNDI = Replace(SChp(RCount), "||", "|")
                        If J = 3 Then If RCount < XCount Then MARDI = Replace(SChp(RCount), "||", "|")
                        If J = 4 Then If RCount < XCount Then MERCREDI = Replace(SChp(RCount), "||", "|")
                        If J = 5 Then If RCount < XCount Then JEUDI = Replace(SChp(RCount), "||", "|")
                        If J = 6 Then If RCount < XCount Then VENDREDI = Replace(SChp(RCount), "||", "|")
                        If J = 7 Then If RCount < XCount Then SAMDI = Replace(SChp(RCount), "||", "|")
                        If J = 8 Then If RCount < XCount Then DIMANCHE = Replace(SChp(RCount), "||", "|")
                    End If
                Next J
                If LUNDI <> "" Or MARDI <> "" Or MERCREDI <> "" Or JEUDI <> "" Or VENDREDI <> "" Or SAMDI <> "" Or DIMANCHE <> "" Then
                    Call Lobj_Tmp.Save_TmpPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, DatePrint, DateAu, TOURNEE, LUNDI, MARDI, _
                        MERCREDI, JEUDI, VENDREDI, SAMDI, DIMANCHE, CNB)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                        Exit Sub
                    End If
                    Set Lobj_Tmp = Nothing
                End If
                LUNDI = ""
                MARDI = ""
                MERCREDI = ""
                JEUDI = ""
                VENDREDI = ""
                SAMDI = ""
                DIMANCHE = ""
                TOURNEE = ""
            Next RCount
        Next i
    End With
    With Grid_Repos
        RXCount = .Rows
        CXCount = 8
        MaxCount = 1
        DateAu = Cda_FinPlg.Value
        For i = 1 To RXCount
            For RCount = 0 To 10
                For J = 2 To CXCount
                    TOURNEE = .CellText(i, 1)
                    If .CellText(i, J) <> "" And .CellText(i, J) <> " " Then
                        SChp = Split(.CellText(i, J), vbCr)
                        XCount = UBound(SChp) + 1
                        If MaxCount < XCount Then MaxCount = XCount
                        If J = 2 Then If RCount < XCount Then LUNDI = Replace(SChp(RCount), "||", "|")
                        If J = 3 Then If RCount < XCount Then MARDI = Replace(SChp(RCount), "||", "|")
                        If J = 4 Then If RCount < XCount Then MERCREDI = Replace(SChp(RCount), "||", "|")
                        If J = 5 Then If RCount < XCount Then JEUDI = Replace(SChp(RCount), "||", "|")
                        If J = 6 Then If RCount < XCount Then VENDREDI = Replace(SChp(RCount), "||", "|")
                        If J = 7 Then If RCount < XCount Then SAMDI = Replace(SChp(RCount), "||", "|")
                        If J = 8 Then If RCount < XCount Then DIMANCHE = Replace(SChp(RCount), "||", "|")
                    End If
                Next J
                If LUNDI <> "" Or MARDI <> "" Or MERCREDI <> "" Or JEUDI <> "" Or VENDREDI <> "" Or SAMDI <> "" Or DIMANCHE <> "" Then
                    Call Lobj_Tmp.Save_TmpPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, DatePrint, DateAu, TOURNEE, LUNDI, MARDI, _
                        MERCREDI, JEUDI, VENDREDI, SAMDI, DIMANCHE, CNB)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                        Exit Sub
                    End If
                    Set Lobj_Tmp = Nothing
                End If
                LUNDI = ""
                MARDI = ""
                MERCREDI = ""
                JEUDI = ""
                VENDREDI = ""
                SAMDI = ""
                DIMANCHE = ""
                TOURNEE = ""
            Next RCount
        Next i
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
Private Sub TmpPLANNING(ByVal DatePrint As Date)
    Dim Lobj_Tmp        As New PLANNING
    Dim MaxCount        As Integer
    Dim RXCount         As Integer
    Dim CXCount         As Integer
    Dim XCount          As Integer
    Dim i               As Integer
    Dim J               As Integer
    Dim DateAu          As Date
    Dim TOURNEE         As String
    Dim SChp()          As String
    Dim RCount As Integer
    Dim LUNDI           As String
    Dim MARDI           As String
    Dim MERCREDI        As String
    Dim JEUDI           As String
    Dim VENDREDI        As String
    Dim SAMDI           As String
    Dim DIMANCHE        As String
    Dim Lrs_ABrC        As New Recordset
    Dim LObj_ABrC       As New Conducteur
    Dim Lrs_ABrV        As New Recordset
    Dim LObj_ABrV       As New VEHICULE
    Dim ABrC            As String
    Dim ABrV            As String
    Dim txt             As String
    Dim XCountS         As Integer
    Dim XChps()         As String
    
On Error GoTo Err
    '========================
    'Deleted ALL DataTemp***
    '========================
    Call Lobj_Tmp.Delete_TmpPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
        Exit Sub
    End If
    Set Lobj_Tmp = Nothing
    '================
    'Selected Data***
    '================
    With Grid_Planning
        RXCount = .Rows
        CXCount = 8
        MaxCount = 1
        DateAu = Cda_FinPlg.Value
        For i = 1 To RXCount
            For RCount = 0 To 10
                For J = 2 To CXCount
                    TOURNEE = .CellText(i, 1)
                    If .CellText(i, J) <> "" And .CellText(i, J) <> " " Then
                        SChp = Split(.CellText(i, J), vbCr)
                        XCount = UBound(SChp) + 1
                        If RCount < XCount Then
                          XChps = Split(SChp(RCount), "||")
                          XCountS = UBound(XChps) + 1
                          Set Lrs_ABrC = LObj_ABrC.GetRow_ABrByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, Trim(XChps(0)), CNB)
                          If ErrNumber <> 0 Then
                              ErrNumber = 0
                              MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                              Exit Sub
                          End If
                          Set LObj_ABrC = Nothing
                          If Not Lrs_ABrC.EOF Then
                              If Trim(Lrs_ABrC("Abreviation")) <> "" And Not (IsNull(Lrs_ABrC("Abreviation"))) Then
                                  ABrC = Lrs_ABrC("Abreviation")
                              Else
                                    MsgBox "Le conducteur :  '" & Trim(XChps(0)) & "'  sans Abréviation!..." & vbCr & "Vérifier fichier de Base 'Personnel'", vbExclamation, App.ProductName
                                    ErreurABr = True
                                    Exit Sub
                                    'ABrC = XChps(0)
                              End If
                          Else
                                MsgBox "Le conducteur :  '" & Trim(XChps(0)) & "'  sans Abréviation!..." & vbCr & "Vérifier fichier de Base 'Personnel'", vbExclamation, App.ProductName
                                ErreurABr = True
                                Exit Sub
                                'ABrC = XChps(0)
                          End If
                          Set Lrs_ABrC = Nothing
                          If XCountS > 1 Then
                              Set Lrs_ABrV = LObj_ABrV.GetRow_AbrByMatricule(ErrNumber, ErrDescription, ErrSourceDetail, Trim(XChps(1)), CNB)
                              If ErrNumber <> 0 Then
                                  ErrNumber = 0
                                  MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                                  Exit Sub
                              End If
                              Set LObj_ABrV = Nothing
                              
                              If Not Lrs_ABrV.EOF Then
                                  If Trim(Lrs_ABrV("Abreviation")) <> "" And Not (IsNull(Lrs_ABrV("Abreviation"))) Then
                                      ABrV = Lrs_ABrV("Abreviation")
                                  Else
                                        If Len(Trim(XChps(1))) > 5 Then
                                            MsgBox "Le véhicule :  '" & Trim(XChps(1)) & "'  sans Abréviation!..." & vbCr & "Vérifier fichier de Base 'Véhicule'", vbExclamation, App.ProductName
                                            ErreurABr = True
                                            Exit Sub
                                        Else
                                            ABrV = XChps(1)
                                        End If
                                  End If
                              Else
                                    If Len(Trim(XChps(1))) > 5 Then
                                        MsgBox "Le véhicule :  '" & Trim(XChps(1)) & "'  sans Abréviation!..." & vbCr & "Vérifier fichier de Base 'Véhicule'", vbExclamation, App.ProductName
                                        ErreurABr = True
                                        Exit Sub
                                    Else
                                        ABrV = XChps(1)
                                    End If
                              End If
                              Set Lrs_ABrV = Nothing
                          End If
                          If XCountS = 1 Then
                              txt = ABrC
                          ElseIf XCountS = 2 Then
                              txt = ABrC & " | " & ABrV
                          ElseIf XCountS = 3 Then
                              txt = ABrC & " | " & ABrV & " | " & XChps(2)
                          ElseIf XCountS = 4 Then
                              txt = ABrC & " | " & ABrV & " | " & XChps(2) & " | " & XChps(3)
                          End If
                        '  TxT = Replace(SChp(RCount), "||", "|")
                          If MaxCount < XCount Then MaxCount = XCount
                          If J = 2 Then If RCount < XCount Then LUNDI = txt
                          If J = 3 Then If RCount < XCount Then MARDI = txt
                          If J = 4 Then If RCount < XCount Then MERCREDI = txt
                          If J = 5 Then If RCount < XCount Then JEUDI = txt
                          If J = 6 Then If RCount < XCount Then VENDREDI = txt
                          If J = 7 Then If RCount < XCount Then SAMDI = txt
                          If J = 8 Then If RCount < XCount Then DIMANCHE = txt
                        End If
                    End If
                Next J
                If LUNDI <> "" Or MARDI <> "" Or MERCREDI <> "" Or JEUDI <> "" Or VENDREDI <> "" Or SAMDI <> "" Or DIMANCHE <> "" Then
                    Call Lobj_Tmp.Save_TmpPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, DatePrint, DateAu, TOURNEE, LUNDI, MARDI, _
                        MERCREDI, JEUDI, VENDREDI, SAMDI, DIMANCHE, CNB)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                        Exit Sub
                    End If
                    Set Lobj_Tmp = Nothing
                End If
                LUNDI = ""
                MARDI = ""
                MERCREDI = ""
                JEUDI = ""
                VENDREDI = ""
                SAMDI = ""
                DIMANCHE = ""
                TOURNEE = ""
                txt = ""
                ABrC = ""
                ABrV = ""
            Next RCount
        Next i
    End With
    With Grid_Repos
        RXCount = .Rows
        CXCount = 8
        MaxCount = 1
        DateAu = Cda_FinPlg.Value
        For i = 1 To RXCount
            For RCount = 0 To 10
                For J = 2 To CXCount
                    TOURNEE = .CellText(i, 1)
                    If .CellText(i, J) <> "" And .CellText(i, J) <> " " Then
                        SChp = Split(.CellText(i, J), vbCr)
                        XCount = UBound(SChp) + 1
                        If RCount < XCount Then
                          XChps = Split(SChp(RCount), "||")
                          XCountS = UBound(XChps) + 1
                          Set Lrs_ABrC = LObj_ABrC.GetRow_ABrByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, Trim(XChps(0)), CNB)
                          If ErrNumber <> 0 Then
                              ErrNumber = 0
                              MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                              Exit Sub
                          End If
                          Set LObj_ABrC = Nothing
                          If Not Lrs_ABrC.EOF Then
                              If Trim(Lrs_ABrC("Abreviation")) <> "" And Not (IsNull(Lrs_ABrC("Abreviation"))) Then
                                  ABrC = Lrs_ABrC("Abreviation")
                              Else
                                  ABrC = XChps(0)
                              End If
                          Else
                              ABrC = XChps(0)
                          End If
                          Set Lrs_ABrC = Nothing
                          If XCountS > 1 Then
                              Set Lrs_ABrV = LObj_ABrV.GetRow_AbrByMatricule(ErrNumber, ErrDescription, ErrSourceDetail, Trim(XChps(1)), CNB)
                              If ErrNumber <> 0 Then
                                  ErrNumber = 0
                                  MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                                  Exit Sub
                              End If
                              Set LObj_ABrV = Nothing
                              
                              If Not Lrs_ABrV.EOF Then
                                  If Trim(Lrs_ABrV("Abreviation")) <> "" And Not (IsNull(Lrs_ABrV("Abreviation"))) Then
                                      ABrV = Lrs_ABrV("Abreviation")
                                  Else
                                      ABrV = XChps(1)
                                  End If
                              Else
                                  ABrV = XChps(1)
                              End If
                              Set Lrs_ABrV = Nothing
                          End If
                          If XCountS = 1 Then
                              txt = ABrC
                          ElseIf XCountS = 2 Then
                              txt = ABrC & " | " & ABrV
                          ElseIf XCountS = 3 Then
                              txt = ABrC & " | " & ABrV & " | " & XChps(2)
                          ElseIf XCountS = 4 Then
                              txt = ABrC & " | " & ABrV & " | " & XChps(2) & " | " & XChps(3)
                          End If
                           ' TxT = Replace(SChp(RCount), "||", "|")
                            If MaxCount < XCount Then MaxCount = XCount
                            If J = 2 Then If RCount < XCount Then LUNDI = txt
                            If J = 3 Then If RCount < XCount Then MARDI = txt
                            If J = 4 Then If RCount < XCount Then MERCREDI = txt
                            If J = 5 Then If RCount < XCount Then JEUDI = txt
                            If J = 6 Then If RCount < XCount Then VENDREDI = txt
                            If J = 7 Then If RCount < XCount Then SAMDI = txt
                            If J = 8 Then If RCount < XCount Then DIMANCHE = txt
                        End If
                    End If
                Next J
                If LUNDI <> "" Or MARDI <> "" Or MERCREDI <> "" Or JEUDI <> "" Or VENDREDI <> "" Or SAMDI <> "" Or DIMANCHE <> "" Then
                    Call Lobj_Tmp.Save_TmpPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, DatePrint, DateAu, TOURNEE, LUNDI, MARDI, _
                        MERCREDI, JEUDI, VENDREDI, SAMDI, DIMANCHE, CNB)
                    If ErrNumber <> 0 Then
                        ErrNumber = 0
                        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                        Exit Sub
                    End If
                    Set Lobj_Tmp = Nothing
                End If
                LUNDI = ""
                MARDI = ""
                MERCREDI = ""
                JEUDI = ""
                VENDREDI = ""
                SAMDI = ""
                DIMANCHE = ""
                TOURNEE = ""
                txt = ""
                ABrC = ""
                ABrV = ""
            Next RCount
        Next i
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Sub
'==================================================================================================================================
'Date PLANNING***
'================
Public Function DatePlanning(ByVal DateSearch As Date) As Date
    Dim LObj_Find As New PLANNING, Lrs_Date As New Recordset, Journee As String
On Error GoTo Err
    Journee = UCase(Format(DateSearch, "dddd"))
    If Journee <> "LUNDI" Then
        While Journee <> "LUNDI"
           Set Lrs_Date = LObj_Find.GetDate_NewPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, DateSearch, -1, CNB)
           If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
                Exit Function
            End If
            Set LObj_Find = Nothing
            If Not Lrs_Date.EOF Then
                Journee = UCase(Format((Lrs_Date.Fields("datedebut")), "dddd"))
                DateSearch = Lrs_Date.Fields("datedebut")
                DatePlanning = DateSearch
            End If
            Set Lrs_Date = Nothing
        Wend
    Else
        DatePlanning = DateSearch
    End If
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Function
'Date New PLANNING***
'====================
Private Function DateNewPlanning(ByVal DatePGN As Date) As Date
    Dim LObj_Find As New PLANNING, Lrs_Date As New Recordset
    Dim Journee As String, xDate As Date
On Error GoTo Err
    While Journee <> "LUNDI"
       Set Lrs_Date = LObj_Find.GetDate_NewPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, DatePGN, 1, CNB)
       If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
            Exit Function
        End If
        Set LObj_Find = Nothing
        If Not Lrs_Date.EOF Then
            Journee = UCase(Format((Lrs_Date.Fields("datedebut")), "dddd"))
            DatePGN = Lrs_Date.Fields("datedebut")
            DateNewPlanning = DatePGN
        Else
            DateNewPlanning = DatePGN
        End If
        Set Lrs_Date = Nothing
    Wend
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Function
'Date WEnd***
'============
Private Function DateWEnd(ByVal DatePlanning As Date) As Date
    Dim LObj_Find As New PLANNING, Lrs_Date As New Recordset
On Error GoTo Err
    Set Lrs_Date = LObj_Find.GetDate_WEnd(ErrNumber, ErrDescription, ErrSourceDetail, DatePlanning, 6, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
        Exit Function
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Date.EOF Then DateWEnd = Lrs_Date.Fields("DateWEnd")
    Set Lrs_Date = Nothing
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Function
'Date Cellule***
'===============
Private Function DateCell(ByVal DateSearch As Date, ByVal Colm As Integer) As Date
    Dim LObj_Find As New PLANNING, Lrs_Date As New Recordset
On Error GoTo Err
    Set Lrs_Date = LObj_Find.GetDate_NewPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, DateSearch, Colm - 2, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion, App.ProductName
        Exit Function
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Date.EOF Then DateCell = Lrs_Date.Fields("datedebut")
    Set Lrs_Date = Nothing
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description, vbExclamation, App.ProductName
End Function

