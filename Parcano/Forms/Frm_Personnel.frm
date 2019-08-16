VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Begin VB.Form Frm_Personnel 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   14325
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   2520
      ScaleHeight     =   2985
      ScaleWidth      =   4185
      TabIndex        =   53
      Top             =   3240
      Width           =   4215
      Begin SToolBox.SGrid SGrid1 
         Height          =   3015
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5318
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
         DisableIcons    =   -1  'True
         MaxVisibleRows  =   0
      End
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   4680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab Tab_Personnel 
      Height          =   7455
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   696
      BackColor       =   16777215
      TabCaption(0)   =   "Informations Personnel"
      TabPicture(0)   =   "Frm_Personnel.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Pic_Controlbox"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4080
         ScaleHeight     =   375
         ScaleWidth      =   7215
         TabIndex        =   47
         Top             =   0
         Width           =   7215
         Begin VB.Label Lbl_Supp 
            BackStyle       =   0  'Transparent
            Caption         =   "=> Conducteur est supprimé, Voulez-Vous ré-ajouter?..."
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
            Left            =   120
            TabIndex        =   48
            Top             =   0
            Width           =   5535
         End
         Begin VB.Image Cmd_Supp 
            Height          =   375
            Left            =   5520
            Picture         =   "Frm_Personnel.frx":001C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.PictureBox Pic_Controlbox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6735
         Left            =   120
         ScaleHeight     =   6735
         ScaleWidth      =   11295
         TabIndex        =   15
         Top             =   600
         Width           =   11295
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   0
            ScaleHeight     =   825
            ScaleWidth      =   4065
            TabIndex        =   33
            Top             =   5760
            Width           =   4095
            Begin VB.Label Label19 
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
               TabIndex        =   39
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label24 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "2"
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
               TabIndex        =   38
               Top             =   480
               Width           =   255
            End
            Begin VB.Label Label23 
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
               TabIndex        =   37
               Top             =   240
               Width           =   3615
            End
            Begin VB.Label Label21 
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
               TabIndex        =   36
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label20 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ": 8 Chiffres Maximum..."
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
               TabIndex        =   35
               Top             =   480
               Width           =   3495
            End
            Begin VB.Label Label18 
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
               TabIndex        =   34
               Top             =   0
               Width           =   2055
            End
         End
         Begin VB.TextBox Txt_Abr 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   2
            Top             =   1440
            Width           =   2175
         End
         Begin SToolBox.SCheckBox Chk_Permi 
            Height          =   255
            Left            =   7200
            TabIndex        =   8
            Top             =   3240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Caption         =   "SCheckBox1"
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
         Begin VB.Frame Group_Permi 
            Caption         =   "Informations Permi"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   6360
            TabIndex        =   27
            Top             =   3600
            Width           =   4575
            Begin VB.TextBox txt_lieuPermi 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   2280
               MaxLength       =   30
               TabIndex        =   11
               Top             =   1200
               Width           =   2175
            End
            Begin VB.TextBox txt_permie 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   2220
               MaxLength       =   30
               TabIndex        =   9
               Top             =   240
               Width           =   2175
            End
            Begin MSComCtl2.DTPicker cda_DateLivrPermi 
               Height          =   375
               Left            =   2280
               TabIndex        =   10
               Top             =   720
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarBackColor=   12632256
               Format          =   93388801
               CurrentDate     =   42857
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N° Permi :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   120
               TabIndex        =   30
               Top             =   360
               Width           =   960
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lieu de livraison :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   120
               TabIndex        =   29
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date de livr. permi :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   120
               TabIndex        =   28
               Top             =   840
               Width           =   1920
            End
         End
         Begin VB.Frame Frame1 
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
            Height          =   2775
            Left            =   6360
            TabIndex        =   26
            Top             =   240
            Width           =   4575
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   375
               Left            =   240
               ScaleHeight     =   315
               ScaleWidth      =   3675
               TabIndex        =   31
               Top             =   240
               Width           =   3735
               Begin VB.TextBox Txt_Photo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "MS Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   0
                  MaxLength       =   30
                  TabIndex        =   12
                  Tag             =   "M"
                  Top             =   0
                  Width           =   3620
               End
            End
            Begin SToolBox.SCommand Cmd_Photo 
               Height          =   375
               Left            =   2280
               TabIndex        =   13
               Top             =   2040
               Width           =   1815
               _ExtentX        =   3201
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
            Begin VB.Image Img_Conducteur 
               Height          =   1575
               Left            =   240
               Picture         =   "Frm_Personnel.frx":11D3E
               Stretch         =   -1  'True
               Top             =   840
               Width           =   1935
            End
         End
         Begin VB.CheckBox chk_Actif 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2880
            TabIndex        =   7
            Top             =   3960
            Width           =   255
         End
         Begin VB.TextBox txt_CIN 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   3
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox txt_mobile 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   4
            Top             =   2400
            Width           =   2175
         End
         Begin VB.TextBox txt_fonction 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   2160
            MaxLength       =   30
            TabIndex        =   6
            Top             =   3360
            Width           =   3975
         End
         Begin VB.TextBox txt_Telephone 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   5
            Top             =   2880
            Width           =   2175
         End
         Begin VB.TextBox txt_Nom 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "M"
            Top             =   960
            Width           =   3975
         End
         Begin VB.TextBox txt_Matricule 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Top             =   240
            Width           =   2295
         End
         Begin SToolBox.SCommand cmdFindMatricule 
            Height          =   495
            Left            =   4560
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
            Picture         =   "Frm_Personnel.frx":1221B
            ButtonType      =   1
         End
         Begin VB.Label Label26 
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
            Left            =   1440
            TabIndex        =   46
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label25 
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
            Left            =   1920
            TabIndex        =   45
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " 1"
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
            Left            =   4320
            TabIndex        =   44
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " 2"
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
            Left            =   4320
            TabIndex        =   43
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " 2"
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
            Left            =   4320
            TabIndex        =   42
            Top             =   2760
            Width           =   255
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   " 2"
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
            Left            =   4320
            TabIndex        =   41
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label Label17 
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
            TabIndex        =   40
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Abréviation :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   240
            TabIndex        =   32
            Top             =   1440
            Width           =   1275
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Actif : O/N"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1680
            TabIndex        =   24
            Top             =   3960
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C.I.N :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   240
            TabIndex        =   23
            Top             =   1920
            Width           =   570
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   240
            TabIndex        =   22
            Top             =   2400
            Width           =   750
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Permi :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   6480
            TabIndex        =   21
            Top             =   3240
            Width           =   675
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   240
            TabIndex        =   20
            Top             =   2880
            Width           =   1125
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fonction :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   240
            TabIndex        =   19
            Top             =   3360
            Width           =   945
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nom et prénom :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   360
            TabIndex        =   18
            Top             =   960
            Width           =   1605
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Matricule"
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
            Left            =   360
            TabIndex        =   17
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   11040
      TabIndex        =   49
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
      Picture         =   "Frm_Personnel.frx":1256E
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   10320
      TabIndex        =   50
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
      Picture         =   "Frm_Personnel.frx":126F0
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   10680
      TabIndex        =   51
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
      Picture         =   "Frm_Personnel.frx":12A43
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   9960
      TabIndex        =   52
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
      Picture         =   "Frm_Personnel.frx":12D96
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche personnel"
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
      Left            =   240
      TabIndex        =   25
      Top             =   240
      Width           =   2580
   End
   Begin VB.Image PicBox_Header 
      Height          =   1455
      Left            =   0
      Picture         =   "Frm_Personnel.frx":12F18
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "Frm_Personnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim StrPicture      As String
    Dim ChangPic        As Boolean

Private Sub Form_Load()
Call InitialGrid
Picture4.Visible = False
    Lbl_Supp.Visible = False
    Cmd_Supp.Visible = False
    Group_Permi.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
    ChangPic = False
    StrPicture = ""
    cda_DateLivrPermi.Value = Date
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Dim WidthForm   As Integer
    WidthForm = Frm_Main.ACB_Main.Width
    Tab_Personnel.Width = WidthForm - 3000
    PicBox_Header.Width = WidthForm - 1000
    Pic_Controlbox.Width = WidthForm - 3300
         CmdAdd.Left = WidthForm - 5500
        CmdDelete.Left = WidthForm - 5100
        CmdFind.Left = WidthForm - 4700
        CmdSave.Left = WidthForm - 4300
    Picture3.Left = WidthForm - 10500
End Sub
Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
'    If MsgBox("Voulez-vous vraiment quitter?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Cancel = True Else Unload Me
End Sub
Private Sub EnbDisb(ByVal TYP As Boolean)
    txt_Nom.Enabled = TYP
    txt_CIN.Enabled = TYP
    Txt_Abr.Enabled = TYP
    txt_fonction.Enabled = TYP
    txt_Telephone.Enabled = TYP
    txt_mobile.Enabled = TYP
    txt_permie.Enabled = TYP
    cda_DateLivrPermi.Enabled = TYP
    txt_lieuPermi.Enabled = TYP
    chk_Actif.Enabled = TYP
    Cmd_Photo.Enabled = TYP
    CmdDelete.Enabled = TYP
    CmdSave.Enabled = TYP
    txt_Matricule.Enabled = TYP
    Img_Conducteur.Enabled = TYP
    CmdDelete.Enabled = TYP
    CmdSave.Enabled = TYP
    Chk_Permi.Enabled = TYP
End Sub

Private Sub SGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
Dim VCode
    On Error GoTo Err
    
     VCode = SGrid1.CellText(lRow, 1)
     cmdFindMatricule.SetFocus
     Frm_Personnel.AfficheRow (VCode)
     
     Picture4.Visible = False
     Exit Sub
Err:
        MsgBox Err.Description, vbInformation
End Sub

Private Sub SGrid1_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
Dim VCode
On Error GoTo Err
    VCode = SGrid1.CellText(SGrid1.SelectedRow, 1)
    Select Case KeyCode
        Case vbKeyReturn
            cmdFindMatricule.SetFocus
            Frm_Personnel.AfficheRow (VCode)
            Picture4.Visible = False
        Case vbKeyEscape
            Picture4.Visible = False
    End Select
 Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub SGrid1_LostFocus()
    Picture4.Visible = False
End Sub

'# ControlBox
Private Sub txt_lieuPermi_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° matricule obligatoire      ", vbInformation, App.ProductName
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub Txt_Abr_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° matricule obligatoire      ", vbInformation, App.ProductName
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_CIN_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° matricule obligatoire      ", vbInformation, App.ProductName
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_fonction_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° matricule obligatoire      ", vbInformation, App.ProductName
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_mobile_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° matricule obligatoire      ", vbInformation, App.ProductName
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_Nom_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° matricule obligatoire      ", vbInformation, App.ProductName
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_permie_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° matricule obligatoire      ", vbInformation, App.ProductName
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub txt_telephone_GotFocus()
    If Len(Trim(txt_Matricule.Text)) = 0 Then
        MsgBox "N° matricule obligatoire      ", vbInformation, App.ProductName
        txt_Matricule.SetFocus
    End If
End Sub
Private Sub Txt_Abr_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Len(Txt_Abr.Text) > 7 And KeyAscii <> 8 And KeyAscii <> 127 Then KeyAscii = 0
    If Not (Chr(KeyAscii) Like "[0123456789AZERTYUIOPQSDFGHJKLMWXCVBNazertyuiopqsdfghjklmwxcvbn.]") _
        And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txt_fonction_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txt_lieuPermi_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txt_Nom_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txt_CIN_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Len(txt_CIN.Text) > 7 And KeyAscii <> 8 And KeyAscii <> 127 Then KeyAscii = 0
    If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txt_Telephone_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Len(txt_Telephone.Text) > 7 And KeyAscii <> 8 And KeyAscii <> 127 Then KeyAscii = 0
    If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txt_mobile_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Len(txt_mobile.Text) > 7 And KeyAscii <> 8 And KeyAscii <> 127 Then KeyAscii = 0
    If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub Txt_Abr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_mobile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_Nom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_permie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_telephone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_CIN_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_fonction_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_lieuPermi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If IsNumeric(Trim(txt_Matricule.Text)) Then
    If KeyCode = vbKeyReturn Then
        If Len(Trim(txt_Matricule.Text)) > 0 And Trim(txt_Matricule.Text) <> "Auto" Then FindByCode
    End If
Else
    If KeyCode = vbKeyRight Then
        If Len(Trim(txt_Matricule.Text)) > 0 And Trim(txt_Matricule.Text) <> "Auto" Then FindByCode
    End If
End If
End Sub
Private Sub txt_Matricule_GotFocus()
    Call ViderZone(Frm_Personnel)
    Img_Conducteur.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Personnel\user.jpg")
End Sub
Private Sub txt_Nom_Change()
    Txt_Photo.Text = txt_Nom.Text
End Sub
Private Sub Chk_Permi_Click()
    If Chk_Permi.Value = vbChecked Then
       Group_Permi.Enabled = True
    Else
        Group_Permi.Enabled = False
    End If
End Sub
Private Sub cmdFindMatricule_Click()
    On Error Resume Next
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
    End If
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "Personnel_actif"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName

End Sub
Private Sub FindByCode()
    Dim VCode   As String
    VCode = txt_Matricule.Text
    If Len(Trim(VCode)) > 0 Then AfficheRow (VCode)
End Sub
'# Choisir Photo Conducteur***
Private Sub Img_Conducteur_Click()
    Call Cmd_Photo_Click
End Sub
Private Sub Cmd_Photo_Click()
    With CDlg
        .DialogTitle = "Séléctionner Photo.."
        .FileName = ""
        .Filter = "Image (*.jpg; *.bmp)|*.jpg; *.bmp"
        .ShowOpen
        If Len(Trim(.FileName)) < 1 Then Exit Sub
        Img_Conducteur.Picture = LoadPicture(.FileName)
        ChangPic = True
    End With
End Sub
'# Nouveau
Private Sub CmdAdd_Click()
On Error GoTo Err
    If (CHECK_ACCES("Ins_Personnel", LInt_UserId) = False) Then
        MsgBox "Insertion n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-etre pas des autorisations nécessaires pour ajouter un Conducteur", vbInformation, App.ProductName
        Exit Sub
    End If
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
    End If
    Call EnbDisb(True)
    Call ViderZone(Frm_Personnel)
    Img_Conducteur.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Personnel\user.jpg")
    txt_Matricule.Text = "Auto"
    txt_Nom.SetFocus
    Lbl_Supp.Visible = False
    Cmd_Supp.Visible = False
    cda_DateLivrPermi.Value = Date
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'# Supprission
Private Sub CmdDelete_Click()
    Dim Lobj_Conducteur As New Conducteur
    Dim VCode           As String
On Error GoTo Err
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
            Exit Sub
        Else
            txt_Matricule.SetFocus
            Exit Sub
        End If
    ElseIf txt_Matricule.Text <> "Auto" Then
        If (CHECK_ACCES("Supp_personnel", LInt_UserId) = False) Then
                MsgBox "Supprission n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-etre pas des autorisations nécessaires pour supprime un Conducteur", vbInformation, App.ProductName
            Exit Sub
        End If
    End If
    If MsgBox("Confirmez vous la suppression", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
        VCode = txt_Matricule.Text
        Call Lobj_Conducteur.Delete_Restaurer(ErrNumber, ErrDescription, ErrSourceDetail, VCode, LInt_UserId, "O", CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set Lobj_Conducteur = Nothing
        MsgBox "Conducteur Supprime avec succes!...", vbInformation, App.ProductName
        Call ViderZone(Frm_Personnel)
    On Error Resume Next
        Img_Conducteur.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Personnel\user.jpg")
    On Error GoTo Err
        Call EnbDisb(True)
        Lbl_Supp.Visible = False
        Cmd_Supp.Visible = False
        CmdDelete.Enabled = False
        CmdSave.Enabled = False
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'# Ré-ajouter
Private Sub Cmd_Supp_Click()
    Dim Lobj_Conducteur As New Conducteur
    Dim VCode           As String
On Error GoTo Err
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
            Exit Sub
        Else
            txt_Matricule.SetFocus
            Exit Sub
        End If
    ElseIf txt_Matricule.Text <> "Auto" Then
        If (CHECK_ACCES("Supp_personnel", LInt_UserId) = False) Then
            MsgBox "Ré-ajoute n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-etre pas des autorisations nécessaires pour supprime un Conducteur", vbInformation, App.ProductName
            Exit Sub
        End If
    End If
    If MsgBox("Confirmez vous la ré-ajouter", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
        VCode = txt_Matricule.Text
        Call Lobj_Conducteur.Delete_Restaurer(ErrNumber, ErrDescription, ErrSourceDetail, VCode, LInt_UserId, "N", CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set Lobj_Conducteur = Nothing
        MsgBox "Conducteur Ré-ajouter avec succes!...", vbInformation, App.ProductName
        Call ViderZone(Frm_Personnel)
    On Error Resume Next
        Img_Conducteur.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Personnel\user.jpg")
    On Error GoTo Err
        Call EnbDisb(True)
        Lbl_Supp.Visible = False
        Cmd_Supp.Visible = False
        CmdDelete.Enabled = False
        CmdSave.Enabled = False
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
'# Rechercher***
Private Sub CmdFind_Click()
    On Error Resume Next
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
    End If
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "Personnel_base"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName

End Sub
Public Sub AfficheRow(ByVal VCode As String)
    Dim LOBJ_Cond   As New Conducteur
    Dim Lrs_Cond    As New Recordset
On Error GoTo Err
    Call ViderZone(Frm_Personnel)
    SGrid1.ClearRows
    Picture4.Visible = False
    Set Lrs_Cond = LOBJ_Cond.GetRow_Conducteur_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LOBJ_Cond = Nothing

If Lrs_Cond.RecordCount = 1 Then
        txt_Matricule.Text = Lrs_Cond("Code")
        txt_Nom.Text = Lrs_Cond("Libelle")
        txt_CIN.Text = Lrs_Cond("CIN")
        txt_fonction.Text = Lrs_Cond("Fonction")
        txt_Telephone.Text = Lrs_Cond("telephone")
        txt_mobile.Text = Lrs_Cond("mobile")
        txt_permie.Text = Lrs_Cond("permie")
        If Not (IsNull(Lrs_Cond("Abreviation"))) Then Txt_Abr.Text = Lrs_Cond("Abreviation") Else Txt_Abr.Text = ""
        If Lrs_Cond("datlivr") <> "01/01/1900" Then cda_DateLivrPermi.Value = Lrs_Cond("datlivr")
        txt_lieuPermi.Text = Lrs_Cond("lieulivr")
        chk_Actif.Value = Lrs_Cond("Actif")
    On Error Resume Next
        Txt_Photo.Text = Lrs_Cond("PicBox")
    On Error GoTo Err
        If (txt_lieuPermi.Text <> "" Or txt_permie.Text <> "") Then
            Chk_Permi.Value = vbChecked
            Group_Permi.Enabled = True
        Else
            Chk_Permi.Value = vbUnchecked
            Group_Permi.Enabled = False
        End If
        If Lrs_Cond("Supp") = "O" Then
            Lbl_Supp.Visible = True
            Cmd_Supp.Visible = True
            Call EnbDisb(False)
        Else
            Lbl_Supp.Visible = False
            Cmd_Supp.Visible = False
            Call EnbDisb(True)
        End If
ElseIf Not Lrs_Cond.EOF Then
        While Not Lrs_Cond.EOF
        SGrid1.Redraw = False
            With SGrid1
            .AddRow
            .CellDetails .Rows, .ColumnIndex("Code"), Lrs_Cond.Fields("code"), DT_RIGHT
            .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Cond.Fields("libelle"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("Cin"), Lrs_Cond.Fields("Cin"), DT_RIGHT
                    
            End With
        Lrs_Cond.MoveNext
        Wend
        SGrid1.Redraw = True
        Picture4.Visible = True
        cmdFindMatricule.SetFocus
        If SGrid1.Rows > 0 Then
            SGrid1.SelectedRow = 1
            SGrid1.SetFocus
        End If

    
    Else
        MsgBox "Code introuvable", vbInformation
     
    End If
On Error Resume Next
    Img_Conducteur.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Personnel\" & Lrs_Cond("PicBox"))
    Set Lrs_Cond = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Public Sub InitialGrid()
 With SGrid1
         ' Allow the grid to be grouped, but
      ' don't show the grouping box
      .HideGroupingBox = True
      .AllowGrouping = True

      ' Group rows will be shown by
      ' a gradient underline
      .GroupRowBackColor = vbWindowBackground
      .GroupRowForeColor = vbWindowText
      
      'Couleur de Grid
      .GridLineColor = vbWindowBackground
      .GridFillLineColor = vbWindowBackground
      .GridLines = True
      
      'Ajout des Colonnes
      .SelectionAlphaBlend = True
      .SelectionOutline = True
      .DrawFocusRectangle = False
      
      .AddColumn "Code", "Code"
      .AddColumn "Libelle", "Libelle", , , 140
'      .AddColumn "tht", "Prix _HT"
'      .AddColumn "tva", "Taux_Tva"
      .AddColumn "Cin", "Cin"
      .StretchLastColumnToFit = True

   End With


End Sub
'# Enregistre***
Private Sub CmdSave_Click()
    Dim LOBJ_Cond           As New Conducteur
    Dim Lrs_Cond            As New Recordset
On Error Resume Next
    StrPicture = txt_Nom.Text & ".Bmp"
    '# Vérifier Abreviation...
    If Len(Txt_Abr.Text) < 2 Or Trim(Txt_Abr.Text) = "" Then
        MsgBox "Abreviation Obligatoir!...           ", vbExclamation + vbOKOnly + vbDefaultButton2, App.ProductName
        Exit Sub
    End If
    '# Vérifier CIN / Téléphone / Mobile...
    If (Len(txt_CIN.Text) < 8 And Len(txt_CIN.Text) > 0) Then
        MsgBox "CIN Invalide!...           "
        Exit Sub
    End If
    If (Len(txt_Telephone.Text) < 8 And Len(txt_Telephone.Text) > 0) Then
        MsgBox "Téléphone Invalide!...           "
        Exit Sub
    End If
    If Len(txt_mobile.Text) < 8 Or txt_mobile.Text = "00000000" Then
        MsgBox "Mobile Invalide!...           "
        Exit Sub
    End If
    If Chk_Permi.Value = vbChecked Then
        If cda_DateLivrPermi.Value = "__/__/____" Or txt_permie = "" Or txt_lieuPermi = "" Or cda_DateLivrPermi.Value = "" Then
            MsgBox "Vérifier les informations de permi!..."
            Exit Sub
        End If
    End If
    '# Confirmer Abreviation existe ou non.
    Set Lrs_Cond = LOBJ_Cond.GetAll_Abreviation(ErrNumber, ErrDescription, ErrSourceDetail, Trim(Txt_Abr.Text), Trim(txt_Matricule.Text), CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LOBJ_Cond = Nothing
    If Not Lrs_Cond.EOF Then
        MsgBox "Abreviation existe déja!... Choisir un autre...", vbExclamation, App.ProductName
        Exit Sub
    End If
    Set Lrs_Cond = Nothing
    If txt_Matricule.Text <> "Auto" Then
        Update_Personnel
    Else
        Insert_Personnel
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Private Sub Update_Personnel()
    Dim Lrs_Cond        As New Recordset
    Dim LOBJ_Cond       As New Conducteur
    Dim VCode           As String
On Error Resume Next
    If (CHECK_ACCES("Maj_Personnel", LInt_UserId) = False) Then
        MsgBox "Modification n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier les informartions Conducteur", vbInformation, App.Path
        Exit Sub
    End If
    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
        VCode = txt_Matricule.Text
        'Insertion enregistrement
        Set Lrs_Cond = CreateEmptyRS_Conducteur()
        With Lrs_Cond
            .AddNew
            .Fields("Libelle") = txt_Nom.Text
            .Fields("ABr") = UCase(Txt_Abr.Text)
            If Trim(txt_CIN.Text) <> "" Then .Fields("CIN") = txt_CIN.Text Else .Fields("CIN") = "00000000"
            .Fields("Fonction") = txt_fonction.Text
            If Trim(txt_Telephone.Text) <> "" Then .Fields("Telephone") = txt_Telephone.Text Else .Fields("Telephone") = "00000000"
            .Fields("Mobile") = txt_mobile.Text
            If Chk_Permi.Value = vbChecked Then
                .Fields("Permie") = txt_permie.Text
                .Fields("DateLivr") = cda_DateLivrPermi.Value
                .Fields("LieuLivr") = txt_lieuPermi.Text
            Else
                .Fields("Permie") = ""
                .Fields("DateLivr") = ""
                .Fields("LieuLivr") = ""
            End If
            .Fields("Actif") = chk_Actif.Value
            If ChangPic = True Then
                .Fields("PicBox") = StrPicture
                Txt_Photo = StrPicture
            End If
            .Fields("UserUpdate") = LInt_UserId
        End With
        Set LOBJ_Cond = New Conducteur
        Call LOBJ_Cond.Update_Conducteur(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB, Lrs_Cond)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LOBJ_Cond = Nothing
        Set Lrs_Cond = Nothing
    On Error Resume Next
        If ChangPic = True Then SavePicture Img_Conducteur.Picture, "\\srv-files\Parcano\Image Parcano\Personnel\" & StrPicture
    On Error GoTo Err
        MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
        Call ViderZone(Frm_Personnel)
        Img_Conducteur.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Personnel\user.jpg")
        txt_Matricule.Text = "Auto"
        txt_Nom.SetFocus
        Lbl_Supp.Visible = False
        Cmd_Supp.Visible = False
        Call EnbDisb(True)
        cda_DateLivrPermi.Value = Date
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
Private Sub Insert_Personnel()
    Dim LInt_NumCompteur    As Long
    Dim Lrs_Cond            As New Recordset
    Dim LOBJ_Cond           As New Conducteur
    Dim VCode               As String
On Error Resume Next
    If (CHECK_ACCES("Ins_Personnel", LInt_UserId) = False) Then
        MsgBox "Insertion n'est pas accessible!.." & vbNewLine & " Vous ne disposez peut-etre pas des autorisations nécessaires pour ajouter un Conducteur", vbInformation, App.ProductName
        Exit Sub
    End If
    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
        LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_Personnel")
        If ErrNumber <> 0 Then
           MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
           ErrNumber = 0
           Exit Sub
        End If
        'Insertion enregistrement assiette
        txt_Matricule.Text = Format(LInt_NumCompteur, "00000")
        VCode = Format(LInt_NumCompteur, "00000")
        Set Lrs_Cond = CreateEmptyRS_Conducteur()
        With Lrs_Cond
            .AddNew
            .Fields("Numero") = VCode
            .Fields("Libelle") = txt_Nom.Text
            .Fields("ABr") = UCase(Txt_Abr.Text)
            .Fields("CIN") = txt_CIN.Text
            .Fields("Fonction") = txt_fonction.Text
            .Fields("Telephone") = txt_Telephone.Text
            .Fields("Mobile") = txt_mobile.Text
            If Chk_Permi.Value = vbChecked Then
                .Fields("Permie") = txt_permie.Text
                .Fields("DateLivr") = cda_DateLivrPermi.Value
                .Fields("LieuLivr") = txt_lieuPermi.Text
            Else
                .Fields("Permie") = ""
                .Fields("DateLivr") = ""
                .Fields("LieuLivr") = ""
            End If
            .Fields("Actif") = chk_Actif.Value
            .Fields("Disponible") = "O"
            If ChangPic = True Then
                .Fields("PicBox") = StrPicture
            Else
                .Fields("PicBox") = "User.Jpg"
            End If
            .Fields("UserInsert") = LInt_UserId
        End With
        Call LOBJ_Cond.Insert_Conducteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Cond)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LOBJ_Cond = Nothing
        Set Lrs_Cond = Nothing
    On Error Resume Next
        If ChangPic = True Then SavePicture Img_Conducteur.Picture, "\\srv-files\Parcano\Image Parcano\Personnel\" & StrPicture
    On Error GoTo Err
        MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
        Call ViderZone(Frm_Personnel)
        Img_Conducteur.Picture = LoadPicture("\\srv-files\Parcano\Image Parcano\Personnel\user.jpg")
        txt_Matricule.Text = "Auto"
        txt_Nom.SetFocus
        Lbl_Supp.Visible = False
        Cmd_Supp.Visible = False
        Call EnbDisb(True)
        cda_DateLivrPermi.Value = Date
    End If
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub
