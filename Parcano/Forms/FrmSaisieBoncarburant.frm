VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmSaisieBoncarburant 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saisie ligne bon carburant"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSaisieBoncarburant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_anomali 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11160
      Top             =   6000
   End
   Begin VB.TextBox Txt_Observ 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7560
      TabIndex        =   7
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "détails bon carburant"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1935
      Left            =   120
      TabIndex        =   36
      Top             =   6240
      Width           =   11535
      Begin VB.TextBox txt_NbreLitre 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   430
         Left            =   5040
         MaxLength       =   5
         TabIndex        =   4
         Tag             =   "M"
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         ScaleHeight     =   375
         ScaleWidth      =   1335
         TabIndex        =   61
         Top             =   1440
         Width           =   1335
         Begin VB.TextBox Txt_MoyConsom 
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
            Height          =   375
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.TextBox txt_Valeur 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   405
         Left            =   7440
         MaxLength       =   9
         TabIndex        =   6
         Text            =   "0,000"
         Top             =   360
         Width           =   1575
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   3255
         TabIndex        =   47
         Top             =   360
         Width           =   3255
         Begin VB.TextBox txt_prixLitre 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   64
            Tag             =   "M"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txt_compteur 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   15
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prix TTC  de 1 L :"
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
            TabIndex        =   65
            Top             =   600
            Width           =   1365
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Anc.Compteur :"
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
            TabIndex        =   48
            Top             =   120
            Width           =   1290
         End
      End
      Begin VB.TextBox txt_Ncompteur 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Left            =   5040
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "M"
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Nbre Litre "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   3600
         TabIndex        =   66
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Lbl_MoyConsom 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Moyenne consommation pour 6 mois "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Lbl_Anomalie 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Height          =   1335
         Left            =   9480
         TabIndex        =   60
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valeur :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   6720
         TabIndex        =   54
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Txt_tva 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   51
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label txt_ht 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label LBL_DIF_COMP 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   495
         Left            =   6360
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Lbl_Consommation 
         BackColor       =   &H80000014&
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
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   7440
         TabIndex        =   17
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nouv.Compteur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   3600
         TabIndex        =   37
         Top             =   480
         Width           =   1305
      End
   End
   Begin VB.Frame Fram_vidge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vidange"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1575
      Left            =   120
      TabIndex        =   35
      Top             =   4560
      Width           =   7095
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   6855
         TabIndex        =   43
         Top             =   240
         Width           =   6855
         Begin VB.TextBox txt_KlmVidange 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   13
            Top             =   600
            Width           =   1180
         End
         Begin VB.TextBox Txt_ComptVdg 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   14
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label LBL_VIDANGE 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
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
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   3600
            TabIndex        =   55
            Top             =   360
            Width           =   3045
            WordWrap        =   -1  'True
         End
         Begin VB.Label SDate_vdg 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   5040
            TabIndex        =   52
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   0
            TabIndex        =   46
            Top             =   720
            Width           =   1320
         End
         Begin VB.Image Im_Vid 
            Height          =   240
            Left            =   3240
            Picture         =   "FrmSaisieBoncarburant.frx":0ECA
            Stretch         =   -1  'True
            Top             =   600
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
            Left            =   3000
            TabIndex        =   18
            Top             =   600
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label Lbl_CptVdg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Lbl_date 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Date vidange :"
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
            Left            =   3720
            TabIndex        =   44
            Top             =   0
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Fram_papier 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Papiers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1815
      Left            =   6000
      TabIndex        =   34
      Top             =   1800
      Width           =   4815
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   4575
         TabIndex        =   38
         Top             =   240
         Width           =   4575
         Begin VB.Label cda_fin_tax 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1920
            TabIndex        =   58
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label cda_FinVisite 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1920
            TabIndex        =   57
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label cda_FinAssur 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1920
            TabIndex        =   56
            Top             =   240
            Width           =   1335
         End
         Begin VB.Image Im_tax 
            Height          =   240
            Left            =   3720
            Picture         =   "FrmSaisieBoncarburant.frx":11D4
            Stretch         =   -1  'True
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Im_Vis 
            Height          =   240
            Left            =   3720
            Picture         =   "FrmSaisieBoncarburant.frx":14DE
            Stretch         =   -1  'True
            Top             =   720
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Im_ass 
            Height          =   240
            Left            =   3720
            Picture         =   "FrmSaisieBoncarburant.frx":17E8
            Stretch         =   -1  'True
            Top             =   240
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
            Left            =   3480
            TabIndex        =   19
            Top             =   240
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
            Left            =   3480
            TabIndex        =   20
            Top             =   720
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
            Left            =   3480
            TabIndex        =   21
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date fin taxe :"
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
            TabIndex        =   41
            Top             =   1200
            Width           =   1185
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date fin visite :"
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
            TabIndex        =   40
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date fin assurance :"
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
            Top             =   240
            Width           =   1665
         End
      End
   End
   Begin VB.Frame Fram_Veh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Véhicule"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   2895
      Left            =   120
      TabIndex        =   32
      Top             =   1560
      Width           =   5295
      Begin SToolBox.SBiCombo Cbo_Matricule 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1935
         ScaleWidth      =   5055
         TabIndex        =   42
         Top             =   840
         Width           =   5055
         Begin VB.TextBox Txt_CptVeh 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   12
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txt_Energie 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   11
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox txt_libelle 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   9
            Top             =   120
            Width           =   2415
         End
         Begin VB.TextBox txt_Type 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   10
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Lbl_CptVeh 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Compteur Traffic :"
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
            TabIndex        =   27
            Top             =   1680
            Width           =   1695
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
            Left            =   0
            TabIndex        =   28
            Top             =   240
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
            Left            =   0
            TabIndex        =   26
            Top             =   1200
            Width           =   720
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
            Left            =   0
            TabIndex        =   25
            Top             =   720
            Width           =   510
         End
      End
      Begin SToolBox.SCommand cmdFindMatricule 
         Height          =   345
         Left            =   4440
         TabIndex        =   33
         Top             =   480
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
         Picture         =   "FrmSaisieBoncarburant.frx":1AF2
         ButtonType      =   1
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
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Timer T_ASS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   0
   End
   Begin VB.Timer T_VIS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   720
   End
   Begin VB.Timer T_TAX 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   0
   End
   Begin VB.Timer T_VID 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6000
      Top             =   0
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1680
      ScaleHeight     =   615
      ScaleWidth      =   6975
      TabIndex        =   30
      Top             =   960
      Width           =   6975
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4800
         ScaleHeight     =   375
         ScaleWidth      =   2295
         TabIndex        =   31
         Top             =   0
         Width           =   2295
         Begin SToolBox.SDateBox cda_Create 
            Height          =   285
            Left            =   840
            TabIndex        =   1
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date :"
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
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.TextBox txt_Numero 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   405
         Left            =   1440
         TabIndex        =   0
         Top             =   90
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Numéro bon"
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
         Left            =   0
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11820
      TabIndex        =   29
      Top             =   8250
      Width           =   11820
      Begin VB.CommandButton Cmd_annul 
         Caption         =   "&Annuler"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   6480
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_ok 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   4440
         TabIndex        =   5
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.Label Lbl_anomaliConso 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   10920
      TabIndex        =   59
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Lbl_Obser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Observation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   7560
      TabIndex        =   53
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ligne bon carburant"
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
      TabIndex        =   49
      Top             =   360
      Width           =   3300
   End
   Begin VB.Image PicBox_Header 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmSaisieBoncarburant.frx":1E45
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   12615
   End
End
Attribute VB_Name = "FrmSaisieBoncarburant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Okay As Boolean
Public ii As Integer
Dim flap
Dim flap1
Dim flap2
Dim flap3
Dim anomaliCons As Integer

'Afficher liste des véhicule
Private Sub cmdFindMatricule_Click()

T_ASS.Enabled = False
T_VIS.Enabled = False
T_TAX.Enabled = False
T_VID.Enabled = False
Timer_anomali.Enabled = False
LBL_VIDANGE.Visible = False
LBL_ALERT_ASSURANCE.Visible = False
LBL_ALERT_VIDANGE.Visible = False
LBL_ALERT_VISITE.Visible = False
LBL_ALERT_TAXE.Visible = False
Lbl_Anomalie.Visible = False
Im_ass.Visible = False
Im_Vis.Visible = False
Im_tax.Visible = False
Im_Vid.Visible = False
Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Véhicule"
    .Show vbModal
End With
End Sub

'Charger les données concernat le véhicule dans les champs correspondants
'Faire appel à cette fonction par dbl_click sur la liste des véhicules dans Frmfind_Fils
Public Sub AfficheRow_Vehicule(ByVal VCode As String)

Dim LOBJ_BonCarburant As BonCarburant
Dim LOBJ_BonVidange As BonVidange
Dim Lobj_Vehicule As VEHICULE
Dim Energie As String
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim AA As Long
Dim AnCompteur As Long
Dim Name_Tab As String
Dim id

AA = 0
txt_NbreLitre.Text = "0,00"
txt_Valeur.Text = "0,000"

Set LOBJ_BonCarburant = New BonCarburant
If txt_Numero.Text <> "Auto" Then
    'Return la valeur ancienne du compteurCarb du bon en cours de modification d'un véhicule
    Set rs = LOBJ_BonCarburant.Get_AnComptCar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt_Numero.Text, VCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs.EOF Then AnCompteur = rs("maxCpt")
    rs.Close
End If

Set Lobj_Vehicule = New VEHICULE
Set rs = Lobj_Vehicule.GetVehiculeByCode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    Cbo_Matricule.Text = rs("Code") & "  -  " & rs("Matricule")
    If Not IsNull(rs("Matricule")) Then txt_libelle.Text = rs("Matricule")
    If Not IsNull(rs("TYPE")) Then txt_Type.Text = rs("TYPE")
    If Not IsNull(rs("Energie")) Then txt_Energie.Text = rs("Energie")
    
    'Ancien CompteurCarburant
    If txt_Numero.Text = "Auto" Then
        If Not IsNull(rs("CompteurCarburant")) Then txt_compteur.Text = rs("CompteurCarburant")
    Else
        If Not IsNull(AnCompteur) Then txt_compteur.Text = AnCompteur
    End If
    
    Call RET_PRIX_ENERGIE(rs("Energie"))
    If Not IsNull(rs("CompteurCarburant")) Then txt_Ncompteur.Text = rs("CompteurCarburant")
    
    'Dernier compteur du véhicule en entrant (ficheTraffic)
    Name_Tab = "FicheTraffic"
    Set rs1 = Lobj_Vehicule.Get_DerCompt(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Name_Tab, rs("Code"))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs1.EOF Then
        If Not IsNull(rs1("maxCpt")) Then Txt_CptVeh.Text = rs1("maxCpt")
    End If
    rs1.Close
    ' Afficher le dernier vidange au lieu du type à partir du BonVidange
       'Dernier vidange
    Set LOBJ_BonVidange = New BonVidange
    Set rs1 = LOBJ_BonVidange.Get_DerBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("Code"))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs1.EOF Then
        If Not IsNull(rs1("DateDoc")) Then SDate_vdg.Caption = rs1("DateDoc")
        If Not IsNull(rs1("NBKLMvid")) Then txt_KlmVidange.Text = rs1("NBKLMvid")
        If Not IsNull(rs1("CompteurVidange")) Then
            Txt_ComptVdg.Text = rs1("CompteurVidange")
            'Alert vidange
            AA = txt_Ncompteur.Text - rs1("CompteurVidange")
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
    
    If Not IsNull(rs("DateFinAssur")) And rs("DateFinAssur") <> "01/01/1900" Then cda_FinAssur.Caption = rs("DateFinAssur")
    If Not IsNull(rs("DAteFinVisite")) And rs("DAteFinVisite") <> "01/01/1900" Then cda_FinVisite.Caption = rs("DAteFinVisite")
    
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
    If Not IsNull(rs("DAteFinVisite")) And rs("DAteFinVisite") <> "01/01/1900" Then cda_FinVisite.Caption = rs("DAteFinVisite")
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
    If Not IsNull(rs("DateFinTax")) And rs("DateFinTax") <> "01/01/1900" Then cda_fin_tax.Caption = rs("DateFinTax")
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
    Txt_MoyConsom.Text = Format(Calcule_MoyConsom(VCode), "#0.00")
Else
    MsgBox "Code véhicule introuvable", vbInformation
    Cbo_Matricule.Text = ""
    Cbo_Matricule.SetFocus
    Exit Sub
End If
rs.Close

End Sub

'Affiche détails dans la forme FrmSaisieBoncarburant par Dbclick sur la listeBox du FrmAllBoncarburant
Public Sub AfficheRow_Vehicule_sansPrix(ByVal VCode As String)

Dim LOBJ_BonCarburant As BonCarburant
Dim LOBJ_BonVidange As BonVidange
Dim Lobj_Vehicule As VEHICULE
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim AA As Long
Dim AnCompt As Long
Dim Name_Tab As String

AA = 0
txt_NbreLitre.Text = "0,00"
txt_Valeur.Text = "0,000"

'return Ancien CompteurCarburant pour ce véhicule : le compteurCar du bon avant ce bon
Set LOBJ_BonCarburant = New BonCarburant
If txt_Numero.Text <> "Auto" Then
    Set rs = LOBJ_BonCarburant.Get_AnComptCar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt_Numero.Text, VCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs.EOF Then
        If Not IsNull(rs("maxCpt")) Then AnCompt = rs("maxCpt")
    End If
    rs.Close
End If

Set Lobj_Vehicule = New VEHICULE
Set rs = Lobj_Vehicule.GetVehicule(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    Cbo_Matricule.Text = rs("Code") & " - " & rs("Matricule")
    If Not IsNull(rs("Matricule")) Then txt_libelle.Text = rs("Matricule")
    If Not IsNull(rs("marque")) Then txt_Type.Text = rs("TYPE")
    If Not IsNull(rs("Energie")) Then txt_Energie.Text = rs("Energie")
     
    'Ancien CompteurCarburant
    If txt_Numero.Text = "Auto" Then
        If Not IsNull(rs("CompteurCarburant")) Then txt_compteur.Text = rs("CompteurCarburant")
    Else
        If Not IsNull(AnCompt) Then txt_compteur.Text = AnCompt
    End If
    
    If Not IsNull(rs("DAteFinAssur")) Then cda_FinAssur.Caption = rs("DAteFinAssur")
    If Not IsNull(rs("DAteFinVisite")) Then cda_FinVisite.Caption = rs("DAteFinVisite")
    If Not IsNull(rs("DAteFintax")) Then cda_fin_tax.Caption = rs("DAteFintax")
    'Changer le champs type de vidange par dernier vidange
    If Not IsNull(rs("CompteurCarburant")) Then txt_Ncompteur.Text = rs("CompteurCarburant")
    'Dernier compteur du véhicule en entrant (ficheTraffic)
    Name_Tab = "FicheTraffic"
    Set rs1 = Lobj_Vehicule.Get_DerCompt(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Name_Tab, rs("Code"))
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs1.EOF Then
        If Not IsNull(rs1("maxCpt")) Then Txt_CptVeh.Text = rs1("maxCpt")
    End If
    rs1.Close
    'Date du dernier vidange
    Set LOBJ_BonVidange = New BonVidange
    Set rs1 = LOBJ_BonVidange.Get_DerBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    If Not rs1.EOF Then
        If Not IsNull(rs1("DateDoc")) Then SDate_vdg.Caption = rs1("DateDoc")
        If Not IsNull(rs1("NBKLMvid")) Then txt_KlmVidange.Text = rs1("NBKLMvid")
        If Not IsNull(rs1("CompteurVidange")) Then
            Txt_ComptVdg.Text = rs1("CompteurVidange")
            'Alert vidange
            AA = txt_Ncompteur.Text - rs1("CompteurVidange")
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
    
    'Alert assurance
    If Not IsNull(rs("DateFinAssur")) And rs("DateFinAssur") <> "01/01/1900" Then cda_FinAssur.Caption = rs("DateFinAssur")
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
    If Not IsNull(rs("DAteFinVisite")) And rs("DAteFinVisite") <> "01/01/1900" Then cda_FinVisite.Caption = rs("DAteFinVisite")
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
    If Not IsNull(rs("DateFinTax")) And rs("DateFinTax") <> "01/01/1900" Then cda_fin_tax.Caption = rs("DateFinTax")
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
    Txt_MoyConsom.Text = Format(Calcule_MoyConsom(VCode), "#0.00")
Else
    MsgBox "Code véhicule introuvable", vbInformation
    Cbo_Matricule.SetFocus
    Exit Sub
End If
rs.Close

End Sub

'Commande OK : Charger les données saisie dans FrmAllBonCarburant et faire les calculs
Private Sub Cmd_ok_Click()  'OK

Dim itmX As ListItem
Dim Energie As String
Dim Hiem As Boolean

On Error GoTo Err
'Control même energie
If Left(CheckMandatory(FrmSaisieBoncarburant), 1) = 1 Then
   Exit Sub
End If
Energie = txt_Energie.Text

If Cbo_Matricule.Text = "" Then
    MsgBox "Véhicule obligatoire ! ", vbInformation
    Cbo_Matricule.SetFocus
    Exit Sub
End If

If Val(txt_Ncompteur.Text) < Val(txt_compteur.Text) Then
    MsgBox "Nouveau CompteurCarburant invalid", vbInformation, "Parcano..."
    txt_Ncompteur.SetFocus
    Exit Sub
End If
If Val(txt_Ncompteur.Text) - Val(txt_compteur.Text) > 1200 Then
    If MsgBox("Nouveau CompteurCarburant invalid : Plus que 1200 klm" & vbNewLine & "Vlouez vous malgré ça l'accepter.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        txt_Ncompteur.SetFocus
        Exit Sub
    End If
End If

If txt_NbreLitre <= 0 Then
    MsgBox "Nbr litre invalid", vbInformation
    txt_NbreLitre.SetFocus
    Exit Sub
End If
Call Calcul
Hiem = False
If Okay = True Then
    With FrmAllBonCarburant
        For ii = 1 To .Lsv_Client.ListItems.Count
            If .Lsv_Client.ListItems(ii).SubItems(2) = Cbo_Matricule.FirstValue Then
               Hiem = True
               Exit For
            End If
        Next
        If Hiem = True Then
            If MsgBox("Véhicule existe déja dans ce bon " & vbNewLine & "Voulez vous l'accepté", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
        End If
    End With
    'si on click sur ok avant que
    Dim i
    With FrmAllBonCarburant
        i = .Lsv_Client.ListItems.Count + 1
        Set itmX = .Lsv_Client.ListItems.Add(, , CStr(.txt_Numero.Text))
        itmX.SubItems(1) = CStr(.cda_Create.Caption)
        itmX.SubItems(2) = CStr(get_CodeVeh(txt_libelle.Text))
        itmX.SubItems(3) = CStr(txt_libelle.Text)
        itmX.SubItems(4) = CStr(txt_Energie.Text)
        itmX.SubItems(5) = CStr(txt_compteur.Text)
        itmX.SubItems(6) = CStr(txt_Ncompteur.Text)
        itmX.SubItems(7) = CStr(txt_NbreLitre.Text)
        itmX.SubItems(8) = CStr(txt_prixLitre.Text)
        itmX.SubItems(9) = CStr(txt_Valeur.Text)
        itmX.SubItems(10) = CStr(txt_ht.Caption)
        itmX.SubItems(11) = CStr(Txt_tva.Caption)
        itmX.SubItems(12) = CStr(LBL_DIF_COMP.Caption)
        itmX.SubItems(13) = CStr(Lbl_Consommation.Caption)
        itmX.SubItems(14) = CStr(Txt_Observ.Text)
        itmX.SubItems(15) = CStr(Format(Lbl_anomaliConso.Caption, "#,##0.00"))
        If Lbl_anomaliConso.Caption <> "" Then
            If CDbl(Lbl_anomaliConso.Caption) >= 2 Then
                With .Lsv_Client.ListItems(i)
                    .Bold = True
                    .ListSubItems(3).Bold = True
                    .ListSubItems(6).Bold = True
                    .ListSubItems(7).Bold = True
                    .ListSubItems(13).Bold = True
                    .ListSubItems(15).Bold = True
                    .ForeColor = &HFF&
                    .ListSubItems(3).ForeColor = &HFF&
                    .ListSubItems(6).ForeColor = &HFF&
                    .ListSubItems(7).ForeColor = &HFF&
                    .ListSubItems(13).ForeColor = &HFF&
                    .ListSubItems(15).ForeColor = &HFF&
                End With
            End If
        End If
    End With
Else
    With FrmAllBonCarburant
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(2) = CStr(get_CodeVeh(txt_libelle.Text))
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(3) = CStr(txt_libelle.Text)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(4) = CStr(txt_Energie.Text)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(5) = CStr(txt_compteur.Text)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(6) = CStr(txt_Ncompteur.Text)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(7) = CStr(txt_NbreLitre.Text)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(8) = CStr(txt_prixLitre.Text)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(9) = CStr(txt_Valeur.Text)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(10) = CStr(txt_ht.Caption)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(11) = CStr(Txt_tva.Caption)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(12) = CStr(LBL_DIF_COMP.Caption)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(13) = CStr(Lbl_Consommation.Caption)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(14) = CStr(Txt_Observ.Text)
        .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index).SubItems(15) = CStr(Format(Lbl_anomaliConso.Caption, "#,##0.00"))
        
        If CDbl(Lbl_anomaliConso.Caption) >= 2 Then
            With .Lsv_Client.ListItems(.Lsv_Client.SelectedItem.Index)
                .Bold = True
                .ListSubItems(3).Bold = True
                .ListSubItems(6).Bold = True
                .ListSubItems(7).Bold = True
                .ListSubItems(13).Bold = True
                .ListSubItems(15).Bold = True
                .ForeColor = &HFF&
                .ListSubItems(3).ForeColor = &HFF&
                .ListSubItems(6).ForeColor = &HFF&
                .ListSubItems(7).ForeColor = &HFF&
                .ListSubItems(13).ForeColor = &HFF&
                .ListSubItems(15).ForeColor = &HFF&
            End With
        End If
    
    End With

End If

Unload Me
FrmAllBonCarburant.AppCalcul
FrmAllBonCarburant.Get_Details
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub Cmd_Annul_Click()
If MsgBox("Voulez vous annuler l'opération en cours ?", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
Unload Me
End Sub


Private Sub cbo_Matricule_GotFocus()

LBL_ALERT_ASSURANCE.Visible = False
LBL_ALERT_VISITE.Visible = False
LBL_ALERT_TAXE.Visible = False
LBL_ALERT_VIDANGE.Visible = False

T_ASS.Enabled = False
T_TAX.Enabled = False
T_VID.Enabled = False
T_VIS.Enabled = False

End Sub

Private Sub Cbo_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then Call AfficheRow_Vehicule(Cbo_Matricule.FirstValue)

End Sub


Private Sub Cbo_Matricule_LostFocus()

On Error GoTo Err

If Len(Trim(Cbo_Matricule.FirstValue)) > 0 Then Call AfficheRow_Vehicule(Cbo_Matricule.FirstValue)

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

'Retourne PrixLitre de l'energie utilisé
Private Sub RET_PRIX_ENERGIE(txt As String)

Dim LOBJ_Energie As Energie
Dim rs As New Recordset

Set LOBJ_Energie = New Energie
Set rs = LOBJ_Energie.Get_PrixEnergie(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    txt_prixLitre.Text = Format(rs("prix"), "#,##0.000")
    txt_ht.Caption = Format(rs("tht"), "#,##0.000")
    Txt_tva.Caption = Format(rs("tva"), "#0.00")
End If
rs.Close
End Sub

Private Sub Form_Load()
Call Affiche_Matricule_SBCombo(Cbo_Matricule)
End Sub






Private Sub Timer_anomali_Timer()

Timer_anomali.Enabled = True

'If Lbl_Anomalie.Visible = True Then
'    Lbl_Anomalie.Visible = False
'Else
'    Lbl_Anomalie.Visible = True
'End If
End Sub

Private Sub txt_NbreLitre_GotFocus()
On Error Resume Next
If Cbo_Matricule.Text = "" Then
    MsgBox "Véhicule obligatoire ! ", vbInformation
    Cbo_Matricule.SetFocus
    Exit Sub
End If
txt_NbreLitre.SelStart = 0
txt_NbreLitre.SelLength = Len(txt_NbreLitre.Text)
End Sub

Private Sub txt_NbreLitre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
'    Call txt_NbreLitre_LostFocus
   
End If
End Sub

Private Sub txt_NbreLitre_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub

'Calcul de nombre de kilometre parcourus et de consomation d'energie par 100 km
Public Sub Calcul()

Dim P As Double
Dim L As Double
Dim V As Double
Dim Consom As Double
Dim NbKM As Long

On Error GoTo Err
P = txt_prixLitre.Text

If txt_NbreLitre.Text = "" Then txt_NbreLitre.Text = "0,00"
If Val(txt_Valeur.Text) <> 0 Then
    V = CDbl(txt_Valeur.Text)
    L = V / P
    txt_NbreLitre.Text = Format(L, "#,##0.00")
Else
    L = CDbl(txt_NbreLitre.Text)
    V = P * L
End If
txt_Valeur.Text = CStr(Format(V, "#,##0.000"))

NbKM = txt_Ncompteur.Text - txt_compteur.Text
LBL_DIF_COMP.Caption = NbKM & " KM "
Consom = Calcule_Consommation(txt_NbreLitre.Text, NbKM)
Lbl_Consommation.Caption = Format(Consom, "#,##0.00") & " L/100km"
'la différence entre la consommation de L/100km dans ce bon et la moyenne durant 6 mois
'Si la différence est - : c'est que la consommation est inféreur à la moyenne
'si la différence > 2 donc il y'a une anomalie
If CDbl(Txt_MoyConsom.Text) <> 0 Then
    Lbl_anomaliConso.Caption = CStr(Format(Consom - CDbl(Txt_MoyConsom.Text), "#,##0.00"))
    Call Anomalie
Else
    Lbl_anomaliConso.Caption = CStr(Format(0, "#,##0.00"))
End If

Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Public Sub Anomalie()
If Lbl_anomaliConso.Caption <> "" Then
    If CDbl(Lbl_anomaliConso.Caption) >= 2 Then
       Lbl_Anomalie.Caption = "Moyenne de consommation dépassée par " & Lbl_anomaliConso.Caption & " L/100Km"
       Lbl_Anomalie.Visible = True
    '   Call Timer_anomali_Timer
    Else
    '    Timer_anomali.Enabled = False
        Lbl_Anomalie.Caption = ""
        Lbl_Anomalie.Visible = False
    End If
End If
End Sub

Private Sub txt_NbreLitre_LostFocus()

Dim P As Double
Dim L As Double
Dim V As Double
Dim Consom As Double
Dim NbKM As Long

On Error GoTo Err
If Cbo_Matricule.Text = "" Then
    Cbo_Matricule.SetFocus
    Exit Sub
End If
    P = txt_prixLitre.Text
    L = txt_NbreLitre.Text
    V = P * L
    txt_Valeur.Text = Format(V, "#,##0.000")
    NbKM = txt_Ncompteur.Text - txt_compteur.Text
    Consom = Calcule_Consommation(txt_NbreLitre.Text, NbKM)
    Lbl_Consommation.Caption = CStr(Format(Consom, "#,##0.00")) & " L/100km"
    
    If CDbl(Txt_MoyConsom.Text) <> 0 Then
        Lbl_anomaliConso.Caption = CStr(Format(Consom - CDbl(Txt_MoyConsom.Text), "#,##0.00"))
        Call Anomalie
    End If

Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub txt_Ncompteur_GotFocus()

On Error Resume Next

If Cbo_Matricule.Text = "" Then
    MsgBox "Véhicule obligatoire ! ", vbInformation
    Cbo_Matricule.SetFocus
    Exit Sub
End If
If (txt_Ncompteur = "") Then txt_Ncompteur.Text = "0000"

txt_Ncompteur.SelStart = Len(txt_Ncompteur.Text)
'txt_Ncompteur.SelLength = Len(txt_Ncompteur.Text)

End Sub

Private Sub txt_Ncompteur_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call txt_Ncompteur_LostFocus
End If

End Sub

Private Sub txt_Ncompteur_KeyPress(KeyAscii As Integer)

On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub

Private Sub txt_Ncompteur_LostFocus()

If Cbo_Matricule.Text = "" Then
    Cbo_Matricule.SetFocus
    Exit Sub
End If

If Val(txt_Ncompteur.Text) < Val(txt_compteur.Text) Then
    MsgBox "Nouveau CompteurCarburant invalid", vbInformation
    Exit Sub
End If
If Val(txt_Ncompteur.Text) - Val(txt_compteur.Text) > 1200 Then
    If MsgBox("Nouveau CompteurCarburant invalid : Plus que 1200 klm" & vbNewLine & "Voulez vous malgré ça l'accepter.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
End If

Call Calcul

Exit Sub
End Sub

Private Sub txt_Numero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Valeur_GotFocus()
On Error Resume Next
If Cbo_Matricule.Text = "" Then
    MsgBox "Véhicule obligatoire ! ", vbInformation
    Cbo_Matricule.SetFocus
    Exit Sub
End If
txt_Valeur.SelStart = 0
txt_Valeur.SelLength = Len(txt_Valeur.Text)

End Sub

Private Sub txt_Valeur_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Valeur_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub

Private Sub txt_Valeur_LostFocus()

On Error GoTo Err

If Cbo_Matricule.Text = "" Then
    Cbo_Matricule.SetFocus
    Exit Sub
End If

If txt_Valeur.Text = "" Then txt_Valeur.Text = "0,000"

Call Calcul
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Public Function Calcule_Consommation(ByVal NbLitre As Long, ByVal Kilometrage As Long) As Double

If Kilometrage <= 0 Then
    Calcule_Consommation = 0
Else
    Calcule_Consommation = CDbl(Format((NbLitre * 100 / Kilometrage), "#,##0.00"))
End If
End Function

Public Function get_CodeVeh(ByVal VCode As String) As String

Dim Lobj_Vehicule As VEHICULE
Dim rs As New Recordset

Set Lobj_Vehicule = New VEHICULE
Set rs = Lobj_Vehicule.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
If Not rs.EOF Then
    get_CodeVeh = rs("Code")
End If
rs.Close

End Function

'Calculer la moyenne de consommation du carburant par 100 km pour un véhicule durant 6 mois
Public Function Calcule_MoyConsom(ByVal VEHICULE As String) As Double

Dim LOBJ_BonCarburant As BonCarburant
Dim rs As New Recordset
Dim consomation As Double
Dim VCode As String
Dim i As Long

consomation = 0
i = 0

'vcode = get_CodeVeh(Vehicule)
Set LOBJ_BonCarburant = New BonCarburant
Set rs = LOBJ_BonCarburant.Get_MoyConsom(ErrNumber, ErrDescription, ErrSourceDetail, CNB, CDate(cda_Create.Text), VEHICULE)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
If Not rs.EOF Then
    While Not rs.EOF
        consomation = Format(CDbl(consomation) + CDbl(rs("Consom")), "#,##0.000")
        i = i + 1
        rs.MoveNext
    Wend
Calcule_MoyConsom = Format(CDbl(consomation / i), "#,##0.00")
Else
    Calcule_MoyConsom = 0
End If
rs.Close

End Function

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

'Private Sub txt_Valeur_Change()
'
'Dim P As Double
'Dim L As Double
'Dim V As Double
'Dim NbKM As Long
'
'On Error GoTo Err
'If Not (txt_Numero.Text = "Auto") Then
'    If txt_NbreLitre.Text = "" Then txt_NbreLitre.Text = "0,00"
''    P = txt_prixLitre.Text
''    L = (txt_NbreLitre.Text)
''    V = P * L
''    txt_Valeur.Text = Format(V, "#,##0.000")
''    NbKM = txt_Ncompteur.Text - txt_compteur.Text
''    Lbl_Consommation.Caption = CStr(Format(Calcule_Consommation(txt_NbreLitre.Text, NbKM), "#,#.0")) & " L/100km"
'
'    Call Calcul
'End If
'Exit Sub
'
'Err:
'MsgBox Err.Description, vbInformation
'
'End Sub
