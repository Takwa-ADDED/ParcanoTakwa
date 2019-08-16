VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form Frm_ConsultBV 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Détails bon vidange"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
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
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   10200
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7215
      ScaleWidth      =   10215
      TabIndex        =   9
      Top             =   1560
      Width           =   10215
      Begin VB.ComboBox Cbo_Vidange 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         TabIndex        =   41
         Tag             =   "M"
         Top             =   4440
         Width           =   5055
      End
      Begin VB.ComboBox Cbo_Conducteur 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         TabIndex        =   40
         Tag             =   "M"
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txt_MatriculeStation 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1860
         TabIndex        =   39
         Tag             =   "M"
         Top             =   2820
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   720
         ScaleHeight     =   1095
         ScaleWidth      =   4095
         TabIndex        =   32
         Top             =   3240
         Width           =   4095
         Begin VB.TextBox txt_ville 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1140
            TabIndex        =   35
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txt_adresse 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1140
            TabIndex        =   34
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txt_rsocial 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1140
            TabIndex        =   33
            Top             =   0
            Width           =   2895
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ville :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   735
            TabIndex        =   38
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adresse :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   405
            TabIndex        =   37
            Top             =   360
            Width           =   690
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Raison sociale :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   0
         ScaleHeight     =   1455
         ScaleWidth      =   10455
         TabIndex        =   17
         Top             =   840
         Width           =   10455
         Begin VB.TextBox txt_compteur 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5760
            TabIndex        =   22
            Top             =   0
            Width           =   1215
         End
         Begin VB.TextBox txt_Type 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1860
            TabIndex        =   21
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox txt_libelle 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1860
            TabIndex        =   20
            Top             =   0
            Width           =   2295
         End
         Begin VB.TextBox txt_Energie 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1860
            TabIndex        =   19
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txt_KlmVidange 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1860
            TabIndex        =   18
            Tag             =   "M"
            Top             =   1080
            Width           =   1335
         End
         Begin SToolBox.SDateBox cda_FinAssur 
            Height          =   285
            Left            =   5760
            TabIndex        =   23
            Top             =   360
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_FinVisite 
            Height          =   285
            Left            =   5760
            TabIndex        =   24
            Top             =   720
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4200
            TabIndex        =   31
            Top             =   360
            Width           =   1500
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4200
            TabIndex        =   30
            Top             =   720
            Width           =   1500
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   29
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compteur :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4200
            TabIndex        =   28
            Top             =   0
            Width           =   1500
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   27
            Top             =   720
            Width           =   1800
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   26
            Top             =   0
            Width           =   1500
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NB KM Vidange :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   25
            Top             =   1080
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4080
         ScaleHeight     =   375
         ScaleWidth      =   2895
         TabIndex        =   14
         Top             =   6720
         Width           =   2895
         Begin VB.TextBox txt_Valeur 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   15
            Tag             =   "M"
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valeur :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   960
            TabIndex        =   16
            Top             =   0
            Width           =   555
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   900
         ScaleHeight     =   375
         ScaleWidth      =   2295
         TabIndex        =   11
         Top             =   0
         Width           =   2295
         Begin SToolBox.SDateBox cda_Create 
            Height          =   285
            Left            =   960
            TabIndex        =   12
            Tag             =   "M"
            Top             =   0
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.TextBox txt_Matricule 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1860
         TabIndex        =   10
         Tag             =   "M"
         Top             =   420
         Width           =   2295
      End
      Begin SToolBox.SGrid grid 
         Height          =   1815
         Left            =   1860
         TabIndex        =   42
         Top             =   4800
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3201
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
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox txt_Numero 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   1860
      TabIndex        =   3
      Tag             =   "M"
      Top             =   990
      Width           =   2295
   End
   Begin VB.PictureBox PIC_NFACT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5640
      ScaleHeight     =   495
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ce bon est inseré dans une facture N° : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   2910
      End
      Begin VB.Label LBL_NFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1250"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   1
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Station "
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
      Left            =   840
      TabIndex        =   8
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conducteur"
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
      Height          =   240
      Left            =   420
      TabIndex        =   7
      Top             =   3960
      Width           =   1125
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
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bon de sortie vidange"
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
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3120
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type Vidange"
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
      Height          =   240
      Left            =   255
      TabIndex        =   4
      Top             =   6000
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "Frm_ConsultBV.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   10215
   End
End
Attribute VB_Name = "Frm_ConsultBV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Cbo_Conducteur_GotFocus()
If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
End If
End Sub

Private Sub Cbo_Vidange_GotFocus()
If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
End If
End Sub

Private Sub Cbo_Vidange_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub cmdFindVidange_Click()
Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "LubrifiantVidange"
    .Show
End With

End Sub

Private Sub CmdPrint_Click()
On Error GoTo Err

If MsgBox("Imprimer ce bon        ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Dim F As Form
    Set F = New Frm_Rpt_Apercus
    With F
        .Numero = txt_Numero.Text
        Call .PrintOutAndApercu_BV(0)
        .Show
    End With
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Form_Load()
Me.Height = 9345
Me.Width = 10395
Me.Move 0, 0
InitGrid
Call Affiche_Personnel_Combo(Cbo_Conducteur)
End Sub

Private Sub InitGrid()
With grid
    ' Allow the grid to be grouped, but
    ' don't show the grouping box
    .HideGroupingBox = True
    .AllowGrouping = True
    ' Group rows will be shown by
    ' a gradient underline
    .GroupRowBackColor = vbWindowBackground
    .GroupRowForeColor = vbWindowText
    
    .GridLineColor = vbWindowBackground
    .GridFillLineColor = vbWindowBackground
    .GridLines = True
    
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .DrawFocusRectangle = False
    
    .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Libelle", "Libelle", , , 200
    .AddColumn "Qte", "Qte", , , 60
    .AddColumn "Prix", "Prix.TTC", , , 80
  
    .AddColumn "Q", "", , , 5
    .StretchLastColumnToFit = True
End With



End Sub
Private Function RET_CODE_CONDUCTEUR(txt As String) As String
Dim SQL As String
Dim rs As New ADODB.Recordset
RET_CODE_CONDUCTEUR = ""
SQL = "select code from personnel where libelle = " & SQLText(txt)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    RET_CODE_CONDUCTEUR = rs(0)
End If
rs.Close
End Function

Private Function RET_PRIX_ENERGIE(txt As String) As Double
Dim SQL As String
Dim rs As New ADODB.Recordset
RET_PRIX_ENERGIE = 0
SQL = "select Prix from energie where libelle = " & SQLText(txt)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    RET_PRIX_ENERGIE = rs(0)
End If
rs.Close
End Function

Private Sub Cbo_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub CmdAdd_Click()
On Error GoTo Err
Call ViderZone(FrmBonVidange)
cda_Create.Text = Date
txt_Numero.Text = "Auto"
txt_Matricule.SetFocus
grid.ClearRows
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim vcode As String

On Error GoTo Err

    If PIC_NFACT.Visible = True Then
    MsgBox "Maj impossible", vbInformation
    Exit Sub
    End If

    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "bon vidange", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
        vcode = txt_Numero.Text
        SQL = "Delete from BonVidange where Numero =" & SQLText(vcode)
        CNB.Execute SQL
        txt_Numero.SetFocus
    End If
    
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub


Private Sub CmdFind_Click()
Unload FrmFind
With FrmFind
    .StrSource = "BonVidange"
    .Show
End With
End Sub


Private Sub CmdFindConducteur_Click()
Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "PersonnelVidange"
    .Show
End With
End Sub

Private Sub cmdFindMatricule_Click()
Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "VéhiculeVidange"
    .Show
End With
End Sub

Private Sub cmdFindNumero_Click()
Unload FrmFind
With FrmFind
    .StrSource = "BonVidange"
    .Show
End With
End Sub


Private Sub CmdFindStation_Click()
Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "StationVidange"
    .Show
End With

End Sub

Private Sub CmdSave_Click()

Dim SQL As String
Dim rs As New ADODB.Recordset

    If Left(CheckMandatory(FrmBonVidange), 1) = 1 Then
       Exit Sub
    End If
    
On Error GoTo Err
    If PIC_NFACT.Visible = True Then
    MsgBox "Maj impossible", vbInformation
    Exit Sub
    End If
    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    'Delete l'enregistrement
    vcode = txt_Numero.Text
    CNB.BeginTrans
    SQL = "Delete from BonVidange where Numero =" & SQLText(vcode)
    CNB.Execute SQL
    If vcode = "Auto" Then
    LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "BonVidange")
    If ErrNumber <> 0 Then
       MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
       ErrNumber = 0
       Exit Sub
    End If
    Set LObj_Compteur = Nothing
    'Insertion enregistrement assiette
    txt_Numero.Text = Format(LInt_NumCompteur, "00000")
    End If
    'Insertion enregistrement
    SQL = "Insert into BonVidange  (Numero,DateDoc,Vehicule,Station,Conducteur,Vidange,valeur) values ("
    SQL = SQL & SQLText(txt_Numero.Text)
    SQL = SQL & "," & SQLText(cda_Create.Text)
    SQL = SQL & "," & SQLText(txt_Matricule.Text)
    SQL = SQL & "," & SQLText(txt_MatriculeStation.Text)
    SQL = SQL & "," & SQLText(RET_CODE_CONDUCTEUR(Cbo_Conducteur.Text))
    SQL = SQL & "," & SQLText(RET_CODE_VIDANGE(Cbo_Vidange.Text))
    SQL = SQL & "," & Replace(txt_Valeur.Text, ",", ".")
    SQL = SQL & ")"
    CNB.Execute SQL
    CNB.CommitTrans
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
    txt_Matricule.SetFocus
    End If
Exit Sub
Err:
CNB.RollbackTrans
MsgBox Err.Description, vbInformation
End Sub


Public Sub AfficheRow_Vehicule(ByVal vcode As String)

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim rQ As New ADODB.Recordset

grid.ClearRows
txt_libelle.Text = ""
txt_Type.Text = ""
txt_Energie.Text = ""
txt_compteur.Text = ""
cda_FinAssur.Text = ""
cda_FinVisite.Text = ""
Cbo_Vidange.Text = ""
    
    
SQL = "Select * from vehicule where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    If Not IsNull(rs("Matricule")) Then txt_libelle.Text = rs("Matricule")
    If Not IsNull(rs("marque")) Then txt_Type.Text = rs("TYPE")
    If Not IsNull(rs("Energie")) Then txt_Energie.Text = rs("Energie")
    If Not IsNull(rs("compteur")) Then txt_compteur.Text = rs("compteur")
    If Not IsNull(rs("DAteFinAssur")) Then cda_FinAssur.Text = rs("DAteFinAssur")
    If Not IsNull(rs("DAteFinVisite")) Then cda_FinVisite.Text = rs("DAteFinVisite")
    If Not IsNull(rs("typeVid")) Then Cbo_Vidange.Text = rs("typeVid")
    If Not IsNull(rs("NBKLMvid")) Then txt_KlmVidange.Text = rs("NBKLMvid")
    
    SQL = "select Code from lubrifiant where libelle = " & SQLText(rs("typeVid"))
    rQ.Open SQL, CNB, adOpenKeyset
    If Not rQ.EOF Then
        Call AfficheRow_Lubrifiant(rQ("Code"))
    End If
    rQ.Close
Else
    MsgBox "Code introuvable", vbInformation
    txt_Matricule.SetFocus
    Exit Sub
End If
rs.Close

End Sub


Public Sub AfficheRow_Vehicule_sansPrix(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from vehicule where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    If Not IsNull(rs("Matricule")) Then txt_libelle.Text = rs("Matricule")
    If Not IsNull(rs("marque")) Then txt_Type.Text = rs("TYPE")
    If Not IsNull(rs("Energie")) Then txt_Energie.Text = rs("Energie")
    If Not IsNull(rs("compteur")) Then txt_compteur.Text = rs("compteur")
    If Not IsNull(rs("DAteFinAssur")) Then cda_FinAssur.Text = rs("DAteFinAssur")
    If Not IsNull(rs("DAteFinVisite")) Then cda_FinVisite.Text = rs("DAteFinVisite")
    If Not IsNull(rs("NBKLMvid")) Then txt_KlmVidange.Text = rs("NBKLMvid")
End If
rs.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo erreur
   Dim i As Integer
   Dim MSG ' Déclare la variable.
   ' Définit le texte du message.
   MSG = "Voulez-vous vraiment quitter?"
   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
   If MsgBox(MSG, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
      Cancel = True
   Else
   Unload Me
   End If
   
   Exit Sub
erreur:
   MsgBox Err.Description, 48
End Sub

Private Sub txt_Matricule_GotFocus()
If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
End If
End Sub

Private Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Matricule_LostFocus()
On Error GoTo Err

If Len(Trim(txt_Matricule.Text)) > 0 Then Call AfficheRow_Vehicule(txt_Matricule.Text)


Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub txt_MatriculeStation_GotFocus()
If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
End If
End Sub

Private Sub txt_MatriculeStation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_NbreLitre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub txt_NbreLitre_LostFocus()

Dim P As Double
Dim L As Integer
Dim V As Double

On Error GoTo Err

P = txt_prixLitre.Text
L = Val(txt_NbreLitre.Text)
V = P * L
txt_Valeur.Text = Format(V, "#,##0.000")

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub


Private Sub txt_MatriculeStation_LostFocus()
If Len(Trim(txt_MatriculeStation.Text)) > 0 Then Call AfficheRow_Station(txt_MatriculeStation.Text)
End Sub

Private Sub txt_Numero_GotFocus()
End Sub

Private Sub txt_Numero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Public Sub AfficheRow_Station(ByVal vcode As String)

Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from station where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_MatriculeStation.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
Else
    MsgBox "Code introuvable", vbInformation
    txt_MatriculeStation.SetFocus
    Exit Sub
End If

End Sub


Public Sub AfficheRow(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from BonVidange where Numero = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Numero.Text = rs("Numero")
    If Not IsNull(rs("VEHICULE")) Then txt_libelle.Text = rs("VEHICULE")
    If Not IsNull(rs("STATION")) Then txt_Type.Text = rs("STATION")
    If Not IsNull(rs("DATEDOC")) Then cda_Create.Text = rs("DATEDOC")
    If Not IsNull(rs("CONDUCTEUR")) Then txt_Energie.Text = rs("CONDUCTEUR")
    If Not IsNull(rs("VALEUR")) Then txt_Valeur.Text = Format(rs("VALEUR"), "#,##0.000")
    Call AfficheRow_Vehicule_sansPrix(rs("VEHICULE"))
    Call AfficheRow_Station(rs("STATION"))
    Call AfficheRow_Conducteur(rs("CONDUCTEUR"))
    Call AfficheRow_Lubrifiant(rs("Vidange"))
    Cbo_Vidange.Text = RET_LIBELLE_VIDANGE(rs("Vidange"))
    If rs("Transf") = "O" Then
        LBL_NFact.Caption = rs("NumFact")
        PIC_NFACT.Visible = True
    Else
        PIC_NFACT.Visible = False
    End If
End If
rs.Close

End Sub
Public Sub AfficheRow_Conducteur(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from personnel where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Libelle")) Then Cbo_Conducteur.Text = rs("Libelle")
End If
rs.Close

End Sub
Public Sub AfficheRow_Lubrifiant(ByVal vcode As String)

Dim SQL As String
Dim rs As New ADODB.Recordset
grid.ClearRows
'Charge details
SQL = "SELECT DetLubrifiant.CodeProduit, Produit.Libelle, Produit.Prix,DetLubrifiant.qte"
SQL = SQL & " From DetLubrifiant"
SQL = SQL & " INNER JOIN  Produit ON DetLubrifiant.CodeProduit = Produit.Code"
SQL = SQL & " WHERE DetLubrifiant.CodeLubrifiant = " & SQLText(vcode)
Dim som As Double
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    While Not rs.EOF
        With grid
            .AddRow
            .CellDetails .Rows, 1, rs("CodeProduit")
            .CellDetails .Rows, 2, rs("Libelle")
            .CellDetails .Rows, 3, rs("Qte")
            .CellDetails .Rows, 4, Format(rs("Prix"), "#,##0.000"), DT_RIGHT
        End With
        som = som + (rs("Qte") * rs("Prix"))
        rs.MoveNext
    Wend
End If
rs.Close
txt_Valeur.Text = Format(som, "#,##0.000")

End Sub


Private Function RET_CODE_VIDANGE(txt As String) As String
Dim SQL As String
Dim rs As New ADODB.Recordset
RET_CODE_VIDANGE = ""
SQL = "select code from Lubrifiant where libelle = " & SQLText(txt)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    RET_CODE_VIDANGE = rs(0)
End If
rs.Close
End Function

Private Function RET_LIBELLE_VIDANGE(txt As String) As String
Dim SQL As String
Dim rs As New ADODB.Recordset
RET_LIBELLE_VIDANGE = ""
SQL = "select Libelle from Lubrifiant where Code = " & SQLText(txt)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    RET_LIBELLE_VIDANGE = rs(0)
End If
rs.Close
End Function

