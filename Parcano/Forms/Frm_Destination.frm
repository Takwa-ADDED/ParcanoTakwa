VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form Frm_Destination 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Destinations"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9915
   ScaleWidth      =   13290
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab TabDestination 
      Height          =   8655
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Information Destination"
      TabPicture(0)   =   "Frm_Destination.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Pic_Controlbox"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox Pic_Controlbox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8175
         Left            =   120
         ScaleHeight     =   8175
         ScaleWidth      =   11655
         TabIndex        =   13
         Top             =   360
         Width           =   11655
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   3255
            Left            =   3240
            ScaleHeight     =   3225
            ScaleWidth      =   5985
            TabIndex        =   34
            Top             =   1920
            Width           =   6015
            Begin SToolBox.SGrid SGrid1 
               Height          =   3855
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   6800
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
               DisableIcons    =   -1  'True
               MaxVisibleRows  =   0
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   3240
            ScaleHeight     =   1065
            ScaleWidth      =   4185
            TabIndex        =   29
            Top             =   6000
            Width           =   4215
            Begin SToolBox.SOptionButton Opt_Jour 
               Height          =   195
               Left            =   240
               TabIndex        =   30
               Top             =   840
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   344
               BackStyle       =   0
               Caption         =   "Journée"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin SToolBox.SOptionButton Opt_soir 
               Height          =   195
               Left            =   240
               TabIndex        =   31
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   344
               BackStyle       =   0
               Caption         =   "Après_midi"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin SToolBox.SOptionButton Opt_matin 
               Height          =   375
               Left            =   240
               TabIndex        =   32
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               BackStyle       =   0
               Caption         =   "Matin"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.TextBox Txt_Ord 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "0"
            Top             =   5280
            Width           =   2055
         End
         Begin VB.TextBox Txt_Destination 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   3240
            MaxLength       =   30
            TabIndex        =   2
            Top             =   2520
            Width           =   4095
         End
         Begin VB.TextBox Txt_MinCompteur 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            MaxLength       =   30
            TabIndex        =   4
            Text            =   "1"
            Top             =   3840
            Width           =   2055
         End
         Begin VB.TextBox Txt_MaxCompteur 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            MaxLength       =   30
            TabIndex        =   5
            Text            =   "1"
            Top             =   4560
            Width           =   2055
         End
         Begin VB.ComboBox cb_Type 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3240
            TabIndex        =   1
            Top             =   1920
            Width           =   4095
         End
         Begin VB.TextBox txt_Numero 
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
            Height          =   615
            Left            =   3240
            TabIndex        =   0
            Text            =   "Auto"
            Top             =   1080
            Width           =   3615
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   375
            Left            =   5280
            TabIndex        =   14
            Top             =   4560
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin SToolBox.STimeBox Txt_MaxDuree 
            Height          =   285
            Left            =   3240
            TabIndex        =   3
            Top             =   3240
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            BackColor       =   14737632
         End
         Begin SToolBox.SCheckBox chk_Actif 
            Height          =   375
            Left            =   3240
            TabIndex        =   15
            Top             =   7320
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Value           =   1
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
            BackColor       =   -2147483633
            BackStyle       =   0
         End
         Begin SToolBox.SCommand cmdFindReparation 
            Height          =   615
            Left            =   6840
            TabIndex        =   16
            Top             =   1080
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   1085
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
            Picture         =   "Frm_Destination.frx":001C
            ButtonType      =   1
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   375
            Left            =   5280
            TabIndex        =   17
            Top             =   3840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Lbl_tmp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Temps"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   960
            TabIndex        =   33
            Top             =   6360
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ordre"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   960
            TabIndex        =   28
            Top             =   5280
            Width           =   2175
         End
         Begin VB.Image Cmd_ReAjouter 
            Height          =   375
            Left            =   6600
            Picture         =   "Frm_Destination.frx":036F
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Distance Min "
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   960
            TabIndex        =   27
            Top             =   3960
            Width           =   2175
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Km"
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
            Left            =   5520
            TabIndex        =   26
            Top             =   3960
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Km"
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
            Left            =   5520
            TabIndex        =   25
            Top             =   4680
            Width           =   855
         End
         Begin VB.Label Lbl_RaAjouter 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination est déja supprime, Voulez-vous ré-ajouter?..."
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
            Left            =   1080
            TabIndex        =   24
            Top             =   480
            Width           =   5535
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Désignantion"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   960
            TabIndex        =   23
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Actif?"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   2040
            TabIndex        =   22
            Top             =   7320
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Numero°"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   960
            TabIndex        =   21
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   960
            TabIndex        =   20
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Distance Max "
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   960
            TabIndex        =   19
            Top             =   4560
            Width           =   2175
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Max Durée"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   960
            TabIndex        =   18
            Top             =   3240
            Width           =   2175
         End
      End
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   6720
      TabIndex        =   10
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
      Picture         =   "Frm_Destination.frx":12091
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   6000
      TabIndex        =   8
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
      Picture         =   "Frm_Destination.frx":12213
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   6360
      TabIndex        =   9
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
      Picture         =   "Frm_Destination.frx":12566
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   5640
      TabIndex        =   7
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
      Picture         =   "Frm_Destination.frx":128B9
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Destinations"
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
      TabIndex        =   11
      Top             =   360
      Width           =   2535
   End
   Begin VB.Image PicBox_Header 
      Height          =   1095
      Left            =   0
      Picture         =   "Frm_Destination.frx":12A3B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Frm_Destination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim thekey As Integer
    Dim theshift As Integer

'===================================================================================================================================
'Chargement la Forme***
'===================================================================================================================================
Private Sub Form_Load()
    Call InitialGrid
    Picture2.Visible = False
    Lbl_RaAjouter.Visible = False
    Cmd_ReAjouter.Visible = False
    Txt_MaxDuree.Text = ""
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
End Sub
Private Sub Form_Resize()
    Dim WidthForm As Integer
    WidthForm = Frm_Main.ACB_Main.Width
        PicBox_Header.Width = WidthForm - 1000
        TabDestination.Width = WidthForm - 3000
        Pic_Controlbox.Width = WidthForm - 3300
        CmdAdd.Left = WidthForm - 5500
        CmdDelete.Left = WidthForm - 5100
        CmdFind.Left = WidthForm - 4700
        CmdSave.Left = WidthForm - 4300
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    Dim Msg As String
'
'On Error GoTo erreur
'
'    Msg = "Voulez-vous vraiment quitter?"
'    If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'        Cancel = True
'    Else
'        Unload Me
'    End If
'
'Exit Sub
'erreur:
'   MsgBox Err.Description, 48
End Sub
'=================================================
'ControlBox***
'=================================================
Private Sub cb_Type_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Txt_Destination.SetFocus
End Sub
Private Sub cb_Type_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub

Private Sub SGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
Dim VCode
On Error GoTo Err

 VCode = SGrid1.CellText(lRow, 1)
 AfficheRow (VCode)
 cmdFindReparation.SetFocus
 Picture2.Visible = False
 Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub SGrid1_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
Dim VCode
On Error GoTo Err
    VCode = SGrid1.CellText(SGrid1.SelectedRow, 1)
    Select Case KeyCode
        Case vbKeyF2, vbKeyReturn
                    AfficheRow (VCode)
                    cmdFindReparation.SetFocus
                    Picture2.Visible = False
        Case vbKeyEscape
        Picture2.Visible = False
        End Select
 Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub SGrid1_LostFocus()
    Picture2.Visible = False
End Sub

Private Sub Txt_MaxCompteur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub Txt_MaxCompteur_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
    If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub Txt_MaxDuree_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub Txt_minCompteur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub Txt_minCompteur_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
    If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txt_Numero_GotFocus()
    Call ViderZone(Frm_Destination)
    Txt_MaxCompteur.Text = 1
    Txt_MinCompteur.Text = 1
End Sub
Private Sub txt_Numero_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
'    If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub UpDown1_DownClick()
    If Val(Txt_MaxCompteur.Text) > 1 Then Txt_MaxCompteur.Text = Val(Txt_MaxCompteur.Text) - 1
End Sub
Private Sub UpDown1_UpClick()
    Txt_MaxCompteur.Text = Val(Txt_MaxCompteur.Text) + 1
End Sub
Private Sub UpDown2_DownClick()
    If Val(Txt_MinCompteur.Text) > 1 Then Txt_MinCompteur.Text = Val(Txt_MinCompteur.Text) - 1
End Sub
Private Sub UpDown2_UpClick()
    Txt_MinCompteur.Text = Val(Txt_MinCompteur.Text) + 1
End Sub
Private Sub txt_Numero_KeyDown(KeyCode As Integer, Shift As Integer)
If IsNumeric(Trim(txt_Numero.Text)) Then
   If KeyCode = vbKeyReturn Then
        On Error GoTo Err
        If Len(Trim(txt_Numero.Text)) > 0 Then Call AfficheRow(txt_Numero.Text)
   End If
   Exit Sub
Else
    If KeyCode = vbKeyRight Then
        On Error GoTo Err
        If Len(Trim(txt_Numero.Text)) > 0 Then Call AfficheRowbylib(txt_Numero.Text)
    End If
Exit Sub
End If
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub CmdFind_Click()
    Unload FrmFind_Fils
    With FrmFind_Fils
        .StrSource = "Destination"
        .Show vbModal
    End With
End Sub
Private Sub cmdFindReparation_Click()
    Unload FrmFind
With FrmFind_Actif
    .StrSource = "Destination"
    .Show vbModal
End With
End Sub
Private Sub cb_Type_GotFocus()

On Error GoTo Err

    If Len(Trim(txt_Numero.Text)) = 0 Then
        MsgBox "N° bon obligatoire      ", vbInformation
        txt_Numero.SetFocus
        Else
        Call Affiche_Type_Combo(cb_Type)
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'=================================================
'Afficher Destination***
'=================================================
Public Sub AfficheRow(ByVal VCode As String)
    Dim Lobj_Destination As DESTINATION
    Dim Lrs_Destination As Recordset
    
On Error GoTo Err
    
    Call ViderZone(Frm_Destination)
    Opt_Jour.Value = vbUnchecked
    Opt_Matin.Value = vbUnchecked
    Opt_soir.Value = vbUnchecked
    
    Set Lobj_Destination = New DESTINATION
    Set Lrs_Destination = Lobj_Destination.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set Lobj_Destination = Nothing
    
    If Not Lrs_Destination.EOF Then
        'Charge
        txt_Numero.Text = Lrs_Destination("Numero")
        cb_Type.Text = Lrs_Destination("Type")
        Txt_Destination.Text = Lrs_Destination("Libelle")
        Txt_MaxCompteur.Text = Lrs_Destination("MaxCompteur")
        Txt_MinCompteur.Text = Lrs_Destination("MinCompteur")
        Txt_MaxDuree.Text = Format(Lrs_Destination("MaxDuree"), "hh:mm:ss")
        chk_Actif.Value = Lrs_Destination("actif")
        If Lrs_Destination("Temps") = "Après midi" Then
            Opt_soir.Value = vbChecked
        ElseIf Lrs_Destination("Temps") = "Matin" Then
            Opt_Matin.Value = vbChecked
        ElseIf Lrs_Destination("Temps") = "Journée" Then
            Opt_Jour.Value = vbChecked
        End If
        
        If Lrs_Destination("Supp") = "O" Then
            Lbl_RaAjouter.Visible = True
            Cmd_ReAjouter.Visible = True
            txt_Numero.Enabled = False
            cb_Type.Enabled = False
            Txt_Destination.Enabled = False
            chk_Actif.Enabled = False
            Txt_MaxCompteur.Enabled = False
            Txt_MinCompteur.Enabled = False
            Txt_MaxDuree.Enabled = False
            CmdSave.Enabled = False
            CmdDelete.Enabled = False
        Else
            Lbl_RaAjouter.Visible = False
            Cmd_ReAjouter.Visible = False
            txt_Numero.Enabled = True
            cb_Type.Enabled = True
            Txt_Destination.Enabled = True
            chk_Actif.Enabled = True
            Txt_MaxCompteur.Enabled = True
            Txt_MinCompteur.Enabled = True
            Txt_MaxDuree.Enabled = True
            CmdSave.Enabled = True
            CmdDelete.Enabled = True
        End If
    Else
        MsgBox "Code introuvable", vbInformation
        Call ViderZone(Frm_Destination)
        Lbl_RaAjouter.Visible = False
        Cmd_ReAjouter.Visible = False
        txt_Numero.Enabled = True
        cb_Type.Enabled = True
        Txt_Destination.Enabled = True
        chk_Actif.Enabled = True
        Txt_MaxCompteur.Enabled = True
        Txt_MinCompteur.Enabled = True
        Txt_MaxDuree.Enabled = True
        CmdSave.Enabled = True
        CmdDelete.Enabled = True
        txt_Numero.Text = "Auto"
        cb_Type.SetFocus
        Txt_MaxCompteur.Text = 1
        Txt_MinCompteur.Text = 1
        Txt_MaxDuree.Text = ""
        Exit Sub
    End If
    Set Lrs_Destination = Nothing

Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub
'=================================================
'Afficher Destination par libelle***
'=================================================
Public Sub AfficheRowbylib(ByVal VCode As String)
    Dim Lobj_Destination As DESTINATION
    Dim Lrs_Destination As Recordset
    
On Error GoTo Err
    
    Call ViderZone(Frm_Destination)
    SGrid1.ClearRows
    Opt_Jour.Value = vbUnchecked
    Opt_Matin.Value = vbUnchecked
    Opt_soir.Value = vbUnchecked
    
    Set Lobj_Destination = New DESTINATION
    Set Lrs_Destination = Lobj_Destination.GetRow_Destination_ByLib(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set Lobj_Destination = Nothing
    
    If Lrs_Destination.RecordCount = 1 Then
        'Charge
        txt_Numero.Text = Lrs_Destination("Numero")
        cb_Type.Text = Lrs_Destination("Type")
        Txt_Destination.Text = Lrs_Destination("Libelle")
        Txt_MaxCompteur.Text = Lrs_Destination("MaxCompteur")
        Txt_MinCompteur.Text = Lrs_Destination("MinCompteur")
        Txt_MaxDuree.Text = Format(Lrs_Destination("MaxDuree"), "hh:mm:ss")
        chk_Actif.Value = Lrs_Destination("actif")
        If Lrs_Destination("Temps") = "Après midi" Then
            Opt_soir.Value = vbChecked
        ElseIf Lrs_Destination("Temps") = "Matin" Then
            Opt_Matin.Value = vbChecked
        ElseIf Lrs_Destination("Temps") = "Journée" Then
            Opt_Jour.Value = vbChecked
        End If
        
        If Lrs_Destination("Supp") = "O" Then
            Lbl_RaAjouter.Visible = True
            Cmd_ReAjouter.Visible = True
            txt_Numero.Enabled = False
            cb_Type.Enabled = False
            Txt_Destination.Enabled = False
            chk_Actif.Enabled = False
            Txt_MaxCompteur.Enabled = False
            Txt_MinCompteur.Enabled = False
            Txt_MaxDuree.Enabled = False
            CmdSave.Enabled = False
            CmdDelete.Enabled = False
        
        Else
            Lbl_RaAjouter.Visible = False
            Cmd_ReAjouter.Visible = False
            txt_Numero.Enabled = True
            cb_Type.Enabled = True
            Txt_Destination.Enabled = True
            chk_Actif.Enabled = True
            Txt_MaxCompteur.Enabled = True
            Txt_MinCompteur.Enabled = True
            Txt_MaxDuree.Enabled = True
            CmdSave.Enabled = True
            CmdDelete.Enabled = True
        End If
    ElseIf Not Lrs_Destination.EOF Then
    While Not Lrs_Destination.EOF
    SGrid1.Redraw = False
    With SGrid1
        .AddRow
        .CellDetails .Rows, .ColumnIndex("Numero"), Lrs_Destination.Fields("numero"), DT_RIGHT
        .CellDetails .Rows, .ColumnIndex("Type"), Lrs_Destination.Fields("type"), DT_LEFT
        .CellDetails .Rows, .ColumnIndex("Libelle"), Lrs_Destination.Fields("libelle"), DT_LEFT
        .CellDetails .Rows, .ColumnIndex("Actif"), Lrs_Destination.Fields("actif"), DT_RIGHT
    End With
    Lrs_Destination.MoveNext
    Wend
SGrid1.Redraw = True
Picture2.Visible = True
If SGrid1.Rows > 0 Then
    SGrid1.SelectedRow = 1
    SGrid1.SetFocus
End If

Else
    
    
        MsgBox "Code introuvable", vbInformation
        Call ViderZone(Frm_Destination)
        Lbl_RaAjouter.Visible = False
        Cmd_ReAjouter.Visible = False
        txt_Numero.Enabled = True
        cb_Type.Enabled = True
        Txt_Destination.Enabled = True
        chk_Actif.Enabled = True
        Txt_MaxCompteur.Enabled = True
        Txt_MinCompteur.Enabled = True
        Txt_MaxDuree.Enabled = True
        CmdSave.Enabled = True
        CmdDelete.Enabled = True
        txt_Numero.Text = "Auto"
        cb_Type.SetFocus
        Txt_MaxCompteur.Text = 1
        Txt_MinCompteur.Text = 1
        Txt_MaxDuree.Text = ""
        Exit Sub
    End If
    Set Lrs_Destination = Nothing

Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub
'=================================================
'Ajouter Nouveau***
'=================================================
Private Sub CmdAdd_Click()
    Dim LObj_FindUser As Utilisateur
    Dim Lrs_User As Recordset
    
On Error GoTo Err

    'Controle des droits***
    Set LObj_FindUser = New Utilisateur
    Set Lrs_User = LObj_FindUser.GetRow_User_Ins_Fournisseur(ErrNumber, ErrDescription, ErrSourceDetail, LInt_UserId, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_FindUser = Nothing

    If Lrs_User.EOF Then
        Set Lrs_User = Nothing
        MsgBox "Insertion n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour ajouter une destination", vbExclamation, App.ProductName
        Exit Sub
    End If

    If txt_Numero.Text = "Auto" Then
        If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
    End If

    '===================================
    'Initialise Control box***
    '===================================
    Call ViderZone(Frm_Destination)
    Lbl_RaAjouter.Visible = False
    Cmd_ReAjouter.Visible = False
    txt_Numero.Enabled = True
    cb_Type.Enabled = True
    Txt_Destination.Enabled = True
    chk_Actif.Enabled = True
    Txt_MaxCompteur.Enabled = True
    Txt_MinCompteur.Enabled = True
    Txt_MaxDuree.Enabled = True
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    txt_Numero.Text = "Auto"
    cb_Type.SetFocus
    Txt_MaxCompteur.Text = 1
    Txt_MinCompteur.Text = 1
    Txt_MaxDuree.Text = ""

Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'=================================================
'Supprimer Destination***
'=================================================
Private Sub CmdDelete_Click()
    Dim LObj_FindUser As Utilisateur
    Dim LObj_Desitination As DESTINATION
    Dim Lrs_User As Recordset
    Dim VCode As String

On Error GoTo Err

    'Controle des droits***
    Set LObj_FindUser = New Utilisateur
    Set Lrs_User = LObj_FindUser.GetRow_User_Supp_Fournisseur(ErrNumber, ErrDescription, ErrSourceDetail, LInt_UserId, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_FindUser = Nothing

    If Lrs_User.EOF Then
        Set Lrs_User = Nothing
        MsgBox "Suppression n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Supprime une destination", vbExclamation, App.ProductName
        Exit Sub
    End If
    Set Lrs_User = Nothing
    
   If txt_Numero.Text = "Auto" Then
        If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
            Exit Sub
        Else
            txt_Numero.SetFocus
            Exit Sub
        End If
    End If

    If txt_Numero.Text <> "Auto" And txt_Numero.Text <> "" Then
        If MsgBox("Confirmez vous la suppression de ce Destination", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
            VCode = txt_Numero.Text
            
            Set LObj_Desitination = New DESTINATION
            Call LObj_Desitination.Delete_Add_Destination(ErrNumber, ErrDescription, ErrSourceDetail, VCode, "O", LInt_UserId, CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            Set LObj_Desitination = Nothing
            
            MsgBox "Destination Supprime Aavec succes!...", vbInformation, App.ProductName
            Call ViderZone(Frm_Destination)
            Lbl_RaAjouter.Visible = False
            Cmd_ReAjouter.Visible = False
            txt_Numero.Enabled = True
            cb_Type.Enabled = True
            Txt_Destination.Enabled = True
            chk_Actif.Enabled = True
            Txt_MaxCompteur.Enabled = True
            Txt_MinCompteur.Enabled = True
            Txt_MaxDuree.Enabled = True
            CmdSave.Enabled = True
            CmdDelete.Enabled = True
            txt_Numero.Text = "Auto"
            cb_Type.SetFocus
            Txt_MaxCompteur.Text = 1
            Txt_MinCompteur.Text = 1
            Txt_MaxDuree.Text = ""
        End If
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'=================================================
'Ré-ajouter Destination***
'=================================================
Private Sub Cmd_ReAjouter_Click()
    Dim LObj_FindUser As Utilisateur
    Dim LObj_Desitination As DESTINATION
    Dim Lrs_User As Recordset
    Dim VCode As String

On Error GoTo Err

    'Controle des droits***
    Set LObj_FindUser = New Utilisateur
    Set Lrs_User = LObj_FindUser.GetRow_User_Supp_Fournisseur(ErrNumber, ErrDescription, ErrSourceDetail, LInt_UserId, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_FindUser = Nothing

    If Lrs_User.EOF Then
        Set Lrs_User = Nothing
        MsgBox "Ré-ajouter n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Ré-Ajouter une destination", vbExclamation, App.ProductName
        Exit Sub
    End If
    Set Lrs_User = Nothing
    
   If txt_Numero.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
            Exit Sub
        Else
            txt_Numero.SetFocus
            Exit Sub
        End If
    End If

    If txt_Numero.Text <> "Auto" Then
        If MsgBox("Confirmez vous le ré-ajout de cette Destination", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
            VCode = txt_Numero.Text
            
            Set LObj_Desitination = New DESTINATION
            Call LObj_Desitination.Delete_Add_Destination(ErrNumber, ErrDescription, ErrSourceDetail, VCode, "N", LInt_UserId, CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            Set LObj_Desitination = Nothing
            
            MsgBox "Destination ré-ajouter avec succèes!...", vbInformation, App.ProductName
            Call AfficheRow(VCode)
        End If
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'=================================================
'Enregistre Destination***
'=================================================
Private Sub CmdSave_Click()
    Dim Lobj_Destination As DESTINATION
    Dim LObj_FindUser As Utilisateur
    Dim Lrs_User As Recordset
    Dim Lrs_Destination As Recordset
    Dim LInt_NumCompteur As Long
    Dim VCode As String
    Dim Msg As VbMsgBoxResult
    Dim temp As String
On Error GoTo Err

    If txt_Numero = "" Or cb_Type.Text = "" Or Txt_Destination = "" Or Txt_MaxDuree.Text = "" Or Txt_MaxCompteur = "" Or Txt_MinCompteur.Text = "" Or Txt_MaxDuree.Text = "__:__:__" Then
        MsgBox "Champ(s) Vide(s), Vérifier le(s) Champ(s)!...", vbExclamation, App.ProductName
        Exit Sub
    End If
    If Txt_MinCompteur.Text = 0 Or Txt_MaxCompteur.Text = 0 Then
        Msg = MsgBox("Compteur invalide '0' ..." & vbCr & "Voulez-Vous Confirmer", vbExclamation + vbYesNo, App.ProductName)
        If Msg = vbNo Then Exit Sub
    End If
    If Txt_MaxDuree.Text = "00:00:00" Then
        MsgBox "Durée invalide '00:00:00'...", vbExclamation, App.ProductName
        Exit Sub
    End If
    If Val(Txt_MaxCompteur.Text) < Val(Txt_MinCompteur.Text) Then
        MsgBox "Vérifier compteur entre Max & Min!...", vbExclamation, App.ProductName
        Exit Sub
    End If
    If cb_Type.Text = "Planning" Then
        If Opt_Jour.Value = vbUnchecked And Opt_Matin.Value = vbUnchecked And Opt_soir.Value = vbUnchecked Then
            MsgBox "Choisir temps du tournée!...", vbExclamation, App.ProductName
            Exit Sub
        End If
    End If
    
    VCode = txt_Numero.Text
    temp = ""
    If Opt_Jour.Value = vbChecked Then
        temp = "Journée"
    ElseIf Opt_Matin.Value = vbChecked Then
        temp = "Matin"
    ElseIf Opt_soir.Value = vbChecked Then
        temp = "Après midi"
    End If
    '--Modification***
    If VCode <> "Auto" Then
    
        '--Droit d'accès***
        Set LObj_FindUser = New Utilisateur
        Set Lrs_User = LObj_FindUser.GetRow_User_Maj_Fournisseur(ErrNumber, ErrDescription, ErrSourceDetail, LInt_UserId, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_FindUser = Nothing
        If Lrs_User.EOF Then
            Set Lrs_User = Nothing
            MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier une destination", vbExclamation, App.ProductName
            Exit Sub
        End If
        Set Lrs_User = Nothing
        
        If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
             Set Lrs_Destination = CreateEmptyRS_Destination()
             With Lrs_Destination
                 .AddNew
                 .Fields("Type") = cb_Type.Text
                 .Fields("Libelle") = Txt_Destination
                 .Fields("Actif") = chk_Actif.Value
                 .Fields("MaxDuree") = Txt_MaxDuree.Text
                 .Fields("MaxCompteur") = Txt_MaxCompteur
                 .Fields("MinCompteur") = Txt_MinCompteur
                 .Fields("UserUpdate") = LInt_UserId
                 .Fields("Temps") = temp
             End With
             Set Lobj_Destination = New DESTINATION
             Call Lobj_Destination.UpDate_Destination(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Destination, VCode)
             If ErrNumber <> 0 Then
                 ErrNumber = 0
                 MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                 Exit Sub
             End If
             Set Lobj_Destination = Nothing
             Set Lrs_Destination = Nothing
        
             MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
        End If
    '--Ajouter***
    ElseIf VCode = "Auto" Then
        '--Droit d'accès***
        Set LObj_FindUser = New Utilisateur
        Set Lrs_User = LObj_FindUser.GetRow_User_Ins_Fournisseur(ErrNumber, ErrDescription, ErrSourceDetail, LInt_UserId, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_FindUser = Nothing
        If Lrs_User.EOF Then
            Set Lrs_User = Nothing
            MsgBox "Inseration n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour ajouter une destination", vbExclamation, App.ProductName
            Exit Sub
        End If
        Set Lrs_User = Nothing
        
        If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbYes Then
            LInt_NumCompteur = return_Compteur() + 1
            txt_Numero.Text = Format(LInt_NumCompteur, "00000")
            VCode = Format(LInt_NumCompteur, "00000")
            
            Set Lrs_Destination = CreateEmptyRS_Destination()
             With Lrs_Destination
                 .AddNew
                .Fields("Numero") = VCode
                 .Fields("Type") = cb_Type.Text
                 .Fields("Libelle") = Txt_Destination
                 .Fields("Actif") = chk_Actif.Value
                 .Fields("MaxDuree") = Txt_MaxDuree.Text
                 .Fields("MaxCompteur") = Txt_MaxCompteur
                 .Fields("MinCompteur") = Txt_MinCompteur
                 .Fields("UserInsert") = LInt_UserId
                 .Fields("Temps") = temp
             End With
             Set Lobj_Destination = New DESTINATION
             Call Lobj_Destination.Save_Destination(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Destination)
             If ErrNumber <> 0 Then
                 ErrNumber = 0
                 MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                 Exit Sub
             End If
             Set Lobj_Destination = Nothing
             Set Lrs_Destination = Nothing
        
             MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
        End If
    End If
    
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Function return_Compteur() As Long
    Dim Lobj_Destination As DESTINATION
    Dim Lrs_Destination As Recordset
    
On Error GoTo Err
    
    Set Lobj_Destination = New DESTINATION
    Set Lrs_Destination = Lobj_Destination.GetNumero_MaxCompteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Function
    End If
    Set Lobj_Destination = Nothing
    If Not Lrs_Destination.EOF Then
        return_Compteur = Lrs_Destination(0)
    End If
    Set Lrs_Destination = Nothing
    
Exit Function
Err:
    MsgBox Err.Description, vbExclamation
End Function

Public Sub InitialGrid()
 With SGrid1

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
    
    .AddColumn "Numero", "Numero", , , 60, , , , , , , CCLSortNumeric
    .AddColumn "Type", "Type", , , 140
    .AddColumn "Libelle", "Libelle", , , 140
    .AddColumn "Actif", "Actif", , , 60
  
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With


End Sub


