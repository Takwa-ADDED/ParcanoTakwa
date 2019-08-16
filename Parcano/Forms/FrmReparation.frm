VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBCReparation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "BC Réparation"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11955
   Begin MSComctlLib.ListView grid 
      Height          =   3495
      Left            =   240
      TabIndex        =   36
      Top             =   5520
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6165
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Numéro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Désignation"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Quantité"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Véhicule"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Observation"
         Object.Width           =   2540
      EndProperty
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   3840
      TabIndex        =   30
      Top             =   1680
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
      Picture         =   "FrmReparation.frx":0000
      ButtonType      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   480
      Top             =   120
   End
   Begin VB.Frame Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   11655
      Begin VB.ComboBox cbo_MatriculeStation 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   1320
         Width           =   2895
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1560
         ScaleHeight     =   495
         ScaleWidth      =   2775
         TabIndex        =   32
         Top             =   120
         Width           =   2775
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
            Height          =   555
            Left            =   0
            MaxLength       =   50
            TabIndex        =   0
            Tag             =   "M"
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   4575
         TabIndex        =   18
         Top             =   1920
         Width           =   4575
         Begin VB.TextBox txt_ville 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   21
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox txt_adresse 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   20
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox txt_rsocial 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   19
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   24
            Top             =   1080
            Width           =   435
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   23
            Top             =   600
            Width           =   780
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   1290
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4560
         ScaleHeight     =   495
         ScaleWidth      =   3375
         TabIndex        =   16
         Top             =   120
         Width           =   3375
         Begin MSComCtl2.DTPicker cda_Create 
            Height          =   375
            Left            =   1920
            TabIndex        =   35
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   149094401
            CurrentDate     =   42875
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Création  :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   17
            Top             =   120
            Width           =   1545
         End
      End
      Begin VB.ComboBox Cbo_Conducteur 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7560
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin SToolBox.SCommand CmdFindConducteur 
         Height          =   375
         Left            =   10440
         TabIndex        =   25
         Top             =   1320
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
         Picture         =   "FrmReparation.frx":0353
         ButtonType      =   1
      End
      Begin SToolBox.SCommand CmdFindStation 
         Height          =   375
         Left            =   4560
         TabIndex        =   26
         Top             =   1320
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
         Picture         =   "FrmReparation.frx":06A6
         ButtonType      =   1
      End
      Begin VB.Label Lbl_user 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BC saisi par :"
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
         Height          =   255
         Left            =   8400
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Lbl_UserSaisi 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   9720
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Numéro  :"
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
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conducteur  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6000
         TabIndex        =   28
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Fournisseur 
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
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.PictureBox PIC_NFACT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5760
      ScaleHeight     =   495
      ScaleWidth      =   4455
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label LBL_NFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1250"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3840
         TabIndex        =   14
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ce bon est inseré dans une facture N° : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -720
         TabIndex        =   13
         Top             =   120
         Width           =   4380
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   11280
      ScaleHeight     =   3495
      ScaleWidth      =   615
      TabIndex        =   8
      Top             =   5400
      Width           =   615
      Begin SToolBox.SCommand Cmd_SupDet 
         Height          =   495
         Left            =   120
         TabIndex        =   9
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
         Picture         =   "FrmReparation.frx":09F9
      End
      Begin SToolBox.SCommand Cmd_SaisiDet 
         Height          =   495
         Left            =   120
         TabIndex        =   3
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
         Picture         =   "FrmReparation.frx":0B7B
      End
      Begin SToolBox.SCommand Cmd_ModifDet 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   840
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
         Picture         =   "FrmReparation.frx":0CFD
      End
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   10320
      TabIndex        =   4
      Top             =   480
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
      Picture         =   "FrmReparation.frx":1050
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   9360
      TabIndex        =   5
      Top             =   480
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
      Picture         =   "FrmReparation.frx":11D2
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   9840
      TabIndex        =   6
      Top             =   480
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
      Picture         =   "FrmReparation.frx":1525
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   480
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
      Picture         =   "FrmReparation.frx":1878
   End
   Begin SToolBox.SCommand SCmd_Print 
      Height          =   495
      Left            =   10800
      TabIndex        =   10
      Top             =   480
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
      Picture         =   "FrmReparation.frx":19FA
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "BC Réparation"
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
      Left            =   720
      TabIndex        =   31
      Top             =   480
      Width           =   2535
   End
   Begin VB.Image PicBox_Header 
      Height          =   1335
      Left            =   0
      Picture         =   "FrmReparation.frx":1D4D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "FrmBCReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim thekey As Integer
Dim theshift As Integer

Public Sub AfficheRow_Conducteur(ByVal VCode As String)

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_CONDUCTEUR(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Libelle")) Then Cbo_Conducteur.Text = rs("Libelle")
Else
    MsgBox "Code introuvable, vérifier votre saisie.", vbInformation, App.ProductName
    Cbo_Conducteur.Text = ""
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
    start = Len(Cbo_Conducteur.Text)
    For i = 0 To Cbo_Conducteur.ListCount - 1
        If Left(Cbo_Conducteur.List(i), start) = Cbo_Conducteur.Text Then
            Cbo_Conducteur.Text = Cbo_Conducteur.List(i)
        End If
    Next
    Cbo_Conducteur.SelStart = start
    Cbo_Conducteur.SelLength = Len(Cbo_Conducteur.Text)
    End If
End Sub

Private Sub cbo_Conducteur_Click()
If Len(Trim(Cbo_Conducteur.Text)) > 0 Then Call AfficheRow_Conducteur(Cbo_Conducteur.Text)

End Sub

Private Sub Cbo_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Cbo_Conducteur_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub

Private Sub cbo_conducteur_LostFocus()
On Error GoTo Err
If Len(Trim(Cbo_Conducteur.Text)) > 0 Then Call AfficheRow_Conducteur(Cbo_Conducteur.Text)
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub ExistData(ByVal cbo As ComboBox)

On Error GoTo Err

    Dim RCount As Integer, i As Integer, Existe As Boolean, tcbo As String
    
    RCount = cbo.ListCount
    tcbo = cbo.Text
    
    For i = 0 To RCount - 1
        cbo.ListIndex = i
        If tcbo = cbo.Text Then
            Existe = True
            Exit For
        Else
            Existe = False
        End If
    Next i
    If i = RCount Then
        If Existe = False Then
            MsgBox "Saisie non Valide!...     ", vbExclamation, App.ProductName
            cbo.Text = ""
            txt_rsocial.Text = ""
            txt_adresse.Text = ""
            txt_ville.Text = ""
            Exit Sub
        End If
    End If
    Existe = False
    
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cda_Create_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub CmdAdd_Click()

On Error GoTo Err

Dim LOBJ_Personnel As personnel

Set LOBJ_Personnel = New personnel
If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Ins_BCR", LInt_UserId) Then
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If
    
If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If
Picture2.Enabled = True
Picture5.Enabled = True
grid.Enabled = True
Call ViderZone(FrmBCReparation)
CmdDelete.Enabled = True
CmdSave.Enabled = True
SCmd_Print.Enabled = True

txt_Numero.Text = "Auto"
cda_Create.Value = Date
grid.ListItems.Clear
PIC_NFACT.Visible = False
Lbl_UserSaisi.Caption = LStr_NameUser
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()

Dim LOBJ_Personnel As personnel
Dim LOBJ_BonRepar As BCReparation
Dim VCode As String

On Error GoTo Err

If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        txt_Numero.SetFocus
        Exit Sub
    End If
End If
' Ce bon est inseré dans une facture N°: bon payé donc on peut pas ni le supprimer ni modifier
If PIC_NFACT.Visible = True Then
    MsgBox "Maj impossible", vbInformation
    Exit Sub
End If

Set LOBJ_Personnel = New personnel
If txt_Numero.Text <> "Auto" Then
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Supp_BCR", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
End If
    
If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "bon de réparation", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
    
VCode = txt_Numero.Text
Set LOBJ_BonRepar = New BCReparation

Call LOBJ_BonRepar.Delete_DetRepBySup(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

Call LOBJ_BonRepar.Delete_BRepa(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LInt_UserId, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Call ViderZone(FrmBCReparation)
grid.ListItems.Clear
Lbl_UserSaisi.Caption = ""
Picture2 = True
Picture5.Enabled = True
grid.Enabled = True

Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdFind_Click()

On Error GoTo Err

If grid.ListItems.Count > 0 And txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

If Okayy = True Then
    If MsgBox("Annuler le maj en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
End If
Lbl_UserSaisi.Caption = ""
Unload FrmFind
With FrmFind
    .StrSource = "Reparation"
    .Show vbModal
End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub cmdFindConducteur_Click()

On Error GoTo Err

If txt_Numero.Text = "" Then
    Exit Sub
Else
    Okayy = True
    Unload FrmFind_Fils
    With FrmFind_Fils
        .StrSource = "PersonnelReparation"
        .Show vbModal
    End With
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub cmdFindMatricule_Click()

On Error GoTo Err

If grid.ListItems.Count > 0 And txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

If Okayy = True Then
    If MsgBox("Annuler le maj en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
End If
Lbl_UserSaisi.Caption = ""
Unload FrmFind
With FrmFind
    .StrSource = "Reparation"
    .Show vbModal
End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdFindStation_Click()

On Error GoTo Err

If txt_Numero.Text <> "" Then
    Unload FrmFind_Fils
    With FrmFind_Fils
        .StrSource = "Station BCR"
        .Show vbModal
    End With
End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdSave_Click()

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset
On Error GoTo Err

' Ce bon est inseré dans une facture N°: bon payé donc on peut pas ni le supprimer ni modifier
If PIC_NFACT.Visible = True Then
    MsgBox "Maj impossible", vbInformation
    Exit Sub
End If

If Left(CheckMandatory(FrmBCReparation), 1) = 1 Then
   Exit Sub
End If

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_CONDUCTEUR(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Cbo_Conducteur.Text)
If ErrNumber <> 0 Then
   MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
   ErrNumber = 0
   Exit Sub
End If
If rs.EOF Then
    MsgBox "Ce conducteur n'existe pas ", vbInformation
    Cbo_Conducteur.SetFocus
    Exit Sub
    rs.Close
End If

If grid.ListItems.Count = 0 Then
    MsgBox "Veuillez saisir les details du BCReparation", vbInformation
    Exit Sub
End If

If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
    
If txt_Numero.Text <> "Auto" And txt_Numero.Text <> "" Then
    Call modifier_BRep
End If

If txt_Numero.Text = "Auto" Then
    Call Ajouter_BCRep
End If
    
 Okayy = False
If MsgBox("Enregistrement terminé avec succé  " & vbNewLine & "Imprimer ce bon        ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Dim F As Form
    Set F = New Frm_Rpt_Apercus
    With F
        .Numero = txt_Numero.Text
        Call .PrintOutAndApercu_BCReparation(0)
        .Show
    End With
End If
'Call ViderZone(FrmBCReparation)
'grid.ListItems.Clear
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

'Modification d'un bon de réparation
Private Sub modifier_BRep()

Dim LOBJ_Personnel As personnel
Dim LOBJ_BonRepar As BCReparation
Dim LRs_NewRecord As New Recordset

Set LOBJ_Personnel = New personnel
If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Maj_BCR", LInt_UserId) Then
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If
VCode = txt_Numero.Text
Set LOBJ_BonRepar = New BCReparation
Set LRs_NewRecord = CreateEmptyRS_AssBCRepar()
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = txt_Numero.Text
    .Fields("DateCreation") = CDate(cda_Create.Value)
    .Fields("Fournisseur") = cbo_MatriculeStation.Text
    .Fields("Conducteur") = Cbo_Conducteur.Text
    .Fields("UserUpdate") = LInt_UserId
End With
Call LOBJ_BonRepar.Update_BRepar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing
'Supprimer les anciens détails de ce bon et enregistrer les nouveaux
Call LOBJ_BonRepar.Delete_DetRep(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

Call insert_DetBCRep

End Sub

'Ajout d'un nouveau bon de réparation
Private Sub Ajouter_BCRep()

Dim LOBJ_BonRepar As BCReparation
Dim LRs_NewRecord As New Recordset
Dim LInt_NumCompteur As Long

'Incrémenter compteur lors d'un nouveau ajout
LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_BCReparation")
If ErrNumber <> 0 Then
   MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
   ErrNumber = 0
   Exit Sub
End If
'Insertion enregistrement assiette
txt_Numero.Text = Format(LInt_NumCompteur, "00000")

Set LOBJ_BonRepar = New BCReparation
Set LRs_NewRecord = CreateEmptyRS_AssBCRepar()
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = txt_Numero.Text
    .Fields("DateCreation") = CDate(cda_Create.Value)
    .Fields("Fournisseur") = cbo_MatriculeStation.Text
    .Fields("Conducteur") = Cbo_Conducteur.Text
    .Fields("UserInsert") = LInt_UserId
End With
Call LOBJ_BonRepar.Insert_BCRepar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

Call insert_DetBCRep

End Sub

Private Sub insert_DetBCRep()

Dim i As Integer
Dim LRs_NewRecord As New Recordset
Dim LOBJ_BonRepar As BCReparation

Set LOBJ_BonRepar = New BCReparation
Set LRs_NewRecord = CreateEmptyRS_DetBCRepar
For i = 1 To grid.ListItems.Count
    With LRs_NewRecord
        .AddNew
        .Fields("Numero") = txt_Numero.Text   'Numero BCrepar
        .Fields("désignation") = grid.ListItems(i).SubItems(1)
        .Fields("Qté") = grid.ListItems(i).SubItems(2)
        .Fields("Vehicule") = grid.ListItems(i).SubItems(3)
        .Fields("Observation") = grid.ListItems(i).SubItems(4)
    End With
    grid.ListItems(i).Text = txt_Numero.Text
Next
Call LOBJ_BonRepar.Insert_DetBCRepar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing
End Sub

Private Function RET_CODE_CONDUCTEUR(txt As String) As String

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset

RET_CODE_CONDUCTEUR = ""
Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.GetCODE_CONDUCTEUR(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
If Not rs.EOF Then
    RET_CODE_CONDUCTEUR = rs(0)
End If
rs.Close

End Function

Private Sub Form_Load()

On Error GoTo Err
CmdDelete.Enabled = False
CmdSave.Enabled = False
SCmd_Print.Enabled = False
Call Affiche_StatRep_Combo(cbo_MatriculeStation)
Call Affiche_Personnel_Combo(Cbo_Conducteur)
Me.Width = 11715
Me.Height = 8625
Me.Move 0, 0
cda_Create.Value = Date
Me.WindowState = 2
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub Form_Resize()

Dim WidthForm As Integer
WidthForm = Frm_Main.ACB_Main.Width
PicBox_Header.Width = WidthForm - 1000
CmdAdd.Left = WidthForm - 5500
CmdDelete.Left = WidthForm - 5100
CmdFind.Left = WidthForm - 4700
CmdSave.Left = WidthForm - 4300
SCmd_Print.Left = WidthForm - 3900
End Sub

Private Sub Form_Unload(Cancel As Integer)

'On Error GoTo erreur
'Dim i As Integer
'Dim Msg ' Déclare la variable.
'' Définit le texte du message.
'Msg = "Voulez-vous vraiment quitter?"
'' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
'If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'   Cancel = True
'Else
'Unload Me
'End If
'
'Exit Sub
'erreur:
'   MsgBox Err.Description, 48
End Sub

Private Sub grid_DblClick()

Dim i

On Error GoTo Err

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If

If grid.ListItems.Count <= 0 Then Exit Sub
Okayy = True
With frmDetailBCReparation
    .Okay = False
    .ii = grid.SelectedItem.Index
     i = grid.SelectedItem.Index
    .txt_Numero.Text = txt_Numero.Text
    .txt_libelle.Text = grid.ListItems(i).SubItems(1)
    .txt_Qte.Text = grid.ListItems(i).SubItems(2)
    .cbo_Matricule.Text = grid.ListItems(i).SubItems(3)
    .txt_Observation.Text = grid.ListItems(i).SubItems(4)
    .Show
End With
Err:
Exit Sub
MsgBox Err.Description, vbInformation
End Sub

'Impression du bon
Private Sub SCmd_Print_Click()

'On Error GoTo Err
If txt_Numero.Text = "" Then Exit Sub
If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        txt_Numero.SetFocus
        Exit Sub
    End If
End If

If MsgBox(" Imprimer ce bon    ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Dim F As Form
    Set F = New Frm_Rpt_Apercus
    With F
        .Numero = txt_Numero.Text
        Call .PrintOutAndApercu_BCReparation(0)
        .Show
    End With
End If

Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

'Modification d'un détail du bon
Private Sub Cmd_ModifDet_Click()

Dim i

On Error GoTo Err

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If

If grid.ListItems.Count <= 0 Then Exit Sub
Okayy = True
With frmDetailBCReparation
    .Okay = False
    .ii = grid.SelectedItem.Index
    i = grid.SelectedItem.Index
    .txt_libelle.Text = grid.ListItems(i).SubItems(1)
    .txt_Qte.Text = grid.ListItems(i).SubItems(2)
    .cbo_Matricule.Text = grid.ListItems(i).SubItems(3)
    .txt_Observation.Text = grid.ListItems(i).SubItems(4)
    .Show
End With
Err:
Exit Sub
MsgBox Err.Description, vbInformation
End Sub

'Suppression d'une ligne (détail) de la liste
Private Sub Cmd_SupDet_Click()

Dim i As Integer

On Error GoTo Err
If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If

If grid.ListItems.Count <= 0 Then Exit Sub
Okayy = True
If MsgBox("Confirmez vous la suppression de la ligne en cours.?", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    i = grid.SelectedItem.Index
    grid.ListItems.Remove i
    'Call AppCalcul
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

'saisir un nouveau detail du bon reparation
Private Sub Cmd_SaisiDet_Click()

If txt_Numero.Text = "" Then
    If Len(Trim(txt_Numero.Text)) = 0 Then
        MsgBox "N° bon obligatoire      ", vbInformation
        txt_Numero.SetFocus
        Exit Sub
    End If
End If
        
If cbo_MatriculeStation.Text = "" Or cbo_MatriculeStation.Text = " " Then
    MsgBox "Station obligatoire      ", vbInformation
    Exit Sub
End If

If Cbo_Conducteur.Text = "" Then
    If Len(Trim(Cbo_Conducteur.Text)) = 0 Then
        MsgBox "Conducteur obligatoire      ", vbInformation
        Exit Sub
    End If
End If
    
On Error GoTo Err

Okayy = True
If txt_Numero.Text = "" Then
    If Len(Trim(txt_Numero.Text)) = 0 Then
        MsgBox "N° bon obligatoire      ", vbInformation
        txt_Numero.SetFocus
        Exit Sub
    End If
  
    With frmDetailBCReparation
        .txt_Numero.Text = Me.txt_Numero.Text
        .Okay = True
        .Show vbModal
    End With
Else
    With frmDetailBCReparation
        .txt_Numero.Text = Me.txt_Numero.Text
        .Okay = True
        .Show vbModal
    End With
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Function return_Compteur() As Long

Dim rD As New Recordset
Dim LOBJ_BCRepar As BCReparation

return_Compteur = 0

Set LOBJ_BCRepar = New BCReparation
Set rD = LOBJ_BCRepar.Get_MaxNum(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If

If Not rD.EOF Then
    return_Compteur = rD("maxNum")
End If

rD.Close
End Function

Public Sub AfficheRow(ByVal VCode As String)

Dim LOBJ_BCRepar As BCReparation
Dim rs As New Recordset

Call ViderZone(FrmBCReparation)
grid.ListItems.Clear
Set LOBJ_BCRepar = New BCReparation
'Assiette BCR
Set rs = LOBJ_BCRepar.Get_AssBRepar(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    txt_Numero.Text = rs("Numero")
    If Not IsNull(rs("DateCreation")) Then cda_Create.Value = rs("DateCreation")
    If Not IsNull(rs("Fournisseur")) Then Call AfficheRow_Station(rs("Fournisseur"))
    If Not IsNull(rs("Conducteur")) Then Cbo_Conducteur.Text = rs("Conducteur")
    If Not IsNull(rs("UserInsert")) Then Lbl_UserSaisi.Caption = Get_NameUserByCode(rs("UserInsert"))
    If rs("Supp") = "O" Then
        MsgBox "Bon de commande de réparation supprimé par " & Get_NameUserByCode(rs("UserDelete")), vbInformation
        Picture2.Enabled = False
        Picture5.Enabled = False
        grid.Enabled = False
        CmdSave.Enabled = False
        CmdDelete.Enabled = False
        SCmd_Print.Enabled = False
    Else
        Picture2.Enabled = True
        Picture5.Enabled = True
        grid.Enabled = True
        CmdSave.Enabled = True
        CmdDelete.Enabled = True
        SCmd_Print.Enabled = True
    End If
     If rs("Transf") = "O" Then
        LBL_NFact.Caption = rs("NumPR")
        PIC_NFACT.Visible = True
        Picture2.Enabled = False
        Picture5.Enabled = False
        grid.Enabled = False
        CmdSave.Enabled = False
        CmdDelete.Enabled = False
        SCmd_Print.Enabled = False
        Call Timer1_Timer
    Else
        PIC_NFACT.Visible = False
        Timer1.Enabled = False
    End If
Else
    txt_Numero.SetFocus
End If
rs.Close
'Detail BCR
Set rs = LOBJ_BCRepar.Get_DetBRepar(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
            Set itmX = grid.ListItems.Add(, , CStr(txt_Numero.Text))
            itmX.SubItems(1) = rs("Désignation")
            itmX.SubItems(2) = rs("qté")
            itmX.SubItems(3) = rs("Vehicule")
            itmX.SubItems(4) = rs("Observation")
        rs.MoveNext
    Wend
End If
rs.Close

End Sub
Public Sub AfficheRow_Station(ByVal VCode As String)

Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.GetStatByCodeLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    cbo_MatriculeStation.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
Else
    MsgBox "Code introuvable, vérifier votre saisie.", vbInformation, App.ProductName
    cbo_MatriculeStation.Text = ""
End If

End Sub

Private Sub cbo_MatriculeStation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cbo_MatriculeStation_Click()
If Len(Trim(cbo_MatriculeStation.Text)) > 0 Then Call AfficheRow_Station(cbo_MatriculeStation.Text)

End Sub
Private Sub txt_Numero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub Timer1_Timer()

    Timer1.Enabled = True
    Timer1.Interval = 600
    If PIC_NFACT.Visible = True Then
        PIC_NFACT.Visible = False
    Else
        PIC_NFACT.Visible = True
    End If
End Sub

