VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsultBC 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Details bon carburant"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11595
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
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8040
      ScaleHeight     =   375
      ScaleWidth      =   3015
      TabIndex        =   19
      Top             =   1680
      Width           =   3015
      Begin SToolBox.SDateBox cda_Operation 
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Tag             =   "M"
         Top             =   0
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Text            =   ""
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Operation:"
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
         TabIndex        =   21
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   0
      ScaleHeight     =   6735
      ScaleWidth      =   11535
      TabIndex        =   3
      Top             =   1680
      Width           =   11535
      Begin VB.ComboBox Cbo_Conducteur 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6480
         TabIndex        =   23
         Tag             =   "M"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txt_Valeur 
         Alignment       =   1  'Right Justify
         Height          =   435
         Left            =   6600
         MaxLength       =   50
         TabIndex        =   14
         Tag             =   "M"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4920
         ScaleHeight     =   375
         ScaleWidth      =   2655
         TabIndex        =   13
         Top             =   0
         Width           =   2655
         Begin VB.Label cda_Create 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1320
            TabIndex        =   28
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Creation :"
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
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   4575
         TabIndex        =   6
         Top             =   2040
         Width           =   4575
         Begin VB.TextBox txt_ville 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   9
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox txt_adresse 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   8
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txt_rsocial 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   7
            Top             =   120
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   120
            Width           =   1290
         End
      End
      Begin VB.TextBox txt_MatriculeStation 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "M"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txt_Numero 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "M"
         Top             =   0
         Width           =   2175
      End
      Begin MSComctlLib.ListView Lsv_Client 
         Height          =   3015
         Left            =   0
         TabIndex        =   15
         Top             =   3840
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5318
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
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "N"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1853
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Station"
            Object.Width           =   2471
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "CodeMat"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Immatriculation"
            Object.Width           =   3177
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Energie"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Conducteur"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Compteur"
            Object.Width           =   1589
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "N.Litre"
            Object.Width           =   1272
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Prix.TTC"
            Object.Width           =   1766
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Valeur"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Lbl_User 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3000
         TabIndex        =   27
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bon de carburant saisi par : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conducteur :"
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
         Left            =   5040
         TabIndex        =   24
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valeur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   5400
         TabIndex        =   18
         Top             =   2400
         Width           =   1005
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
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   735
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
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox PIC_NFACT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label LBL_NFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1250"
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
         Left            =   3480
         TabIndex        =   2
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ce bon est inseré dans une facture N° : "
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
         TabIndex        =   1
         Top             =   120
         Width           =   3300
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bon de sortie carburant"
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
      TabIndex        =   25
      Top             =   600
      Width           =   3855
   End
   Begin VB.Image PicBox_Header 
      Height          =   1695
      Left            =   0
      Picture         =   "frmConsult.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "frmConsultBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Afficher assiette et détails bon du carburant
Public Sub AfficheRow(ByVal VCode As String)

Dim LOBJ_BC As BonCarburant
Dim rs As New Recordset

Lsv_Client.ListItems.Clear

Set LOBJ_BC = New BonCarburant
Set rs = LOBJ_BC.Get_AssBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    txt_Numero.Text = rs("Numero")
    If Not IsNull(rs("STATION")) Then txt_MatriculeStation.Text = rs("STATION")
    If Not IsNull(rs("DATEDOC")) Then cda_Create.Caption = rs("DATEDOC")
    If Not IsNull(rs("dateop")) Then cda_Operation.Text = rs("dateop")
    If Not IsNull(rs("CONDUCTEUR")) Then cbo_conducteur.Text = rs("CONDUCTEUR")
    If Not IsNull(rs("UserInsert")) Then Lbl_user.Caption = Get_NameUserByCode(rs("UserInsert"))
    If Not IsNull(rs("VALEUR")) Then txt_Valeur.Text = Format(rs("VALEUR"), "#,##0.000")
    Call AfficheRow_Station(rs("STATION"))
    Call AfficheRow_Conducteur(rs("CONDUCTEUR"))
    If rs("Transf") = "O" Then
        LBL_NFact.Caption = rs("NumFact")
        PIC_NFACT.Visible = True
    Else
        PIC_NFACT.Visible = False
    End If
End If
rs.Close

Set rs = LOBJ_BC.Get_DetailBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
            Set itmX = Lsv_Client.ListItems.Add(, , CStr(txt_Numero.Text))
            itmX.SubItems(1) = CStr(cda_Create.Caption)
            itmX.SubItems(2) = CStr(txt_MatriculeStation.Text)
            itmX.SubItems(3) = CStr(txt_rsocial.Text)
            itmX.SubItems(4) = CStr(rs("Vehicule"))
            itmX.SubItems(5) = CStr(rs("Matricule"))
            itmX.SubItems(6) = CStr(rs("Energie"))
            itmX.SubItems(7) = CStr(cbo_conducteur.Text)
            itmX.SubItems(8) = CStr(rs("CompteurCarburant"))
            itmX.SubItems(9) = CStr(rs("Litre"))
            itmX.SubItems(10) = CStr(Format(rs("prixLitre"), "#,##0.000"))
            itmX.SubItems(11) = Format(rs("Litre") * rs("prixLitre"), "#,##0.000")
        rs.MoveNext
    Wend
End If
rs.Close

End Sub

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
    If Not IsNull(rs("Libelle")) Then cbo_conducteur.Text = rs("Libelle")
End If
rs.Close

End Sub

Public Sub AfficheRow_Station(ByVal VCode As String)

Dim LOBJ_Stat As Station
Dim rs As New Recordset

Set LOBJ_Stat = New Station
Set rs = LOBJ_Stat.GetStationByCode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    txt_MatriculeStation.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
End If

End Sub

Private Sub Form_Load()
PicBox_Header.Width = Me.Width
Me.Height = 9210
Me.Width = 11715
Me.Move 0, 0
cda_Create.Caption = Date
Me.Move 0, 0

End Sub

'Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo erreur
'   Dim i As Integer
'   Dim MSG ' Déclare la variable.
'   ' Définit le texte du message.
'   MSG = "Voulez-vous vraiment quitter?"
'   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
'   If MsgBox(MSG, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'      Cancel = True
'   Else
'   Unload Me
'   End If
'
'   Exit Sub
'erreur:
'   MsgBox Err.Description, 48
'
'End Sub


