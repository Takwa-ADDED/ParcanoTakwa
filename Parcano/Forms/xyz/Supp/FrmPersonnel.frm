VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmPersonnel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Parcano"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPersonnel.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleMode       =   0  'User
   ScaleWidth      =   1.02601e9
   Begin VB.CheckBox chk_Actif 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      Top             =   5280
      Width           =   255
   End
   Begin VB.TextBox txt_CIN 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2085
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txt_mobile 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txt_permie 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2100
      MaxLength       =   30
      TabIndex        =   6
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txt_lieuPermi 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2100
      MaxLength       =   30
      TabIndex        =   8
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox txt_fonction 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2100
      MaxLength       =   30
      TabIndex        =   5
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txt_Telephone 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   4440
      TabIndex        =   14
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
      Picture         =   "FrmPersonnel.frx":0ECA
      ButtonType      =   1
   End
   Begin VB.TextBox txt_Nom 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2100
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "M"
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox txt_Matricule 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      Left            =   2100
      TabIndex        =   0
      Tag             =   "M"
      Top             =   1560
      Width           =   2295
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   9720
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
      Picture         =   "FrmPersonnel.frx":121D
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   8880
      TabIndex        =   11
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
      Picture         =   "FrmPersonnel.frx":139F
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   9240
      TabIndex        =   12
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
      Picture         =   "FrmPersonnel.frx":16F2
   End
   Begin SToolBox.SDateBox cda_DateLivrPermi 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   4440
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   8520
      TabIndex        =   23
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
      Picture         =   "FrmPersonnel.frx":1A45
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5280
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
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   120
      TabIndex        =   22
      Top             =   2640
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
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   120
      TabIndex        =   21
      Top             =   3360
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
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   675
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
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   1920
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
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "telephone :"
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
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1110
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
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   120
      TabIndex        =   16
      Top             =   3720
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
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   2280
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
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche personnel"
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
      TabIndex        =   9
      Top             =   240
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "FrmPersonnel.frx":1BC7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "FrmPersonnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAdd_Click()

On Error GoTo Err

If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If
Dim rQ As New ADODB.Recordset
Dim SQL As String
SQL = "Select * from utilisateur where Ins_Personnel = 1 and code= " & LInt_UserId
rQ.Open SQL, CNB, adOpenDynamic
If rQ.EOF Then
    rQ.Close
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If

Call ViderZone(FrmPersonnel)
'txt_Matricule.Enabled = False
txt_Matricule.Text = "Auto"
txt_Nom.SetFocus

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim vcode As String

On Error GoTo Err
   If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        Else
            txt_Matricule.SetFocus
            Exit Sub
        End If
    End If
    If txt_Matricule.Text <> "Auto" Then
        Dim rQ As New ADODB.Recordset
        SQL = "Select * from utilisateur where Supp_personnel = 1 and code= " & LInt_UserId
        rQ.Open SQL, CNB, adOpenDynamic
        If rQ.EOF Then
            rQ.Close
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
    End If
    If MsgBox("Confirmez vous la suppression", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    vcode = txt_Matricule.Text
    SQL = "Delete from personnel where code =" & SQLText(vcode)
    CNB.Execute SQL
    txt_Matricule.SetFocus
    End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdFind_Click()
On Error Resume Next
If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Unload FrmFind
With FrmFind
    .StrSource = "Personnel"
    .Show
End With
End Sub

Private Sub cmdFindMatricule_Click()
Unload FrmFind_Actif
With FrmFind_Actif
    .StrSource = "Personnel"
    .Show
End With
End Sub

Private Sub CmdSave_Click()

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim vcode As String

On Error GoTo Err
If txt_Matricule.Text <> "Auto" Then
        Dim rQ As New ADODB.Recordset
        SQL = "Select * from utilisateur where Maj_Personnel = 1 and code= " & LInt_UserId
        rQ.Open SQL, CNB, adOpenDynamic
        If rQ.EOF Then
            rQ.Close
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
    End If
    If Left(CheckMandatory(FrmPersonnel), 1) = 1 Then
       Exit Sub
    End If
    
    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    'Delete l'enregistrement

    CNB.BeginTrans
    vcode = txt_Matricule.Text
    SQL = "Delete from Personnel where code =" & SQLText(vcode)
    CNB.Execute SQL
    If vcode = "Auto" Then
    LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_Personnel")
    If ErrNumber <> 0 Then
       MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
       ErrNumber = 0
       Exit Sub
    End If
    Set LObj_Compteur = Nothing
    'Insertion enregistrement assiette
    txt_Matricule.Text = Format(LInt_NumCompteur, "00000")
    End If
    Dim D
    If cda_DateLivrPermi.Text = "__/__/____" Then
    
    Else
    D = cda_DateLivrPermi.Text
    End If
    'Insertion enregistrement
    SQL = "Insert into personnel  (Code,Libelle,CIN,Fonction,telephone,mobile,permie,datlivr,lieulivr, Actif,disponible) values ("
    SQL = SQL & SQLText(txt_Matricule.Text)
    SQL = SQL & "," & SQLText(txt_Nom.Text)
    SQL = SQL & "," & SQLText(txt_CIN.Text)
    SQL = SQL & "," & SQLText(txt_fonction.Text)
    SQL = SQL & "," & SQLText(txt_telephone.Text)
    SQL = SQL & "," & SQLText(txt_mobile.Text)
    SQL = SQL & "," & SQLText(txt_permie.Text)
    SQL = SQL & "," & SQLText(D)
    SQL = SQL & "," & SQLText(txt_lieuPermi.Text)
    SQL = SQL & "," & chk_Actif.Value
    SQL = SQL & "," & SQLText("O")
    SQL = SQL & ")"
    CNB.Execute SQL
    CNB.CommitTrans
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
    txt_Matricule.SetFocus
    End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Form_Load()
Me.Width = 10380
Me.Height = 7935
Me.Move 0, 0
Me.WindowState = 2
End Sub

Private Sub Form_Resize()
On Error Resume Next
'Image1.Width = Me.Width
'CmdSave.Left = Me.Width - 700
'CmdFind.Left = Me.Width - 1100
'CmdDelete.Left = Me.Width - 1500
'CmdAdd.Left = Me.Width - 1900
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

Private Sub txt_CIN_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_CIN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_fonction_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_fonction_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_lieuPermi_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_lieuPermi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Matricule_GotFocus()
Call ViderZone(FrmPersonnel)
End Sub

Private Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Matricule_LostFocus()
Dim SQL As String
Dim rs As New ADODB.Recordset

On Error GoTo Err
If Len(Trim(txt_Matricule.Text)) > 0 Then
SQL = "Select * from personnel where code = " & SQLText(txt_Matricule.Text)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Nom.Text = rs("Libelle")
    txt_CIN.Text = rs("CIN")
    txt_fonction.Text = rs("Fonction")
    txt_telephone.Text = rs("telephone")
    txt_mobile.Text = rs("mobile")
    txt_permie.Text = rs("permie")
    cda_DateLivrPermi.Text = rs("datlivr")
    txt_lieuPermi.Text = rs("lieulivr")
'Else
'    MsgBox "Code introuvable", vbInformation
'    txt_Matricule.SetFocus
'    Exit Sub
End If
rs.Close
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub txt_mobile_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_mobile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Nom_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_Nom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_permie_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_permie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_telephone_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_telephone_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Public Sub AfficheRow(vcode As String)

Dim SQL As String
Dim rs As New ADODB.Recordset

Call ViderZone(FrmPersonnel)
SQL = "Select * from personnel where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    txt_Nom.Text = rs("Libelle")
    txt_CIN.Text = rs("CIN")
    txt_fonction.Text = rs("Fonction")
    txt_telephone.Text = rs("telephone")
    txt_mobile.Text = rs("mobile")
    txt_permie.Text = rs("permie")
    If rs("datlivr") <> "01/01/1900" Then cda_DateLivrPermi.Text = rs("datlivr")
    txt_lieuPermi.Text = rs("lieulivr")
    chk_Actif.Value = rs("Actif")
Else
    MsgBox "Code introuvable", vbInformation
    txt_Matricule.SetFocus
End If
rs.Close

End Sub
