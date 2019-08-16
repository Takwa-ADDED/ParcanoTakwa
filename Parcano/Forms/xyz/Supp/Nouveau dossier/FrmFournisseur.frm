VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmFournisseur 
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
   Icon            =   "FrmFournisseur.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   10260
   Begin VB.TextBox txt_codePostal 
      Height          =   315
      Left            =   1620
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txt_ville 
      Height          =   315
      Left            =   1620
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "M"
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox txt_adresse 
      Height          =   315
      Left            =   1605
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "M"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox txt_mobile 
      Height          =   315
      Left            =   1620
      MaxLength       =   10
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txt_activite 
      Height          =   315
      Left            =   1620
      MaxLength       =   50
      TabIndex        =   5
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox txt_email 
      Height          =   315
      Left            =   1620
      MaxLength       =   30
      TabIndex        =   9
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox txt_fax 
      Height          =   315
      Left            =   1620
      MaxLength       =   10
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txt_telbureau 
      Height          =   315
      Left            =   1620
      MaxLength       =   10
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   3960
      TabIndex        =   15
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
      Picture         =   "FrmFournisseur.frx":0ECA
      ButtonType      =   1
   End
   Begin VB.TextBox txt_libelle 
      Height          =   315
      Left            =   1620
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "M"
      Top             =   2280
      Width           =   7695
   End
   Begin VB.TextBox txt_Matricule 
      Alignment       =   2  'Center
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
      Left            =   1620
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
      Picture         =   "FrmFournisseur.frx":121D
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   8880
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
      Picture         =   "FrmFournisseur.frx":139F
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   9240
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
      Picture         =   "FrmFournisseur.frx":16F2
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   8520
      TabIndex        =   25
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
      Picture         =   "FrmFournisseur.frx":1A45
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code postal :"
      Height          =   195
      Left            =   600
      TabIndex        =   24
      Top             =   3360
      Width           =   960
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ville :"
      Height          =   195
      Left            =   1185
      TabIndex        =   23
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adresse :"
      Height          =   195
      Left            =   855
      TabIndex        =   22
      Top             =   2640
      Width           =   690
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile(GSM) :"
      Height          =   195
      Left            =   570
      TabIndex        =   21
      Top             =   4440
      Width           =   990
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activité :"
      Height          =   195
      Left            =   915
      TabIndex        =   20
      Top             =   3720
      Width           =   645
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail :"
      Height          =   195
      Left            =   1035
      TabIndex        =   19
      Top             =   5160
      Width           =   525
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "telephone bureau :"
      Height          =   195
      Left            =   180
      TabIndex        =   18
      Top             =   4080
      Width           =   1380
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax :"
      Height          =   195
      Left            =   1185
      TabIndex        =   17
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Raison sociale :"
      Height          =   195
      Left            =   450
      TabIndex        =   16
      Top             =   2280
      Width           =   1110
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Matricule"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   14
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche fournisseur"
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
      TabIndex        =   13
      Top             =   240
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmFournisseur.frx":1BC7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "FrmFournisseur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAdd_Click()
On Error GoTo Err

Dim rQ As New ADODB.Recordset
Dim sql As String
sql = "Select * from utilisateur where INS_FR = 1 and code= " & LInt_UserId
rQ.Open sql, CNB, adOpenDynamic
If rQ.EOF Then
    rQ.Close
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If


If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Call ViderZone(FrmFournisseur)
'txt_Matricule.Enabled = False
txt_Matricule.Text = "Auto"
txt_libelle.SetFocus
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()
Dim sql As String
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
        sql = "Select * from utilisateur where MAJ_FR = 1 and code= " & LInt_UserId
        rQ.Open sql, CNB, adOpenDynamic
        If rQ.EOF Then
            rQ.Close
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
    End If
    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "fournisseur", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    vcode = txt_Matricule.Text
    sql = "Delete from fournisseur where code =" & SQLText(vcode)
    CNB.Execute sql
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
    .StrSource = "Fournisseur"
    .Show
End With
End Sub

Private Sub cmdFindMatricule_Click()
Unload FrmFind
With FrmFind
    .StrSource = "Fournisseur"
    .Show
End With
End Sub

Private Sub CmdSave_Click()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim vcode As String

On Error GoTo Err
    If Left(CheckMandatory(FrmFournisseur), 1) = 1 Then
       Exit Sub
    End If
     If txt_Matricule.Text <> "Auto" Then
        Dim rQ As New ADODB.Recordset
        sql = "Select * from utilisateur where MAJ_FR = 1 and code= " & LInt_UserId
        rQ.Open sql, CNB, adOpenDynamic
        If rQ.EOF Then
            rQ.Close
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
    End If
    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    'Delete l'enregistrement
    vcode = txt_Matricule.Text

    CNB.BeginTrans
    sql = "Delete from fournisseur where code =" & SQLText(vcode)
    CNB.Execute sql
    If vcode = "Auto" Then
    LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_Fournisseur")
    If ErrNumber <> 0 Then
       MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
       ErrNumber = 0
       Exit Sub
    End If
    Set LObj_Compteur = Nothing
    'Insertion enregistrement assiette
    txt_Matricule.Text = Format(LInt_NumCompteur, "00000")
    End If
    'Insertion enregistrement
    sql = "Insert into fournisseur (Code,Libelle,Adresse,Ville,CPOSTAL,Activite,telephone,mobile,fax,email) values ("
    sql = sql & SQLText(txt_Matricule.Text)
    sql = sql & "," & SQLText(txt_libelle.Text)
    If Len(Trim(txt_adresse.Text)) = 0 Then sql = sql & ",NULL" Else sql = sql & "," & SQLText(txt_adresse.Text)
    If Len(Trim(txt_ville.Text)) = 0 Then sql = sql & ",NULL" Else sql = sql & "," & SQLText(txt_ville.Text)
    If Len(Trim(txt_codepostal.Text)) = 0 Then sql = sql & ",0" Else sql = sql & "," & Val((txt_codepostal.Text))
    If Len(Trim(txt_activite.Text)) = 0 Then sql = sql & ",NULL" Else sql = sql & "," & SQLText(txt_activite.Text)
    If Len(Trim(txt_telbureau.Text)) = 0 Then sql = sql & ",NULL" Else sql = sql & "," & SQLText(txt_telbureau.Text)
    If Len(Trim(txt_mobile.Text)) = 0 Then sql = sql & ",NULL" Else sql = sql & "," & SQLText(txt_mobile.Text)
    If Len(Trim(txt_fax.Text)) = 0 Then sql = sql & ",NULL" Else sql = sql & "," & SQLText(txt_fax.Text)
    If Len(Trim(txt_email.Text)) = 0 Then sql = sql & ",NULL" Else sql = sql & "," & SQLText(txt_email.Text)
    sql = sql & ")"
    
    CNB.Execute sql
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
End Sub

Private Sub Form_Resize()
On Error Resume Next
Image1.Width = Me.Width
CmdSave.Left = Me.Width - 700
CmdFind.Left = Me.Width - 1100
CmdDelete.Left = Me.Width - 1500
CmdAdd.Left = Me.Width - 1900
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

Private Sub txt_activite_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub
Private Sub txt_activite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_adresse_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub
Private Sub txt_adresse_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_codepostal_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub
Private Sub txt_codepostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_email_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub
Private Sub txt_email_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_fax_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub
Private Sub txt_fax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_libelle_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub
Private Sub txt_libelle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_Matricule_GotFocus()
Call ViderZone(FrmFournisseur)
End Sub
Private Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Matricule_LostFocus()

On Error GoTo Err

If Len(Trim(txt_Matricule.Text)) > 0 Then Call AfficheRow(txt_Matricule.Text)

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
Private Sub txt_telbureau_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub
Private Sub txt_telbureau_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txt_ville_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub
Private Sub txt_ville_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Public Sub AfficheRow(ByVal vcode As String)
Dim sql As String
Dim rs As New ADODB.Recordset
Call ViderZone(FrmFournisseur)
sql = "Select * from fournisseur where code = " & SQLText(vcode)
rs.Open sql, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_libelle.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
    If Not IsNull(rs("CPOSTAL")) Then txt_codepostal.Text = rs("CPOSTAL")
    If Not IsNull(rs("Activite")) Then txt_activite.Text = rs("Activite")
    If Not IsNull(rs("telephone")) Then txt_telbureau.Text = rs("telephone")
    If Not IsNull(rs("mobile")) Then txt_mobile.Text = rs("mobile")
    If Not IsNull(rs("fax")) Then txt_fax.Text = rs("fax")
    If Not IsNull(rs("email")) Then txt_email.Text = rs("email")
Else
    MsgBox "Code introuvable", vbInformation
    txt_Matricule.SetFocus
    Exit Sub
End If

End Sub


