VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmArticles 
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
   Icon            =   "FrmArticles.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   10260
   Begin VB.OptionButton Op_Lub 
      BackColor       =   &H80000009&
      Caption         =   "Lubrifiant"
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
      Left            =   4440
      TabIndex        =   6
      Top             =   5760
      Width           =   1575
   End
   Begin VB.OptionButton Op_Prod 
      BackColor       =   &H80000009&
      Caption         =   "Produit"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CheckBox chk_Actif 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   6480
      Width           =   255
   End
   Begin VB.TextBox txt_tva 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "M"
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox txt_tht 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "M"
      Top             =   3600
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2160
      ScaleHeight     =   615
      ScaleWidth      =   3495
      TabIndex        =   14
      Top             =   4920
      Width           =   3500
      Begin VB.TextBox txt_prix 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "M"
         Top             =   120
         Width           =   3015
      End
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   2040
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
      Picture         =   "FrmArticles.frx":0ECA
      ButtonType      =   1
   End
   Begin VB.TextBox txt_libelle 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "M"
      Top             =   2880
      Width           =   3015
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "M"
      Top             =   2040
      Width           =   3015
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   9720
      TabIndex        =   8
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
      Picture         =   "FrmArticles.frx":121D
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   8760
      TabIndex        =   10
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
      Picture         =   "FrmArticles.frx":139F
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   9240
      TabIndex        =   9
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
      Picture         =   "FrmArticles.frx":16F2
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   8280
      TabIndex        =   18
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
      Picture         =   "FrmArticles.frx":1A45
   End
   Begin VB.Label Lbl_typ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche Articles"
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
      TabIndex        =   20
      Top             =   600
      Width           =   2235
   End
   Begin VB.Image PicBox_Header 
      Height          =   1335
      Left            =   0
      Picture         =   "FrmArticles.frx":1BC7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Actif : O/N"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prix ht:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   480
      TabIndex        =   17
      Top             =   3600
      Width           =   915
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taux tva:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   480
      TabIndex        =   16
      Top             =   4320
      Width           =   1170
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prix ttc:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   480
      TabIndex        =   15
      Top             =   5040
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Libelle :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   480
      TabIndex        =   13
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Matricule :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "FrmArticles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Saisi nouveaux Produits et lubrifiants

Private Sub CmdAdd_Click()

Dim LOBJ_Personnel As Personnel

On Error GoTo Err
If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
    If MsgBox("Annuler le MAJ en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Set LOBJ_Personnel = New Personnel
' Vérifier les droits d'accès de l'utilisateur : s'il a le droit d'ajouter un nouveau bon.
If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Ins_Produit", LInt_UserId) Then
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If

Call ViderZone(FrmArticles)
txt_Matricule.Text = "Auto"
txt_libelle.SetFocus
Op_Prod.Value = False
Op_Lub.Value = False
chk_Actif.Value = vbUnchecked
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()

Dim LOBJ_Personnel As Personnel
Dim LOBJ_Prod As Produit_Lubrifiant

On Error GoTo Err

If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        txt_Matricule.SetFocus
        Exit Sub
    End If
End If
If txt_Matricule.Text <> "Auto" Then
    Set LOBJ_Personnel = New Personnel
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Supp_Produit", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
End If

If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "produit", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
Set LOBJ_Prod = New Produit_Lubrifiant
Call LOBJ_Prod.Delete_ProdLub(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt_Matricule.Text)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

txt_Matricule.SetFocus

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdFind_Click()

On Error Resume Next
If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Unload FrmFind
Unload FrmFind_Actif
Unload FrmFind_Fils
With FrmFind
    .StrSource = "Produits"
    .Show vbModal
End With
End Sub

Private Sub cmdFindMatricule_Click()
On Error GoTo Err
Unload FrmFind
Unload FrmFind_Actif
Unload FrmFind_Fils
With FrmFind_Actif
    .StrSource = "Produits"
    .Show
End With
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdSave_Click()

Dim LOBJ_Personnel As Personnel
Dim rs As New Recordset

If Left(CheckMandatory(FrmArticles), 1) = 1 Then
   Exit Sub
End If
    
On Error GoTo Err

If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub

If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
    Set LOBJ_Personnel = New Personnel
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Maj_produit", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
   ' Modification du produit
    Call Modif_Produit
ElseIf txt_Matricule.Text = "Auto" Then
    'Insertion enregistrement
    Call Ajout_Produit
End If
    
MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
txt_Matricule.SetFocus
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Ajout_Produit()

Dim LInt_NumCompteur As Long
Dim LRs_NewRecord As New Recordset
Dim LOBJ_Produit As Produit_Lubrifiant

Set LOBJ_Produit = New Produit_Lubrifiant

'Insertion enregistrement
Set LRs_NewRecord = CreateEmptyRS_Prod_Lub()
With LRs_NewRecord
    .AddNew
    .Fields("Libelle") = txt_libelle.Text
    .Fields("prixht") = CDbl(txt_tht.Text)
    .Fields("DatePrix") = Format(Date)
    .Fields("tva") = CDbl(txt_tva.Text)
    If Op_Prod.Value = True Then
        .Fields("Type_PL") = "Produit"
    ElseIf Op_Lub.Value = True Then
        .Fields("Type_PL") = "Lubrifiant"
    End If
    If (chk_Actif.Value = 1) Then
        .Fields("Actif") = "O"
    ElseIf (chk_Actif.Value = 0) Then
         .Fields("Actif") = "N"
    End If
     .Fields("OperateurSaisi") = LStr_NameUser
End With
Call LOBJ_Produit.Insert_ProdLub(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
txt_Matricule.Text = Format(return_Compteur(), "00000")
Set LRs_NewRecord = Nothing

End Sub

Private Sub Modif_Produit()

Dim LRs_NewRecord As New Recordset
Dim LOBJ_Produit As Produit_Lubrifiant

Set LOBJ_Produit = New Produit_Lubrifiant

Set LRs_NewRecord = CreateEmptyRS_Prod_Lub()
With LRs_NewRecord
    .AddNew
    .Fields("Numero") = txt_Matricule.Text
    .Fields("Libelle") = txt_libelle.Text
    .Fields("prixht") = CDbl(txt_tht.Text)
    .Fields("DatePrix") = Format(Date)
    .Fields("tva") = CDbl(txt_tva.Text)
    If Op_Prod.Value = True Then
        .Fields("Type_PL") = "Produit"
    ElseIf Op_Lub.Value = True Then
        .Fields("Type_PL") = "Lubrifiant"
    End If
    If (chk_Actif.Value = 1) Then
        .Fields("Actif") = "O"
    ElseIf (chk_Actif.Value = 0) Then
         .Fields("Actif") = "N"
    End If
    .Fields("OperateurSaisi") = LStr_NameUser
End With
Call LOBJ_Produit.Update_ProdLub(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing

End Sub

Private Function return_Compteur() As Long

Dim rs As New Recordset
Dim LOBJ_Prod As Produit_Lubrifiant
return_Compteur = 0
Set LOBJ_Prod = New Produit_Lubrifiant
Set rs = LOBJ_Prod.Get_MaxNum(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
If Not rs.EOF Then
    return_Compteur = rs("maxnum")
End If

rs.Close
End Function

Private Sub Form_Load()
Me.Width = 10380
Me.Height = 7935
Me.Move 0, 0
Me.WindowState = 2
End Sub

Private Sub Form_Resize()

On Error Resume Next
PicBox_Header.Width = Me.Width
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

Private Sub txt_libelle_GotFocus()

On Error GoTo Err
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Txt_Libelle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Matricule_GotFocus()
Call ViderZone(FrmArticles)
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

Private Sub txt_prix_GotFocus()
On Error GoTo Err
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub txt_prix_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Public Sub AfficheRow(ByVal vcode As String)

Dim LOBJ_Prod As Produit_Lubrifiant
Dim rs As New Recordset

Call ViderZone(FrmArticles)
Set LOBJ_Prod = New Produit_Lubrifiant
Set rs = LOBJ_Prod.Get_ProdLubBycode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vcode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Numero")
    txt_libelle.Text = rs("Libelle")
    txt_tht.Text = Format(rs("prixht"), "##0.000")
    txt_tva.Text = rs("tva")
    txt_prix.Text = Format(rs("prixht") + (rs("prixht") * rs("tva") / 100), "##0.000")
    
    If (rs("Type_PL") = "Produit") Then
        Op_Prod.Value = True
    ElseIf (rs("Type_PL") = "Lubrifiant") Then
        Op_Lub.Value = True
    End If
    
    If (rs("Actif") = "O") Then
        chk_Actif.Value = 1
    ElseIf (rs("Actif") = "N") Then
        chk_Actif.Value = 0
    End If
Else
    MsgBox "Code introuvable", vbInformation
    txt_Matricule.SetFocus
End If
rs.Close

End Sub

Private Sub txt_prix_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_tht_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
    Exit Sub
End If
txt_tht.SelStart = 0
txt_tht.SelLength = Len(txt_tht.Text)
End Sub

Private Sub txt_tht_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"

End Sub

Private Sub txt_tht_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub

Private Sub txt_tht_LostFocus()

Dim tht As Double
Dim tva As Double
Dim ttc As Double

On Error GoTo Err
    ttc = 0
    tht = 0
    tva = 0
    If Len(Trim(txt_tva.Text)) = 0 Then
        tva = 0
    Else
        tva = txt_tva.Text
    End If
    If Len(Trim(txt_tht.Text)) = 0 Then
        tht = 0
        txt_tht.Text = 0
    Else
        tht = txt_tht.Text
    End If
    ttc = tht + (tht * (tva / 100))
    txt_prix.Text = Format(ttc, "##0.000")
    txt_tht.Text = Format(tht, "##0.000")
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Txt_tva_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If

End Sub

Private Sub Txt_tva_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"

End Sub

Private Sub Txt_tva_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub

Private Sub txt_tva_LostFocus()

Dim tht As Double
Dim tva As Double
Dim ttc As Double

On Error GoTo Err
    ttc = 0
    tht = 0
    tva = 0
    If Len(Trim(txt_tva.Text)) = 0 Then
        tva = 0
    Else
        tva = txt_tva.Text
    End If
    If Len(Trim(txt_tht.Text)) = 0 Then
        tht = 0
        txt_tht.Text = 0
    Else
        tht = txt_tht.Text
    End If
    ttc = tht + (tht * (tva / 100))
    txt_tva.Text = Format(txt_tva.Text, "##0.00")
    txt_prix.Text = Format(ttc, "##0.000")
Exit Sub
Err:
MsgBox Err.Description, vbInformation


End Sub
