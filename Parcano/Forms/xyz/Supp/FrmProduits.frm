VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmProduits 
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
   Icon            =   "FrmProduits.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   10260
   Begin VB.CheckBox chk_Actif 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox txt_tva 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Tag             =   "M"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txt_tht 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Tag             =   "M"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      ScaleHeight     =   375
      ScaleWidth      =   1335
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
      Begin VB.TextBox txt_prix 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Tag             =   "M"
         Top             =   0
         Width           =   1215
      End
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   3960
      TabIndex        =   8
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
      Picture         =   "FrmProduits.frx":0ECA
      ButtonType      =   1
   End
   Begin VB.TextBox txt_libelle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "M"
      Top             =   2520
      Width           =   4935
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
      Left            =   1560
      TabIndex        =   0
      Tag             =   "M"
      Top             =   1560
      Width           =   2295
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   9360
      TabIndex        =   4
      Top             =   720
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
      Picture         =   "FrmProduits.frx":121D
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   8400
      TabIndex        =   6
      Top             =   720
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
      Picture         =   "FrmProduits.frx":139F
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   8880
      TabIndex        =   5
      Top             =   720
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
      Picture         =   "FrmProduits.frx":16F2
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   7920
      TabIndex        =   15
      Top             =   720
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
      Picture         =   "FrmProduits.frx":1A45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche lubrifiant"
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
      Left            =   480
      TabIndex        =   18
      Top             =   480
      Width           =   2535
   End
   Begin VB.Image PicBox_Header 
      Height          =   1335
      Left            =   0
      Picture         =   "FrmProduits.frx":1BC7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Actif :O/N"
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
      Left            =   240
      TabIndex        =   17
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prix ht:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   705
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "taux tva:"
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
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prix ttc:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Libelle :"
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
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   735
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
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "FrmProduits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Lubrifiant

Private Sub CmdAdd_Click()

On Error GoTo Err

Dim LOBJ_Personnel As Personnel

Set LOBJ_Personnel = New Personnel
If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Ins_Lub", LInt_UserId) Then
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If

If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Call ViderZone(FrmProduits)
txt_Matricule.Text = "Auto"
txt_libelle.SetFocus
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdDelete_Click()

Dim LOBJ_Personnel As Personnel
Dim LOBJ_Lubr As Lubrifiant

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
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Supp_Lub", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
End If

If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "produit", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    Set LOBJ_Lubr = New Lubrifiant
    Call LOBJ_Lubr.Delete_Lubrif(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt_Matricule.Text)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    txt_Matricule.SetFocus
End If
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
With FrmFind
    .StrSource = "Produit"
    .Show vbModal
End With
End Sub

Private Sub cmdFindMatricule_Click()
Unload FrmFind
Unload FrmFind_Actif
With FrmFind_Actif
    .StrSource = "Produit"  'Lubrifiant
    .Show
End With
End Sub

Private Sub CmdSave_Click()

Dim SQL As String
Dim rs As New ADODB.Recordset

    If Left(CheckMandatory(FrmProduits), 1) = 1 Then
       Exit Sub
    End If
    
On Error GoTo Err

    If txt_Matricule.Text <> "Auto" Then
        Dim rQ As New ADODB.Recordset
        SQL = "Select * from utilisateur where Maj_Lub = 1 and code= " & LInt_UserId
        rQ.Open SQL, CNB, adOpenDynamic
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
    SQL = "Delete from produit where code =" & SQLText(vcode)
    CNB.Execute SQL
    If vcode = "Auto" Then
    LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_Produit")
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
    SQL = "Insert into produit  (Code,libelle,tht,tva,Prix, Actif) values ("
    SQL = SQL & SQLText(txt_Matricule.Text)
    SQL = SQL & "," & SQLText(txt_libelle.Text)
    SQL = SQL & "," & Replace(txt_tht.Text, ",", ".")
    SQL = SQL & "," & Replace(txt_tva.Text, ",", ".")
    SQL = SQL & "," & Replace(txt_prix.Text, ",", ".")
    If (chk_Actif.Value = 1) Then
        SQL = SQL & "," & SQLText("O")
    ElseIf (chk_Actif.Value = 0) Then
        SQL = SQL & "," & SQLText("N")
    End If
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

Private Sub txt_libelle_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If

End Sub

Private Sub Txt_Libelle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Matricule_GotFocus()
Call ViderZone(FrmProduits)
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
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If

End Sub

Private Sub txt_prix_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Public Sub AfficheRow(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset
Call ViderZone(FrmProduits)
SQL = "Select * from produit where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    txt_libelle.Text = rs("Libelle")
    txt_tht.Text = Format(rs("tht"), "##0.000")
    txt_tva.Text = rs("tva")
    txt_prix.Text = Format(rs("prix"), "##0.000")
    If (rs("actif") = "O") Then
        chk_Actif.Value = 1
    ElseIf (rs("actif") = "N") Then
        chk_Actif.Value = 0
    End If
    
Else
    MsgBox "Code introuvable", vbInformation
    txt_Matricule.SetFocus
End If
rs.Close

End Sub
Private Sub txt_tht_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If

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

Private Sub txt_tva_GotFocus()
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
    txt_prix.Text = Format(ttc, "##0.000")
Exit Sub
Err:
MsgBox Err.Description, vbInformation


End Sub
