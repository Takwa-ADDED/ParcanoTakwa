VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmTournee 
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   2
      Tag             =   "M"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txt_libelle 
      Height          =   315
      Left            =   1620
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "M"
      Top             =   2880
      Width           =   5655
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   2160
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
      Picture         =   "FrmTournee.frx":0000
      ButtonType      =   1
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   9720
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
      Picture         =   "FrmTournee.frx":0353
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   8880
      TabIndex        =   4
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
      Picture         =   "FrmTournee.frx":04D5
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   9240
      TabIndex        =   5
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
      Picture         =   "FrmTournee.frx":0828
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   8520
      TabIndex        =   6
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
      Picture         =   "FrmTournee.frx":0B7B
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche Tournée"
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
      Width           =   2010
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
      TabIndex        =   8
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Libelle :"
      Height          =   195
      Left            =   1020
      TabIndex        =   7
      Top             =   2880
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmTournee.frx":0CFD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "FrmTournee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdd_Click()

On Error GoTo Err

'Dim rQ As New ADODB.Recordset
'Dim SQL As String
'SQL = "Select * from utilisateur where INS_LUB = 1 and code= " & LInt_UserId
'rQ.Open SQL, CNB, adOpenDynamic
'If rQ.EOF Then
'    rQ.Close
'    MsgBox "Accès refusé.", vbExclamation
'    Exit Sub
'End If

If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Call ViderZone(FrmTournee)
'txt_Matricule.Enabled = False
txt_Matricule.Text = "Auto"
txt_libelle.SetFocus
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
'    If txt_Matricule.Text <> "Auto" Then
'        Dim rQ As New ADODB.Recordset
'        SQL = "Select * from utilisateur where MAJ_lub = 1 and code= " & LInt_UserId
'        rQ.Open SQL, CNB, adOpenDynamic
'        If rQ.EOF Then
'            rQ.Close
'            MsgBox "Accès refusé.", vbExclamation
'            Exit Sub
'        End If
'    End If

    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "produit", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    vcode = txt_Matricule.Text
    SQL = "Delete from Tournee where Numero =" & SQLText(vcode)
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
    .StrSource = "Tournee"
    .Show
End With
End Sub


Private Sub cmdFindMatricule_Click()
Unload FrmFind
With FrmFind
    .StrSource = "Tournee"""
    .Show
End With
End Sub

Public Sub AfficheRow(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset
Call ViderZone(FrmProduits)
SQL = "Select * from Tournee where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Numero")
    txt_libelle.Text = rs("Libelle")
Else
    MsgBox "Code introuvable", vbInformation
    txt_Matricule.SetFocus
End If
rs.Close

End Sub

Private Sub CmdSave_Click()
Dim LInt_NumCompteur As Long

Dim SQL As String
Dim rs As New ADODB.Recordset

'    If Left(CheckMandatory(FrmProduits), 1) = 1 Then
'       Exit Sub
'    End If
    
On Error GoTo Err

'    If txt_Matricule.Text <> "Auto" Then
'        Dim rQ As New ADODB.Recordset
'        SQL = "Select * from utilisateur where MAJ_lub = 1 and code= " & LInt_UserId
'        rQ.Open SQL, CNB, adOpenDynamic
'        If rQ.EOF Then
'            rQ.Close
'            MsgBox "Accès refusé.", vbExclamation
'            Exit Sub
'        End If
'    End If


    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    'Delete l'enregistrement
    vcode = txt_Matricule.Text
    CNB.BeginTrans
    SQL = "Delete from Tournee where Numero =" & SQLText(vcode)
    CNB.Execute SQL
    
    If vcode = "Auto" Then
    LInt_NumCompteur = Return_Compteur() + 1
    txt_Matricule.Text = Format(LInt_NumCompteur, "00000")
    End If
    'Insertion enregistrement
    SQL = "Insert into Tournee  (Numero,libelle,tht,tva,Prix) values ("
    SQL = SQL & SQLText(txt_Matricule.Text)
    SQL = SQL & "," & SQLText(txt_libelle.Text)
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


Private Function Return_Compteur() As Long
Dim rD As New ADODB.Recordset
Dim SQL As String
Return_Compteur = 0
SQL = "select Max(Numero) from Tournee "
rD.Open SQL, CNB, adOpenKeyset
If Not rD.EOF Then
Return_Compteur = rD(0)
End If

rD.Close
End Function

