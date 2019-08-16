VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmDestination 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Destinations"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleMode       =   0  'User
   ScaleWidth      =   13585.72
   Begin VB.CheckBox chk_Actif 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Actif ?"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2280
      TabIndex        =   12
      Top             =   3960
      Width           =   2775
   End
   Begin VB.ComboBox cb_Type 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txt_Numero 
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
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Text            =   "Auto"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox Txt_Destination 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
   End
   Begin SToolBox.SCommand cmdFindReparation 
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   1800
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
      Picture         =   "FrmTypeReparation.frx":0000
      ButtonType      =   1
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   7080
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
      Picture         =   "FrmTypeReparation.frx":0353
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   6360
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
      Picture         =   "FrmTypeReparation.frx":04D5
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   6720
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
      Picture         =   "FrmTypeReparation.frx":0828
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   6000
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
      Picture         =   "FrmTypeReparation.frx":0B7B
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type _ _ _ _ _ __ _ _ _ :"
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
      TabIndex        =   11
      Top             =   2640
      Width           =   2340
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Désignantion  _ _ _:"
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
      TabIndex        =   7
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro _  _ _ _ _:"
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
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Destinations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmTypeReparation.frx":0CFD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "FrmDestination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Okay As Boolean
Public ii As Integer
Dim thekey As Integer
Dim theshift As Integer

Public Sub AfficheRow(ByVal vcode As String)

Dim SQL As String
Dim rs As New ADODB.Recordset
Call ViderZone(FrmDestination)

SQL = "Select * from Destination where Numero = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Numero.Text = rs("Numero")
    cb_Type.Text = rs("Type")
    Txt_Destination.Text = rs("Libelle")
    chk_Actif.Value = rs("actif")
Else
    MsgBox "Code introuvable", vbInformation
    cb_Type.SetFocus
    Exit Sub
End If
rs.Close
End Sub




Private Sub cb_Type_Click()
'    If Len(Trim(cb_Type.Text)) > 0 Then Call AfficheRow_Type(cb_Type.Text)
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

Private Sub cb_Type_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"

End Sub

Private Sub cb_Type_KeyUp(KeyCode As Integer, Shift As Integer)
thekey = KeyCode
    theshift = Shift
End Sub

Private Sub CmdAdd_Click()

On Error GoTo Err

Dim rQ As New ADODB.Recordset
Dim SQL As String
SQL = "Select * from utilisateur where Ins_Fournisseur = 1 and code= " & LInt_UserId
rQ.Open SQL, CNB, adOpenDynamic
If rQ.EOF Then
    rQ.Close
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If

If txt_Numero.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Call ViderZone(FrmDestination)
txt_Numero.Text = "Auto"
cb_Type.SetFocus

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdDelete_Click()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim vcode As String

On Error GoTo Err

   If txt_Numero.Text = "Auto" Then
        If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        Else
            txt_Numero.SetFocus
            Exit Sub
        End If
    End If
    If txt_Numero.Text <> "Auto" Then
        Dim rQ As New ADODB.Recordset
        SQL = "Select * from utilisateur where Supp_Fournisseur = 1 and code= " & LInt_UserId
        rQ.Open SQL, CNB, adOpenDynamic
        If rQ.EOF Then
            rQ.Close
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
    End If

    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "Type Reparation", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    vcode = txt_Numero.Text
    SQL = "Delete from Destination where Numero =" & SQLText(vcode)
    CNB.Execute SQL
     Call ViderZone(FrmDestination)
    txt_Numero.SetFocus
    End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub


Private Sub CmdFind_Click()
Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Destination"
    .Show
End With
End Sub

Private Sub cmdFindReparation_Click()
Unload FrmFind_Fils
With FrmFind_Actif
    .StrSource = "Destination"
    .Show
End With
End Sub



Private Sub CmdSave_Click()

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim vcode
Dim LInt_NumCompteur
Dim LObj_Compteur

    If Left(CheckMandatory(FrmDestination), 1) = 1 Then
       Exit Sub
    End If
    
On Error GoTo Err
    If txt_Numero.Text <> "Auto" Then
        Dim rQ As New ADODB.Recordset
        SQL = "Select * from utilisateur where Maj_Fournisseur = 1 and code= " & LInt_UserId
        rQ.Open SQL, CNB, adOpenDynamic
        If rQ.EOF Then
            rQ.Close
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
    End If

    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    'Delete l'enregistrement
    vcode = txt_Numero.Text
    CNB.BeginTrans
    SQL = "Delete from Destination where Numero =" & SQLText(vcode)
    CNB.Execute SQL
    If vcode = "Auto" Then
    LInt_NumCompteur = return_Compteur() + 1
    'Insertion enregistrement assiette
    txt_Numero.Text = Format(LInt_NumCompteur, "00000")
    End If
    'Insertion enregistrement
    SQL = "Insert into Destination  (Numero,Type,Libelle, Actif) values ("
    SQL = SQL & SQLText(txt_Numero.Text)
    SQL = SQL & "," & SQLText(cb_Type.Text)
    SQL = SQL & "," & SQLText(Txt_Destination.Text)
    SQL = SQL & "," & chk_Actif.Value
    SQL = SQL & ")"
    CNB.Execute SQL
    CNB.CommitTrans
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
    txt_Numero.SetFocus
    End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Function return_Compteur() As Long
Dim rD As New ADODB.Recordset
Dim SQL As String
return_Compteur = 0
SQL = "select Max(Numero) from Destination "
rD.Open SQL, CNB, adOpenKeyset
If Not rD.EOF Then
return_Compteur = rD(0)
End If
rD.Close
End Function


Private Sub Form_Load()
Me.WindowState = 2
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

Private Sub txt_Numero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Public Sub AfficheRow_Type(ByVal vcode As String)

Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select Distinct(Type) from Destination )"
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Type")) Then
    cb_Type.Text = rs("Type")
    End If
Else
    MsgBox "Type introuvable", vbInformation
    cb_Type.SetFocus
    Exit Sub
End If
rs.Close

End Sub








