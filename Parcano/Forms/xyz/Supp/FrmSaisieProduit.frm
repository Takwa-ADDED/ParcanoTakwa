VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmSaisieProduit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parcano"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSaisieProduit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   7815
      TabIndex        =   8
      Top             =   2280
      Width           =   7815
      Begin VB.TextBox txt_tva 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1485
         TabIndex        =   15
         Tag             =   "M"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txt_tht 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1485
         TabIndex        =   14
         Tag             =   "M"
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   885
         ScaleHeight     =   375
         ScaleWidth      =   2775
         TabIndex        =   12
         Top             =   1080
         Width           =   2775
         Begin VB.TextBox txt_prix 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   600
            TabIndex        =   13
            Tag             =   "M"
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   645
         ScaleHeight     =   375
         ScaleWidth      =   1335
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txt_libelle 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1485
         TabIndex        =   9
         Tag             =   "M"
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prix ht:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   0
         TabIndex        =   18
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taux tva:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   0
         TabIndex        =   17
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prix ttc:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Libelle :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   630
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   380
      Left            =   4920
      TabIndex        =   7
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Annuler"
      Height          =   380
      Left            =   6000
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txt_Qte 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Tag             =   "M"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txt_Matricule 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
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
      Left            =   1440
      TabIndex        =   1
      Tag             =   "M"
      Top             =   1560
      Width           =   2295
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   3960
      TabIndex        =   0
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
      Picture         =   "FrmSaisieProduit.frx":0ECA
      ButtonType      =   1
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qte:"
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
      Left            =   720
      TabIndex        =   5
      Top             =   3840
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saisie produit lubrifiant"
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
      TabIndex        =   3
      Top             =   240
      Width           =   3390
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
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmSaisieProduit.frx":121D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "FrmSaisieProduit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdDelete_Click()
Dim SQL As String
Dim rs As New ADODB.Recordset

On Error GoTo Err
    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "produit", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    SQL = "Delete from produit where code =" & SQLText(vcode)
    CNB.Execute SQL
    txt_Matricule.SetFocus
    End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdFind_Click()
Unload FrmFind
With FrmFind
    .StrSource = "Produit"
    .Show
End With
End Sub

Private Sub cmdFindMatricule_Click()
Unload FrmFind
With FrmFind
    .StrSource = "ProduitSaisie"
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
    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    'Delete l'enregistrement
    vcode = txt_Matricule.Text
    CNB.BeginTrans
    SQL = "Delete from produit where code =" & SQLText(vcode)
    CNB.Execute SQL
    'Insertion enregistrement
    SQL = "Insert into produit  (Code,libelle,Prix) values ("
    SQL = SQL & SQLText(txt_Matricule.Text)
    SQL = SQL & "," & SQLText(txt_libelle.Text)
    SQL = SQL & "," & Replace(txt_prix.Text, ",", ".")
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

Private Sub Command1_Click()

Dim ii
Dim Okay As Boolean

On Error GoTo Err
    
     
    Okay = False
    With FrmLubrifiant.Grid
        For ii = 1 To .Rows
            If txt_Matricule.Text = .CellText(ii, 1) Then
                Okay = True
                Exit For
            End If
        Next
    End With
    
    If Val(txt_Qte.Text) = 0 Then
        MsgBox "Qte invalid  ", vbInformation
        Exit Sub
    End If
    
    If Okay = False Then
        With FrmLubrifiant
            .Grid.AddRow
            .Grid.CellDetails .Grid.Rows, 1, txt_Matricule.Text
            .Grid.CellDetails .Grid.Rows, 2, txt_libelle.Text
            .Grid.CellDetails .Grid.Rows, 3, txt_Qte.Text
            .Grid.CellDetails .Grid.Rows, 4, txt_prix.Text
        End With
    Else
        MsgBox "Produit déja saisie ", vbInformation
    End If
    Unload Me

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Width = 7400
Me.Height = 5130
End Sub
Private Sub Form_Resize()
On Error Resume Next
Image1.Width = Me.Width
CmdSave.Left = Me.Width - 700
CmdFind.Left = Me.Width - 1100
CmdDelete.Left = Me.Width - 1500
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

Private Sub txt_mobile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_mobile_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub


Private Sub txt_rsocial_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub txt_ville_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
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

SQL = "Select * from produit where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    txt_libelle.Text = rs("Libelle")
    txt_tht.Text = Format(rs("tht"), "##0.000")
    Txt_tva.Text = rs("tva")
    txt_prix.Text = Format(rs("prix"), "##0.000")
End If
rs.Close

End Sub

Private Sub txt_Qte_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub

