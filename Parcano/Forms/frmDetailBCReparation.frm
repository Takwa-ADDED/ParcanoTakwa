VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "STOOLBOX.OCX"
Begin VB.Form frmDetailBCReparation 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detail BC Reparation"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2520
      ScaleHeight     =   615
      ScaleWidth      =   3135
      TabIndex        =   14
      Top             =   1560
      Width           =   3135
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
         TabIndex        =   15
         Text            =   "Auto"
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.TextBox txt_Observation 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   2520
      TabIndex        =   3
      Top             =   4680
      Width           =   4935
   End
   Begin VB.ComboBox cbo_Matricule 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox Txt_Libelle 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox txt_Qte 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   3015
   End
   Begin VB.CommandButton CmdAjout 
      Appearance      =   0  'Flat
      Caption         =   "Ajouter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton CmdAnnul 
      Appearance      =   0  'Flat
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   6360
      Width           =   2175
   End
   Begin SToolBox.SCommand CmdFindDesi 
      Height          =   345
      Left            =   5640
      TabIndex        =   11
      Top             =   2520
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
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
      Picture         =   "frmDetailBCReparation.frx":0000
      ButtonType      =   1
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   345
      Index           =   1
      Left            =   5640
      TabIndex        =   12
      Top             =   3960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
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
      Picture         =   "frmDetailBCReparation.frx":0353
      ButtonType      =   1
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Detail BC Reparation"
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
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   360
      Width           =   3855
   End
   Begin VB.Image PicBox_Header 
      Height          =   1215
      Left            =   0
      Picture         =   "frmDetailBCReparation.frx":06A6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observation :"
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
      Left            =   480
      TabIndex        =   10
      Top             =   4800
      Width           =   1320
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Véhicule :"
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
      Left            =   480
      TabIndex        =   9
      Top             =   3960
      Width           =   945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro :"
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
      Left            =   480
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Désignation  :"
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
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantité : "
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
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   1035
   End
End
Attribute VB_Name = "frmDetailBCReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Okay As Boolean
Public ii As Integer
Dim thekey As Integer
Dim theshift As Integer

Private Sub Cbo_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cbo_Matricule_Click()
If Len(Trim(Cbo_Matricule.Text)) > 0 Then Call AfficheRow_Vehicule(Cbo_Matricule.Text)

End Sub

Private Sub cbo_Matricule_GotFocus()

On Error GoTo Err

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
    Else
    Call Affiche_Matricule_Combo(Cbo_Matricule)
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub cbo_Matricule_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub

Private Sub Cbo_Matricule_LostFocus()

On Error GoTo Err

If Len(Trim(Cbo_Matricule.Text)) > 0 Then Call AfficheRow_Vehicule(Cbo_Matricule.Text)

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdAnnul_Click()
If MsgBox("Voulez vous annuler l'opération en cours ?", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
Unload Me
End Sub

'Afficher liste des véhicules
Private Sub cmdFindMatricule_Click(Index As Integer)

On Error GoTo Err

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Véhicule Detail BC Reparation"
    .Show vbModal
End With
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

'ajouter les détails de la réparation dans le grid ds FrmBCReparation
Private Sub CmdAjout_Click()   'Ajouter

Dim itmX As ListItem
Dim LOBJ_Veh As VEHICULE
Dim rs As New Recordset
On Error GoTo Err
'Control meme reparation

If Left(CheckMandatory(frmDetailBCReparation), 1) = 1 Then
   Exit Sub
End If

If txt_Numero.Text = "" Or txt_Numero.Text = "Auto" Then
    If Len(Trim(txt_Numero.Text)) = 0 Then
        MsgBox "Matricule  obligatoire      ", vbInformation
        txt_Numero.SetFocus
        Exit Sub
    End If
End If

If txt_Qte.Text = "" Or txt_Qte.Text = "0" Then
    MsgBox "  Quantité non valide      ", vbInformation
    txt_Qte.SetFocus
    Exit Sub
End If

If Cbo_Matricule.Text = "" Then
    MsgBox "  Vehicule obligatoire      ", vbInformation
    txt_Qte.SetFocus
    Exit Sub
End If

If Okay = True Then
    With FrmBCReparation
        Set itmX = .grid.ListItems.Add(, , (.txt_Numero.Text))
        itmX.SubItems(1) = CStr(txt_libelle.Text)
        itmX.SubItems(2) = CStr(txt_Qte.Text)
        itmX.SubItems(3) = CStr(Cbo_Matricule.Text)
        itmX.SubItems(4) = CStr(txt_Observation.Text)
    End With
Else
    With FrmBCReparation
        .grid.ListItems(.grid.SelectedItem.Index).SubItems(1) = CStr(txt_libelle.Text)
        .grid.ListItems(.grid.SelectedItem.Index).SubItems(2) = CStr(txt_Qte.Text)
        .grid.ListItems(.grid.SelectedItem.Index).SubItems(3) = CStr(Cbo_Matricule.Text)
        .grid.ListItems(.grid.SelectedItem.Index).SubItems(4) = CStr(txt_Observation.Text)
        
    End With
    
End If
Unload Me
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

'liste désignation des pièce de réparation
Private Sub CmdFindDesi_Click()

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Detail Reparation"
    .Show vbModal
End With
End Sub

Public Sub AfficheRow_Vehicule(ByVal VCode As String)

Dim Lobj_Vehicule As VEHICULE
Dim rs As New Recordset

Set Lobj_Vehicule = New VEHICULE
Set rs = Lobj_Vehicule.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Matricule")) Then
    Cbo_Matricule.Text = rs("Matricule")
    End If
Else
    MsgBox "Code véhicule introuvable", vbInformation
    Cbo_Matricule.SetFocus
    Exit Sub
End If
rs.Close

End Sub

'Recherche des produits et lubrifiants par initial
Private Sub SearchByInitial(ByVal Initial As String)

Dim LOBJ_Prod As Produit_Lubrifiant
Dim rs As Recordset

Set LOBJ_Prod = New Produit_Lubrifiant
Set rs = LOBJ_Prod.Get_ProdLubByInit(ErrNumber, ErrDescription, ErrSourceDetail, Initial, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LOBJ_Prod = Nothing
If Not rs.EOF Then
    If rs.RecordCount = 1 Then
        txt_libelle.Text = rs("Libelle")
    ElseIf rs.RecordCount > 1 Then
        Unload FrmFind_Fils
        With FrmFind_Fils
            .StrSource = "searchBCRepar"
            .Show vbModal
        End With
    End If
Else
    MsgBox "Aucun résultat !!"
End If
rs.Close
End Sub


Private Sub Txt_Libelle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
If KeyCode = vbKeyRight Then
     Call SearchByInitial(txt_libelle.Text)
End If
End Sub

Private Sub txt_Numero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Txt_Observation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Qte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Qte_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub
