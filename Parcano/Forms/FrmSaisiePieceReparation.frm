VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "STOOLBOX.OCX"
Begin VB.Form FrmSaisiePieceReparation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pict_TTC 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   3015
      TabIndex        =   24
      Top             =   6960
      Width           =   3015
      Begin VB.TextBox txt_ttc 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.PictureBox Pict_THT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   3015
      TabIndex        =   23
      Top             =   5760
      Width           =   3015
      Begin VB.TextBox txt_TotHT 
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.TextBox Txt_Remise 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2520
      TabIndex        =   5
      Text            =   "0"
      Top             =   5160
      Width           =   3015
   End
   Begin VB.TextBox txt_PUHT 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2520
      TabIndex        =   4
      Top             =   4560
      Width           =   3015
   End
   Begin VB.CommandButton Cmd_Annul 
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   10
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton Cmd_Ajout 
      Caption         =   "Ajouter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   9
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox txt_Numero 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Auto"
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txt_Qte 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2520
      TabIndex        =   3
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox Txt_tva 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2520
      TabIndex        =   7
      Text            =   "0"
      Top             =   6360
      Width           =   3015
   End
   Begin VB.TextBox Txt_Designation 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2880
      Width           =   3015
   End
   Begin VB.ComboBox cbo_Matricule 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   3480
      Width           =   3015
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   345
      Left            =   5760
      TabIndex        =   11
      Top             =   3480
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
      Picture         =   "FrmSaisiePieceReparation.frx":0000
      ButtonType      =   1
   End
   Begin SToolBox.SCommand SCmd_FindDesign 
      Height          =   345
      Left            =   5760
      TabIndex        =   12
      Top             =   2880
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
      Picture         =   "FrmSaisiePieceReparation.frx":0353
      ButtonType      =   1
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Saisie Pièce Réparation"
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
      TabIndex        =   22
      Top             =   360
      Width           =   4095
   End
   Begin VB.Image PicBox_Header 
      Height          =   1215
      Left            =   0
      Picture         =   "FrmSaisiePieceReparation.frx":06A6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total TTC :"
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
      Height          =   240
      Left            =   480
      TabIndex        =   21
      Top             =   7080
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remise (%):"
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
      Height          =   240
      Left            =   480
      TabIndex        =   20
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totale HT:"
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
      Height          =   240
      Left            =   480
      TabIndex        =   19
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PUHT:"
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
      Height          =   240
      Left            =   480
      TabIndex        =   18
      Top             =   4680
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantité: "
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
      Height          =   240
      Left            =   480
      TabIndex        =   17
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TVA (%):"
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
      Height          =   240
      Left            =   480
      TabIndex        =   16
      Top             =   6480
      Width           =   915
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
      Height          =   240
      Left            =   480
      TabIndex        =   15
      Top             =   3000
      Width           =   1335
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
      TabIndex        =   14
      Top             =   2160
      Width           =   1935
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
      Height          =   240
      Left            =   480
      TabIndex        =   13
      Top             =   3480
      Width           =   945
   End
End
Attribute VB_Name = "FrmSaisiePieceReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Okay As Boolean
Public ii As Integer
Dim thekey As Integer
Dim theshift As Integer
Public choix As Boolean

Private Sub Cbo_Matricule_LostFocus()
Call ExistDonneeCbo(cbo_Matricule)
End Sub

Private Sub cmdFindMatricule_Click()

On Error GoTo Err

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Véhicule Detail Piece Reparation"
    .Show vbModal
End With
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub Cmd_Ajout_Click()

Dim itmX As ListItem
Dim Calculer As Boolean
Dim LOBJ_ProdLub As Produit_Lubrifiant
Dim rs As New Recordset
'
On Error GoTo Err
'Control meme reparation
If Left(CheckMandatory(FrmSaisiePieceReparation), 1) = 1 Then
   Exit Sub
End If
If txt_Numero.Text = "" Or txt_Numero.Text = "Auto" Then
    If Len(Trim(txt_Numero.Text)) = 0 Then
        MsgBox "Matricule  obligatoire      ", vbInformation
        txt_Numero.SetFocus
        Exit Sub
    End If
End If
Call Calcul
Set LOBJ_ProdLub = New Produit_Lubrifiant
Set rs = LOBJ_ProdLub.Get_ProdLubByLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Txt_Designation.Text)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If rs.EOF Then
    Frm_Msgbox.Show vbModal
    If choix = False Then Exit Sub
End If

If Okay = True Then
    With FrmPieceReparation
        Set itmX = .Lsv_Detail.ListItems.Add(, , (.txt_Numero.Text))
        itmX.SubItems(1) = CStr(Txt_Designation.Text)
        itmX.SubItems(2) = CStr(txt_Qte.Text)
        itmX.SubItems(3) = CStr(cbo_Matricule.Text)
        itmX.SubItems(4) = CStr(txt_PUHT.Text)
        itmX.SubItems(5) = CStr(Txt_Remise.Text)
        itmX.SubItems(6) = CStr(txt_TotHT.Text)
        itmX.SubItems(7) = CStr(Txt_tva.Text)
        itmX.SubItems(8) = CStr(txt_ttc.Text)
    End With
Else
    With FrmPieceReparation
        .Lsv_Detail.ListItems(.Lsv_Detail.SelectedItem.Index).SubItems(1) = CStr(Txt_Designation.Text)
        .Lsv_Detail.ListItems(.Lsv_Detail.SelectedItem.Index).SubItems(2) = CStr(txt_Qte.Text)
        .Lsv_Detail.ListItems(.Lsv_Detail.SelectedItem.Index).SubItems(3) = CStr(cbo_Matricule.Text)
        .Lsv_Detail.ListItems(.Lsv_Detail.SelectedItem.Index).SubItems(4) = CStr(txt_PUHT.Text)
        .Lsv_Detail.ListItems(.Lsv_Detail.SelectedItem.Index).SubItems(5) = CStr(Txt_Remise.Text)
        .Lsv_Detail.ListItems(.Lsv_Detail.SelectedItem.Index).SubItems(6) = CStr(txt_TotHT.Text)
        .Lsv_Detail.ListItems(.Lsv_Detail.SelectedItem.Index).SubItems(7) = CStr(Txt_tva.Text)
        .Lsv_Detail.ListItems(.Lsv_Detail.SelectedItem.Index).SubItems(8) = CStr(txt_ttc.Text)
    End With
End If

Unload Me
FrmPieceReparation.Pict_TRP.Enabled = True
Calculer = True
For ii = 1 To FrmPieceReparation.Lsv_Detail.ListItems.Count
    If (FrmPieceReparation.Lsv_Detail.ListItems(ii).SubItems(7) = "") Then
        Calculer = False
        Exit For
    End If
Next
If Calculer = True Then
    FrmPieceReparation.AppCalcul
End If

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Public Sub Update_ProdLub(ByVal VCode As String, ByVal TYP As String)

Dim LOBJ_Produit As Produit_Lubrifiant
Dim LRs_NewRecord As New Recordset
Dim rs As New Recordset
Dim LInt_NumCompteur As Long

Set LOBJ_Produit = New Produit_Lubrifiant
Set rs = LOBJ_Produit.Get_ProdLubByLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'MAJ du produit (type de reparation) s'il existe
    Set LRs_NewRecord = CreateEmptyRS_Prod_Lub()
    With LRs_NewRecord
        .AddNew
        .Fields("Numero") = txt_Numero.Text
        .Fields("Libelle") = Txt_Designation.Text
        .Fields("tva") = CDbl(Txt_tva.Text)
        If IsNull(rs("DatePrix")) Then
            .Fields("DatePrix") = Format(Date)
        Else
            .Fields("DatePrix") = rs("DatePrix")
        End If
        .Fields("prixht") = CDbl(txt_PUHT.Text)
        .Fields("OperateurSaisi") = LStr_NameUser
    End With
    Call LOBJ_Produit.Update_ProdLub(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LRs_NewRecord = Nothing
Else
    Set LRs_NewRecord = CreateEmptyRS_Prod_Lub()
    With LRs_NewRecord
        .AddNew
        .Fields("Libelle") = Txt_Designation.Text
        .Fields("tva") = CDbl(Txt_tva.Text)
        .Fields("prixht") = CDbl(txt_PUHT.Text)
        .Fields("DatePrix") = Format(Date)
        .Fields("Type_PL") = TYP
        .Fields("Actif") = "O"
        .Fields("OperateurSaisi") = LStr_NameUser
    End With
    Call LOBJ_Produit.Insert_ProdLub(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LRs_NewRecord = Nothing
    LInt_NumCompteur = return_Compteur()
    txt_Numero.Text = Format(LInt_NumCompteur, "00000")
End If

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

Private Sub Cmd_Ajout_GotFocus()

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Numero.SetFocus
End If

If Len(Trim(Txt_Designation.Text)) = 0 Then
    MsgBox "Désignation obligatoire      ", vbInformation
    Txt_Designation.SetFocus
End If

If Len(Trim(cbo_Matricule.Text)) = 0 Then
    MsgBox "Matricule obligatoire      ", vbInformation
    cbo_Matricule.SetFocus
End If

If Len(Trim(txt_Qte.Text)) = 0 Then
    MsgBox "Quantité obligatoire      ", vbInformation
    txt_Qte.SetFocus
End If

If Len(Trim(txt_PUHT.Text)) = 0 Then
    MsgBox "PUHT obligatoire      ", vbInformation
    txt_PUHT.SetFocus
End If

End Sub

Private Sub Cmd_Annul_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Affiche_Matricule_Combo(cbo_Matricule)
End Sub

Private Sub SCmd_FindDesign_Click()

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Piece Reparation"
    .Show vbModal
End With
End Sub

Public Sub AfficheRow_Vehicule(ByVal VCode As String)

Dim LOBJ_Veh As VEHICULE
Dim rs As New Recordset

Set LOBJ_Veh = New VEHICULE
Set rs = LOBJ_Veh.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Matricule")) Then
    cbo_Matricule.Text = rs("Matricule")
    End If
Else
    MsgBox "Code introuvable", vbInformation
    cbo_Matricule.SetFocus
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
        txt_Numero.Text = rs("Numero")
        Txt_Designation.Text = rs("Libelle")
        txt_PUHT.Text = rs("prixht")
        Txt_tva.Text = rs("tva")
    ElseIf rs.RecordCount > 1 Then
        Unload FrmFind_Fils
        With FrmFind_Fils
            .StrSource = "searchPieceRepar"
            .Show vbModal
        End With
    End If
Else
    MsgBox "Aucune résultat !!"
End If
rs.Close
End Sub

Private Sub Calcul()

Dim TotHtNet As Double
Dim TotTTC As Double
Dim Qte As Double
Dim PUHT As Double
Dim Remise As Double
Dim tva As Double

On Error GoTo Err
TotTTC = 0
TotHtNet = 0
Qte = 0
PUHT = 0
Remise = 0
tva = 0

If Len(Trim(txt_ttc.Text)) = 0 Then
    TotTTC = 0
Else
    TotTTC = txt_ttc.Text
End If

If Len(Trim(txt_TotHT.Text)) = 0 Then
    TotHtNet = 0
Else
    TotHtNet = txt_TotHT.Text
End If

If Len(Trim(txt_Qte.Text)) = 0 Then
    Qte = 0
Else
    Qte = txt_Qte.Text
End If

If Len(Trim(txt_PUHT.Text)) = 0 Then
    PUHT = 0
Else
    PUHT = txt_PUHT.Text
End If

If Len(Trim(Txt_Remise.Text)) = 0 Then
    Remise = 0
Else
    Remise = Txt_Remise.Text
End If

If Len(Trim(Txt_tva.Text)) = 0 Then
    tva = 0
Else
    tva = Txt_tva.Text
End If

Txt_tva.Text = Format(Txt_tva.Text, "#0.00")
If Txt_Remise.Text = "" Then Txt_Remise.Text = "0"
Txt_Remise.Text = Format(Txt_Remise.Text, "##0.00")

TotHtNet = Return_TotHT(Qte, PUHT, Remise)
txt_TotHT.Text = Format(TotHtNet, "##0.000")

TotTTC = TotHtNet + (TotHtNet * (tva / 100))
txt_ttc.Text = Format(TotTTC, "##0.000")
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub txt_Designation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
If KeyCode = vbKeyRight Then
     Call SearchByInitial(Txt_Designation.Text)
End If
End Sub

Private Sub Txt_Designation_LostFocus()
txt_PUHT.Text = Format(txt_PUHT.Text, "##0.000")
End Sub

Private Sub txt_Numero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_PUHT_GotFocus()
txt_PUHT.SelStart = 0
txt_PUHT.SelLength = Len(txt_PUHT.Text)
End Sub

Private Sub txt_PUHT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
    Call Calcul
End If
End Sub

Private Sub txt_PUHT_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Or InStr(txt_PUHT.Text, ",") <> 0 And Chr(KeyAscii) = "," Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_PUHT_LostFocus()

Call Calcul
txt_PUHT.Text = Format(txt_PUHT.Text, "##0.000")
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub txt_Qte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
    Call Calcul
End If
End Sub

Private Sub txt_Qte_KeyPress(KeyAscii As Integer)
If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_Qte_LostFocus()

Call Calcul
    
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Txt_Remise_GotFocus()
Txt_Remise.SelStart = 0
Txt_Remise.SelLength = Len(Txt_Remise.Text)
End Sub

Private Sub Txt_Remise_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Txt_Remise.Text = Format(Txt_Remise, "##0.00")
    SendKeys "{tab}"
    Call Calcul
End If

End Sub

Private Sub Txt_Remise_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Or InStr(Txt_Remise.Text, ",") <> 0 And Chr(KeyAscii) = "," Then
    KeyAscii = 0
End If
End Sub

Private Sub Txt_Remise_LostFocus()

Call Calcul
Txt_Remise.Text = Format(Txt_Remise, "##0.00")
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub txt_TotHT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_ttc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Txt_tva_GotFocus()
Txt_tva.SelStart = 0
Txt_tva.SelLength = Len(Txt_tva.Text)
End Sub

Private Sub Txt_tva_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Public Function Return_TotHT(ByVal Qte As Integer, _
                                ByVal PUHT As Double, _
                                ByVal Remise As Double) As Double
Dim TotHTBrut As Double

On Error GoTo Err
TotHTBrut = Qte * PUHT

Return_TotHT = TotHTBrut - (TotHTBrut * Remise / 100)

Exit Function
Err:
    MsgBox Err.Description, vbInformation

End Function

Private Sub Txt_tva_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Chr(KeyAscii) = "." Then KeyAscii = Asc(",")
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Or InStr(Txt_tva.Text, ",") <> 0 And Chr(KeyAscii) = "," Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_tva_LostFocus()

Call Calcul
Txt_tva.Text = Format(Txt_tva.Text, "#0.00")
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub cbo_Matricule_Change()

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
    start = Len(cbo_Matricule.Text)
    For i = 0 To cbo_Matricule.ListCount - 1
        If Left(cbo_Matricule.List(i), start) = cbo_Matricule.Text Then
            cbo_Matricule.Text = cbo_Matricule.List(i)
        End If
    Next
    cbo_Matricule.SelStart = start
    cbo_Matricule.SelLength = Len(cbo_Matricule.Text)
End If
End Sub

Private Sub cbo_Matricule_GotFocus()

If Len(Trim(txt_Numero.Text)) = 0 Then
    MsgBox "N° bon obligatoire      ", vbInformation
    txt_Numero.SetFocus
End If
End Sub

Private Sub Cbo_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cbo_Matricule_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
