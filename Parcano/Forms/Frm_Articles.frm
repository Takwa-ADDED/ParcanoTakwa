VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form Frm_Articles 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Parcano"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Articles.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   13320
   Begin VB.TextBox txt_libelle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   17
      Tag             =   "M"
      Top             =   3120
      Width           =   3855
   End
   Begin VB.OptionButton Op_Lub 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lubrifiant"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   6000
      Width           =   1935
   End
   Begin VB.OptionButton Op_Prod 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Produit"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CheckBox chk_Actif 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   6840
      Width           =   255
   End
   Begin VB.TextBox txt_tva 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "M"
      Top             =   4680
      Width           =   3855
   End
   Begin VB.TextBox txt_tht 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "M"
      Top             =   3960
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2760
      ScaleHeight     =   615
      ScaleWidth      =   4215
      TabIndex        =   9
      Top             =   5280
      Width           =   4215
      Begin VB.TextBox txt_prix 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "M"
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.TextBox txt_Matricule 
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
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "M"
      Top             =   2400
      Width           =   3855
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   11040
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
      Picture         =   "Frm_Articles.frx":0ECA
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   11520
      TabIndex        =   19
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
      Picture         =   "Frm_Articles.frx":104C
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   12000
      TabIndex        =   20
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
      Picture         =   "Frm_Articles.frx":139F
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   12480
      TabIndex        =   21
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
      Picture         =   "Frm_Articles.frx":16F2
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   6960
      TabIndex        =   22
      Top             =   2400
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
      Picture         =   "Frm_Articles.frx":1874
      ButtonType      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produits"
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
      Left            =   0
      TabIndex        =   16
      Top             =   240
      Width           =   2535
   End
   Begin VB.Image PicBox_Header 
      Height          =   1095
      Left            =   0
      Picture         =   "Frm_Articles.frx":1BC7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
   Begin VB.Image Cmd_ReAjouter 
      Height          =   375
      Left            =   11280
      Picture         =   "Frm_Articles.frx":9B01D
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Lbl_RaAjouter 
      BackStyle       =   0  'Transparent
      Caption         =   "Article supprimé, Voulez le ré-ajouter?..."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7320
      TabIndex        =   15
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label Lbl_typ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Actif : O/N"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prix ht:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   600
      TabIndex        =   12
      Top             =   3960
      Width           =   1320
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taux tva:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   585
      TabIndex        =   11
      Top             =   4680
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prix ttc:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   585
      TabIndex        =   10
      Top             =   5400
      Width           =   1485
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Libelle :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   570
      TabIndex        =   8
      Top             =   3240
      Width           =   1485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Code :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "Frm_Articles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_ReAjouter_Click()
    Dim LOBJ_Prod As New Produit_Lubrifiant

On Error GoTo Err

    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        Else
            txt_Matricule.SetFocus
            Exit Sub
        End If
    End If
    If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
      
        If (CHECK_ACCES("Supp_Produit", LInt_UserId) = False) Then
            MsgBox "Suppression n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Supprime un produit", vbExclamation, App.ProductName
            Exit Sub
        End If
    End If
    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "produit", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
    Call LOBJ_Prod.Delete_Add_ProdLub(ErrNumber, ErrDescription, ErrSourceDetail, txt_Matricule.Text, "N", CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LOBJ_Prod = Nothing
    MsgBox "Article ré-ajouter avec succes!...", vbInformation
    AfficheRow (txt_Matricule.Text)
    EndDisb (True)
    Lbl_RaAjouter.Visible = False
    Cmd_ReAjouter.Visible = False
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

'Saisi nouveaux Produits et lubrifiants

Private Sub CmdAdd_Click()
On Error GoTo Err
    If (CHECK_ACCES("Ins_Produit", LInt_UserId) = False) Then
        MsgBox "Insertion n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Ajouter un produit", vbExclamation, App.ProductName
        Exit Sub
    End If
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
        If MsgBox("Annuler le MAJ en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    EndDisb (True)

    Lbl_RaAjouter.Visible = False
    Cmd_ReAjouter.Visible = False
    Call ViderZone(Frm_Articles)
    txt_Matricule.Text = "Auto"
    txt_libelle.SetFocus
    Op_Prod.Value = False
    Op_Lub.Value = False
    chk_Actif.Value = vbUnchecked
    CmdDelete.Enabled = False
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdAdd1_Click()
On Error GoTo Err
    If (CHECK_ACCES("Ins_Produit", LInt_UserId) = False) Then
        MsgBox "Insertion n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Ajouter un produit", vbExclamation, App.ProductName
        Exit Sub
    End If
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
        If MsgBox("Annuler le MAJ en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    EndDisb (True)

    Lbl_RaAjouter.Visible = False
    Cmd_ReAjouter.Visible = False
    Call ViderZone(Frm_Articles)
    txt_Matricule.Text = "Auto"
    txt_libelle.SetFocus
    Op_Prod.Value = False
    Op_Lub.Value = False
    chk_Actif.Value = vbUnchecked
    CmdDelete.Enabled = False
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

'=========================
'Suppression***
'=========================
Private Sub CmdDelete_Click()
   Dim LOBJ_Prod As Produit_Lubrifiant
    Dim LOBJ_Veh As VEHICULE
    Dim rs As New Recordset

On Error GoTo Err

Set LOBJ_Prod = New Produit_Lubrifiant
Set LOBJ_Veh = New VEHICULE
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        Else
            txt_Matricule.SetFocus
            Exit Sub
        End If
    End If
    If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
        If (CHECK_ACCES("Supp_Produit", LInt_UserId) = False) Then
            MsgBox "Suppression n'est pas possible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Supprime un produit", vbExclamation, App.ProductName
            Exit Sub
        End If
    End If
'Vérifier si ce lubrifiant est associé à un véhicule dans la table de vidange "Vehicule_vidange"
Set rs = LOBJ_Veh.Get_VdgVehByLub(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt_Matricule.Text)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
  If Not rs.EOF Then
    MsgBox "Suppression de ce lubrifiant impossible , " & vbNewLine & "il est associé au vidange d'un véhicule ", vbInformation
    Exit Sub
  End If
    
    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "produit", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
    
    Call LOBJ_Prod.Delete_Add_ProdLub(ErrNumber, ErrDescription, ErrSourceDetail, txt_Matricule.Text, "O", CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LOBJ_Prod = Nothing
    MsgBox "Article supprimer avec succes!...", vbInformation
    EndDisb (True)
    Call ViderZone(Frm_Articles)
    Lbl_RaAjouter.Visible = False
    Cmd_ReAjouter.Visible = False
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    txt_Matricule.SetFocus

Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdDelete1_Click()
   Dim LOBJ_Prod As Produit_Lubrifiant
    Dim LOBJ_Veh As VEHICULE
    Dim rs As New Recordset

On Error GoTo Err

Set LOBJ_Prod = New Produit_Lubrifiant
Set LOBJ_Veh = New VEHICULE
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        Else
            txt_Matricule.SetFocus
            Exit Sub
        End If
    End If
    If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
        If (CHECK_ACCES("Supp_Produit", LInt_UserId) = False) Then
            MsgBox "Suppression n'est pas possible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Supprime un produit", vbExclamation, App.ProductName
            Exit Sub
        End If
    End If
'Vérifier si ce lubrifiant est associé à un véhicule dans la table de vidange "Vehicule_vidange"
Set rs = LOBJ_Veh.Get_VdgVehByLub(ErrNumber, ErrDescription, ErrSourceDetail, CNB, txt_Matricule.Text)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
  If Not rs.EOF Then
    MsgBox "Suppression de ce lubrifiant impossible , " & vbNewLine & "il est associé au vidange d'un véhicule ", vbInformation
    Exit Sub
  End If
    
    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "produit", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
    
    Call LOBJ_Prod.Delete_Add_ProdLub(ErrNumber, ErrDescription, ErrSourceDetail, txt_Matricule.Text, "O", CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LOBJ_Prod = Nothing
    MsgBox "Article supprimer avec succes!...", vbInformation
    EndDisb (True)
    Call ViderZone(Frm_Articles)
    Lbl_RaAjouter.Visible = False
    Cmd_ReAjouter.Visible = False
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
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

Private Sub CmdFind1_Click()
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

Call CmdFind_Click

End Sub

Private Sub cmdFindMatricule1_Click()
Call CmdFind_Click

End Sub

Private Sub CmdSave_Click()
Dim rs As New Recordset

If Left(CheckMandatory(Frm_Articles), 1) = 1 Then
   Exit Sub
End If
    
On Error GoTo Err
If txt_tht.Text = 0 Then
    MsgBox "Prix invalide !!", vbInformation
    Exit Sub
End If

If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub

If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
    If (CHECK_ACCES("Maj_produit", LInt_UserId) = False) Then
        MsgBox "Modification n'est pas possible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier un produit", vbExclamation, App.ProductName
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
    .Fields("tva") = CDbl(Txt_tva.Text)
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
    .Fields("tva") = CDbl(Txt_tva.Text)
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

Private Sub CmdSave1_Click()
Dim rs As New Recordset

If Left(CheckMandatory(Frm_Articles), 1) = 1 Then
   Exit Sub
End If
    
On Error GoTo Err
If txt_tht.Text = 0 Then
    MsgBox "Prix invalide !!", vbInformation
    Exit Sub
End If

If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub

If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
    If (CHECK_ACCES("Maj_produit", LInt_UserId) = False) Then
        MsgBox "Modification n'est pas possible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier un produit", vbExclamation, App.ProductName
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

Private Sub Form_Load()
Me.WindowState = 2
Lbl_RaAjouter.Visible = False
Cmd_ReAjouter.Visible = False
CmdDelete.Enabled = False
CmdSave.Enabled = False
End Sub

Private Sub Form_Resize()
'Dim WidthForm As Integer
'On Error Resume Next
'PicBox_Header.Width = Me.Width
On Error Resume Next
    Dim WidthForm As Integer
        WidthForm = Frm_Main.ACB_Main.Width
        PicBox_Header.Width = WidthForm - 1000
        CmdAdd.Left = WidthForm - 5500
        CmdDelete.Left = WidthForm - 5100
        CmdFind.Left = WidthForm - 4700
        CmdSave.Left = WidthForm - 4300

End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo erreur
'   Dim i As Integer
'   Dim Msg ' Déclare la variable.
'   ' Définit le texte du message.
'   Msg = "Voulez-vous vraiment quitter?"
'   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
'   If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'      Cancel = True
'   Else
'   Unload Me
'
'   End If
'
'   Exit Sub
'erreur:
'   MsgBox Err.Description, 48
End Sub

Private Sub txt_Libelle_GotFocus()

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
        Call AfficheRow(rs("Numero"))
    ElseIf rs.RecordCount > 1 Then
        Unload FrmFind_Fils
        With FrmFind_Fils
            .StrSource = "searchArticle"
            .Show vbModal
        End With
    End If
Else
    MsgBox "Aucune résultat !!"
End If
rs.Close
End Sub

Private Sub txt_Matricule_GotFocus()
Call ViderZone(Frm_Articles)
End Sub

Private Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
Dim LOBJ_Prod As Produit_Lubrifiant
Dim rs As New Recordset

If Len(txt_Matricule.Text) > 0 And IsNumeric(Trim(txt_Matricule.Text)) Then
    If KeyCode = vbKeyReturn Then
        Call SearchByInitial(txt_Matricule.Text)
    End If
Else
    If KeyCode = vbKeyRight Then
       cmdFindMatricule.SetFocus
       Call SearchByInitial(txt_Matricule.Text)
    End If
End If
End Sub

Private Sub txt_Matricule_LostFocus()

On Error GoTo Err

'If Len(Trim(txt_Matricule.Text)) > 0 Then Call SearchByInitial(txt_Matricule.Text)
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

Public Sub AfficheRow(ByVal VCode As String)

Dim LOBJ_Prod As Produit_Lubrifiant
Dim rs As New Recordset

Call ViderZone(Frm_Articles)

Set LOBJ_Prod = New Produit_Lubrifiant
Set rs = LOBJ_Prod.Get_ProdLubBycode(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
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
    Txt_tva.Text = Format(rs("tva"), "##0.00")
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
    
    If (rs("supp") = "O") Then
        Call EndDisb(False)
        Cmd_ReAjouter.Visible = True
        Lbl_RaAjouter.Visible = True
    ElseIf (rs("supp") = "N") Then
        EndDisb (True)
        Cmd_ReAjouter.Visible = False
        Lbl_RaAjouter.Visible = False
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
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Or InStr(txt_tht.Text, ",") <> 0 And Chr(KeyAscii) = "," Then
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
    If Len(Trim(Txt_tva.Text)) = 0 Then
        tva = 0
    Else
        tva = Txt_tva.Text
    End If
    If Len(Trim(txt_tht.Text)) = 0 Or txt_tht = "," Then
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
If Not (Chr(KeyAscii) Like "[0123456789.,]") And KeyAscii <> 13 And KeyAscii <> 8 Or InStr(Txt_tva.Text, ",") <> 0 And Chr(KeyAscii) = "," Then
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
    If Len(Trim(Txt_tva.Text)) = 0 Or Txt_tva.Text = "," Then
        tva = 0
        Txt_tva.Text = 0
    Else
        tva = Txt_tva.Text
    End If
    If Len(Trim(txt_tht.Text)) = 0 Or txt_tht = "," Then
        tht = 0
        txt_tht.Text = 0
    Else
        tht = txt_tht.Text
    End If
    ttc = tht + (tht * (tva / 100))
    Txt_tva.Text = Format(Txt_tva.Text, "##0.00")
    txt_prix.Text = Format(ttc, "##0.000")
Exit Sub
Err:
MsgBox Err.Description, vbInformation


End Sub


Private Sub EndDisb(ByVal TYP As Boolean)
    txt_Matricule.Enabled = TYP
    txt_libelle.Enabled = TYP
    txt_tht.Enabled = TYP
    Txt_tva.Enabled = TYP
    Op_Prod.Enabled = TYP
    Op_Lub.Enabled = TYP
    chk_Actif.Enabled = TYP
    CmdSave.Enabled = TYP
    CmdDelete.Enabled = TYP
End Sub
