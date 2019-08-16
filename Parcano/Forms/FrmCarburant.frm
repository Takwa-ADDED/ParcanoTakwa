VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmCarburant 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Energie"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
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
   Icon            =   "FrmCarburant.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   10845
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   3960
      ScaleHeight     =   2625
      ScaleWidth      =   3705
      TabIndex        =   20
      Top             =   3240
      Width           =   3735
      Begin SToolBox.SGrid SGrid1 
         Height          =   3015
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5318
         RowMode         =   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisableIcons    =   -1  'True
         MaxVisibleRows  =   0
      End
   End
   Begin VB.PictureBox Pic_Lib 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   3975
      TabIndex        =   19
      Top             =   3240
      Width           =   3975
      Begin VB.TextBox txt_libelle 
         Appearance      =   0  'Flat
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
         Left            =   0
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "M"
         Top             =   0
         Width           =   3735
      End
   End
   Begin VB.PictureBox Pict_TTC 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   2400
      ScaleHeight     =   735
      ScaleWidth      =   3855
      TabIndex        =   16
      Top             =   5400
      Width           =   3855
      Begin VB.TextBox txt_prix 
         Appearance      =   0  'Flat
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
         Left            =   0
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "M"
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.TextBox txt_tht 
      Appearance      =   0  'Flat
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "M"
      Top             =   3960
      Width           =   3735
   End
   Begin VB.TextBox txt_tva 
      Appearance      =   0  'Flat
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "M"
      Top             =   4680
      Width           =   3735
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   6240
      TabIndex        =   9
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
      Picture         =   "FrmCarburant.frx":0ECA
      ButtonType      =   1
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "M"
      Top             =   2400
      Width           =   3735
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   9000
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
      Picture         =   "FrmCarburant.frx":121D
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   8040
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
      Picture         =   "FrmCarburant.frx":139F
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   8520
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
      Picture         =   "FrmCarburant.frx":16F2
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   7560
      TabIndex        =   7
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
      Picture         =   "FrmCarburant.frx":1A45
   End
   Begin SToolBox.SCommand Cmd_ReAjouter 
      Height          =   375
      Left            =   6720
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Oui!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483635
      ForeColor       =   14737632
   End
   Begin VB.Label Lbl_RaAjouter 
      BackStyle       =   0  'Transparent
      Caption         =   "Carburant supprimé, Voulez le  ré-ajouter?..."
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
      Left            =   2400
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche energie"
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
      Left            =   720
      TabIndex        =   14
      Top             =   480
      Width           =   2190
   End
   Begin VB.Image PicBox_Header 
      Height          =   1335
      Left            =   0
      Picture         =   "FrmCarburant.frx":1BC7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
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
      Left            =   600
      TabIndex        =   13
      Top             =   5520
      Width           =   1485
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
      Left            =   600
      TabIndex        =   12
      Top             =   4800
      Width           =   1485
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
      TabIndex        =   11
      Top             =   4080
      Width           =   1320
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
      Left            =   600
      TabIndex        =   10
      Top             =   3360
      Width           =   1485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
End
Attribute VB_Name = "FrmCarburant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Réajouter energie déjà supprimer

Private Sub Cmd_ReAjouter_Click()
    Dim LOBJ_Energ As Energie

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


    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "produit (Carburant)", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
    
    Set LOBJ_Energ = New Energie
    Call LOBJ_Energ.Delete_Add_Energie(ErrNumber, ErrDescription, ErrSourceDetail, txt_Matricule.Text, "N", CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LOBJ_Energ = Nothing
    MsgBox "Energie ré-ajouter avec succes!...", vbInformation
    AfficheRow (txt_Matricule.Text)
    EndDisb (True)
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    Lbl_RaAjouter.Visible = False
    Cmd_ReAjouter.Visible = False
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdAdd_Click()
On Error GoTo Err
    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
        If (CHECK_ACCES("Ins_TC", LInt_UserId) = False) Then
            MsgBox "Insertion n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Ajouter un Type de Carburant", vbExclamation, App.ProductName
            Exit Sub
        End If
    Pic_Lib.Enabled = True
    Call ViderZone(FrmCarburant)
    Lbl_RaAjouter.Visible = False
    Cmd_ReAjouter.Visible = False
    Call EndDisb(True)
    CmdDelete.Enabled = False
    txt_Matricule.Text = "Auto"
    txt_libelle.SetFocus
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub CmdDelete_Click()

Dim LOBJ_Personnel As personnel
Dim LOBJ_Energ As Energie

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
        If (CHECK_ACCES("Supp_TC", LInt_UserId) = False) Then
            MsgBox "Suppression n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Supprime un Type de Carburant", vbExclamation, App.ProductName
            Exit Sub
        End If
    End If
    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "energie carburant", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
        
        Set LOBJ_Energ = New Energie
        Call LOBJ_Energ.Delete_Add_Energie(ErrNumber, ErrDescription, ErrSourceDetail, txt_Matricule.Text, "O", CNB)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LOBJ_Energ = Nothing
    
        txt_Matricule.SetFocus
    End If
    
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdFind_Click()

On Error GoTo Err

    If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
        
    Unload FrmFind
    With FrmFind
        .StrSource = "Energie"
        .Show
    End With

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub cmdFindMatricule_Click()
Unload FrmFind
With FrmFind
    .StrSource = "Energie"
    .Show
End With
End Sub

Private Sub CmdSave_Click()
Dim rs As New Recordset

If Left(CheckMandatory(FrmCarburant), 1) = 1 Then
   Exit Sub
End If

On Error GoTo Err

If txt_tht.Text = 0 Then
    MsgBox "Prix invalide !!", vbInformation
    Exit Sub
End If
    
If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub

If txt_Matricule.Text <> "Auto" Then
    If (CHECK_ACCES("Maj_TC", LInt_UserId) = False) Then
        MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier un Type de Carburant", vbExclamation
        Exit Sub
    End If

    'Modification de l'energie
    Call Modif_Energie
End If

If txt_Matricule.Text = "Auto" Then
   Call Ajout_energie
End If

MsgBox "Enregistrement terminé avec succé  ", vbQuestion
txt_Matricule.Enabled = True
txt_Matricule.SetFocus

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub Ajout_energie()

Dim LInt_NumCompteur As Long
Dim LOBJ_Energie As Energie

LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_Energie")
If ErrNumber <> 0 Then
   MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
   ErrNumber = 0
   Exit Sub
End If

'Insertion enregistrement assiette
txt_Matricule.Text = Format(LInt_NumCompteur, "00000")
Set LOBJ_Energie = New Energie
Call LOBJ_Energie.Insert_Energie(ErrNumber, ErrDescription, ErrSourceDetail, CNB, remplisage_energie)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
End Sub

Private Sub Modif_Energie()

Dim LOBJ_Energie As Energie

Set LOBJ_Energie = New Energie
Call LOBJ_Energie.Update_Energie(ErrNumber, ErrDescription, ErrSourceDetail, CNB, remplisage_energie)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

End Sub

Private Function remplisage_energie() As Recordset

Dim LRs_NewRecord As New Recordset

Set LRs_NewRecord = CreateEmptyRS_Energie()
With LRs_NewRecord
    .AddNew
    .Fields("Code") = txt_Matricule.Text
    .Fields("Libelle") = txt_libelle.Text
    .Fields("tht") = CDbl(txt_tht.Text)
    .Fields("tva") = CDbl(Txt_tva.Text)
    .Fields("prix") = CDbl(txt_prix.Text)
    .Fields("UserInsert") = LInt_UserId
End With
Set remplisage_energie = LRs_NewRecord
Set LRs_NewRecord = Nothing

End Function

Private Sub Form_Load()
Picture1.Visible = False
Call InitialGrid
Me.Width = 10380
Me.Height = 7935
Me.Move 0, 0
Me.WindowState = 2
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
End Sub
Private Sub Form_Resize()
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

Public Sub InitialGrid()
 With SGrid1
         ' Allow the grid to be grouped, but
      ' don't show the grouping box
      .HideGroupingBox = True
      .AllowGrouping = True

      ' Group rows will be shown by
      ' a gradient underline
      .GroupRowBackColor = vbWindowBackground
      .GroupRowForeColor = vbWindowText
      
      'Couleur de Grid
      .GridLineColor = vbWindowBackground
      .GridFillLineColor = vbWindowBackground
      .GridLines = True
      
      'Ajout des Colonnes
      .SelectionAlphaBlend = True
      .SelectionOutline = True
      .DrawFocusRectangle = False
      
      .AddColumn "code", "Code", , , 60
      .AddColumn "Libelle", "Libelle", , , 120
'      .AddColumn "tht", "Prix _HT"
'      .AddColumn "tva", "Taux_Tva"
      .AddColumn "prix", "Prix", , , 60
      .AddColumn "Q", ""
      .StretchLastColumnToFit = True

   End With
End Sub

Private Sub SGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
Dim VCode
On Error GoTo Err

 VCode = SGrid1.CellText(lRow, 1)
 FrmCarburant.AfficheRow (VCode)
 cmdFindMatricule.SetFocus
 Picture1.Visible = False
 Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub SGrid1_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
Dim VCode
On Error GoTo Err
    VCode = SGrid1.CellText(SGrid1.SelectedRow, 1)
    Select Case KeyCode
        Case vbKeyF2, vbKeyReturn
                    FrmCarburant.AfficheRow (VCode)
                    cmdFindMatricule.SetFocus
                    Picture1.Visible = False
        Case vbKeyEscape
        Picture1.Visible = False
        End Select
 Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub SGrid1_LostFocus()
    Picture1.Visible = False
End Sub

Private Sub txt_Libelle_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If

End Sub

Private Sub Txt_Libelle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txt_tht.SetFocus
End Sub

Private Sub txt_Matricule_GotFocus()
Call ViderZone(FrmCarburant)
End Sub

Private Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If IsNumeric(Trim(txt_Matricule.Text)) Then
    If KeyCode = vbKeyReturn Then
        On Error GoTo Err
        If Len(Trim(txt_Matricule.Text)) > 0 Then Call AfficheRow(txt_Matricule.Text)
    End If
    Exit Sub
Else
    If KeyCode = vbKeyRight Then
        On Error GoTo Err
        If Len(Trim(txt_Matricule.Text)) > 0 Then Call AfficheRowbylib(txt_Matricule.Text)
    End If
    Exit Sub
End If
Err:
    MsgBox Err.Description, vbInformation
    End Sub

Private Sub txt_Matricule_LostFocus()

On Error GoTo Err

If Len(Trim(txt_Matricule.Text)) > 0 Then Call AfficheRow(txt_Matricule.Text)

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub
Public Sub AfficheRowbylib(ByVal Libelle As String)

Dim LOBJ_Energ As Energie
Dim rs As New Recordset

Call ViderZone(FrmCarburant)
SGrid1.ClearRows
Pic_Lib.Enabled = False
Set LOBJ_Energ = New Energie
Set rs = LOBJ_Energ.Get_EnergByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Libelle)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If rs.RecordCount = 1 Then
    'Charge
    txt_Matricule.Text = rs("Code")
    txt_libelle.Text = rs("Libelle")
    txt_tht.Text = Format(rs("tht"), "#,##0.000")
    Txt_tva.Text = Format(rs("tva"), "#,##0.00")
    txt_prix.Text = Format(rs("prix"), "#,##0.000")
    
    If (rs("supp") = "O") Then
        Call EndDisb(False)
        Cmd_ReAjouter.Visible = True
        Lbl_RaAjouter.Visible = True
    ElseIf (rs("supp") = "N") Then
        Call EndDisb(True)
        Cmd_ReAjouter.Visible = False
        Lbl_RaAjouter.Visible = False
    End If
Else

If Not rs.EOF Then
While Not rs.EOF
SGrid1.Redraw = False
    With SGrid1
        .AddRow
        .CellDetails .Rows, .ColumnIndex("code"), rs.Fields("code"), DT_RIGHT
        .CellDetails .Rows, .ColumnIndex("Libelle"), rs.Fields("libelle"), DT_RIGHT
        .CellDetails .Rows, .ColumnIndex("prix"), rs.Fields("prix"), DT_RIGHT
    End With
rs.MoveNext
Wend
SGrid1.Redraw = True
Picture1.Visible = True
If SGrid1.Rows > 0 Then
    SGrid1.SelectedRow = 1
    SGrid1.SetFocus
End If

Else
    MsgBox "Code introuvable", vbInformation
    txt_Matricule.SetFocus
End If
rs.Close
End If

End Sub


Public Sub AfficheRow(ByVal VCode As String)

Dim LOBJ_Energ As Energie
Dim rs As New Recordset

Call ViderZone(FrmCarburant)
Pic_Lib.Enabled = False
Set LOBJ_Energ = New Energie
Set rs = LOBJ_Energ.Get_EnergByLiborCod(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    txt_libelle.Text = rs("Libelle")
    txt_tht.Text = Format(rs("tht"), "#,##0.000")
    Txt_tva.Text = Format(rs("tva"), "#,##0.00")
    txt_prix.Text = Format(rs("prix"), "#,##0.000")
    
    If (rs("supp") = "O") Then
        Call EndDisb(False)
        Cmd_ReAjouter.Visible = True
        Lbl_RaAjouter.Visible = True
    ElseIf (rs("supp") = "N") Then
        Call EndDisb(True)
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
End If

End Sub

Private Sub txt_tht_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Txt_tva.SetFocus

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
If KeyCode = vbKeyReturn Then

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
    If Len(Trim(txt_tht.Text)) = 0 Then
        tht = 0
        txt_tht.Text = 0
    Else
        tht = txt_tht.Text
    End If
    ttc = tht + (tht * (tva / 100))
    txt_prix.Text = Format(ttc, "##0.000")
    Txt_tva.Text = Format(tva, "##0.00")
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
    End If

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
    If Len(Trim(Txt_tva.Text)) = 0 Then
        tva = 0
    Else
        tva = Txt_tva.Text
    End If
    If Len(Trim(txt_tht.Text)) = 0 Then
        tht = 0
        txt_tht.Text = 0
    Else
        tht = txt_tht.Text
    End If
    ttc = tht + (tht * (tva / 100))
    txt_prix.Text = Format(ttc, "##0.000")
    Txt_tva.Text = Format(tva, "##0.00")
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub EndDisb(ByVal TYP As Boolean)

    txt_Matricule.Enabled = TYP
    txt_libelle.Enabled = TYP
    txt_tht.Enabled = TYP
    Txt_tva.Enabled = TYP
    CmdSave.Enabled = TYP
    CmdDelete.Enabled = TYP
End Sub

