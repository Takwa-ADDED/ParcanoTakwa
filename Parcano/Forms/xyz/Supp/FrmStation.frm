VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmStation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Parcano"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
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
   Icon            =   "FrmStation.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   10260
   Begin VB.Timer Timer1 
      Left            =   7320
      Top             =   1800
   End
   Begin VB.CheckBox chk_Actif 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   7320
      Width           =   255
   End
   Begin VB.ComboBox Cbo_type 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FrmStation.frx":0ECA
      Left            =   1800
      List            =   "FrmStation.frx":0ECC
      TabIndex        =   2
      Tag             =   "M"
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox txt_codepostal 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   4440
      Width           =   3855
   End
   Begin VB.TextBox txt_ville 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox txt_adresse 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "M"
      Top             =   3480
      Width           =   3855
   End
   Begin VB.TextBox txt_mobile 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   8
      Top             =   5880
      Width           =   3855
   End
   Begin VB.TextBox txt_activite 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   6
      Top             =   4920
      Width           =   3855
   End
   Begin VB.TextBox txt_email 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   10
      Top             =   6840
      Width           =   3855
   End
   Begin VB.TextBox txt_fax 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   9
      Top             =   6360
      Width           =   3855
   End
   Begin VB.TextBox txt_telephone 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   7
      Top             =   5400
      Width           =   3855
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   4080
      TabIndex        =   16
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
      Picture         =   "FrmStation.frx":0ECE
      ButtonType      =   1
   End
   Begin VB.TextBox txt_rsocial 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "M"
      Top             =   2520
      Width           =   3855
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
      Left            =   1800
      TabIndex        =   0
      Tag             =   "M"
      Top             =   1800
      Width           =   2295
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   9720
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
      Picture         =   "FrmStation.frx":1221
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   8760
      TabIndex        =   13
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
      Picture         =   "FrmStation.frx":13A3
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   9240
      TabIndex        =   15
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
      Picture         =   "FrmStation.frx":16F6
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   8280
      TabIndex        =   14
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
      Picture         =   "FrmStation.frx":1A49
   End
   Begin SToolBox.SCommand Cmd_ReAjouter 
      Height          =   255
      Left            =   6600
      TabIndex        =   30
      Top             =   1440
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
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
   Begin VB.Label Lbl_Supp 
      BackStyle       =   0  'Transparent
      Caption         =   "Station supprimée, Voulez-vous ré-ajouter?..."
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
      Left            =   2280
      TabIndex        =   31
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche Fournisseur"
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
      TabIndex        =   29
      Top             =   360
      Width           =   2880
   End
   Begin VB.Image PicBox_Header 
      Height          =   1215
      Left            =   -360
      Picture         =   "FrmStation.frx":1BCB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Actif :O/N "
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
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
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
      Left            =   240
      TabIndex        =   27
      Top             =   3000
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code postal :"
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
      Left            =   240
      TabIndex        =   26
      Top             =   4440
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ville :"
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
      Left            =   240
      TabIndex        =   25
      Top             =   3960
      Width           =   525
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adresse :"
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
      Left            =   240
      TabIndex        =   24
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile(GSM) :"
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
      Left            =   240
      TabIndex        =   23
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activité :"
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
      Left            =   240
      TabIndex        =   22
      Top             =   4920
      Width           =   900
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail :"
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
      Left            =   240
      TabIndex        =   21
      Top             =   6840
      Width           =   705
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tel bureau :"
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
      Left            =   240
      TabIndex        =   20
      Top             =   5400
      Width           =   1155
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax :"
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
      Left            =   240
      TabIndex        =   19
      Top             =   6360
      Width           =   450
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Raison sociale :"
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
      Left            =   240
      TabIndex        =   18
      Top             =   2520
      Width           =   1500
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
      TabIndex        =   17
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "FrmStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Fiche fournisseur

Private Sub CmdAdd_Click()

Dim LOBJ_Personnel As Personnel

On Error GoTo Err
If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cours.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If
Set LOBJ_Personnel = New Personnel
If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Ins_Fournisseur", LInt_UserId) Then
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If

EndDisb (True)
Call ViderZone(FrmStation)
Timer1.Enabled = False
Lbl_Supp.Visible = False
Cmd_ReAjouter.Visible = False
txt_Matricule.Text = "Auto"
txt_rsocial.SetFocus
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()

Dim LOBJ_Personnel As Personnel
Dim LOBJ_Stat As Station

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
    If MsgBox("Confirmez vous la suppression de cette " & vbNewLine & "station", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
        Set LOBJ_Personnel = New Personnel
        If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Supp_Fournisseur", LInt_UserId) Then
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
        Set LOBJ_Personnel = Nothing

        Set LOBJ_Stat = New Station
        Call LOBJ_Stat.Delete_Stat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "O", LInt_UserId, txt_Matricule.Text)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LOBJ_Stat = Nothing
        MsgBox "Station Supprimer avec succes!...", vbInformation, App.ProductName
        Call EndDisb(True)
        Call ViderZone(FrmStation)
        Lbl_Supp.Visible = False
        Cmd_ReAjouter.Visible = False
        CmdDelete.Enabled = False
        CmdSave.Enabled = False
        txt_Matricule.SetFocus
    End If
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Cmd_ReAjouter_Click()
Dim LOBJ_Personnel As Personnel
Dim LOBJ_Stat As Station
Dim vcode As String

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
    If MsgBox("Voulez-vous ré-ajouter cette " & vbNewLine & "station", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
        Set LOBJ_Personnel = New Personnel
        If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Supp_Fournisseur", LInt_UserId) Then
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
        Set LOBJ_Personnel = Nothing
        vcode = txt_Matricule.Text
        Set LOBJ_Stat = New Station
        Call LOBJ_Stat.Delete_Stat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "N", LInt_UserId, txt_Matricule.Text)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LOBJ_Stat = Nothing
        MsgBox "Station ré-ajouter avec succes!...", vbInformation, App.ProductName
        AfficheRow (vcode)
    End If
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
    .StrSource = "Station"
    .Show
End With
End Sub

Private Sub cmdFindMatricule_Click()
Unload FrmFind
With FrmFind_Actif
    .StrSource = "Station"
    .Show
End With
End Sub

Private Sub CmdSave_Click()

Dim LOBJ_Personnel As Personnel
Dim txt As String
Dim Okay As Boolean
Dim I

On Error GoTo Err

If Left(CheckMandatory(FrmStation), 1) = 1 Then
   Exit Sub
End If

If txt_Matricule.Text <> "Auto" And txt_Matricule.Text <> "" Then
    Set LOBJ_Personnel = New Personnel
    If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "Maj_Fournisseur", LInt_UserId) Then
        MsgBox "Accès refusé.", vbExclamation
        Exit Sub
    End If
End If

Okay = True
txt = Cbo_type.Text
For I = 0 To Cbo_type.ListCount - 1
    If txt = Cbo_type.List(I) Then
        Okay = False
        Exit For
    End If
Next
If Okay = True Then
    MsgBox "Veuillez sélectionner un type de la liste", vbInformation
    Cbo_type.SetFocus
    Exit Sub
End If

If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub

If txt_Matricule.Text = "Auto" Then
    Call Ajout_stat
Else
    Call modifier_Stat
End If

MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
txt_Matricule.SetFocus


Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub Ajout_stat()

Dim LRs_NewRecord As New Recordset
Dim LOBJ_Stat As Station
Dim LInt_NumCompteur

LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_Station")
If ErrNumber <> 0 Then
   MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
   ErrNumber = 0
   Exit Sub
End If
txt_Matricule.Text = Format(LInt_NumCompteur, "00000")

Set LRs_NewRecord = CreateEmptyRS_Station()
With LRs_NewRecord
    .AddNew
    .Fields("Code") = txt_Matricule.Text
    .Fields("Libelle") = txt_rsocial.Text
    .Fields("Type") = Cbo_type.Text
    .Fields("Adresse") = txt_adresse.Text
    .Fields("Ville") = txt_ville.Text
    .Fields("CPOSTAL") = Val(txt_codepostal.Text)
    .Fields("Activite") = txt_activite.Text
    .Fields("telephone") = txt_Telephone.Text
    .Fields("mobile") = txt_mobile.Text
    .Fields("fax") = txt_fax.Text
    .Fields("email") = txt_email.Text
    .Fields("Actif") = Val(chk_Actif.Value)
    .Fields("UserInsert") = LInt_UserId
End With
Set LOBJ_Stat = New Station
Call LOBJ_Stat.Insert_Stat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing
    
End Sub

Private Sub modifier_Stat()

Dim LRs_NewRecord As New Recordset
Dim LOBJ_Stat As Station

Set LRs_NewRecord = CreateEmptyRS_Station()
With LRs_NewRecord
    .AddNew
    .Fields("Code") = txt_Matricule.Text
    .Fields("Libelle") = txt_rsocial.Text
    .Fields("Type") = Cbo_type.Text
    .Fields("Adresse") = txt_adresse.Text
    .Fields("Ville") = txt_ville.Text
    .Fields("CPOSTAL") = Val(txt_codepostal.Text)
    .Fields("Activite") = txt_activite.Text
    .Fields("telephone") = txt_Telephone.Text
    .Fields("mobile") = txt_mobile.Text
    .Fields("fax") = txt_fax.Text
    .Fields("email") = txt_email.Text
    .Fields("Actif") = Val(chk_Actif.Value)
    .Fields("UserUpdate") = LInt_UserId
End With
Set LOBJ_Stat = New Station
Call LOBJ_Stat.Update_Stat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, LRs_NewRecord)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LRs_NewRecord = Nothing
End Sub

Public Sub AfficheRow(ByVal vcode As String)

Dim LOBJ_Stat As Station
Dim rs As New Recordset

Call ViderZone(FrmStation)
Set LOBJ_Stat = New Station
Set rs = LOBJ_Stat.GetStatByCodeLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, vcode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Type")) Then Cbo_type.Text = rs("Type")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
    If Not IsNull(rs("CPOSTAL")) Then txt_codepostal.Text = rs("CPOSTAL")
    If Not IsNull(rs("Activite")) Then txt_activite.Text = rs("Activite")
    If Not IsNull(rs("telephone")) Then txt_Telephone.Text = rs("telephone")
    If Not IsNull(rs("mobile")) Then txt_mobile.Text = rs("mobile")
    If Not IsNull(rs("fax")) Then txt_fax.Text = rs("fax")
    If Not IsNull(rs("email")) Then txt_email.Text = rs("email")
    If Not IsNull(rs("actif")) Then chk_Actif.Value = rs("Actif")
    If (rs("supp") = "O") Then
        Call EndDisb(False)
        Call Timer1_Timer
        Cmd_ReAjouter.Visible = True
    ElseIf (rs("supp") = "N") Then
        EndDisb (True)
        Timer1.Enabled = False
        Lbl_Supp.Visible = False
        Cmd_ReAjouter.Visible = False
    End If
Else
    MsgBox "Code introuvable", vbInformation
    txt_Matricule.SetFocus
End If
rs.Close
End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = True
Timer1.Interval = 600

If Lbl_Supp.Visible = True Then
    Lbl_Supp.Visible = False
Else
    Lbl_Supp.Visible = True
End If
End Sub

Private Sub EndDisb(ByVal TYP As Boolean)
    txt_Matricule.Enabled = TYP
    txt_rsocial.Enabled = TYP
    Cbo_type.Enabled = TYP
    txt_adresse.Enabled = TYP
    txt_ville.Enabled = TYP
    txt_codepostal.Enabled = TYP
    txt_activite.Enabled = TYP
    txt_Telephone.Enabled = TYP
    txt_mobile.Enabled = TYP
    txt_fax.Enabled = TYP
    txt_email.Enabled = TYP
    chk_Actif.Enabled = TYP
    CmdDelete.Enabled = TYP
    CmdSave.Enabled = TYP
End Sub

Private Sub Form_Load()
    Lbl_Supp.Visible = False
    Cmd_ReAjouter.Visible = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
Me.WindowState = 2

With Cbo_type
    .Clear
    .AddItem "Station carburant"
    .AddItem "Fournisseur Achat"
    .AddItem "Fournisseur"
    .AddItem "Autre"
End With
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

On Error GoTo erreur
   Dim I As Integer
   Dim MSG ' Déclare la variable.
   ' Définit le texte du message.
   MSG = "Voulez-vous vraiment quitter?"
   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
   If MsgBox(MSG, vbQuestion + vbYesNo + vbDefaultButton2, Label1.Caption) = vbNo Then
      Cancel = True
   Else
    Unload FrmStation
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

Private Sub txt_codepostal_KeyPress(KeyAscii As Integer)

If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

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

Private Sub txt_fax_KeyPress(KeyAscii As Integer)
If Not (Chr(KeyAscii) Like "[0123456789+]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_Matricule_GotFocus()
'Call ViderZone(FrmStation)
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

Private Sub txt_mobile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_mobile_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub txt_mobile_KeyPress(KeyAscii As Integer)

If Not (Chr(KeyAscii) Like "[0123456789+]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txt_rsocial_GotFocus()
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

Private Sub txt_Telephone_KeyPress(KeyAscii As Integer)

If Not (Chr(KeyAscii) Like "[0123456789+]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

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
