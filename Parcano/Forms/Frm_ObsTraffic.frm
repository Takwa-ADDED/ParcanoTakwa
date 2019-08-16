VERSION 5.00
Begin VB.Form Frm_ObsTraffic 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Confirmez vous l'enregistrement..."
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txt_Obs 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   7335
   End
   Begin VB.Timer Timer_Alerte 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Cmd_Cancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H0000C0C0&
      Caption         =   "Enregistre"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Lbl_DifDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Diff.Durée"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Lbl_DrMax 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Durée Max"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Compteur d'entrée :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Lbl_CmptEtr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Lbl_DistMax 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Lbl_Diff 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Lbl_CmptSort 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Diff.Dist :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Distance Max :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Compteur sortie :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "( Au moins 10 Lettres )"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Lbl_Alerte 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Error"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label Lbl_Obs 
      BackStyle       =   0  'Transparent
      Caption         =   "Observation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_ObsTraffic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mise à jour le 23/09/2017
'=========================
Option Explicit
    Public VEHICULE_                    As String
    Public HEURE_                       As String
    Public ALERTE_                      As String
    Dim FlapAlerte                      As Integer
    Public COMPT_ENTRE_
    Public maxDure As String
    
Private Sub Form_Load()
    Lbl_Alerte.Caption = ALERTE_
    Cmd_Save.Enabled = False
End Sub
Private Sub Cmd_Cancel_Click()
    With Frm_Trafic
        Call .Affiche_Vehicule
        Call .Affiche_Conducteur
        Call .Affiche_Destination
        .Txt_KM.Text = ""
        Call .grid_vehicule_ColumnClick(3)
        Call .grid_Conducteur_ColumnClick(3)
    End With
    Lbl_Alerte.Caption = ""
    ALERTE_ = ""
    HEURE_ = ""
    COMPT_ENTRE_ = 0
    VEHICULE_ = ""
    Unload Me
    Frm_Trafic.grid_Conducteur.SetFocus
End Sub
Private Sub Cmd_Save_Click()
    Call SaveTraffic
    Lbl_Alerte.Caption = ""
    ALERTE_ = ""
    HEURE_ = ""
    COMPT_ENTRE_ = 0
    VEHICULE_ = ""
    Unload Me
End Sub

Private Sub Timer_Alerte_Timer()
    If FlapAlerte = 0 Then
        Lbl_Alerte.ForeColor = &HFF&
        FlapAlerte = 1
    Else
        Lbl_Alerte.ForeColor = &HFF0000
        FlapAlerte = 0
    End If
End Sub
Private Sub Txt_Obs_Change()
    If Len(txt_Obs.Text) >= 7 Then Cmd_Save.Enabled = True Else Cmd_Save.Enabled = False
End Sub
Private Sub txt_Obs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then If Len(Trim(txt_Obs.Text)) >= 7 Then Cmd_Save.SetFocus
End Sub
Public Sub SaveTraffic()
    Dim Lrs_Find            As New Recordset
    Dim LObj_FindV          As New VEHICULE
    Dim LObj_FindT          As New Traffic
    Dim Code_vehicule       As String
    Dim NumeroFiche         As String
    Dim NumeroTxt           As String
    Dim Operation()         As String
    Dim Obv                 As String
On Error GoTo Err
    '================
    '-- CODE VEHICULE
    '================
    Set Lrs_Find = LObj_FindV.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VEHICULE_)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_FindV = Nothing
    If Not Lrs_Find.EOF Then Code_vehicule = Lrs_Find("Code")
    Set Lrs_Find = Nothing
    '===============================
    'GET OPERATION & NUMERO DE FICHE
    '===============================
    NumeroFiche = "Auto"
    ReDim Operation(1)
    Operation = ReturnOperation(Code_vehicule)
    If (Operation(0) = "E") Then
        NumeroFiche = Operation(1)
    End If
    NumeroTxt = Format(CStr(NumeroFiche), "00000")
    If Len(Trim(txt_Obs.Text)) = 0 Then
        MsgBox "Vous devez insérer une observation!", vbInformation, App.ProductName
        txt_Obs.SetFocus
        Exit Sub
    End If
    '========================
    'INSERTION ENREGISTREMENT
    '========================
    Set Lrs_Find = New Recordset
    Set Lrs_Find = CreateEmptyRS_Traffic()
    With Lrs_Find
        .AddNew
        .Fields("HeureEntre") = Format(HEURE_, "dd/mm/yyyy hh:mm:ss")
        .Fields("CompteurEntre") = COMPT_ENTRE_
        .Fields("OperateurEntre") = LStr_NameUser
        .Fields("Observation") = txt_Obs.Text
    End With
    Call LObj_FindT.UpDate_Traffic_VE(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Find, NumeroTxt)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_FindT = Nothing
    Set Lrs_Find = Nothing
    Set LObj_FindV = New VEHICULE
    Call LObj_FindV.UpdateCompteurFT_Vehicule(ErrNumber, ErrDescription, ErrSourceDetail, VEHICULE_, COMPT_ENTRE_, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_FindV = Nothing
    MsgBox "Enregistrement terminé avec succé!...", vbInformation, App.ProductName
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub
