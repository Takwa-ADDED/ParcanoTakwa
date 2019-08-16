VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmBonCarburant 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bon carburant"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   10215
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   660
      ScaleHeight     =   375
      ScaleWidth      =   2295
      TabIndex        =   48
      Top             =   2040
      Width           =   2295
      Begin SToolBox.SDateBox cda_Create 
         Height          =   285
         Left            =   960
         TabIndex        =   49
         Tag             =   "M"
         Top             =   0
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Text            =   ""
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         Height          =   195
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1020
      ScaleHeight     =   375
      ScaleWidth      =   1935
      TabIndex        =   45
      Top             =   7080
      Width           =   1935
      Begin VB.TextBox txt_Valeur 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   46
         Tag             =   "M"
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valeur :"
         Height          =   195
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   555
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   -240
      ScaleHeight     =   1455
      ScaleWidth      =   10455
      TabIndex        =   22
      Top             =   2880
      Width           =   10455
      Begin VB.TextBox txt_Obs 
         Height          =   1035
         Left            =   8160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   0
         Width           =   2175
      End
      Begin VB.TextBox txt_compteur 
         Height          =   315
         Left            =   5760
         TabIndex        =   27
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txt_Type 
         Height          =   315
         Left            =   1860
         TabIndex        =   26
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txt_libelle 
         Height          =   315
         Left            =   1860
         TabIndex        =   25
         Top             =   0
         Width           =   2295
      End
      Begin VB.TextBox txt_Energie 
         Height          =   315
         Left            =   1860
         TabIndex        =   24
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txt_prixLitre 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1860
         TabIndex        =   23
         Tag             =   "M"
         Top             =   1080
         Width           =   1215
      End
      Begin SToolBox.SDateBox cda_FinAssur 
         Height          =   285
         Left            =   5760
         TabIndex        =   35
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Text            =   ""
      End
      Begin SToolBox.SDateBox cda_FinVisite 
         Height          =   285
         Left            =   5760
         TabIndex        =   36
         Top             =   720
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Text            =   ""
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date fin assurance :"
         Height          =   195
         Left            =   4200
         TabIndex        =   38
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date fin visite :"
         Height          =   195
         Left            =   4200
         TabIndex        =   37
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observation :"
         Height          =   195
         Left            =   7080
         TabIndex        =   34
         Top             =   0
         Width           =   990
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         Height          =   195
         Left            =   300
         TabIndex        =   33
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compteur :"
         Height          =   195
         Left            =   4200
         TabIndex        =   32
         Top             =   0
         Width           =   1500
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Energie :"
         Height          =   195
         Left            =   0
         TabIndex        =   31
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Matricule :"
         Height          =   195
         Left            =   300
         TabIndex        =   30
         Top             =   0
         Width           =   1500
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prix de 1 Litre TTC :"
         Height          =   195
         Left            =   375
         TabIndex        =   29
         Top             =   1080
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   480
      ScaleHeight     =   1095
      ScaleWidth      =   4215
      TabIndex        =   21
      Top             =   5280
      Width           =   4215
      Begin VB.TextBox txt_ville 
         Height          =   315
         Left            =   1140
         TabIndex        =   41
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txt_adresse 
         Height          =   315
         Left            =   1140
         TabIndex        =   40
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txt_rsocial 
         Height          =   315
         Left            =   1140
         TabIndex        =   39
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ville :"
         Height          =   195
         Left            =   735
         TabIndex        =   44
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse :"
         Height          =   195
         Left            =   405
         TabIndex        =   43
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Raison sociale :"
         Height          =   195
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   1110
      End
   End
   Begin VB.TextBox txt_MatriculeStation 
      BackColor       =   &H80000016&
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
      Height          =   360
      Left            =   1620
      TabIndex        =   3
      Tag             =   "M"
      Top             =   4860
      Width           =   2295
   End
   Begin VB.TextBox txt_NbreLitre 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   465
      Left            =   1620
      TabIndex        =   4
      Tag             =   "M"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.ComboBox Cbo_Conducteur 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1620
      TabIndex        =   2
      Tag             =   "M"
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txt_Numero 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   465
      Left            =   1620
      TabIndex        =   0
      Tag             =   "M"
      Top             =   1470
      Width           =   2295
   End
   Begin VB.TextBox txt_Matricule 
      BackColor       =   &H80000016&
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
      Height          =   345
      Left            =   1620
      TabIndex        =   1
      Tag             =   "M"
      Top             =   2490
      Width           =   2295
   End
   Begin SToolBox.SCommand CmdSave 
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
      Picture         =   "FrmBonCarburant.frx":0000
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   8520
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
      Picture         =   "FrmBonCarburant.frx":0182
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   8880
      TabIndex        =   7
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
      Picture         =   "FrmBonCarburant.frx":04D5
   End
   Begin SToolBox.SCommand cmdFindNumero 
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   1470
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   661
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
      Picture         =   "FrmBonCarburant.frx":0828
      ButtonType      =   1
   End
   Begin SToolBox.SCommand CmdFindConducteur 
      Height          =   315
      Left            =   3960
      TabIndex        =   15
      Top             =   4440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
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
      Picture         =   "FrmBonCarburant.frx":0B7B
      ButtonType      =   1
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   345
      Left            =   3960
      TabIndex        =   18
      Top             =   2490
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
      Picture         =   "FrmBonCarburant.frx":0ECE
      ButtonType      =   1
   End
   Begin SToolBox.SCommand CmdEdit 
      Height          =   495
      Left            =   8160
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
      Picture         =   "FrmBonCarburant.frx":1221
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   7800
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
      Picture         =   "FrmBonCarburant.frx":13A3
   End
   Begin SToolBox.SCommand CmdPrint 
      Height          =   495
      Left            =   9720
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
      Picture         =   "FrmBonCarburant.frx":1525
   End
   Begin SToolBox.SCommand CmdFindStation 
      Height          =   360
      Left            =   3960
      TabIndex        =   19
      Top             =   4860
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   635
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
      Picture         =   "FrmBonCarburant.frx":1878
      ButtonType      =   1
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Station "
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
      Left            =   840
      TabIndex        =   20
      Top             =   4860
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Nbre Litre "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conducteur :"
      Height          =   195
      Left            =   615
      TabIndex        =   16
      Top             =   4440
      Width           =   945
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro bon"
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
      Left            =   420
      TabIndex        =   14
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bon de sortie carburant"
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
      TabIndex        =   12
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Immatriculation"
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
      Left            =   60
      TabIndex        =   11
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmBonCarburant.frx":1BCB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "FrmBonCarburant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function RET_CODE_CONDUCTEUR(txt As String) As String
Dim SQL As String
Dim rs As New ADODB.Recordset
RET_CODE_CONDUCTEUR = ""
SQL = "select code from personnel where libelle = " & SQLText(txt)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    RET_CODE_CONDUCTEUR = rs(0)
End If
rs.Close
End Function

Private Function RET_PRIX_ENERGIE(txt As String) As Double
Dim SQL As String
Dim rs As New ADODB.Recordset
RET_PRIX_ENERGIE = 0
SQL = "select Prix from energie where libelle = " & SQLText(txt)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    RET_PRIX_ENERGIE = rs(0)
End If
rs.Close
End Function

Private Sub Cbo_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub CmdAdd_Click()
On Error GoTo Err
Call ViderZone(FrmBonCarburant)
cda_Create.Text = Date
txt_Matricule.SetFocus
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdDelete_Click()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim vcode As String
On Error GoTo Err
    If MsgBox("Confirmez vous la suppression de ce " & vbNewLine & "bon carburant", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    vcode = txt_Numero.Text
    SQL = "Delete from BonCarburant where Numero =" & SQLText(vcode)
    CNB.Execute SQL
    txt_Numero.SetFocus
    End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub


Private Sub CmdFind_Click()
txt_Numero.SetFocus
End Sub


Private Sub CmdFindConducteur_Click()
With FrmFind_Fils
    .StrSource = "Personnel"
    .Show
End With
End Sub

Private Sub cmdFindMatricule_Click()
With FrmFind_Fils
    .StrSource = "Véhicule"
    .Show
End With
End Sub

Private Sub cmdFindNumero_Click()
With FrmFind
    .StrSource = "BonCarburant"
    .Show
End With
End Sub


Private Sub CmdFindStation_Click()
With FrmFind_Fils
    .StrSource = "Station"
    .Show
End With

End Sub

Private Sub CmdSave_Click()

Dim SQL As String
Dim rs As New ADODB.Recordset

    If Left(CheckMandatory(FrmBonCarburant), 1) = 1 Then
       Exit Sub
    End If
    
On Error GoTo Err
    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    'Delete l'enregistrement
    vcode = txt_Numero.Text
    CNB.BeginTrans
    SQL = "Delete from BonCarburant where Numero =" & SQLText(vcode)
    CNB.Execute SQL
    'Insertion enregistrement
    SQL = "Insert into BonCarburant  (Numero,DateDoc,Vehicule,Station,Conducteur,litre,valeur,prixLitre) values ("
    SQL = SQL & SQLText(txt_Numero.Text)
    SQL = SQL & "," & SQLText(cda_Create.Text)
    SQL = SQL & "," & SQLText(txt_Matricule.Text)
    SQL = SQL & "," & SQLText(txt_MatriculeStation.Text)
    SQL = SQL & "," & SQLText(RET_CODE_CONDUCTEUR(Cbo_Conducteur.Text))
    SQL = SQL & "," & (txt_NbreLitre.Text)
    SQL = SQL & "," & Replace(txt_Valeur.Text, ",", ".")
    SQL = SQL & "," & Replace(txt_prixLitre.Text, ",", ".")
    SQL = SQL & ")"
    CNB.Execute SQL
    CNB.CommitTrans
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
    txt_Matricule.SetFocus
    End If
Exit Sub
Err:
CNB.RollbackTrans
MsgBox Err.Description, vbInformation
End Sub

Private Sub Form_Load()
Me.Width = 10380
Me.Height = 7935
cda_Create.Text = Date
End Sub
Public Sub AfficheRow_Vehicule(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from vehicule where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    If Not IsNull(rs("Matricule")) Then txt_libelle.Text = rs("Matricule")
    If Not IsNull(rs("marque")) Then txt_Type.Text = rs("TYPE")
    If Not IsNull(rs("Energie")) Then txt_Energie.Text = rs("Energie")
    If Not IsNull(rs("compteur")) Then txt_compteur.Text = rs("compteur")
    If Not IsNull(rs("DAteFinAssur")) Then cda_FinAssur.Text = rs("DAteFinAssur")
    If Not IsNull(rs("DAteFinVisite")) Then cda_FinVisite.Text = rs("DAteFinVisite")
    txt_prixLitre = Format(RET_PRIX_ENERGIE(rs("Energie")), "#,##0.000")
End If
rs.Close

End Sub


Public Sub AfficheRow_Vehicule_sansPrix(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from vehicule where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    If Not IsNull(rs("Matricule")) Then txt_libelle.Text = rs("Matricule")
    If Not IsNull(rs("marque")) Then txt_Type.Text = rs("TYPE")
    If Not IsNull(rs("Energie")) Then txt_Energie.Text = rs("Energie")
    If Not IsNull(rs("compteur")) Then txt_compteur.Text = rs("compteur")
    If Not IsNull(rs("DAteFinAssur")) Then cda_FinAssur.Text = rs("DAteFinAssur")
    If Not IsNull(rs("DAteFinVisite")) Then cda_FinVisite.Text = rs("DAteFinVisite")
End If
rs.Close

End Sub

Private Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Matricule_LostFocus()
On Error GoTo Err

Call AfficheRow_Vehicule(txt_Matricule.Text)

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub txt_MatriculeStation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_NbreLitre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub txt_NbreLitre_LostFocus()

Dim P As Double
Dim L As Integer
Dim V As Double

On Error GoTo Err

P = txt_prixLitre.Text
L = Val(txt_NbreLitre.Text)
V = P * L
txt_Valeur.Text = Format(V, "#,##0.000")

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub


Private Sub txt_Numero_GotFocus()
Call ViderZone(FrmBonCarburant)
End Sub

Private Sub txt_Numero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Public Sub AfficheRow_Station(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from station where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_MatriculeStation.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
'    If Not IsNull(rs("CPOSTAL")) Then txt_codepostal.Text = rs("CPOSTAL")
'    If Not IsNull(rs("Activite")) Then txt_activite.Text = rs("Activite")
'    If Not IsNull(rs("telephone")) Then txt_telephone.Text = rs("telephone")
'    If Not IsNull(rs("mobile")) Then txt_mobile.Text = rs("mobile")
'    If Not IsNull(rs("fax")) Then txt_fax.Text = rs("fax")
'    If Not IsNull(rs("email")) Then txt_email.Text = rs("email")
    
End If

End Sub

Private Sub txt_Numero_LostFocus()

On Error GoTo Err

Call AfficheRow(txt_Numero.Text)

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub
Public Sub AfficheRow(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from BonCarburant where Numero = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Numero.Text = rs("Numero")
    If Not IsNull(rs("VEHICULE")) Then txt_libelle.Text = rs("VEHICULE")
    If Not IsNull(rs("STATION")) Then txt_Type.Text = rs("STATION")
    If Not IsNull(rs("DATEDOC")) Then cda_Create.Text = rs("DATEDOC")
    If Not IsNull(rs("CONDUCTEUR")) Then txt_Energie.Text = rs("CONDUCTEUR")
    If Not IsNull(rs("LITRE")) Then txt_NbreLitre.Text = rs("LITRE")
    If Not IsNull(rs("VALEUR")) Then txt_Valeur.Text = Format(rs("VALEUR"), "#,##0.000")
    If Not IsNull(rs("PrixLitre")) Then txt_prixLitre.Text = Format(rs("PrixLitre"), "#,##0.000")
    Call AfficheRow_Vehicule_sansPrix(rs("VEHICULE"))
    Call AfficheRow_Station(rs("STATION"))
    Call AfficheRow_Conducteur(rs("CONDUCTEUR"))
    
End If
rs.Close

End Sub
Public Sub AfficheRow_Conducteur(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from personnel where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Libelle")) Then Cbo_Conducteur.Text = rs("Libelle")
End If
rs.Close

End Sub
