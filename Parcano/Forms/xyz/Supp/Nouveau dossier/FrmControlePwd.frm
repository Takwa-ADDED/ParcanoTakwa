VERSION 5.00
Begin VB.Form Frm_ControlePwd 
   BorderStyle     =   0  'None
   Caption         =   "Controle pwd"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5939.637
   ScaleMode       =   0  'User
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpasse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   960
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Lbl_exit 
      BackStyle       =   0  'Transparent
      Caption         =   "Annuler (Echap)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Entre Votre Mot de Passe"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Lbl_Valider 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Lbl_Quitter 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Image img_background 
      Height          =   3975
      Left            =   0
      Picture         =   "FrmControlePwd.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Frm_ControlePwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Public vCode As String
    Public Sible As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyEscape Then Unload Me
    If KeyAscii = vbKeyReturn Then Lbl_Valider_Click
End Sub
Private Sub txtpasse_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Lbl_Valider_Click
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
Private Sub Lbl_exit_Click()
    Unload Me
End Sub
Private Sub Lbl_Quitter_Click()
    Unload Me
End Sub
Private Sub Lbl_Valider_Click()
    Dim LObj_FindUser As Utilisateur
    Dim Lrs_User As Recordset
    Dim i As Integer
    Dim LInt_Code As Integer
    Dim exist As Boolean
    Dim w, k

On Error GoTo Err

    'Verifier si l'utilisateur est actif
    w = UCase(txtpasse.Text)
    k = ""
    For i = 1 To Len(w)
        k = k & Asc(Mid(w, i, 1))
    Next
    
    Set LObj_FindUser = New Utilisateur
    Set Lrs_User = LObj_FindUser.GetRow_UsersByPwd(ErrNumber, ErrDescription, ErrSourceDetail, UCase(k), CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_FindUser = Nothing
    
    exist = False
    If Not Lrs_User.EOF Then exist = True
    Set Lrs_User = Nothing
     
    If (exist) Then
        If Sible = "MajTraffic" Then
            Set LObj_FindUser = New Utilisateur
            Set Lrs_User = LObj_FindUser.GetRow_User_Maj_FT_ByPwd(ErrNumber, ErrDescription, ErrSourceDetail, UCase(k), CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            Set LObj_FindUser = Nothing

            If Lrs_User.EOF Then
                Set Lrs_User = Nothing
                MsgBox "Accès refusé.", vbExclamation, App.ProductName
                Exit Sub
            End If
            Set Lrs_User = Nothing
         
            Frm_MajTrafic.NumFiche = vCode
            Unload Me
            Frm_MajTrafic.Show
            Exit Sub
        ElseIf Sible = "Disponibilité" Then
            Set LObj_FindUser = New Utilisateur
            Set Lrs_User = LObj_FindUser.GetRow_User_Maj_Disp_ByPwd(ErrNumber, ErrDescription, ErrSourceDetail, UCase(k), CNB)
            If ErrNumber <> 0 Then
                ErrNumber = 0
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                Exit Sub
            End If
            Set LObj_FindUser = Nothing
            
            If Lrs_User.EOF Then
                Set Lrs_User = Nothing
                MsgBox "Accès refusé.", vbExclamation
                Exit Sub
            End If
            Set Lrs_User = Nothing
         
            Frm_Trafic.MajDisp
            Unload Me
            Frm_Trafic.Show
            Exit Sub
        End If
    Else
        MsgBox "Mot passe incorrect .?  ", vbExclamation
        With txtpasse
            .SetFocus
            .SelStart = 0
            .SelLength = Len(txtpasse.Text)
        End With
    End If
    
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
