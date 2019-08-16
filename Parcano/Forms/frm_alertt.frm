VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form Frm_Alertt 
   BackColor       =   &H000000C0&
   Caption         =   " ."
   ClientHeight    =   6345
   ClientLeft      =   660
   ClientTop       =   4950
   ClientWidth     =   13320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_alertt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   13320
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.ListBox LSV_CPTBC 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3300
      ItemData        =   "frm_alertt.frx":0ECA
      Left            =   10920
      List            =   "frm_alertt.frx":0ED1
      TabIndex        =   15
      Top             =   1560
      Width           =   1155
   End
   Begin VB.ListBox LSV_CPTBV 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3300
      ItemData        =   "frm_alertt.frx":0EE0
      Left            =   9720
      List            =   "frm_alertt.frx":0EE7
      TabIndex        =   14
      Top             =   1560
      Width           =   1155
   End
   Begin VB.ListBox LSV_CPTFT 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3300
      ItemData        =   "frm_alertt.frx":0EF6
      Left            =   8520
      List            =   "frm_alertt.frx":0EFD
      TabIndex        =   13
      Top             =   1560
      Width           =   1155
   End
   Begin VB.ListBox lsv_Code 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3030
      ItemData        =   "frm_alertt.frx":0F0C
      Left            =   3000
      List            =   "frm_alertt.frx":0F13
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ListBox Lsv_Visite 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3300
      ItemData        =   "frm_alertt.frx":0F23
      Left            =   5280
      List            =   "frm_alertt.frx":0F2A
      TabIndex        =   8
      Top             =   1560
      Width           =   1035
   End
   Begin VB.ListBox lsv_Taxe 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3300
      ItemData        =   "frm_alertt.frx":0F3A
      Left            =   6360
      List            =   "frm_alertt.frx":0F41
      TabIndex        =   7
      Top             =   1560
      Width           =   1035
   End
   Begin VB.ListBox Lsv_VID 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3300
      ItemData        =   "frm_alertt.frx":0F4F
      Left            =   0
      List            =   "frm_alertt.frx":0F56
      TabIndex        =   2
      Top             =   1560
      Width           =   1275
   End
   Begin VB.ListBox Lst_Vehicule 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3300
      ItemData        =   "frm_alertt.frx":0F63
      Left            =   1320
      List            =   "frm_alertt.frx":0F6A
      TabIndex        =   1
      Top             =   1560
      Width           =   3915
   End
   Begin VB.ListBox Lsv_ASS 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   3300
      ItemData        =   "frm_alertt.frx":0F7C
      Left            =   7440
      List            =   "frm_alertt.frx":0F83
      TabIndex        =   0
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   120
      Top             =   120
   End
   Begin SToolBox.SCommand SCommand1 
      Height          =   255
      Left            =   11880
      TabIndex        =   5
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   ":"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483634
      ButtonType      =   1
   End
   Begin SToolBox.SCommand SCommand2 
      Height          =   255
      Left            =   11520
      TabIndex        =   6
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "x"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483634
      ButtonType      =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Echap pour sortir"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   240
      Left            =   10440
      TabIndex        =   21
      Top             =   6000
      Width           =   2085
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   975
      Left            =   1680
      Top             =   4920
      Width           =   8415
   End
   Begin VB.Label lbl_designation 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Veuillez contactez vos supérieurs  pour gérer ces véhicules !"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   930
      Left            =   1680
      TabIndex        =   20
      Top             =   4920
      Width           =   8520
      WordWrap        =   -1  'True
   End
   Begin VB.Image Pic_Footer 
      Height          =   855
      Left            =   0
      Picture         =   "frm_alertt.frx":0F90
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   12375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   1800
      X2              =   -120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   12120
      X2              =   8400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   735
      Left            =   1800
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label lbl_urgent 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URGENT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   3840
      TabIndex        =   19
      Top             =   120
      Width           =   3750
      WordWrap        =   -1  'True
   End
   Begin VB.Image Pic_Header 
      Height          =   855
      Left            =   -120
      Picture         =   "frm_alertt.frx":29BD2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CPT.BC"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   10920
      TabIndex        =   18
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CPT.BV"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   9720
      TabIndex        =   17
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CPT.FT"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8400
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ASSUR"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TAXE"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VISITE"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VIDANG"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VEHICULE"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   12120
      X2              =   12120
      Y1              =   600
      Y2              =   5880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   0
      Y1              =   600
      Y2              =   5880
   End
End
Attribute VB_Name = "Frm_Alertt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i
Dim iii
Dim k

Private Sub Form_Load()
    Me.Move 0, 0
    Call Alerte
    iii = 36
    k = 0
End Sub
Public Sub Alerte()

    Dim LObj_Find As VEHICULE
    Dim Lrs_Vehicule As Recordset
    Dim kk As Integer
    Dim AA
    Dim Okay As Boolean

    Set LObj_Find = New VEHICULE
    Set Lrs_Vehicule = LObj_Find.GetRow_Vehicule_ByDateSortie(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing

    With Frm_Alertt
        .Lst_Vehicule.Clear
        .Lsv_ASS.Clear
        .lsv_Taxe.Clear
        .Lsv_Visite.Clear
        .Lsv_VID.Clear
        .lsv_Code.Clear
        .LSV_CPTFT.Clear
        .LSV_CPTBV.Clear
        .LSV_CPTBC.Clear

        If Not Lrs_Vehicule.EOF Then
            While Not Lrs_Vehicule.EOF
                kk = 0
                'Alert assurance
                If Lrs_Vehicule("DateFinAssur") <> "01/01/1900" And DateAdd("d", 5, Date) >= Lrs_Vehicule("DateFinAssur") Then
                    If (Date - Lrs_Vehicule("DateFinAssur")) > 5 Then
                        .Lsv_ASS.AddItem "+ " & ((Date - Lrs_Vehicule("DateFinAssur")) - 5) & " JLrs_Vehicule"
                    Else
                        .Lsv_ASS.AddItem -(Date - Lrs_Vehicule("DateFinAssur"))
                    End If
                    kk = 1
                End If
                'Fin assurance

                'Alert visite
                If Lrs_Vehicule("DAteFinVisite") <> "01/01/1900" And DateAdd("d", 5, Date) >= Lrs_Vehicule("DAteFinVisite") Then
                    If (Date - Lrs_Vehicule("DAteFinVisite")) > 5 Then
                        .Lsv_Visite.AddItem "+ " & ((Date - Lrs_Vehicule("DAteFinVisite")) - 5) & " JLrs_Vehicule"
                    Else
                        .Lsv_Visite.AddItem -(Date - Lrs_Vehicule("DAteFinVisite"))
                    End If
                    kk = 1

                End If
                'Fin visite

                'Alert tax
                If Lrs_Vehicule("DateFinTax") <> "01/01/1900" And DateAdd("d", 5, Date) >= Lrs_Vehicule("DateFinTax") Then
                    If (Date - Lrs_Vehicule("DateFinTax")) > 5 Then
                        .lsv_Taxe.AddItem "+ " & ((Date - Lrs_Vehicule("DateFinTax")) - 5) & " JLrs_Vehicule"
                    Else
                        .lsv_Taxe.AddItem -(Date - Lrs_Vehicule("DateFinTax"))
                    End If

                    kk = 1
                End If
                'Fin tax

                'Alert vidange
                If Lrs_Vehicule("CompteurVidange") <> 0 Then
                    AA = Lrs_Vehicule("CompteurFT") - Lrs_Vehicule("CompteurVidange")
                    If AA + 500 >= Lrs_Vehicule("NBKLMvidange") Then
                        If AA > Lrs_Vehicule("NBKLMvidange") Then
                            .Lsv_VID.AddItem "+ " & (AA - Lrs_Vehicule("NBKLMvidange")) & " klm"
                        Else
                            .Lsv_VID.AddItem "- " & (Lrs_Vehicule("NBKLMvidange") - AA) & " klm"
                        End If

                        kk = 1
                    End If
                End If
                'Fin

                If kk = 1 Then
                    If Not (IsNull(Lrs_Vehicule("Marque"))) Then .Lst_Vehicule.AddItem Lrs_Vehicule("Marque") & "   " & Lrs_Vehicule("Matricule")
                    'CompteuLrs_Vehicule
                    If Not (IsNull(Lrs_Vehicule("CompteurFT"))) Then .LSV_CPTFT.AddItem (Lrs_Vehicule("CompteurFT"))
                    If Not (IsNull(Lrs_Vehicule("CompteurVidange"))) Then .LSV_CPTBV.AddItem (Lrs_Vehicule("CompteurVidange"))
                    If Not (IsNull(Lrs_Vehicule("CompteurCarburant"))) Then LSV_CPTBC.AddItem (Lrs_Vehicule("CompteurCarburant"))
                    If Not (IsNull(Lrs_Vehicule("Code"))) Then .lsv_Code.AddItem Lrs_Vehicule("Code")
                    AA = .Lst_Vehicule.ListCount
                    If .Lsv_Visite.ListCount < AA Then Lsv_Visite.AddItem ""
                    If .lsv_Taxe.ListCount < AA Then lsv_Taxe.AddItem ""
                    If .Lsv_ASS.ListCount < AA Then Lsv_ASS.AddItem ""
                    If .Lsv_VID.ListCount < AA Then Lsv_VID.AddItem ""

                End If
                Lrs_Vehicule.MoveNext
            Wend
        End If
    End With
End Sub
Private Sub Command1_Click()
    Lst_Vehicule.Clear
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeySpace Then Unload Me
End Sub

Private Sub Form_Resize()
    Dim WidthForm As Integer
    WidthForm = Frm_Main.Width
    Pic_Header.Width = WidthForm
    Pic_Footer.Width = WidthForm
    Label3.Left = WidthForm - 4600
    SCommand1.Left = WidthForm - 3000
    SCommand2.Left = WidthForm - 3300

End Sub

Private Sub SCommand1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()

    On Error Resume Next

    lbl_urgent.Font.Size = iii
    If k = 1 Then
        iii = iii - 2
    Else
        iii = iii + 2
    End If

    If iii = 52 Then k = 1
    If iii = 36 Then k = 0
End Sub
