VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultPieceReception 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consult Piece Reparation"
   ClientHeight    =   10680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   12705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt_typePiece 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   36
      Top             =   2760
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   5880
      ScaleHeight     =   1935
      ScaleWidth      =   4815
      TabIndex        =   23
      Top             =   2400
      Width           =   4815
      Begin VB.TextBox Txt_PMainOeuvre 
         Height          =   285
         Left            =   2760
         TabIndex        =   29
         Text            =   "0"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox Tex_RSP 
         Height          =   285
         Left            =   2760
         TabIndex        =   28
         Text            =   "0"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txt_Timbre 
         Height          =   285
         Left            =   2760
         TabIndex        =   27
         Text            =   "0"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Txt_TvaMO 
         Height          =   285
         Left            =   2760
         TabIndex        =   26
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Lbl_MOeuvre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prix main d'oeuvre :"
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
         Left            =   360
         TabIndex        =   33
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Remise sur pièce (%) :"
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
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Timbre fiscal :"
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
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Lbl_TvaMO 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TVA M.Oeuvre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox txt_Numero 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   9
      Tag             =   "M"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txt_ref 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox txt_BCReparation 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   4560
      Width           =   4695
      Begin VB.TextBox txt_MatriculeStation 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1800
         TabIndex        =   21
         Tag             =   "M"
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox txt_ville 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txt_adresse 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txt_rsocial 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Fournisseur 
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
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ville :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Raison sociale :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   840
         Width           =   1380
      End
   End
   Begin MSComctlLib.ListView Lsv_Detail 
      Height          =   3375
      Left            =   0
      TabIndex        =   10
      Top             =   6840
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Désignation"
         Object.Width           =   3617
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qte"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Véhicule"
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "PU.HT"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Remise (%)"
         Object.Width           =   3177
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tot.HT"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "TVA (%)"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Prix.TTC"
         Object.Width           =   1766
      EndProperty
   End
   Begin MSComctlLib.ListView Lsv_Toto 
      Height          =   1455
      Left            =   5640
      TabIndex        =   11
      Top             =   5160
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tot.HT.Brut"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tot.Remise.L"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "TOT.Remise.P"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tot.HT.Net"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tot.TVA"
         Object.Width           =   1853
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tot.TTC"
         Object.Width           =   1853
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   5775
      TabIndex        =   12
      Top             =   1440
      Width           =   5775
      Begin VB.Label cda_Create 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   35
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label cda_Operation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1680
         TabIndex        =   34
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Piece :"
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
         Left            =   3300
         TabIndex        =   14
         Top             =   0
         Width           =   1170
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Operation"
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
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.Label Lbl_user 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3240
      TabIndex        =   25
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Pièce de réception saisie par : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pièce de Rèception"
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
      TabIndex        =   20
      Top             =   360
      Width           =   3060
   End
   Begin VB.Image PicBox_Header 
      Height          =   1215
      Left            =   0
      Picture         =   "FrmConsultPieceReception.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro pièce"
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
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type de pièce"
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
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAUX EN (DT)"
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
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Référence"
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
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BC Reparation"
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
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "FrmConsultPieceReception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub AfficheRow(ByVal VCode As String)

Dim TotHTBrut As Double
Dim TotTTC As Double
Dim Fcode As String
Dim Qte As Double
Dim PUHT As Double
Dim Remise As Double
Dim tva As Double

Dim LOBJ_PRepar As PieceReparation
Dim rs As New Recordset

Call ViderZone(FrmConsultPieceReception)
Lsv_Detail.ListItems.Clear
Lsv_Toto.ListItems.Clear

Set LOBJ_PRepar = New PieceReparation
Set rs = LOBJ_PRepar.Get_AssPieceReparation(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
'Assiette
If Not rs.EOF Then
    'Charge
    Fcode = rs("Fournisseur")
    txt_Numero.Text = rs("Numero")
    If Not IsNull(rs("Type")) Then Txt_typePiece.Text = rs("Type")
    If Not IsNull(rs("refPiece")) Then txt_ref.Text = rs("refPiece")
    If Not IsNull(rs("DatePiece")) Then cda_Create.Caption = rs("DatePiece")
    If Not IsNull(rs("DateOperation")) Then cda_Operation.Caption = rs("DateOperation")
    If Not IsNull(rs("Fournisseur")) Then Call AfficheRow_Station(rs("Fournisseur"))
    If Not IsNull(rs("RemisePiece")) Then Tex_RSP.Text = Format(rs("RemisePiece"), "#,##0.00")
    If Not IsNull(rs("timbre")) Then txt_Timbre.Text = Format(rs("Timbre"), "#,##0.000")
    If Not IsNull(rs("PrixMOeuvre")) Then Txt_PMainOeuvre.Text = Format(rs("PrixMOeuvre"), "#,##0.000")
    If Not IsNull(rs("TVA_MOeuvre")) Then Txt_TvaMO.Text = rs("TVA_MOeuvre")
    If Not IsNull(rs("UserInsert")) Then Lbl_user.Caption = Get_NameUserByCode(rs("UserInsert"))
Else
    MsgBox "Numéro bon introuvable", vbInformation
    txt_Numero.SetFocus
    Exit Sub
End If
rs.Close

'Detail
Set rs = LOBJ_PRepar.Get_DetPieceReparation(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        TotTTC = 0
        TotHTBrut = 0
        Qte = 0
        PUHT = 0
        Remise = 0
        tva = 0
        
        Qte = rs("Qte")
        PUHT = rs("PUHT")
        Remise = rs("Remise")
        tva = rs("tva")
        
        TotHTBrut = FrmSaisiePieceReparation.Return_TotHT(Qte, PUHT, Remise)
        TotTTC = TotHTBrut + (TotHTBrut * (tva / 100))
            
            Set itmX = Lsv_Detail.ListItems.Add(, , CStr(txt_Numero.Text))
            itmX.SubItems(1) = rs("Designation")
            itmX.SubItems(2) = rs("Qte")
            itmX.SubItems(3) = rs("Vehicule")
            itmX.SubItems(4) = Format(rs("PUHT"), "#,##0.000")
            itmX.SubItems(5) = Format(rs("Remise"), "#0.00")
            itmX.SubItems(6) = Format(TotHTBrut, "#,##0.000")
            itmX.SubItems(7) = Format(rs("tva"), "#0.00")
            itmX.SubItems(8) = Format(TotTTC, "#,##0.000")
        rs.MoveNext
    Wend
    Call AppCalcul
End If
rs.Close

End Sub

Public Sub AfficheRow_Station(ByVal VCode As String)

Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.GetStatByCodeLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    'Charge
    txt_MatriculeStation.Text = rs("Code")
    If Not IsNull(rs("Libelle")) Then txt_rsocial.Text = rs("Libelle")
    If Not IsNull(rs("Adresse")) Then txt_adresse.Text = rs("Adresse")
    If Not IsNull(rs("Ville")) Then txt_ville.Text = rs("Ville")
Else
    MsgBox "Code introuvable", vbInformation
    txt_MatriculeStation.SetFocus
    Exit Sub
End If
rs.Close
End Sub

Public Sub AppCalcul()

Dim ii As Integer

'Ligne de pièce
Dim pu As Double
Dim Qte As Double
Dim HTBrutLigne As Double
Dim RemiseL As Double
Dim HTNetLigne As Double
Dim tvaLigne As Double
Dim ttcLigne As Double

'Totaux Pièce
Dim TotHTBrut As Double
Dim TotRemLigne As Double
Dim TotHtNet As Double
Dim TotTva  As Double
Dim ValRemP As Double
Dim TotTTCSansRP As Double
Dim RemiseP As Double
Dim TotHtNetPiece As Double
Dim Timbre As Double
Dim MainOeuvre As Double
Dim TVa_MOv As Double
Dim TotTTC As Double

'Intit totaux
TotHTBrut = 0
TotRemLigne = 0
RemiseP = 0
TotHtNet = 0
TotTva = 0
ValRemP = 0
Timbre = 0
TotTTC = 0
MainOeuvre = 0
TVa_MOv = 0

Lsv_Toto.ListItems.Clear

For ii = 1 To Lsv_Detail.ListItems.Count
  'Intit Lignes

    pu = 0
    Qte = 0
    HTBrutLigne = 0
    RemiseL = 0
    HTNetLigne = 0
    tvaLigne = 0
    ttcLigne = 0
    
    'TotHTBrut
    Qte = Lsv_Detail.ListItems(ii).SubItems(2)
    pu = Lsv_Detail.ListItems(ii).SubItems(4)
    HTBrutLigne = Qte * pu
    TotHTBrut = TotHTBrut + HTBrutLigne
    
    'TotRemLigne
    RemiseL = Lsv_Detail.ListItems(ii).SubItems(5)
    TotRemLigne = TotRemLigne + (HTBrutLigne * RemiseL / 100)
    
    'TotHtNet
    HTNetLigne = Lsv_Detail.ListItems(ii).SubItems(6)
    RemiseP = RemiseP + (HTNetLigne * CDbl(Tex_RSP.Text) / 100)
    HTNetLigne = HTNetLigne - (HTNetLigne * CDbl(Tex_RSP.Text) / 100)
    TotHtNet = TotHtNet + HTNetLigne
    
    'TotTva
    tvaLigne = Lsv_Detail.ListItems(ii).SubItems(7)
    TotTva = TotTva + (HTNetLigne * tvaLigne / 100)

Next

If Txt_PMainOeuvre.Text = "" Then Txt_PMainOeuvre.Text = 0
Txt_PMainOeuvre.Text = Format(Txt_PMainOeuvre.Text, "##0.000")
MainOeuvre = CDbl(Txt_PMainOeuvre.Text)

'TotHtNet et brut de toute piece
TotHtNet = TotHtNet + MainOeuvre - (MainOeuvre * CDbl(Tex_RSP.Text) / 100)
TotHTBrut = TotHTBrut + MainOeuvre

If Txt_TvaMO.Text = "" Then Txt_TvaMO.Text = "0"
Txt_TvaMO.Text = Format(Txt_TvaMO.Text, "##0.00")
Tva_MOeuvre = CDbl(Txt_TvaMO.Text)
Tva_MOeuvre = (MainOeuvre - (MainOeuvre * CDbl(Tex_RSP.Text) / 100)) * Tva_MOeuvre / 100
MainOeuvre = MainOeuvre + Tva_MOeuvre

'TotTva de toute la pièce
TotTva = TotTva + Tva_MOeuvre

'RemiseP
If Tex_RSP.Text = "" Then Tex_RSP.Text = "0"
Tex_RSP.Text = Format(Tex_RSP.Text, "##0.00")
RemiseP = RemiseP + MainOeuvre * CDbl(Tex_RSP.Text) / 100

'Timbre
If txt_Timbre.Text = "" Then txt_Timbre.Text = "0"
txt_Timbre.Text = Format(txt_Timbre.Text, "##0.000")
Timbre = CDbl(txt_Timbre.Text)

'TotHtNetPiece
TotHtNetPiece = TotHTBrut - TotRemLigne - RemiseP

'TotTTC

TotTTC = TotHtNetPiece + TotTva + Timbre


Set itmX = Lsv_Toto.ListItems.Add(, , Format(TotHTBrut, "#,##0.000"))
        itmX.SubItems(1) = Format(TotRemLigne, "#,##0.000")
        itmX.SubItems(2) = Format(RemiseP, "#,##0.000")
        itmX.SubItems(3) = Format(TotHtNetPiece, "#,##0.000")
        itmX.SubItems(4) = Format(TotTva, "#,##0.000")
        itmX.SubItems(5) = Format(TotTTC, "#,##0.000")

End Sub

