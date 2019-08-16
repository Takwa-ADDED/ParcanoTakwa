VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form FrmFind_BCarb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liste des bons de carburant"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Pict_BonC 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   8175
      TabIndex        =   4
      Top             =   960
      Width           =   8175
      Begin MSComCtl2.DTPicker cda_Db 
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   156893185
         CurrentDate     =   42875
      End
      Begin MSComCtl2.DTPicker cda_fin 
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   156893185
         CurrentDate     =   42875
      End
      Begin VB.ComboBox Cb_typBC 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   2775
      End
      Begin SToolBox.SBiCombo Cmb_station 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image CmdFind 
         Height          =   495
         Left            =   5160
         Picture         =   "FrmFind_BCarb.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Lab_DatD 
         BackStyle       =   0  'Transparent
         Caption         =   "Op:  Du:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label LabDatF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Au :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Lab_typ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type BC :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Lab_Stat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Station :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
   End
   Begin SToolBox.SCommand SCommand1 
      Height          =   255
      Left            =   7800
      TabIndex        =   0
      Top             =   240
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
      ForeColor       =   255
      ButtonType      =   1
   End
   Begin SToolBox.SCommand SCommand2 
      Height          =   255
      Left            =   7560
      TabIndex        =   1
      Top             =   240
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
      ForeColor       =   255
      ButtonType      =   1
   End
   Begin SToolBox.SGrid grid 
      Height          =   4695
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8281
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin VB.Label LBL_Titre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recherche des bons de carburant"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   4710
   End
   Begin VB.Image Pic_Header 
      Height          =   1095
      Left            =   0
      Picture         =   "FrmFind_BCarb.frx":10C02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "FrmFind_BCarb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StrSource As String
Public RETOUR As Integer

Private Sub Cmb_station_LostFocus()
Call ExistDonnee(Cmb_station)
End Sub

'Afficher liste des bonCarburant selon type : facturé ou non facturé supprimé ou non supprimé
Private Sub CmdFind_Click()

Dim typeBC As String
typeBC = Cb_typBC.Text
grid.ClearRows
If cda_Db.Value > cda_fin.Value Then
    MsgBox "Vérifier dates de recherche! ", vbInformation
    Exit Sub
End If
If cda_Db.Value = "" Or cda_fin.Value = "" Then
    MsgBox "Entrer dates de recherche des commandes ", vbInformation
    Exit Sub
End If
'Renvoie ou définit une valeur indiquant que les données du contrôle ont été modifiées
'par un processus autre que l'extraction de données à partir de l'enregistrement en cours.
Call Affiche_BC(cda_Db.Value, cda_fin.Value, Cmb_station.FirstValue, Cb_typBC.Text)
    
End Sub

'initialisation de la forme de recherche des BC
Private Sub init_searchBC()
Pict_BonC.Visible = True
cda_Db.Value = "01/01/" & Year(Date)
cda_fin.Value = Format(Date, "DD/MM/YYYY")
'charger comboBox Items
Cb_typBC.AddItem "Tout BC"
Cb_typBC.AddItem "BC Facturé"
Cb_typBC.AddItem "BC Non Facturé"
Cb_typBC.AddItem "BC Supprimé"
Cb_typBC.ListIndex = 2
End Sub

'Retourne tout les BC d'une station dans un interval de temps
Public Sub Affiche_BC(ByVal DateDu As Date, ByVal DateAu As Date, ByVal Station As String, ByVal TYP As String)

Dim LOBJ_BonCarburant As BonCarburant
Dim rs As New Recordset

grid.ClearRows
If cda_Db.Value = "" Or cda_fin.Value = "" Then
    MsgBox "Entrer dates de recherche des commandes ", vbInformation
    Exit Sub
End If

Set LOBJ_BonCarburant = New BonCarburant
Set rs = LOBJ_BonCarburant.Get_BonCarburant(ErrNumber, ErrDescription, ErrSourceDetail, DateDu, DateAu, Station, TYP, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

If Not rs.EOF Then
    grid.Redraw = False
    While Not rs.EOF
        With grid
            .AddRow
            .CellDetails .Rows, 1, rs("Numero")
            .CellDetails .Rows, .ColumnIndex("Date"), rs("DateDoc")
            .CellDetails .Rows, .ColumnIndex("Station"), rs("Libelle")
            .CellDetails .Rows, .ColumnIndex("Conducteur"), rs("Pers")
            .CellDetails .Rows, .ColumnIndex("Valeur"), Format(rs("VALEUR"), "#,##0.000"), DT_RIGHT
            .CellDetails .Rows, .ColumnIndex("NumFact"), rs("NumFact")
            .CellDetails .Rows, .ColumnIndex("Supp"), rs("Supp")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
    grid.SelectedRow = grid.Rows
    
End If
rs.Close

End Sub

Private Sub Form_Load()

Call init_searchBC
Call Affiche_StatCarb_SBCombo(Cmb_station)
Cmb_station.ListIndex = 0
Call Initgrid_BonCarburant
Call CmdFind_Click
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub grid_DblClick(ByVal lRow As Long, ByVal lCol As Long)

Dim VCode
Dim LOBJ_Personnel As New personnel
On Error GoTo Err

VCode = grid.CellText(lRow, 1)
        If (CHECK_ACCES("MAJ_BC", LInt_UserId) = False) Then
            MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier un Bon de Carburant", vbExclamation
            Exit Sub
        End If
Unload Me
FrmAllBonCarburant.AfficheRow (VCode)

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

Dim VCode
Dim LOBJ_Personnel As New personnel
On Error GoTo Err
VCode = grid.CellText(grid.SelectedRow, 1)
If Not LOBJ_Personnel.Verif_USER_Access(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "MAJ_BC", LInt_UserId) Then
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If
Unload Me
FrmAllBonCarburant.AfficheRow (VCode)

Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub SCommand1_Click()
Unload Me
End Sub

Private Sub Initgrid_BonCarburant()
With grid
.Redraw = False

    .HideGroupingBox = True
    .AllowGrouping = True

    .GroupRowBackColor = vbWindowBackground
    .GroupRowForeColor = vbWindowText
    
    .GridLineColor = vbWindowBackground
    .GridFillLineColor = vbWindowBackground
    .GridLines = True
    
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .DrawFocusRectangle = False
    
    .AddColumn "Code", "Numero", , , 60, , , , , , , CCLSortNumeric
    .AddColumn "Date", "Date", , , 70
    .AddColumn "Station", "Station", , , 120
    .AddColumn "Conducteur", "Conducteur", , , 80
    .AddColumn "Valeur", "Valeur", , , 80
    .AddColumn "NumFact", "Facture", , , 80
    .AddColumn "Supp", "Supprimé", , , 80
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
.Redraw = True
End With
End Sub



