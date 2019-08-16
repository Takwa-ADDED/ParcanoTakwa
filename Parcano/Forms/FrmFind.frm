VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form FrmFind 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "."
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10350
   ClipControls    =   0   'False
   Icon            =   "FrmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pict_BonVdg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1065
      ScaleWidth      =   9705
      TabIndex        =   13
      Top             =   960
      Width           =   9735
      Begin VB.ComboBox Cmb_typeBV 
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   600
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker cda_Db 
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   63766529
         CurrentDate     =   42875
      End
      Begin MSComCtl2.DTPicker cda_fin 
         Height          =   255
         Left            =   4680
         TabIndex        =   15
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   63766529
         CurrentDate     =   42875
      End
      Begin SToolBox.SBiCombo SBC_Stat 
         Height          =   315
         Left            =   4920
         TabIndex        =   17
         Top             =   600
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
      Begin VB.Image Cmd_FindBV 
         Height          =   510
         Left            =   7920
         Picture         =   "FrmFind.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Station :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type Bv :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Au :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Op:  Du:"
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
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Pict_PieceRep 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1065
      ScaleWidth      =   10065
      TabIndex        =   3
      Top             =   960
      Width           =   10095
      Begin SToolBox.SBiCombo Cmb_station 
         Height          =   315
         Left            =   5040
         TabIndex        =   11
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
      Begin VB.ComboBox Cb_typBC 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   480
         Width           =   2775
      End
      Begin SToolBox.SDateBox DateBx_Db 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Text            =   ""
      End
      Begin SToolBox.SDateBox DateBox_fin 
         Height          =   285
         Left            =   5040
         TabIndex        =   7
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Text            =   ""
      End
      Begin VB.Image CmdFind 
         Height          =   495
         Left            =   7920
         Picture         =   "FrmFind.frx":10C0E
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Lab_Stat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Station :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Lab_typ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Etat Pièce :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label LabDatF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Au :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   6
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Lab_DatD 
         BackStyle       =   0  'Transparent
         Caption         =   "Op:  Du:"
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
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
   End
   Begin SToolBox.SCommand SCommand1 
      Height          =   255
      Left            =   8880
      TabIndex        =   0
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
      ForeColor       =   255
      ButtonType      =   1
   End
   Begin SToolBox.SCommand SCommand2 
      Height          =   255
      Left            =   8640
      TabIndex        =   1
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
      ForeColor       =   255
      ButtonType      =   1
   End
   Begin SToolBox.SGrid grid 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8070
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
      Caption         =   "Liste des véhicules"
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
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   3045
   End
   Begin VB.Image PicBox_Header 
      Height          =   1095
      Left            =   0
      Picture         =   "FrmFind.frx":21810
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "FrmFind"
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

Private Sub Cmd_FindBV_Click()

If cda_Db.Value > cda_fin.Value Then
    MsgBox "Vérifier dates de recherche ", vbInformation
    Exit Sub
End If
If cda_Db.Value = "" Or cda_fin.Value = "" Then
    MsgBox "Entrer dates de recherche ", vbInformation
    Exit Sub
End If
grid.ClearRows
If StrSource = "BonVidange2" Then
    Call Affiche_BonVdg(cda_Db.Value, cda_fin.Value, SBC_Stat.FirstValue, Cmb_typeBV.Text)
End If
End Sub

'Afficher liste des pièce réparation selon type : facturé ou non facturé supprimé ou non supprimé
Private Sub CmdFind_Click()

Dim etatBC As String
Dim typePiece As String

etatBC = Cb_typBC.Text
If DateBx_Db.Text > DateBox_fin.Text Then
    MsgBox "Vérifier dates de recherche ", vbInformation
    Exit Sub
End If
If DateBx_Db.Text = "" Or DateBox_fin.Text = "" Then
    MsgBox "Entrer dates de recherche ", vbInformation
    Exit Sub
End If

Select Case StrSource
    Case "BRPieceReparation"  'frmPieceReparation
        typePiece = "Bon Retour"
    Case "BLPieceReparation"
        typePiece = "Piece Reception"
    Case "FacturePieceReparation"
        typePiece = "Facture"
    Case "AvoirPieceReparation"
        typePiece = "Avoir"
    Case "AllPieceReparation"
        typePiece = "Tout"
End Select

If StrSource = "Reparation" Or StrSource = "FIndBCReparation" Then 'frmBCReparation
    
    Call Affiche_Reparation(DateBx_Db.Text, DateBox_fin.Text, Cmb_station.FirstValue, etatBC)
Else
    'Filtrage par date , par station ,  par type
    Call Affiche_PieceReparRech(DateBx_Db.Text, DateBox_fin.Text, Cmb_station.FirstValue, etatBC, typePiece)
End If
End Sub

Private Sub Form_Activate()
 If grid.Rows = 0 Then MsgBox "Pas de données à visualiser", vbInformation
 If grid.Rows > 0 Then
    grid.SelectedRow = 1
    grid.SetFocus
End If
End Sub

'initialisation de la forme de recherche des BC
Private Sub init_searchPRep()

Pict_PieceRep.Visible = True
DateBx_Db.Text = "01/01/" & Year(Date)
DateBox_fin.Text = Format(Date, "DD/MM/YYYY")
Call Affiche_StatRep_SBCombo(Cmb_station)
'charger comboBox Items
Cb_typBC.AddItem "Toute pièces "
Cb_typBC.AddItem "Pièce Facturé"
Cb_typBC.AddItem "Pièce Non Facturé"
Cb_typBC.AddItem "Pièce Supprimé"
Cb_typBC.ListIndex = 2
End Sub

Private Sub init_searchBV()

Pict_BonVdg.Visible = True
cda_Db.Value = "01/01/" & Year(Date)
cda_fin.Value = Format(Date, "DD/MM/YYYY")
Call Affiche_StatCarb_SBCombo(SBC_Stat)
SBC_Stat.ListIndex = 0
Call Initgrid_BonVidange
'charger comboBox Items
Cmb_typeBV.AddItem "Tout Bons "
Cmb_typeBV.AddItem "BV Facturé"
Cmb_typeBV.AddItem "BV Non Facturé"
Cmb_typeBV.AddItem "BV Supprimé"
Cmb_typeBV.ListIndex = 2
End Sub

Private Sub init_searchBC()

Pict_PieceRep.Visible = True
DateBx_Db.Text = "01/01/" & Year(Date)
DateBox_fin.Text = Format(Date, "DD/MM/YYYY")
Call Affiche_Station_SBCombo(Cmb_station)
'charger comboBox Items
Cb_typBC.AddItem "Toute pièces "
Cb_typBC.AddItem "BC transféré"
Cb_typBC.AddItem "BC Non transféré"
Cb_typBC.ListIndex = 2
End Sub

'Afficher la forme selon StrSource : quelle liste doit être afficher
Private Sub Form_Load()

On Error GoTo Err

LBL_Titre.Caption = "Liste des " & StrSource & "s"

Select Case StrSource

    Case "Véhicule"
        Pict_PieceRep.Visible = False
        Pict_BonVdg.Visible = False
        Call Initgrid_Vehicule
        Call Affiche_Vehicule
        
    Case "Station"
        Pict_PieceRep.Visible = False
        Pict_BonVdg.Visible = False
        Call Initgrid_Fournisseur
        Call Affiche_Station
        LBL_Titre.Caption = "Liste des fournisseurs"
        
    Case "Produits"
        Pict_PieceRep.Visible = False
        Pict_BonVdg.Visible = False
        Call Initgrid_Produits
        Call Affiche_Produits
        
    Case "Energie"
        Pict_PieceRep.Visible = False
        Pict_BonVdg.Visible = False
        Call Initgrid_Energie
        Call Affiche_Energie
        
    Case "Personnel"
        Pict_PieceRep.Visible = False
        Pict_BonVdg.Visible = False
        Call Initgrid_Personnel
        Call Affiche_Personnel
        
'    Case "Utilisateur"
'        Pict_PieceRep.Visible = False
'        Pict_BonVdg.Visible = False
'        Call Initgrid_Personnel
'        Call Affiche_Utilisateur
        
    Case "BonVidange2"
    LBL_Titre.Caption = "Liste des Bons de Vidange"
        Pict_PieceRep.Visible = False
        Call init_searchBV
        Call Cmd_FindBV_Click
        'Call Affiche_BonVidange2
        
    Case "FactureCarburant"
    LBL_Titre.Caption = "Liste des factures"
        Pict_PieceRep.Visible = False
        Pict_BonVdg.Visible = False
        Call Initgrid_Facture
        Call Affiche_Facture
        
    Case "Reparation", "FIndBCReparation"
    LBL_Titre.Caption = "Liste des bons de commande de Reparation"
        Pict_PieceRep.Visible = True
        Pict_BonVdg.Visible = False
        Call init_searchBC
        Call Initgrid_Reparation
        Call CmdFind_Click
        
    Case "BLPieceReparation"
    LBL_Titre.Caption = "Liste des Pieces de Réception"
        Pict_PieceRep.Visible = True
        Pict_BonVdg.Visible = False
        Call init_searchPRep
        Call InitGrid_BLPieceReparation
        Call CmdFind_Click '("Piece Reception")
        
     Case "BRPieceReparation"
      LBL_Titre.Caption = "Liste des Bons de retour de Réception"
        Pict_PieceRep.Visible = True
        Pict_BonVdg.Visible = False
        Call init_searchPRep
        Call InitGrid_BLPieceReparation
        Call CmdFind_Click '("Bon Retour")
        
    Case "FacturePieceReparation"
     LBL_Titre.Caption = "Liste des factures de Reparation"
        Pict_PieceRep.Visible = True
        Pict_BonVdg.Visible = False
        Call init_searchPRep
        Call InitGrid_BLPieceReparation
        Call CmdFind_Click '("Facture")
         
    Case "AvoirPieceReparation"
     LBL_Titre.Caption = "Liste des avoirs de Réception"
        Pict_PieceRep.Visible = True
        Pict_BonVdg.Visible = False
        Call init_searchPRep
        Call InitGrid_BLPieceReparation
        Call CmdFind_Click '("Avoir")
      
      Case "AllPieceReparation"
     LBL_Titre.Caption = "Liste des Pièces de Réception"
        Pict_PieceRep.Visible = True
        Pict_BonVdg.Visible = False
        Call init_searchPRep
        Call InitGrid_BLPieceReparation
        Call Affiche_PieceReparation("Tout")
        
     Case "Tournee"
        Pict_PieceRep.Visible = False
        Pict_BonVdg.Visible = False
        Call Initgrid_Tournee
        Call Affiche_Tournee
        LBL_Titre.Caption = "Liste des Tournes"
    
    End Select
 
Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub grid_ColumnClick(ByVal lCol As Long)
Dim sTag As String
Dim i As Long
      
    If (StrSource = "BLPieceReparation" Or StrSource = "FacturePieceReparation") Then
   With grid.SortObject
      .Clear
      .SortColumn(1) = lCol
   
      sTag = grid.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(1) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(1) = CCLOrderDescending
      End If
      grid.ColumnTag(lCol) = sTag
   
    Select Case grid.ColumnKey(lCol)
    Case "Numero"
         .SortType(1) = CCLSortNumeric
    Case "DatePiece"
         .SortType(1) = CCLSortDate
    Case "DateOperation"
         .SortType(1) = CCLSortDate
    Case "Fournisseur"
         .SortType(1) = CCLSortString
      End Select
   End With
   Screen.MousePointer = vbHourglass
   grid.Sort
   Screen.MousePointer = vbDefault
   End If

End Sub

Private Sub grid_DblClick(ByVal lRow As Long, ByVal lCol As Long)

Dim VCode
On Error GoTo Err

 VCode = grid.CellText(lRow, 1)
Select Case StrSource
    Case "Produits"
        Unload Me
        Frm_Articles.AfficheRow (VCode)
        
    Case "Véhicule"
        Unload Me
        Frm_Vehicule.AfficheRow (VCode)
        

    Case "Energie"
        Unload Me
        FrmCarburant.AfficheRow (VCode)
        
    Case "Utilisateur"
        Unload Me
        Frm_Utilisateur.AfficheRow (VCode)
        
    Case "Personnel"
        Unload Me
        Frm_Personnel.AfficheRow (VCode)
        
    Case "BonVidange2"
        If (CHECK_ACCES("Consult_BV", LInt_UserId) = False) Then
            MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour Modifier un Bon de Vidange", vbExclamation
            Exit Sub
        End If
        Unload Me
        FrmBonVidange.AfficheRow (VCode)
        
    Case "FactureCarburant"
        Unload Me
        frmCreationFacture.AfficheRow (VCode)
        
    Case "Station"
        Unload Me
        Frm_Station.AfficheRow (VCode)
        
    Case "Reparation"
        Unload Me
        FrmBCReparation.AfficheRow (VCode)
        
    Case "FIndBCReparation"
        
        With FrmPieceReparation
        
        .AfficheRow_BCR (VCode)
        .txt_ref.Text = "BC: " & CStr(VCode)
        .cda_Create.Caption = CStr(grid.CellText(lRow, 2))
        .AfficheRow_Station (CStr(grid.CellText(lRow, 3)))
        .NumBC.Caption = CStr(grid.CellText(lRow, 1))
        .Pict_stat.Enabled = False
         End With
        Unload Me
        
    Case "BLPieceReparation", "FacturePieceReparation", "AvoirPieceReparation", "BRPieceReparation", "AllPieceReparation"
        Unload Me
        FrmPieceReparation.AfficheRow (VCode)
    End Select

Exit Sub
Err:
    MsgBox Err.Description, vbInformation

End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

Dim VCode
On Error GoTo Err
    VCode = grid.CellText(grid.SelectedRow, 1)
    Select Case KeyCode
        Case vbKeyReturn
            Select Case StrSource
                Case "Véhicule"
                    Unload Me
                    Frm_Vehicule.AfficheRow (VCode)

                Case "Energie"
                    Unload Me
                    FrmCarburant.AfficheRow (VCode)

                Case "BonVidange2"
                    Unload Me
                    FrmBonVidange.AfficheRow (VCode)
       
                Case "FactureCarburant"
                    Unload Me
                    frmCreationFacture.AfficheRow (VCode)
                    
                Case "Station"
                    Unload Me
                    Frm_Station.AfficheRow (VCode)
                
                Case "BLPieceReparation", "FacturePieceReparation", "AllPieceReparation"
                  VCode = grid.CellText(grid, 1)
                    Unload Me
                    FrmPieceReparation.AfficheRow (VCode)
                
                Case "FIndBCReparation"
                    Unload Me
                    FrmPieceReparation.AfficheRow_BCR (VCode)
                    FrmPieceReparation.AfficheRow_Station (grid.CellText(grid.SelectedRow, 3))
                
               
        End Select
         Case vbKeyEscape
                      Unload Me
    End Select
  
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub



Private Sub SBC_Stat_LostFocus()
Call ExistDonnee(SBC_Stat)
End Sub

Private Sub SCommand1_Click()
Unload Me
End Sub

Public Sub Initgrid_Vehicule()
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
    
    .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Libelle", "Matricule", , , 140
    .AddColumn "Marque", "Marque", , , 40
    .AddColumn "Type", "Type", eSortType:=CCLSortStringNoCase
    .AddColumn "Energie", "Energie", eSortType:=CCLSortStringNoCase
    .AddColumn "Puissance", "Puissance", sFmtString:="short date", eSortType:=CCLSortDateDayAccuracy
    .AddColumn "Actif", "Acif", , , 40
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub
Public Sub Affiche_Vehicule()

Dim Lobj_Vehicule As VEHICULE
Dim rs As New Recordset

Set Lobj_Vehicule = New VEHICULE
Set rs = Lobj_Vehicule.Get_AllVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, 1, rs("Code")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Matricule")
            .CellDetails .Rows, .ColumnIndex("Marque"), rs("Marque")
            .CellDetails .Rows, .ColumnIndex("Type"), rs("Type")
            .CellDetails .Rows, .ColumnIndex("Energie"), rs("Energie")
            .CellDetails .Rows, .ColumnIndex("Puissance"), rs("Puissance")
             .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If

End Sub

Private Sub Initgrid_Energie()
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
    
    .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Libelle", "Libelle", , , 200
    .AddColumn "Prix", "Prix.TTC", , , 120
'    .AddColumn "Actif", "Actif", , , 40
  
    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub

Public Sub Affiche_Energie()

Dim LOBJ_Energie As Energie
Dim rs As New Recordset

Set LOBJ_Energie = New Energie
Set rs = LOBJ_Energie.Get_Energ(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, 1, rs("Code")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle")
            .CellDetails .Rows, .ColumnIndex("Prix"), Format(rs("Prix"), "#,##0.000")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

Private Sub Initgrid_Produits()
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
    
    .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Type", "Type", , , 70
    .AddColumn "Libelle", "Libelle", , , 200
    .AddColumn "Prixht", "Prix.HT", , , 120
    .AddColumn "tva", "TVA", , , 120
    .AddColumn "Actif", "Actif", , , 40
    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub

'Affiche liste de tout les Articles
Public Sub Affiche_Produits()

Dim LOBJ_ProdRep As Produit_Lubrifiant
Dim rs As New Recordset

Set LOBJ_ProdRep = New Produit_Lubrifiant
Set rs = LOBJ_ProdRep.Get_AllArticles(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, .ColumnIndex("Type"), rs("Type_PL")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle")
            .CellDetails .Rows, .ColumnIndex("Prixht"), Format(rs("Prixht"), "#,##0.000")
            .CellDetails .Rows, .ColumnIndex("tva"), Format(rs("tva"), "#,##0.00")
            .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

Private Sub Initgrid_Fournisseur()
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
    
    .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Libelle", "Libelle", , , 140
    .AddColumn "Type", "Type", , , 100
    .AddColumn "Activité", "Activité", , , 140
    .AddColumn "Adresse", "Adresse", , , , 140
    .AddColumn "Actif", "Actif", , , , 40
    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
.Redraw = True
End With

End Sub

Public Sub Affiche_Station()

Dim LOBJ_Stat As Station
Dim rs As Recordset

Set LOBJ_Stat = New Station
Set rs = LOBJ_Stat.Get_AllStat(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, 1, rs("Code")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle")
            .CellDetails .Rows, .ColumnIndex("Type"), rs("Type")
            .CellDetails .Rows, .ColumnIndex("Activité"), rs("Activite")
            .CellDetails .Rows, .ColumnIndex("Adresse"), rs("Adresse")
            .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")

        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

Private Sub Initgrid_Facture()

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
    .AddColumn "Date", "Date", , , 40
    .AddColumn "Station", "Station", , , 120
    .AddColumn "Période du", "Période du", , , 100
    .AddColumn "Période au", "Période au", , , 100
    .AddColumn "TTC", "TTC"
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
.Redraw = True
End With

End Sub

Public Sub Affiche_Facture()

Dim LOBJ_Fact As Facture
Dim rs As New Recordset

Set LOBJ_Fact = New Facture
Set rs = LOBJ_Fact.Get_FactFind(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, .ColumnIndex("Période du"), rs("PeriodeDu")
            .CellDetails .Rows, .ColumnIndex("Période au"), rs("Periodeau"), DT_RIGHT
            .CellDetails .Rows, .ColumnIndex("TTC"), Format(rs("TTC"), "#,##0.000"), DT_RIGHT
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
    grid.SelectedRow = grid.Rows
End If
rs.Close

End Sub

Private Sub Initgrid_Reparation()
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
    
    .AddColumn "Numero", "Numero", , , 60, , , , , , , CCLSortNumeric
    .AddColumn "DateCreation", "Date Creation", , , 120
    .AddColumn "CoDFournissseur", "CoDFournissseur", , , 0
    .AddColumn "Fournisseur", "Fournisseur", , , 120
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
.Redraw = True
End With

End Sub

Public Sub Affiche_Reparation(ByVal DateDu As Date, ByVal DateAu As Date, ByVal Stat As String, ByVal vtransf As String)

Dim LOBJ_Repa As BCReparation
Dim rs As New Recordset

grid.ClearRows
Set LOBJ_Repa = New BCReparation
Set rs = LOBJ_Repa.Get_ReparNTrans(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DateDu, DateAu, Stat, vtransf)
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
            .CellDetails .Rows, .ColumnIndex("DateCreation"), rs("DateCreation")
            .CellDetails .Rows, .ColumnIndex("CoDFournissseur"), rs("CoDFournissseur")
            .CellDetails .Rows, .ColumnIndex("Fournisseur"), rs("Fournisseur")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
    grid.SelectedRow = grid.Rows
    
End If
rs.Close

End Sub

Private Sub Initgrid_BonVidange()
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
    .AddColumn "Date", "Date", , , 80
    .AddColumn "Station", "Station", , , 110
    .AddColumn "Conducteur", "Conducteur", , , 110
    .AddColumn "Vehicule", "Vehicule", , , 100
    .AddColumn "Valeur", "Valeur", , , 70
    .AddColumn "Vidange", "Vidange", , , 70
    .AddColumn "Supp", "Supp", , , 60
    .AddColumn "Q", "", , , 5
    .StretchLastColumnToFit = True
.Redraw = True
End With

End Sub

Public Sub Affiche_BonVdg(ByVal DateDu As Date, ByVal DateAu As Date, ByVal Station As String, ByVal TYP As String)

Dim LOBJ_BonViddange As BonVidange
Dim rs As New Recordset

Set LOBJ_BonViddange = New BonVidange
Set rs = LOBJ_BonViddange.Get_BVAfich(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DateDu, DateAu, Station, TYP)
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
            .CellDetails .Rows, .ColumnIndex("Station"), rs("Station")
            .CellDetails .Rows, .ColumnIndex("Conducteur"), rs("Conducteur")
            .CellDetails .Rows, .ColumnIndex("Vehicule"), rs("Matricule")
            .CellDetails .Rows, .ColumnIndex("Valeur"), Format(rs("VALEUR"), "#,##0.000"), DT_RIGHT
            If Val(rs("VALEUR")) = 0 Then
                .CellDetails .Rows, .ColumnIndex("Vidange"), "Simple"
            Else
                .CellDetails .Rows, .ColumnIndex("Vidange"), "avec filtre"
            End If
            .CellDetails .Rows, .ColumnIndex("Supp"), rs("Supp")
  
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
    grid.SelectedRow = grid.Rows
End If
rs.Close

End Sub

Private Sub Initgrid_Personnel()

With grid

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
    
    .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Libelle", "Nom et prénom", , , 140
    .AddColumn "Actif", "Actif", , , 40
    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With

End Sub

Public Sub Affiche_Personnel()

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_AllPers(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, 1, rs("Code")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle")
            .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

Public Sub Affiche_Utilisateur()

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_AllUsers(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, 1, rs("Code")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("NOMPRN")
            .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

Public Sub InitGrid_BLPieceReparation()
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
    
    .AddColumn "Numero", "Numero", , , 60, , , , , , , CCLSortNumeric
    .AddColumn "DatePiece", "DatePiece", , , 120
    .AddColumn "dateOperation", "dateOperation", , , 120
    .AddColumn "Fournisseur", "Fournisseur", , , 120
    .AddColumn "Supp", "Supp", , , 80
    .AddColumn "Facturé", "Facturé", , , 60
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
.Redraw = True
End With
End Sub

Public Sub Affiche_PieceReparation(ByVal typPiece As String)

Dim LOBJ_PieceRepa As PieceReparation
Dim rs As New Recordset

Set LOBJ_PieceRepa = New PieceReparation
Set rs = LOBJ_PieceRepa.Get_PieceReparByTyp(ErrNumber, ErrDescription, ErrSourceDetail, CNB, typPiece)
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
            .CellDetails .Rows, .ColumnIndex("DatePiece"), rs("DatePiece")
            .CellDetails .Rows, .ColumnIndex("dateOperation"), rs("DateOperation")
            .CellDetails .Rows, .ColumnIndex("Fournisseur"), rs("Fournisseur")
            .CellDetails .Rows, .ColumnIndex("Supp"), rs("Supp")
            .CellDetails .Rows, .ColumnIndex("Facturé"), rs("transf")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
    grid.SelectedRow = grid.Rows
End If
End Sub

Public Sub Affiche_PieceReparRech(ByVal DateDu As Date, ByVal DateAu As Date, ByVal Station As String, ByVal TYP As String, ByVal typPiece As String)

Dim LOBJ_PieceRepa As PieceReparation
Dim rs As New Recordset
grid.ClearRows

Set LOBJ_PieceRepa = New PieceReparation
Set rs = LOBJ_PieceRepa.Get_PieceReparation(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DateDu, DateAu, Station, TYP, typPiece)
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
            .CellDetails .Rows, .ColumnIndex("DatePiece"), rs("DatePiece")
            .CellDetails .Rows, .ColumnIndex("dateOperation"), rs("DateOperation")
            .CellDetails .Rows, .ColumnIndex("Fournisseur"), rs("Fournisseur")
            .CellDetails .Rows, .ColumnIndex("Supp"), rs("Supp")
            .CellDetails .Rows, .ColumnIndex("Facturé"), rs("transf")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
    grid.SelectedRow = grid.Rows
End If
End Sub

Private Sub Initgrid_Tournee()
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
    
    .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Libelle", "Libelle", , , 200
    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
    .Redraw = True
End With
End Sub

Public Sub Affiche_Tournee()

Dim Lobj_Dest As DESTINATION
Dim rs As New Recordset

Set Lobj_Dest = New DESTINATION
Set rs = Lobj_Dest.Get_Tournee(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, 1, rs("Code")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub



