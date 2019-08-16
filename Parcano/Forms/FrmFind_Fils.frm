VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmFind_Fils 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
   ClipControls    =   0   'False
   Icon            =   "FrmFind_Fils.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pict_typ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   6975
      TabIndex        =   5
      Top             =   1200
      Width           =   6975
      Begin VB.OptionButton Op_Prod 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Produits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   0
         Width           =   1695
      End
      Begin VB.OptionButton Op_Lub 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lubrifiant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Lab_Typ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   8
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.TextBox txt_Libelle 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   6120
      Width           =   6615
   End
   Begin SToolBox.SGrid grid 
      Height          =   4455
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
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
   Begin SToolBox.SCommand SCommand1 
      Height          =   255
      Left            =   7080
      TabIndex        =   2
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
      Left            =   6840
      TabIndex        =   3
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
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   3045
   End
   Begin VB.Image PicBox_Header 
      Height          =   1215
      Left            =   0
      Picture         =   "FrmFind_Fils.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "FrmFind_Fils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StrSource As String
Public RETOUR As Integer

Private Sub Form_Activate()
If grid.Rows = 0 Then MsgBox "Pas de données à visualiser", vbInformation
 If grid.Rows > 0 Then
    grid.SelectedRow = 1
    grid.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight Then Unload Me
End Sub

'Charger la liste suivant la variable StrSource lors d'ouverture de la forme
Private Sub Form_Load()

On Error GoTo Err

LBL_Titre.Caption = "Liste des " & StrSource & "s"
Select Case StrSource
    Case "Véhicule", "VéhiculeVidange", "VéhiculeReparation", _
    "Véhicule Detail BC Reparation", "Véhicule Detail Piece Reparation", _
            "Véhicule Stat carburant", "Véhicule Stat reparation", _
            "VéhiculeVidange2", "Véhicule Stat Trafic"
        Pict_typ.Visible = False
        Call Initgrid_Vehicule
        Call Affiche_Vehicule
        
        Case "Station", "Stationfacture"
        Pict_typ.Visible = False
        Call Initgrid_Fournisseur
        Call Affiche_Station
        
    Case "Station carburant", "StationVidange"
        Pict_typ.Visible = False
        Call Initgrid_Fournisseur
        Call Affiche_Station_Carburant
     
    Case "Lubrifiant", "LubrifiantVidange2"
        Pict_typ.Visible = True
        Call Initgrid_Lubrifiant
        Call Affiche_Lubrifiant
     
    Case "Energie"
        Pict_typ.Visible = False
        Call Initgrid_Energie
        Call Affiche_Energie
        
    Case "Personnel", "PersonnelVidange2", "PersonnelReparation", "Personnel Stat", "Personnel E/H", "PersoConge", "Personnel_Conducteur"
        Pict_typ.Visible = False
        Call Initgrid_Personnel
        Call Affiche_Personnel
        
    Case "Fournisseur", "Station PR", "Station BCR"
        Pict_typ.Visible = False
        Call Initgrid_Fournisseur
        Call Affiche_Fournisseur
        
    Case "FournisseurReparation", "Fournisseur Reparation"
        Pict_typ.Visible = False
        Call Initgrid_Fournisseur
        Call Affiche_Fournisseur
        
    Case "Type Reparation", "Detail Reparation", "Piece Reparation"
        Pict_typ.Visible = True
        Call Initgrid_TypeReparation
        'Call Affiche_ProdReparation
        Op_Prod.Value = True
        
    Case "searchPieceRepar"
        Pict_typ.Visible = True
        Call Initgrid_TypeReparation
        Call Affiche_ProdByLibelle(FrmSaisiePieceReparation.Txt_Designation.Text)
        
    Case "searchArticle"
        Pict_typ.Visible = True
        Call Initgrid_TypeReparation
        Call Affiche_ProdByLibelle(Frm_Articles.txt_Matricule.Text)
        
        
    Case "searchBCRepar"
        Pict_typ.Visible = True
        Call Initgrid_TypeReparation
        Call Affiche_ProdByLibelle(frmDetailBCReparation.txt_libelle.Text)
            
    Case "Destination"
        Pict_typ.Visible = False
        Call Initgrid_Destination
        Call Affiche_Destination
        
    Case "Destination E/H"
        Pict_typ.Visible = False
        Call Initgrid_Destination
        Call Affiche_DestinationEH
        
End Select

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

' afficher list resultat (recordset) d'une requete
Public Sub list_PRepar(ByVal rs As Recordset)

If Not rs.EOF Then
    rs.MoveFirst
    grid.Redraw = False
    While Not rs.EOF
        With grid
            .AddRow
            .CellDetails .Rows, 1, rs("Numero")
            .CellDetails .Rows, .ColumnIndex("Type"), rs("Type")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Libelle")
            .CellDetails .Rows, .ColumnIndex("Prixht"), Format(rs("Prixht"), "##0.000")
            .CellDetails .Rows, .ColumnIndex("tva"), Format(rs("tva"), "##0.000")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
    grid.SelectedRow = 1
End If
End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub grid_DblClick(ByVal lRow As Long, ByVal lCol As Long)

Dim VCode
Dim vlib
On Error GoTo Err

If grid.Rows = 0 Then Exit Sub

Select Case StrSource
   
    Case "VéhiculeVidange2"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        FrmBonVidange.AfficheRow_Vehicule_sansPrix (VCode)
        
        
    Case "Véhicule"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        FrmSaisieBoncarburant.AfficheRow_Vehicule (VCode)
        FrmSaisieBoncarburant.txt_Ncompteur.SetFocus
        
    Case "Véhicule Stat carburant"
         VCode = grid.CellText(lRow, 1)
        Unload Me
        Frm_Statistiques.AfficheRow_Vehicule (VCode)
    Case "Véhicule Stat reparation"
         VCode = grid.CellText(lRow, 1)
        Unload Me
        Frm_Statistiques.AfficheRow_Vehicule (VCode)
    Case "Véhicule Stat Trafic"
         VCode = grid.CellText(lRow, 1)
        Unload Me
        Frm_Statistiques.AfficheRow_Vehicule (VCode)
    Case "Personnel Stat"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        Frm_Statistiques.AfficheRow_Conducteur (VCode)
    Case "Personnel E/H"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        Frm_Statistiques.AfficheRow_Conducteur (VCode)
    Case "Destination E/H"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        Frm_Statistiques.AfficheRow_Destination (VCode)
         
        
    Case "StationVidange"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        FrmBonVidange.AfficheRow_Station (VCode)
        FrmBonVidange.txt_Ncompteur.SetFocus

    Case "Station carburant"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        FrmAllBonCarburant.AfficheRow_Station (VCode)
        
    Case "Station PR"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        FrmPieceReparation.AfficheRow_Station (VCode)
        
     Case "Station BCR"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        FrmBCReparation.AfficheRow_Station (VCode)
     
    
    Case "Stationfacture"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        frmCreationFacture.AfficheRow_Station (VCode)
   
'    Case "LubrifiantVidange2"
'        vcode = grid.CellText(lRow, 3)
'        Unload Me
'        FrmBonVidange.AfficheRow_Lubrifiant (vcode)
        
    Case "Energie"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        Frm_Vehicule.Cbo_Energie.Text = grid.CellText(lRow, 2)
        
   Case "PersonnelVidange2"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        Call FrmBonVidange.AfficheRow_Conducteur(VCode)
        FrmBonVidange.cbo_MatriculeStation.SetFocus
        
    Case "PersonnelReparation"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        Call FrmBCReparation.AfficheRow_Conducteur(VCode)
    
    Case "Personnel"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        Frm_Personnel.AfficheRow (VCode)
'        FrmAllBonCarburant.cbo_conducteur.Text = grid.CellText(lRow, 2)

Case "Personnel_Conducteur"
        VCode = grid.CellText(lRow, 2)
        Unload Me
        FrmAllBonCarburant.cbo_conducteur.Text = VCode
        
   Case "PersoConge"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        Frm_GestionConge.cbo_conducteur.Text = grid.CellText(lRow, 2)
        
     Case "Detail Reparation", "searchBCRepar"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        frmDetailBCReparation.txt_Numero.Text = grid.CellText(lRow, 1)
        frmDetailBCReparation.txt_libelle.Text = grid.CellText(lRow, 3)
        
    Case "Piece Reparation", "searchPieceRepar"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        FrmSaisiePieceReparation.txt_Numero.Text = grid.CellText(lRow, 1)
        FrmSaisiePieceReparation.Txt_Designation.Text = grid.CellText(lRow, 3)
        FrmSaisiePieceReparation.txt_PUHT.Text = grid.CellText(lRow, 4)
        FrmSaisiePieceReparation.Txt_tva.Text = grid.CellText(lRow, 5)
        
    Case "searchArticle"
         VCode = grid.CellText(lRow, 1)
        Unload Me
        Frm_Articles.txt_Matricule.Text = grid.CellText(lRow, 1)
        Frm_Articles.txt_libelle.Text = grid.CellText(lRow, 3)
        Frm_Articles.txt_tht.Text = grid.CellText(lRow, 4)
        Frm_Articles.Txt_tva.Text = grid.CellText(lRow, 5)
        Frm_Articles.txt_prix.Text = Format(grid.CellText(lRow, 4) + (grid.CellText(lRow, 4) * grid.CellText(lRow, 5) / 100), "##0.000")
        
            If grid.CellText(lRow, 2) = "Produit" Then
                Frm_Articles.Op_Prod.Value = True
            ElseIf grid.CellText(lRow, 2) = "Lubrifiant" Then
                Frm_Articles.Op_Lub.Value = True
            End If
            
            If grid.CellText(lRow, 6) = "O" Then
                Frm_Articles.chk_Actif.Value = 1
            ElseIf grid.CellText(lRow, 6) = "N" Then
                Frm_Articles.chk_Actif.Value = 0
            End If

            
    
    
    Case "Fournisseur Reparation"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        FrmPieceReparation.AfficheRow_Station (VCode)
        
    Case "Véhicule Detail BC Reparation"
          VCode = grid.CellText(lRow, 1)
        Unload Me
        frmDetailBCReparation.AfficheRow_Vehicule (VCode)
        
    Case "Véhicule Detail Piece Reparation"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        FrmSaisiePieceReparation.AfficheRow_Vehicule (VCode)
        
    Case "Destination"
        VCode = grid.CellText(lRow, 1)
        Unload Me
         Call Frm_Destination.AfficheRow(VCode)
End Select

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

Dim VCode
Dim lRow
On Error GoTo Err
Select Case KeyCode
Case vbKeyReturn
    If grid.Rows <> 0 Then lRow = grid.SelectedRow
    Select Case StrSource
        Case "Véhicule"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            FrmSaisieBoncarburant.AfficheRow_Vehicule (VCode)
            
        Case "VéhiculeReparation"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            'frmReparation.AfficheRow_Vehicule (vcode)
            
        Case "Véhicule Stat carburant"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_Statistiques.AfficheRow_Vehicule (VCode)
            
         Case "Véhicule Stat reparation"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_Statistiques.AfficheRow_Vehicule (VCode)
            
        Case "Station carburant"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            FrmAllBonCarburant.AfficheRow_Station (VCode)
       
    '    Case "LubrifiantVidange2"
    '        VCode = grid.CellText(lRow, 3)
    '        Unload Me
    '        FrmBonVidange.AfficheRow_Lubrifiant (VCode)
    
    
        Case "searchArticle"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_Articles.txt_Matricule.Text = grid.CellText(lRow, 1)
            Frm_Articles.txt_libelle.Text = grid.CellText(lRow, 3)
            Frm_Articles.txt_tht.Text = grid.CellText(lRow, 4)
            Frm_Articles.Txt_tva.Text = grid.CellText(lRow, 5)
            Frm_Articles.txt_prix.Text = Format(grid.CellText(lRow, 4) + (grid.CellText(lRow, 4) * grid.CellText(lRow, 5) / 100), "##0.000")
            
                If grid.CellText(lRow, 2) = "Produit" Then
                    Frm_Articles.Op_Prod.Value = True
                ElseIf grid.CellText(lRow, 2) = "Lubrifiant" Then
                    Frm_Articles.Op_Lub.Value = True
                End If
                
                If grid.CellText(lRow, 6) = "O" Then
                    Frm_Articles.chk_Actif.Value = 1
                ElseIf grid.CellText(lRow, 6) = "N" Then
                    Frm_Articles.chk_Actif.Value = 0
                End If
                
            
            
            
        Case "Lubrifiant"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            
        Case "PersoConge"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_GestionConge.cbo_conducteur.Text = grid.CellText(lRow, 2)
            
        Case "Energie"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_Vehicule.Cbo_Energie.Text = grid.CellText(lRow, 2)
    
        Case "Personnel"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_Personnel.AfficheRow (VCode)
    '        FrmAllBonCarburant.cbo_Conducteur.Text = grid.CellText(lRow, 2)
            
        Case "Detail Reparation"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            frmDetailBCReparation.txt_Numero.Text = grid.CellText(lRow, 1)
            frmDetailBCReparation.txt_libelle.Text = grid.CellText(lRow, 3)
        
        Case "Piece Reparation"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            FrmSaisiePieceReparation.txt_Numero.Text = grid.CellText(lRow, 1)
            FrmSaisiePieceReparation.Txt_Designation.Text = grid.CellText(lRow, 3)
            FrmSaisiePieceReparation.txt_PUHT.Text = grid.CellText(lRow, 4)
            FrmSaisiePieceReparation.Txt_tva.Text = grid.CellText(lRow, 5)
            
        Case "Fournisseur Reparation"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            FrmPieceReparation.AfficheRow_Station (VCode)
            
        Case "Véhicule Detail BC Reparation"
              VCode = grid.CellText(lRow, 1)
            Unload Me
            frmDetailBCReparation.AfficheRow_Vehicule (VCode)
            
        Case "Véhicule Detail Piece Reparation"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            FrmSaisiePieceReparation.AfficheRow_Vehicule (VCode)
            
        Case "Destination"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            Call Frm_Destination.AfficheRow(VCode)
            
        Case "Station PR"
           VCode = grid.CellText(lRow, 1)
           Unload Me
           FrmPieceReparation.AfficheRow_Station (VCode)
           
        Case "Station BCR"
           VCode = grid.CellText(lRow, 1)
           Unload Me
           FrmBCReparation.AfficheRow_Station (VCode)
        
        Case "Véhicule Stat carburant"
             VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_Statistiques.AfficheRow_Vehicule (VCode)
            
        Case "Véhicule Stat reparation"
             VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_Statistiques.AfficheRow_Vehicule (VCode)
            
        Case "Véhicule Stat Trafic"
             VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_Statistiques.AfficheRow_Vehicule (VCode)
            
        Case "Personnel Stat"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_Statistiques.AfficheRow_Conducteur (VCode)
            
        Case "Personnel E/H"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_Statistiques.AfficheRow_Conducteur (VCode)
            
        Case "Destination E/H"
            VCode = grid.CellText(lRow, 1)
            Unload Me
            Frm_Statistiques.AfficheRow_Destination (VCode)
    End Select
Case vbKeyEscape
    Unload Me
End Select
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub



Private Sub Op_Lub_Click()

grid.ClearRows
Call Affiche_LubReparation

End Sub

Private Sub Op_Prod_Click()

grid.ClearRows
Call Affiche_ProdReparation

End Sub

Private Sub Op_Prod_Validate(Cancel As Boolean)
Dim i
For i = 1 To grid.Rows
    If grid.CellText(i, 2) = "Produit" Then
        grid.RowVisible(i) = True
        grid.Redraw = True
    Else
        grid.RowVisible(i) = False
        grid.Redraw = True
    End If
Next
End Sub

Private Sub SCommand1_Click()
Unload Me
End Sub

Private Sub Initgrid_Vehicule()
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
    .AddColumn "Libelle", "Matricule", , , 140
    .AddColumn "Marque", "Marque", , , 40
    .AddColumn "Type", "Type", eSortType:=CCLSortStringNoCase
    .AddColumn "Energie", "Energie", eSortType:=CCLSortStringNoCase
    .AddColumn "Puissance", "Puissance", sFmtString:="short date", eSortType:=CCLSortDateDayAccuracy

    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With

End Sub

'Afficher liste des véhicule actifs
Public Sub Affiche_Vehicule()

Dim Lobj_Vehicule As VEHICULE
Dim rs As New Recordset

Set Lobj_Vehicule = New VEHICULE
Set rs = Lobj_Vehicule.GetAllActifVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

'Afficher la liste des stations de type fournisseur dans grid
Public Sub Affiche_Fournisseur()

Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station

If StrSource = "Fournisseur Reparation" Then
    Set rs = LOBJ_Station.GetStat_Fournisseur(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
Else
    Set rs = LOBJ_Station.Get_ActifStatRep(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
End If

If Not rs.EOF Then
    grid.Redraw = False
    While Not rs.EOF
        With grid
            .AddRow
            .CellDetails .Rows, 1, rs("Code")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle")
            .CellDetails .Rows, .ColumnIndex("Activité"), rs("Activite")
            .CellDetails .Rows, .ColumnIndex("Adresse"), rs("Adresse")

        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

'Afficher la liste des personnels
Public Sub Affiche_Personnel()

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_AllActifPers(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, .ColumnIndex("Code"), rs("Code")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle")
            .CellDetails .Rows, .ColumnIndex("Actif"), rs("actif")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

'Afficher liste de tout les types d'energie
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

'Afficher liste des lubrifiant
Public Sub Affiche_Lubrifiant()

Dim LOBJ_Lu As Produit_Lubrifiant
Dim rs As New Recordset

Set LOBJ_Lu = New Produit_Lubrifiant
Set rs = LOBJ_Lu.Get_Lubrifiant(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle")
            .CellDetails .Rows, .ColumnIndex("Prix"), Format(rs("prixht"), "#,##0.000")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

'Afficher liste de tout les stations de tout les types
Public Sub Affiche_Station()

Dim LOBJ_Stat As Station
Dim rs As Recordset

Set LOBJ_Stat = New Station
Set rs = LOBJ_Stat.Get_StatRep(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, .ColumnIndex("Activité"), rs("Activite")
            .CellDetails .Rows, .ColumnIndex("Adresse"), rs("Adresse")

        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

'afficher liste des stations de type station_carburant
Public Sub Affiche_Station_Carburant()

Dim LOBJ_Stat As Station
Dim rs As New Recordset

Set LOBJ_Stat = New Station
Set rs = LOBJ_Stat.Get_StationCarb(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, .ColumnIndex("Activité"), rs("Activite")
            .CellDetails .Rows, .ColumnIndex("Adresse"), rs("Adresse")

        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

'Afficher liste des produits de reparations possible
Public Sub Affiche_ProdReparation()

Dim LOBJ_ProdRep As Produit_Lubrifiant
Dim rs As New Recordset

Set LOBJ_ProdRep = New Produit_Lubrifiant
Set rs = LOBJ_ProdRep.Get_Produits(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Libelle")
            .CellDetails .Rows, .ColumnIndex("Prixht"), Format(rs("Prixht"), "##0.000")
            .CellDetails .Rows, .ColumnIndex("tva"), Format(rs("tva"), "##0.00")
            
                

        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
Set LOBJ_ProdRep = Nothing
Set rs = Nothing

End Sub

Public Sub Affiche_LubReparation()

Dim LOBJ_ProdRep As Produit_Lubrifiant
Dim rs As New Recordset

Set LOBJ_ProdRep = New Produit_Lubrifiant
Set rs = LOBJ_ProdRep.Get_LubActif(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Libelle")
            .CellDetails .Rows, .ColumnIndex("Prixht"), Format(rs("Prixht"), "##0.000")
            .CellDetails .Rows, .ColumnIndex("tva"), Format(rs("tva"), "##0.00")

        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
Set LOBJ_ProdRep = Nothing
Set rs = Nothing

End Sub
'--------------------------
Public Sub Affiche_ProdByLibelle(ByVal Libelle As String)
Dim LOBJ_ProdRep As Produit_Lubrifiant
Dim rs As Recordset


Set LOBJ_ProdRep = New Produit_Lubrifiant
Set rs = LOBJ_ProdRep.Get_ProdLubByInit(ErrNumber, ErrDescription, ErrSourceDetail, Libelle, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Set LOBJ_ProdRep = Nothing
If Not rs.EOF Then
    grid.Redraw = False
    While Not rs.EOF
        With grid
            .AddRow
            .CellDetails .Rows, 1, rs("Numero")
            .CellDetails .Rows, .ColumnIndex("Type"), rs("Type")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Libelle")
            .CellDetails .Rows, .ColumnIndex("Prixht"), Format(rs("Prixht"), "##0.000")
            .CellDetails .Rows, .ColumnIndex("tva"), Format(rs("tva"), "##0.00")
            .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")

        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
    
End If
Set rs = Nothing
End Sub

'Liste des destinations des vehicules
Public Sub Affiche_Destination()

Dim Lobj_Destination As DESTINATION
Dim rs As New Recordset

Set Lobj_Destination = New DESTINATION
Set rs = Lobj_Destination.Get_Destination(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, .ColumnIndex("Type"), rs("Type")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Libelle")
            .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

Public Sub Affiche_DestinationEH()

Dim Lobj_Destination As DESTINATION
Dim rs As New Recordset

Set Lobj_Destination = New DESTINATION
Set rs = Lobj_Destination.Get_DestTrafic(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            .CellDetails .Rows, .ColumnIndex("Type"), rs("Type")
            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Libelle")
            .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub

Private Sub Initgrid_Fournisseur()

With grid
    ' Allow the grid to be grouped, but don't show the grouping box
    .HideGroupingBox = True
    .AllowGrouping = True
    
    .GroupRowBackColor = vbWindowBackground
    .GroupRowForeColor = vbWindowText
    ' Group rows will be shown by a gradient underline
    .GridLineColor = vbWindowBackground
    .GridFillLineColor = vbWindowBackground
    .GridLines = True
    
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .DrawFocusRectangle = False
    
    .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Libelle", "Libelle", , , 140
    .AddColumn "Activité", "Activité", , , 140
    .AddColumn "Adresse", "Adresse", , , , 140
    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With

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
    
    .AddColumn "Code", "Code", , , 60, , , , , , , CCLSortNumeric
    .AddColumn "Libelle", "Nom et prénom", , , 140
    .AddColumn "Actif", "Actif", , , 140

    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With

End Sub

Private Sub Initgrid_Energie()
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
    .AddColumn "Libelle", "Libelle", , , 140
    .AddColumn "Prix", "Prix.TTC", , , 140
     
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With

End Sub

Private Sub Initgrid_Lubrifiant()
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
    .AddColumn "A", "", , , 60, False
    .AddColumn "Libelle", "Libelle", , , 180
    .AddColumn "Prix", "Prix.TTC", , , 140
    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With

End Sub

Public Sub Initgrid_TypeReparation()
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
    
    .AddColumn "Numero", "Numero", , , 40, False, , , , , , CCLSortNumeric
    .AddColumn "Type", "Type", , , 60
    .AddColumn "Libelle", "Libelle", , , 200
    .AddColumn "Prixht", "PrixHT", , , 60
    .AddColumn "tva", "TVA", , , 40
    .AddColumn "Actif", "Actif", , , 40
    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With

End Sub

Private Sub Initgrid_Destination()
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
    
    .AddColumn "Numero", "Numero", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Type", "Type", , , 140
    .AddColumn "Libelle", "Libelle", , , 140
    .AddColumn "Actif", "Actif", , , 60
  
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With

End Sub



Private Sub Txt_Libelle_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Err

If Len(Trim(txt_libelle.Text)) <> 0 Then
    Call FindArticleKeydown(txt_libelle.Text, grid)
End If

Exit Sub

Err:
 MsgBox Err.Description & vbNewLine & Err.Source, vbQuestion
End Sub

Private Sub FindArticleKeydown(ByVal vString As String, vGrid As SGrid)

Dim i As Long
For i = 1 To grid.Rows
    If Len(vString) = 0 Then
        vGrid.RowVisible(i) = True
        vGrid.Redraw = True
    Else
        If UCase(Mid(vGrid.CellText(i, 3), 1, Len(vString))) = UCase(vString) Then
            vGrid.RowVisible(i) = True
            vGrid.Redraw = True
        Else
            vGrid.RowVisible(i) = False
            vGrid.Redraw = True
        End If
    End If
Next
    
End Sub

Private Sub txt_Libelle_KeyPress(KeyAscii As Integer)

Dim rech As String
If KeyAscii <> 0 And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 26 And KeyAscii <> 127 Then
    rech = txt_libelle.Text & Chr(KeyAscii)
    If Len(Trim(rech)) <> 0 Then Call FindArticleKeydown(rech, grid)
Else
    If Trim(txt_libelle.Text) <> "" Then
        rech = Left(txt_libelle.Text, Len(txt_libelle.Text) - 1)
        Call FindArticleKeydown(rech, grid)
    End If
End If
End Sub

