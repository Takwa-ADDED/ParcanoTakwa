VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmFind_Actif 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SToolBox.SCommand SCommand1 
      Height          =   255
      Left            =   7200
      TabIndex        =   0
      Top             =   0
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
      Left            =   6960
      TabIndex        =   1
      Top             =   0
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
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9763
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
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3330
   End
   Begin VB.Image PicBox_Header 
      Height          =   975
      Left            =   0
      Picture         =   "FrmFind_Actif.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "FrmFind_Actif"
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
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Err
LBL_Titre.Caption = "Liste des " & StrSource & "s"

Select Case StrSource

 Case "Destination"
        Call Initgrid_Destination
        Call Affiche_Destination
Case "Véhicule"
        Call Initgrid_Vehicule
        Call Affiche_Vehicule
Case "Station"
        Call Initgrid_Fournisseur
        Call Affiche_Station

 Case "Utilisateur"
        Call Initgrid_Personnel
        Call Affiche_Utilisateur
Case "Personnel"
        Call Initgrid_Personnel
        Call Affiche_Personnel
  Case "Produits"
        Call Initgrid_Produits
        Call Affiche_Produits
'    Case "ConducteurPH"
'        Call Initgrid_Personnel
'        Call Affiche_Personnel
    Case "ConducteurSup"
       Call Initgrid_Personnel
       Call Affiche_Personnel
        
End Select

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub grid_DblClick(ByVal lRow As Long, ByVal lCol As Long)

Dim VCode
On Error GoTo Err

 VCode = grid.CellText(lRow, 1)
Select Case StrSource
     Case "Destination"
        VCode = grid.CellText(lRow, 1)
        Unload Me
        Call Frm_Destination.AfficheRow(VCode)
    Case "Véhicule"
        Unload Me
        Frm_Vehicule.AfficheRow (VCode)
    Case "Station"
        Unload Me
        Frm_Station.AfficheRow (VCode)

    Case "Utilisateur"
        Unload Me
        Frm_Utilisateur.AfficheRow (VCode)
    Case "Personnel"
        Unload Me
        Frm_Personnel.AfficheRow (VCode)
     Case "Produits"
        Unload Me
        Frm_Articles.AfficheRow (VCode)
    Case "ConducteurPH"
'        Unload Me
'        Frm_PrgChauf.AfficheRowconducteurPH (VCode)
    Case "ConducteurSup"
        Unload Me
        Frm_Supervision.AfficheRowconducteurSup (VCode)
End Select

Exit Sub
Err:
MsgBox Err.Description, vbInformation


End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
Dim lRow

Dim VCode
On Error GoTo Err
    VCode = grid.CellText(grid.SelectedRow, 1)
    Select Case KeyCode
        Case vbKeyReturn
            Select Case StrSource
                Case "Destination"
                    VCode = grid.CellText(lRow, 1)
                    Unload Me
                    Call Frm_Destination.AfficheRow(VCode)
                Case "Véhicule"
                    Unload Me
                    Frm_Vehicule.AfficheRow (VCode)
                Case "Station"
                    Unload Me
                    Frm_Station.AfficheRow (VCode)

                Case "Utilisateur"
                    Unload Me
                    Frm_Utilisateur.AfficheRow (VCode)
                Case "Personnel"
                    Unload Me
                    Frm_Personnel.AfficheRow (VCode)
                Case "Produits"
                    Unload Me
                    Frm_Articles.AfficheRow (VCode)
                Case "ConducteurPH"
'                    Unload Me
'                    Frm_PrgChauf.AfficheRowconducteurPH (VCode)
                Case "ConducteurSup"
                    Unload Me
                    Frm_Supervision.AfficheRowconducteurSup (VCode)
                 
            End Select
        Case vbKeyEscape
            Unload Me
    End Select

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub SCommand1_Click()
Unload Me
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

Public Sub Affiche_Destination()

Dim Lobj_Dest As DESTINATION
Dim rs As New Recordset

Set Lobj_Dest = New DESTINATION
Set rs = Lobj_Dest.Get_ActifDest(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
    
    .AddColumn "Code", "Code", , , 60, , , , , , , CCLSortNumeric
    .AddColumn "Libelle", "Matricule", , , 140
    .AddColumn "Marque", "Marque", , , 40
    .AddColumn "Type", "Type", eSortType:=CCLSortStringNoCase
    .AddColumn "Energie", "Energie", eSortType:=CCLSortStringNoCase
    .AddColumn "Puissance", "Puissance", sFmtString:="short date", eSortType:=CCLSortDateDayAccuracy
    .AddColumn "Actif", "Actif", , , 40

    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub

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

Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station

Set rs = LOBJ_Station.Get_ActifStat(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
    .AddColumn "Actif", "Actif", , , 40
  
    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
    .Redraw = True
End With
End Sub

'Public Sub Affiche_LubActif()
'
'Dim LOBJ_Lub As Lubrifiant
'Dim rs As New Recordset
'
'Set LOBJ_Lub = New Lubrifiant
'Set rs = LOBJ_Lub.Get_LubActif(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
'If ErrNumber <> 0 Then
'    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
'    ErrNumber = 0
'    Exit Sub
'End If
'
'If Not rs.EOF Then
'    grid.Redraw = False
'    While Not rs.EOF
'        With grid
'            .AddRow
'            .CellDetails .Rows, 1, rs("Code")
'            .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle")
'            .CellDetails .Rows, .ColumnIndex("Prix"), Format(rs("Prix"), "#,##0.000")
'            .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")
'        End With
'        rs.MoveNext
'    Wend
'    grid.Redraw = True
'End If
'End Sub

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
    If StrSource <> "ConducteurPH" Then .AddColumn "Actif", "Actif", , , 40

    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With

End Sub

Public Sub Affiche_Utilisateur()

Dim LOBJ_Pers As personnel
Dim rs As New Recordset

Set LOBJ_Pers = New personnel
Set rs = LOBJ_Pers.Get_AllActifUsers(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
    rs.Close
    grid.Redraw = True
End If
End Sub

Public Sub Affiche_Personnel()

Dim LOBJ_Pers As personnel
Dim rs As New Recordset

Set LOBJ_Pers = New personnel
Set rs = LOBJ_Pers.Get_AllActifPers(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
            If StrSource <> "ConducteurPH" Then .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")
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
  
    
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub


Public Sub Affiche_Produits()

Dim LOBJ_Prod As Produit_Lubrifiant
Dim rs As New Recordset

Set LOBJ_Prod = New Produit_Lubrifiant
Set rs = LOBJ_Prod.Get_ActifArticles(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
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
        End With
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
End Sub
