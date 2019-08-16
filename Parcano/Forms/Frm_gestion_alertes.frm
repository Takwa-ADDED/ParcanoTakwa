VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form Frm_gestion_alertes 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gestion des alertes"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Gestio_alertes 
      Caption         =   "Gestion des alertes"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9975
      Begin VB.CommandButton Cmd_Print 
         Caption         =   "Imprimer"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin SToolBox.SGrid SGrid1 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   8281
         RowMode         =   -1  'True
         GridLines       =   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         ForeColor       =   0
         GridLineColor   =   0
         GridFillLineColor=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisableIcons    =   -1  'True
         MaxVisibleRows  =   0
      End
      Begin VB.Label Label1 
         Caption         =   "URGENT"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   3720
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Frm_gestion_alertes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Cmd_Print_Click()


    If SGrid1.Rows = 0 Then
        MsgBox "Pas de données à imprimer .", vbInformation
        Exit Sub
    End If

    If MsgBox("Imprimer Les Alertes   ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
            Call Frm_Rpt_Apercus.PrintAlertes(0)
            Frm_Rpt_Apercus.Show
    End If
End Sub

Private Sub Form_Load()
Call InitialGrid
Call Alerte
End Sub

Public Sub InitialGrid()
 With SGrid1
         ' Allow the grid to be grouped, but
      ' don't show the grouping box
      .HideGroupingBox = True
      .AllowGrouping = True

      ' Group rows will be shown by
      ' a gradient underline
      .GroupRowBackColor = vbWindowBackground
      .GroupRowForeColor = vbWindowText
      
      'Couleur de Grid
'      .GridLineColor = vbWindowBackground
'      .GridFillLineColor = vbWindowBackground
      .GridLines = True
      
      'Ajout des Colonnes
      .SelectionAlphaBlend = True
      .SelectionOutline = True
      .DrawFocusRectangle = False
      
      .AddColumn "Véhicule", "Véhicule", , , 200
      .AddColumn "Vidange", "Vidange", , , 60
      .AddColumn "Visite", "Visite", , , 60
      .AddColumn "Assurance", "Assurance", , , 60
      .AddColumn "Taxe", "Taxe", , , 60
      .AddColumn "CPT_FT", "CPT_FT", , , 60
      .AddColumn "CPT_BV", "CPT_BV", , , 60
      .AddColumn "CPT_BC", "CPT_BC", , , 60

      .AddColumn "A", ""
      .StretchLastColumnToFit = True

   End With


End Sub

'Public Sub Alerte()
'    Dim LObj_Find As New VEHICULE
'    Dim LObj_Find_V As New VEHICULE
'    Dim Lrs_Vehicule As Recordset
'    Dim rs As Recordset
'    Dim Vidange As Double
'
''    Set LObj_Find = New VEHICULE
'    Set Lrs_Vehicule = LObj_Find.GetAlerte_Vehicule(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
'            If ErrNumber <> 0 Then
'                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
'                ErrNumber = 0
'                Exit Sub
'            End If
'
'
'
'
'
'    If Not Lrs_Vehicule.EOF Then
'        While Not Lrs_Vehicule.EOF
'
'         Set rs = LObj_Find_V.GetLast_FTParVehicule(ErrNumber, ErrDescription, ErrSourceDetail, Lrs_Vehicule.Fields("code"), CNB)
'                If ErrNumber <> 0 Then
'                    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
'                    ErrNumber = 0
'                    Exit Sub
'                End If
'
'
'                If Not rs.EOF Then
'                    If (Lrs_Vehicule.Fields("CompteurVidange") <> 0) And (Lrs_Vehicule.Fields("CompteurVidange") <> "NULL") And (rs.Fields("CompteurEntre") <> 0) And (rs.Fields("CompteurEntre") <> "NULL") Then
'                        Vidange = (rs.Fields("CompteurEntre") - Lrs_Vehicule.Fields("CompteurVidange"))
'                    Else: Vidange = 0
'                    End If
'                End If
'
'
'        SGrid1.Redraw = False
'         If (Lrs_Vehicule.Fields("Visite") < 0) Or (Lrs_Vehicule.Fields("Assur") < 0) Or (Lrs_Vehicule.Fields("Taxe") < 0) Or (Vidange >= 10000) Then
'
'
'            With SGrid1
'                .AddRow
'
'                    .CellDetails .Rows, .ColumnIndex("Véhicule"), Lrs_Vehicule.Fields("Marque") & Lrs_Vehicule.Fields("Matricule"), DT_LEFT
'
'                    If (Vidange >= 10000) Then
'                        .CellDetails .Rows, .ColumnIndex("Vidange"), "Vidange obligatoire", DT_LEFT
'                    End If
'
'                    If Lrs_Vehicule.Fields("CompteurVidange") <> 0 Then
'                        .CellDetails .Rows, .ColumnIndex("CPT_BV"), Lrs_Vehicule.Fields("CompteurVidange"), DT_LEFT
'                        Else
'                        .CellDetails .Rows, .ColumnIndex("CPT_BV"), "Null", DT_LEFT
'                    End If
'
'                    .CellDetails .Rows, .ColumnIndex("CPT_BC"), Lrs_Vehicule.Fields("CompteurCarburant"), DT_LEFT
'
'                    If Not rs.EOF Then .CellDetails .Rows, .ColumnIndex("CPT_FT"), rs.Fields("CompteurEntre"), DT_LEFT
'
'                    If Lrs_Vehicule.Fields("Visite") < 0 Then
'                        .CellDetails .Rows, .ColumnIndex("Visite"), Lrs_Vehicule.Fields("Visite"), DT_LEFT, , &HFF&, &HFFFFFF
'                    End If
'
'                    If Lrs_Vehicule.Fields("Assur") < 0 Then
'                        .CellDetails .Rows, .ColumnIndex("Assurance"), Lrs_Vehicule.Fields("Assur"), DT_LEFT, , &HFF&, &HFFFFFF
'                    End If
'
'                    If Lrs_Vehicule.Fields("Taxe") < 0 Then
'                        .CellDetails .Rows, .ColumnIndex("Taxe"), Lrs_Vehicule.Fields("Taxe"), DT_LEFT, , &HFF&, &HFFFFFF
'                    End If
'
'
'            End With
'             End If
'        Lrs_Vehicule.MoveNext
'        Wend
'    SGrid1.Redraw = True
'    End If
'End Sub

Public Sub Alerte()
    Dim LObj_Find As New VEHICULE
    Dim Lrs_Vehicule As Recordset
    Dim Vidange As Double

'    Set LObj_Find = New VEHICULE
    Set Lrs_Vehicule = LObj_Find.Get_Alertes(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
            If ErrNumber <> 0 Then
                MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                ErrNumber = 0
                Exit Sub
            End If
    
   
            
    If Not Lrs_Vehicule.EOF Then
        While Not Lrs_Vehicule.EOF
                
                
                If (Lrs_Vehicule.Fields("CompteurVidange") <> 0) And (Lrs_Vehicule.Fields("CompteurVidange") <> "NULL") And (Lrs_Vehicule.Fields("CompteurEntre") <> 0) And (Lrs_Vehicule.Fields("CompteurEntre") <> "NULL") Then
                        Vidange = (Lrs_Vehicule.Fields("CompteurEntre") - Lrs_Vehicule.Fields("CompteurVidange"))
                    Else: Vidange = 0
                End If
                
    
        
        SGrid1.Redraw = False
         If (Lrs_Vehicule.Fields("Visite") < 0) Or (Lrs_Vehicule.Fields("Assur") < 0) Or (Lrs_Vehicule.Fields("Taxe") < 0) Or (Vidange >= 10000) Then

                    
            With SGrid1
                .AddRow
            
                    .CellDetails .Rows, .ColumnIndex("Véhicule"), Lrs_Vehicule.Fields("Marque") & Lrs_Vehicule.Fields("Matricule"), DT_LEFT
                    
                    If (Vidange >= 10000) Then
                        .CellDetails .Rows, .ColumnIndex("Vidange"), "Vidange obligatoire", DT_LEFT
                    End If
                    
                    If Lrs_Vehicule.Fields("CompteurVidange") <> 0 Then
                        .CellDetails .Rows, .ColumnIndex("CPT_BV"), Lrs_Vehicule.Fields("CompteurVidange"), DT_LEFT
                        Else
                        .CellDetails .Rows, .ColumnIndex("CPT_BV"), "Null", DT_LEFT
                    End If
                    
                    .CellDetails .Rows, .ColumnIndex("CPT_BC"), Lrs_Vehicule.Fields("CompteurCarburant"), DT_LEFT
                    
                   .CellDetails .Rows, .ColumnIndex("CPT_FT"), Lrs_Vehicule.Fields("CompteurEntre"), DT_LEFT
                    
                    If Lrs_Vehicule.Fields("Visite") < 0 Then
                        .CellDetails .Rows, .ColumnIndex("Visite"), Lrs_Vehicule.Fields("Visite"), DT_LEFT, , &HFF&, &HFFFFFF
                    End If
                    
                    If Lrs_Vehicule.Fields("Assur") < 0 Then
                        .CellDetails .Rows, .ColumnIndex("Assurance"), Lrs_Vehicule.Fields("Assur"), DT_LEFT, , &HFF&, &HFFFFFF
                    End If
                    
                    If Lrs_Vehicule.Fields("Taxe") < 0 Then
                        .CellDetails .Rows, .ColumnIndex("Taxe"), Lrs_Vehicule.Fields("Taxe"), DT_LEFT, , &HFF&, &HFFFFFF
                    End If

               
            End With
             End If
        Lrs_Vehicule.MoveNext
        Wend
    SGrid1.Redraw = True
    End If
End Sub

