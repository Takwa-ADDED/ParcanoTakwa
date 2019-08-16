VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form Frm_Compteurs 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liset des compteurs"
   ClientHeight    =   6675
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SToolBox.SGrid grid 
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10398
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   -2147483644
      ForeColor       =   0
      NoFocusHighlightBackColor=   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin SToolBox.SCommand CmdPrint 
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BackStyle       =   0
      Caption         =   "Imprimer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_Compteurs.frx":0000
      BackColor       =   12632256
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liset des Compteurs"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3270
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "Frm_Compteurs.frx":0353
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8520
   End
End
Attribute VB_Name = "Frm_Compteurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=================================================
'Chargement Form***
'=================================================
Private Sub Form_Load()
    Call Initgrid
    Call Affiche_Grid
    Me.Caption = Me.Caption
End Sub
Public Sub Initgrid()
    With grid
        .Redraw = False
        ' Allow the grid to be grouped, but
        ' don't show the grouping box
        .HideGroupingBox = True
        .AllowGrouping = True
        ' Group rows will be shown by
        ' a gradient underline
        .GroupRowBackColor = vbWindowBackground
        .GroupRowForeColor = vbWindowText
        
        .GridLineColor = vbWindowBackground
        .GridFillLineColor = vbWindowBackground
        .GridLines = True
        
        .SelectionAlphaBlend = True
        .SelectionOutline = True
        .DrawFocusRectangle = False
        
        .AddColumn "Matricule", "Matricule", , , 140
        .AddColumn "CPTFT", "CPT.FT", , , 120
        .AddColumn "CPTBC", "CPT.BC", , , 120
        .AddColumn "CPTBV", "CPT.BV", , , 120
        
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Public Sub Affiche_Grid()
    Dim LObj_Find As Vehicule
    Dim Lrs_Vehicule As Recordset

    grid.ClearRows
    
    Set LObj_Find = New Vehicule
    Set Lrs_Vehicule = LObj_Find.GetAllActifVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing

    If Not Lrs_Vehicule.EOF Then
        grid.Redraw = False
        While Not Lrs_Vehicule.EOF
            With grid
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Vehicule("Matricule")
                .CellDetails .Rows, .ColumnIndex("CPTFT"), Lrs_Vehicule("CompteurFT")
                .CellDetails .Rows, .ColumnIndex("CPTBC"), Lrs_Vehicule("CompteurCarburant")
                .CellDetails .Rows, .ColumnIndex("CPTBV"), Lrs_Vehicule("CompteurVidange")
            End With
            Lrs_Vehicule.MoveNext
        Wend
        grid.Redraw = True
    End If
    grid.SelectedRow = 1
    Set Lrs_Vehicule = Nothing
End Sub
'=================================================
'Imprimant***
'=================================================
Private Sub CmdPrint_Click()
    Dim F As Form
    If MsgBox("Imprimer la liste!...", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
        Set F = New Frm_Rpt_Apercus
        With F
            Call .PrintOutAndApercu_Compteurs(0)
            .Show
        End With
    End If
End Sub







