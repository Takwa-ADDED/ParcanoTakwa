VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmCompteurs 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Liset des compteurs"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin SToolBox.SGrid grid 
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   7095
      _ExtentX        =   12515
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
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   "Imprimer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmCompteurs.frx":0000
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liset des Compteurs"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "FrmCompteurs.frx":0353
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7080
   End
End
Attribute VB_Name = "FrmCompteurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdPrint_Click()
If MsgBox("Imprimer la liste        ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
    Dim F As Form
    Set F = New Frm_Rpt_Apercus
    With F
        Call .PrintOutAndApercu_Compteurs(0)
        .Show
    End With
End If
End Sub

Private Sub Form_Load()
 Call Initgrid
 Call Affiche_Grid
 
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
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim j As Integer
Dim Couleur As String

grid.ClearRows

SQL = "Select * from vehicule where actif=1 order by Matricule"
rs.Open SQL, CNB, adOpenKeyset

If Not rs.EOF Then
    grid.Redraw = False
    While Not rs.EOF
       
        With grid
            .AddRow
            .CellDetails .Rows, .ColumnIndex("Matricule"), rs("Matricule")
            .CellDetails .Rows, .ColumnIndex("CPTFT"), rs("CompteurFT")
            .CellDetails .Rows, .ColumnIndex("CPTBC"), rs("CompteurCarburant")
            .CellDetails .Rows, .ColumnIndex("CPTBV"), rs("CompteurVidange")
        End With
            rs.MoveNext
    Wend
    grid.Redraw = True
 End If
grid.SelectedRow = 1
End Sub

