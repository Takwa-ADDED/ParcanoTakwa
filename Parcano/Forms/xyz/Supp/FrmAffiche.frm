VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAffiche 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gestion de Traffic Auto"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FrmAffiche.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   11760
   Begin SToolBox.SCommand cmd_r 
      Height          =   615
      Left            =   10200
      TabIndex        =   8
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Caption         =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   240
   End
   Begin MSComctlLib.ListView LSV_Exterieur 
      Height          =   6015
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   10610
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Conducteur"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Matricule"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Destination"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Heure.S"
         Object.Width           =   1164
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Heure.E"
         Object.Width           =   1164
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "CPT.S"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "CPT.E"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "KM"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Durée"
         Object.Width           =   2540
      EndProperty
   End
   Begin SToolBox.SGrid grid_vehicule 
      Height          =   6015
      Left            =   9720
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   10610
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   -2147483644
      ForeColor       =   0
      NoFocusHighlightBackColor=   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin SToolBox.SGrid grid_Conducteur 
      Height          =   6015
      Left            =   7320
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   10610
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      NoFocusHighlightBackColor=   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin MSComctlLib.ListView Lsv_Depot 
      Height          =   3855
      Left            =   -14400
      TabIndex        =   7
      Top             =   4320
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vehicule a l'interieure"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Hors-service"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   11
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "En- service"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Occupé 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Occupé"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Line Line1 
      X1              =   7200
      X2              =   7200
      Y1              =   1560
      Y2              =   8640
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Centra Nord - Bizerte"
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
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4170
   End
   Begin VB.Label Lbl_heure 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Lbl_date 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contrôle Vehicule"
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
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Stretch         =   -1  'True
      Top             =   600
      Width           =   15255
   End
End
Attribute VB_Name = "FrmAffiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Okay As Boolean
Public ii As Integer
Dim Operation() As String

Dim thekey As Integer
Dim theshift As Integer
Dim itmX As ListItem

Private Sub Form_Load()

Dim dat As Date

Dim large As Integer
Dim haut As Integer
large = Screen.Width
haut = Screen.Height
Me.Left = 0
Me.Top = 0
Me.Width = large
Me.Height = haut

Lbl_heure.Caption = Format(Time, "hh:mm:ss")
Timer1.Enabled = True
Timer1.Interval = 1000



Call AfficheExterieur
Call AfficheDepot

Call Initgrid_Vehicule
Call Affiche_Vehicule
Call Initgrid_Conducteur
Call Affiche_Conducteur


Call grid_vehicule_ColumnClick(1)
Call grid_Conducteur_ColumnClick(1)


dat = Date
Lbl_date.Caption = UCase(Format(Now, "dddd-dd-mm-yyyy"))
dat = Time
Lbl_heure.Caption = dat
Me.WindowState = 2
End Sub

Private Sub cmd_r_Click()

Call Affiche_Vehicule
Call Affiche_Conducteur
Call AfficheDepot
Call AfficheExterieur

Call grid_vehicule_ColumnClick(1)
Call grid_Conducteur_ColumnClick(1)
End Sub
Private Sub grid_Conducteur_ColumnClick(ByVal lCol As Long)
Dim sTag As String
Dim i As Long
      
   With grid_Conducteur.SortObject
      .Clear
      .SortColumn(1) = lCol
   
      sTag = grid_Conducteur.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(1) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(1) = CCLOrderAscending
      End If
      grid_Conducteur.ColumnTag(lCol) = sTag
   
      Select Case grid_Conducteur.ColumnKey(lCol)
      Case "Libelle"
         ' sort by backColor:
         .SortType(1) = CCLSortStringNoCase
         .SortType(1) = CCLSortBackColor
      End Select
   End With
   Screen.MousePointer = vbHourglass
   grid_Conducteur.Sort
   Screen.MousePointer = vbDefault


End Sub



Private Sub grid_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"

End Sub



Private Sub grid_vehicule_ColumnClick(ByVal lCol As Long)
Dim sTag As String
Dim i As Long
      
   With grid_vehicule.SortObject
      .Clear
      .SortColumn(1) = lCol
   
      sTag = grid_vehicule.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(1) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(1) = CCLOrderAscending
      End If
      grid_vehicule.ColumnTag(lCol) = sTag
   
      Select Case grid_vehicule.ColumnKey(lCol)
      Case "Matricule"
         .SortType(1) = CCLSortBackColor
      End Select
   End With
   Screen.MousePointer = vbHourglass
   grid_vehicule.Sort
   Screen.MousePointer = vbDefault
   
End Sub


Private Sub grid_vehicule_GotFocus()
If grid_Conducteur.Enabled = False Then grid_Conducteur.Enabled = True

End Sub

Private Sub grid_vehicule_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub Timer1_Timer()
Lbl_heure = Time
End Sub


Public Sub AfficheExterieur()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim Couleur As String
Dim datesys As Date
datesys = Date
LSV_Exterieur.ListItems.Clear

'Voitures en exterieure de plus d'un jours
SQL = " Select * from fichetraffic where CONVERT(VARCHAR,HeureSortie,103)<" & SQLText(datesys) & " And heureEntre is Null Order by HeureEntre ASC,heureSortie ASC"
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then

    While Not rs.EOF

            Set itmX = LSV_Exterieur.ListItems.Add(, , "")
            itmX.SubItems(1) = CStr(rs("Numero"))
            itmX.SubItems(3) = rs("vehicule")
            itmX.SubItems(2) = rs("Conducteur")
            itmX.SubItems(4) = rs("Destination")
            itmX.SubItems(5) = Format(rs("heureSortie"), "hh:mm")
            If Not IsNull(rs("HeureENtre")) Then itmX.SubItems(6) = Format(rs("HeureENtre"), "hh:mm")
            If Not IsNull(rs("CompteurSortie")) Then itmX.SubItems(7) = rs("CompteurSortie")
            If Not IsNull(rs("CompteurEntre")) Then itmX.SubItems(8) = rs("CompteurEntre")
            If Not IsNull(rs("CompteurEntre")) Then itmX.SubItems(9) = Val(rs("CompteurEntre")) - Val(rs("CompteurSortie")) & " KM"
            If IsNull(rs("HeureENtre")) Then itmX.SubItems(10) = Format(rs("heureSortie"), "dd/mm/yyyy hh:mm")

            rs.MoveNext
    Wend
End If
rs.Close

'Detail des fiches d'aujourdhui
SQL = " Select * from fichetraffic where CONVERT(VARCHAR,HeureSortie,103)=" & SQLText(datesys) & " OR CONVERT(VARCHAR,HeureEntre,103)=" & SQLText(datesys) & " Order by HeureEntre"
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then

    While Not rs.EOF

            Set itmX = LSV_Exterieur.ListItems.Add(, , "")
            itmX.SubItems(1) = CStr(rs("Numero"))
            itmX.SubItems(3) = rs("vehicule")
            itmX.SubItems(2) = rs("Conducteur")
            itmX.SubItems(4) = rs("Destination")
            itmX.SubItems(5) = Format(rs("heureSortie"), "hh:mm")
            If Not IsNull(rs("HeureENtre")) Then itmX.SubItems(6) = Format(rs("HeureENtre"), "hh:mm")
            If Not IsNull(rs("CompteurSortie")) Then itmX.SubItems(7) = rs("CompteurSortie")
            If Not IsNull(rs("CompteurEntre")) Then itmX.SubItems(8) = rs("CompteurEntre")
            If Not IsNull(rs("CompteurEntre")) Then itmX.SubItems(9) = Val(rs("CompteurEntre")) - Val(rs("CompteurSortie")) & " KM"
            If Not IsNull(rs("HeureENtre")) Then
            itmX.SubItems(10) = Format(rs("HeureENtre") - rs("heureSortie"), "hh:mm")
            Else
            itmX.SubItems(10) = Format(Now - rs("heureSortie"), "hh:mm")
            End If
            rs.MoveNext
    Wend
End If
rs.Close
End Sub

Public Sub AfficheDepot()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim datesys As Date
Dim N As Integer

datesys = Date
 Lsv_Depot.ListItems.Clear

    SQL = "select Matricule from vehicule Order by matricule"
    rs.Open SQL, CNB, adOpenKeyset

   While Not rs.EOF
                Set itmX = Lsv_Depot.ListItems.Add(, , "")
                itmX.SubItems(1) = CStr(rs("Matricule"))
            rs.MoveNext
    Wend
    
  For j = 1 To LSV_Exterieur.ListItems.Count
  For i = 1 To Lsv_Depot.ListItems.Count - 1
   
           If (Lsv_Depot.ListItems(i).SubItems(1) = LSV_Exterieur.ListItems(j).SubItems(3)) _
                And (Len(LSV_Exterieur.ListItems(j).SubItems(6)) = 0) Then
           Lsv_Depot.ListItems.Remove (i)
           End If
    Next
Next


rs.Close
End Sub




Public Sub Initgrid_Vehicule()
With grid_vehicule
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
    
    .AddColumn "Matricule", "Vehicule", , , 140
    
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub

Public Sub Affiche_Vehicule()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim j As Integer
Dim Couleur As String

grid_vehicule.ClearRows

SQL = "Select * from vehicule where Actif=1 order by Matricule"
rs.Open SQL, CNB, adOpenKeyset

ReDim Operation(1)


If Not rs.EOF Then
    grid_vehicule.Redraw = False
    While Not rs.EOF
        Couleur = "vbRed"
        Operation = ReturnOperation(rs("Matricule"))
       If Operation(0) = "S" Then
                Couleur = "vbGreen"
         End If
            
        With grid_vehicule
            .AddRow
            If Couleur = "vbRed" Then
                .CellDetails .Rows, .ColumnIndex("Matricule"), rs("Matricule"), , , &H8080FF
            Else
            If (rs("disponible") = "O") Then
                .CellDetails .Rows, .ColumnIndex("Matricule"), rs("Matricule"), , , &HC0FFC0
                Else
                 .CellDetails .Rows, .ColumnIndex("Matricule"), rs("Matricule"), , , &HFFFF&
                 End If
            End If
        End With
            rs.MoveNext
    Wend
    grid_vehicule.Redraw = True
 End If
End Sub

Public Sub Initgrid_Conducteur()
With grid_Conducteur
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
    
    .AddColumn "Libelle", "Conducteur", , , 140
    
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub

Public Function CompteurVehicule(ByVal vcode As String) As Long
    Dim rD As New ADODB.Recordset
    Dim SQL As String
    SQL = "Select * from vehicule where Matricule = " & SQLText(vcode)
    rD.Open SQL, CNB, adOpenKeyset
    If Not (IsNull(rD("CompteurFT"))) Then
        CompteurVehicule = rD("CompteurFT")
    End If
End Function

Public Function ReturnOperation(ByVal Matricule As String) As String()
Dim Tableau() As String
ReDim Tableau(1)

Dim rD As New ADODB.Recordset
Dim SQL As String
SQL = "Select * from fichetraffic where Vehicule = " & SQLText(Matricule)
rD.Open SQL, CNB, adOpenKeyset
     Tableau(0) = ("S")
    While Not rD.EOF
        If IsNull(rD("HeureEntre")) Then
            Tableau(0) = ("E")
            Tableau("1") = (rD("Numero"))
        End If
     rD.MoveNext
     Wend
ReturnOperation = Tableau()

rD.Close
End Function

Public Sub Affiche_Conducteur()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim j As Integer
Dim Couleur As String

grid_Conducteur.ClearRows
SQL = "Select * from Personnel where Actif=1  order by Libelle"
rs.Open SQL, CNB, adOpenKeyset

If Not rs.EOF Then
    grid_Conducteur.Redraw = False
    While Not rs.EOF
        Couleur = "vbGreen"
        For j = 1 To LSV_Exterieur.ListItems.Count
                If (rs("Libelle") = LSV_Exterieur.ListItems(j).SubItems(2) And LSV_Exterieur.ListItems(j).SubItems(6) = "" And LSV_Exterieur.ListItems(j).SubItems(4) <> "REPARATION") Then
                Couleur = "vbRed"
            End If
        Next
        With grid_Conducteur
            .AddRow
            If Couleur = "vbRed" Then
                .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Libelle"), , , &H8080FF
            Else
                If (rs("disponible") = "O") Then
                .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Libelle"), , , &HC0FFC0
                Else
                 .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Libelle"), , , &HFFFF&
                 End If
            End If
        End With
            rs.MoveNext
    Wend
    grid_Conducteur.Redraw = True
 End If

End Sub




