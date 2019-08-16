VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmStatService 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Stat Personnel"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   11175
   Begin VB.ComboBox cbo_Conducteur 
      Height          =   315
      ItemData        =   "FrmStatService.frx":0000
      Left            =   2160
      List            =   "FrmStatService.frx":0002
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   420
      Left            =   6120
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin SToolBox.SDateBox cda_Debut 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Tag             =   "M"
      Top             =   2520
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   345
      Left            =   5400
      TabIndex        =   3
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmStatService.frx":0004
      ButtonType      =   1
   End
   Begin SToolBox.SGrid grid 
      Height          =   5655
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9975
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
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
   Begin SToolBox.SDateBox cda_Fin 
      Height          =   285
      Left            =   4650
      TabIndex        =   8
      Tag             =   "M"
      Top             =   2520
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date fin:"
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
      Left            =   3540
      TabIndex        =   9
      Top             =   2520
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conducteur:"
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
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label m 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statistiques Personnel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date debut:"
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
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmStatService.frx":0357
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20895
   End
End
Attribute VB_Name = "FrmStatService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Initgrid_FP()
With grid
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
    
    .AddColumn "Numero", "Numero", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Date", "Date", , , 100
    .AddColumn "Etat", "Etat", , , 0
    .AddColumn "HDebut", "HDebut", , , 100
    .AddColumn "HFin", "HFin", , , 100
    .AddColumn "Dure", "Dure", , , 100
    .AddColumn "Activités", "Activités", , , 500

    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With

End Sub


Public Sub Affiche_FP()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim rQ As New ADODB.Recordset

Dim VdateD As String
Dim vDateF
Dim Conducteur As String

Dim min As Long
Dim heur As Long
Dim Dur As Long
Dim temp As String

Dim Voyage As Long
Dim Distance As Long

Dim m As Long
Dim h As Long
Dim MoyDur As Long
Dim T As String

Dim MoyDis As Long

Dim minutes As Long
Dim hours As Long
Dim Dure As Long
Dim timeElapsed As String

Dim Destination As String

'On Error GoTo Err
' Initialisation des Variables

Conducteur = cbo_Conducteur.Text
VdateD = cda_Debut.Text
vDateF = cda_fin.Text

'Select Case statement !!
Select Case True
   
    Case (Conducteur <> "")
        SQL = "Select * from DispoPerso where Conducteur=" & SQLText(Conducteur)
        SQL = SQL & " And HDebut Between" & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:00:00"))
    Case Else
        MsgBox ("Combinaise Non Valide")
        Exit Sub
    End Select
rs.Open SQL, CNB, adOpenKeyset
grid.ClearRows

'initialisation des variables

Dure = "0"
If Not rs.EOF Then
    grid.Redraw = False
    Dur = 0
    heur = 0
    min = 0
    temp = ""
    While Not rs.EOF
        With grid
            If (rs("Etat") = "En-Service") Then
            .AddRow
                If Not (IsNull(rs("Numero"))) Then .CellDetails .Rows, .ColumnIndex("Numero"), rs("Numero"), , , &HC0FFC0
                If Not (IsNull(rs("Etat"))) Then .CellDetails .Rows, .ColumnIndex("Etat"), rs("Etat"), , , &HC0FFC0
                If Not (IsNull(rs("HDebut"))) Then .CellDetails .Rows, .ColumnIndex("Date"), Format(rs("HDebut"), "dd/mm/yyyy"), , , &HC0FFC0
                If Not (IsNull(rs("HDebut"))) Then .CellDetails .Rows, .ColumnIndex("HDebut"), Format(rs("HDebut"), "hh:mm"), , , &HC0FFC0
                If Not (IsNull(rs("HFin"))) Then .CellDetails .Rows, .ColumnIndex("HFin"), Format(rs("HFin"), "hh:mm"), , , &HC0FFC0
                If Not (IsNull(rs("HFin"))) Then
                'Calcule de durée
                Dur = DateDiff("n", rs("HDebut"), rs("HFin"))
                heur = Dur \ 60
                min = Dur - (heur * 60)
                temp = CStr(heur) & ":" & CStr(min)
                .CellDetails .Rows, .ColumnIndex("Dure"), temp, , , &HC0FFC0
                End If
                
             'Select Destiontion
             If Not (IsNull(rs("HFin"))) Then
                SQL = "select  HeureSortie, Destination from fichetraffic where Conducteur=" & SQLText(Conducteur) & ""
                SQL = SQL & " And HeureSortie Between" & SQLText(rs("HDebut")) & " and " & SQLText(rs("HFin")) & " Order by HeureSortie"
                rQ.Open SQL, CNB, adOpenKeyset
                Destination = ""
                While Not rQ.EOF
                    Destination = Destination & " | " & Format(rQ("HeureSortie"), "hh:mm") & " à " & rQ("Destination")
                    rQ.MoveNext
                Wend
                .CellDetails .Rows, .ColumnIndex("Activités"), Destination, , , &HC0FFC0
                 rQ.Close
            End If
        End If
            
        End With
       
        rs.MoveNext
    Wend
    grid.Redraw = True
End If
'Err:
'MsgBox Err.Description, vbInformation
End Sub


Private Sub Command1_Click()
'On Error GoTo Err
 If Left(CheckMandatory(FrmStatService), 1) = 1 Then
       Exit Sub
    End If
Call Affiche_FP

'Err:
'MsgBox Err.Description, vbInformation


End Sub

Private Sub Form_Load()
Me.Height = 9210
Me.Width = 11715
Me.Move 0, 0

Call Initgrid_FP

cda_Debut.Text = "01/" & Month(Date) & "/" & Year(Date)
cda_fin.Text = Date

Me.WindowState = 2

End Sub

'CBO CONDUCTEUR

Private Sub Cbo_Conducteur_Change()
Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
    Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
 
        ' Nothing to do now !...maybe later ;-)
 
    Else
        start = Len(cbo_Conducteur.Text)
        For i = 0 To cbo_Conducteur.ListCount - 1
            If Left(cbo_Conducteur.List(i), start) = cbo_Conducteur.Text Then
                cbo_Conducteur.Text = cbo_Conducteur.List(i)
            End If
        Next
        cbo_Conducteur.SelStart = start
        cbo_Conducteur.SelLength = Len(cbo_Conducteur.Text)
    End If
End Sub

Private Sub Cbo_Conducteur_GotFocus()
    Call Affiche_Personnel_Combo(cbo_Conducteur)
End Sub


Private Sub Cbo_Conducteur_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub


Private Sub cbo_Conducteur_Click()
If Len(Trim(cbo_Conducteur.Text)) > 0 Then Call AfficheRow_Conducteur(cbo_Conducteur.Text)

End Sub

Public Sub AfficheRow_Conducteur(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from personnel where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Libelle")) Then cbo_Conducteur.Text = rs("Libelle")
End If
rs.Close

End Sub

Private Sub grid_ColumnClick(ByVal lCol As Long)
Dim sTag As String
Dim i As Long
      
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
    Case "Conducteur"
         .SortType(1) = CCLSortString
    Case "Etat"
         .SortType(1) = CCLSortString
    Case "HDebut"
         .SortType(1) = CCLSortDateHourAccuracy
    Case "HFin"
         .SortType(1) = CCLSortDateHourAccuracy
    Case "Dure"
         .SortType(1) = CCLSortDateHourAccuracy
         
  
      End Select
   End With
   Screen.MousePointer = vbHourglass
   grid.Sort
   Screen.MousePointer = vbDefault
   
End Sub


Private Sub grid_DblClick(ByVal lRow As Long, ByVal lCol As Long)
Dim Conducteur As String
Dim DateD As String
Dim DateF As String

Conducteur = cbo_Conducteur.Text
DateD = cda_Debut.Text
DateF = cda_Debut.Text


With FrmStatFT
    .cbo_Vehicule.Text = "Tous"
    .cbo_Conducteur.Text = Conducteur
    .cbo_destination.Text = "Tous"
    .cda_DateDebut.Text = DateD
    .cda_dateFin.Text = DateF
    .Show
End With

End Sub


