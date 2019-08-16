VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmStatFT 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Statistiques Trafic"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleMode       =   0  'User
   ScaleWidth      =   15138.77
   Begin VB.ComboBox cbo_Destination 
      Height          =   315
      Left            =   5880
      TabIndex        =   14
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox cbo_Conducteur 
      Height          =   315
      ItemData        =   "FrmStatFT.frx":0000
      Left            =   3000
      List            =   "FrmStatFT.frx":0002
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox cbo_Vehicule 
      Height          =   315
      ItemData        =   "FrmStatFT.frx":0004
      Left            =   120
      List            =   "FrmStatFT.frx":0006
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
   End
   Begin SToolBox.SCommand SCommand1 
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   2040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      Caption         =   "....."
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
   Begin SToolBox.SDateBox cda_DateDebut 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   2160
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      Picture         =   "FrmStatFT.frx":0008
   End
   Begin SToolBox.SDateBox cda_dateFin 
      Height          =   285
      Left            =   4200
      TabIndex        =   2
      Top             =   2160
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
   End
   Begin SToolBox.SCommand SCommand2 
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   1560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      Picture         =   "FrmStatFT.frx":035B
   End
   Begin SToolBox.SCommand SCommand3 
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   1560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      Picture         =   "FrmStatFT.frx":06AE
   End
   Begin SToolBox.SGrid grid 
      Height          =   5175
      Left            =   0
      TabIndex        =   6
      Top             =   3480
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9128
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
   Begin MSComctlLib.ListView Lsv_Details 
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   2520
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1508
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre des voyages"
         Object.Width           =   1913
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tot.Durée"
         Object.Width           =   1913
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Moy.duré"
         Object.Width           =   1913
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tot.Distance"
         Object.Width           =   1913
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Moy.Distance"
         Object.Width           =   1913
      EndProperty
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Date Fin"
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
      Left            =   3120
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Date Debut"
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
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Destination"
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
      Left            =   6120
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Conducteur"
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
      Left            =   3240
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vehicule"
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
      Left            =   600
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "FrmStatFT.frx":0A01
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "FrmStatFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thekey As Integer
Dim theshift As Integer

Private Sub Initgrid_FT()
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
    .AddColumn "Matricule", "Matricule", , , 100
    .AddColumn "Conducteur", "Conducteur", , , 100
    .AddColumn "Destination", "Destination", , , 140
    .AddColumn "DateFT", "Date", , , 80
    .AddColumn "HeureS", "H.Sortie", , , 60
    .AddColumn "HeureE", "H.Entrée", , , 60
    .AddColumn "CPTS", "CPT.S", , , 60
    .AddColumn "CPTE", "CPT.E", , , 60
    .AddColumn "Distance", "Distance(KM)", , , 40
    .AddColumn "Dure", "Durée(Heure)", , , 60

    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With

End Sub

Private Sub Form_Load()
Me.Height = 9210
Me.Width = 11715
Me.Move 0, 0

Call Initgrid_FT

Dim i As Long

'For i = 2013 To Year(Date)
'    COMBO_ANNEE.AddItem i
'Next
'COMBO_ANNEE.ListIndex = COMBO_ANNEE.ListCount - 1

cbo_Vehicule.AddItem "Tous", 0
Call Affiche_Matricule_Combo(cbo_Vehicule)
cbo_Vehicule.ListIndex = 0
  
cbo_Conducteur.AddItem ("Tous"), 0
Call Affiche_Personnel_Combo(cbo_Conducteur)
cbo_Conducteur.ListIndex = 0

cbo_Destination.AddItem ("Tous"), 0
cbo_Destination.ListIndex = 0

cda_DateDebut.Text = "01/" & Month(Date) & "/" & Year(Date)
cda_dateFin.Text = Date

Me.WindowState = 2
End Sub

' CBO Vehicule

Private Sub cbo_vehicule_Change()
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
        start = Len(cbo_Vehicule.Text)
        For i = 0 To cbo_Vehicule.ListCount - 1
            If Left(cbo_Vehicule.List(i), start) = cbo_Vehicule.Text Then
                cbo_Vehicule.Text = cbo_Vehicule.List(i)
            End If
        Next
        cbo_Vehicule.SelStart = start
        cbo_Vehicule.SelLength = Len(cbo_Vehicule.Text)
    End If
End Sub

Private Sub cbo_Vehicule_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub


Private Sub cbo_Vehicule_Click()
If Len(Trim(cbo_Vehicule.Text)) > 0 Then Call AfficheRow_Vehicule(cbo_Vehicule.Text)

End Sub

Public Sub AfficheRow_Vehicule(ByVal vcode As String)

Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from vehicule where code = " & SQLText(vcode) & " OR Matricule= " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Matricule")) Then
    cbo_Vehicule.Text = rs("Matricule")
    End If
End If
rs.Close

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



Private Sub cbo_destination_Change()
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
        start = Len(cbo_Destination.Text)
        For i = 0 To cbo_Destination.ListCount - 1
            If Left(cbo_Destination.List(i), start) = cbo_Destination.Text Then
                cbo_Destination.Text = cbo_Destination.List(i)
            End If
        Next
        cbo_Destination.SelStart = start
        cbo_Destination.SelLength = Len(cbo_Destination.Text)
    End If
End Sub

Private Sub cbo_destination_GotFocus()
    Call Affiche_Destination_Combo(cbo_Destination)
    cbo_Destination.AddItem ("Tous"), 0
End Sub


Private Sub cbo_destination_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub


Private Sub cbo_Destination_Click()
If Len(Trim(cbo_Destination.Text)) > 0 Then Call AfficheRow_Destination(cbo_Conducteur.Text)

End Sub

Public Sub AfficheRow_Destination(ByVal vcode As String)
Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from Destination where Libelle = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Libelle")) Then cbo_Destination.Text = rs("Libelle")
End If
rs.Close

End Sub

Public Sub Affiche_FT(ByVal Vehicule As String, ByVal Conducteur As String, _
ByVal Destination As String, ByVal VdateD As String, ByVal vDateF As String)
Dim SQL As String
Dim rs As New ADODB.Recordset


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

Dim itmX As ListItem
Dim i As Integer
Dim Name_Table As String


'On Error GoTo Err
' Initialisation des Variables
Vehicule = cbo_Vehicule.Text
Conducteur = cbo_Conducteur.Text
Destination = cbo_Destination.Text
VdateD = cda_DateDebut.Text
vDateF = cda_dateFin.Text

For i = Year(VdateD) To Year(vDateF)

Name_Table = "FicheTraffic"
If i < Year(Date) Then
    Name_Table = "FicheTraffic_" & i
End If

        SQL = SQL & " Select * from  " & Name_Table & "  where HeureSortie Between" & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00"))
        
        If Vehicule <> "Tous" Then
            SQL = SQL & " And vehicule=" & SQLText(Vehicule)
        End If
        
        If Conducteur <> "Tous" Then
            SQL = SQL & " And Conducteur=" & SQLText(Conducteur)
        End If
        
        If Destination <> "Tous" Then
            SQL = SQL & " AND Destination=" & SQLText(Destination)
        End If
        
        If i <> Year(vDateF) Then
            SQL = SQL & " Union all"
        End If
        
Next

        
rs.Open SQL, CNB, adOpenKeyset
grid.ClearRows
'initialisation des variables
Voyage = 0
Dure = "0"
Distance = 0
If Not rs.EOF Then
    grid.Redraw = False
    Dur = 0
    heur = 0
    min = 0
    temp = ""
    While Not rs.EOF
        With grid
            .AddRow
            .CellDetails .Rows, .ColumnIndex("Numero"), rs("Numero")
            .CellDetails .Rows, .ColumnIndex("Matricule"), rs("Vehicule")
            .CellDetails .Rows, .ColumnIndex("Conducteur"), rs("Conducteur")
            .CellDetails .Rows, .ColumnIndex("Destination"), rs("Destination")
            .CellDetails .Rows, .ColumnIndex("DateFT"), Format(rs("heureSortie"), "dd/mm/yyyy")
            .CellDetails .Rows, .ColumnIndex("HeureS"), Format(rs("HeureSortie"), "hh:mm")
            If Not (IsNull(rs("HeureENtre"))) Then .CellDetails .Rows, .ColumnIndex("HeureE"), Format(rs("HeureENtre"), "hh:mm")
            .CellDetails .Rows, .ColumnIndex("CPTS"), rs("CompteurSortie")
            If Not (IsNull(rs("HeureENtre"))) Then .CellDetails .Rows, .ColumnIndex("CPTE"), rs("CompteurEntre")
            If Not (IsNull(rs("HeureENtre"))) Then .CellDetails .Rows, .ColumnIndex("Distance"), Val(rs("CompteurEntre")) - Val(rs("CompteurSortie"))
            If Not (IsNull(rs("HeureENtre"))) Then
            'Calcule de durée
            Dur = DateDiff("n", rs("HeureSortie"), rs("HeureEntre"))
            heur = Dur \ 60
            min = Dur - (heur * 60)
            temp = CStr(heur) & ":" & CStr(min)
            .CellDetails .Rows, .ColumnIndex("Dure"), temp
            End If
            
        End With
        ' Affiche LSv_details
      
        If Not (IsNull(rs("HeureENtre"))) Then
        
            Voyage = Voyage + 1
            Dure = Val(Dure) + (DateDiff("n", rs("HeureSortie"), rs("HeureEntre")))
            Distance = Distance + (Val(rs("CompteurEntre")) - Val(rs("CompteurSortie")))
             
        End If
        rs.MoveNext
    Wend
    grid.Redraw = True
End If

'Lsv_toto
Lsv_Details.ListItems.Clear
Set itmX = Lsv_Details.ListItems.Add(, , CStr(Voyage))
            hours = Dure \ 60
            minutes = Dure - (hours * 60)
            timeElapsed = CStr(hours) & ":" & CStr(minutes)
            
            If Voyage > 0 Then
                MoyDur = Dure \ Voyage
                h = MoyDur \ 60
                m = MoyDur - (h * 60)
                T = CStr(h) & ":" & CStr(m)
            End If
            
            itmX.SubItems(1) = CStr(timeElapsed)
            
            If Distance > 0 Then
                itmX.SubItems(2) = CStr(T)
            Else
                itmX.SubItems(2) = "Voyages égale à zéro"
            End If
            
            itmX.SubItems(3) = CStr(Distance)
            
            If Voyage > 0 Then
                MoyDis = Distance \ Voyage
            End If
            
            If Voyage > 0 Then
                itmX.SubItems(4) = CStr(MoyDis)
            Else
                itmX.SubItems(4) = "Voyages égale à zéro"
            End If
rs.Close
'Err:
'MsgBox Err.Description, vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo erreur
   Dim i As Integer
   Dim MSG ' Déclare la variable.
   ' Définit le texte du message.
   MSG = "Voulez-vous vraiment quitter?"
   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
   If MsgBox(MSG, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
      Cancel = True
   Else
   Unload Me
   End If
   
   Exit Sub
erreur:
   MsgBox Err.Description, 48
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
    Case "Matricule"
         .SortType(1) = CCLSortString
    Case "Conducteur"
         .SortType(1) = CCLSortString
    Case "Destination"
         .SortType(1) = CCLSortString
    Case "DateFT"
         .SortType(1) = CCLSortDate
    Case "HeureS"
         .SortType(1) = CCLSortDateHourAccuracy
    Case "HeureE"
         .SortType(1) = CCLSortDateHourAccuracy
    Case "CPTS"
         .SortType(1) = CCLSortNumeric
    Case "CPTE"
         .SortType(1) = CCLSortNumeric
    Case "Distance"
         .SortType(1) = CCLSortNumeric
    Case "Dure"
         .SortType(1) = CCLSortDateHourAccuracy
         
  
      End Select
   End With
   Screen.MousePointer = vbHourglass
   grid.Sort
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub grid_DblClick(ByVal lRow As Long, ByVal lCol As Long)

On Error GoTo Err

With FrmMajStatFT
        FrmMajStatFT.selectFT (grid.CellText(grid.SelectedRow, grid.ColumnIndex("Numero")))
.Show
    
End With

Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub SCommand1_Click()

If Left(CheckMandatory(FrmStatFT), 1) = 1 Then
   Exit Sub
End If
Call Affiche_FT(cbo_Vehicule.Text, cbo_Conducteur.Text, cbo_Destination.Text, _
cda_DateDebut.Text, cda_dateFin.Text)



End Sub


