VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_StatPLNG 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Statistiques PLANNING"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Tab_FindView 
      Height          =   2535
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Consultation"
      TabPicture(0)   =   "Frm_StatPLNG.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Pic_ConsltConge"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox Pic_ConsltConge 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   7815
         TabIndex        =   3
         Top             =   360
         Width           =   7815
         Begin MSComCtl2.DTPicker cda_fin 
            Height          =   375
            Left            =   6000
            TabIndex        =   4
            Top             =   120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   179896321
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker cda_Db 
            Height          =   375
            Left            =   3600
            TabIndex        =   5
            Top             =   120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   179896321
            CurrentDate     =   42860
         End
         Begin SToolBox.SBiCombo Cbo_Conducteur 
            Height          =   405
            Left            =   1440
            TabIndex        =   6
            Top             =   720
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   714
            FontBold        =   -1  'True
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SToolBox.SBiCombo SBC_Dest 
            Height          =   315
            Left            =   1440
            TabIndex        =   10
            Top             =   1440
            Width           =   4095
            _ExtentX        =   7223
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
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Destination :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   5520
            TabIndex        =   9
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   3120
            TabIndex        =   8
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Conducteur "
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.Image Cmd_FindPLNG 
            Height          =   495
            Left            =   5760
            Picture         =   "Frm_StatPLNG.frx":001C
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   1935
         End
      End
   End
   Begin SToolBox.SGrid grid_detail 
      Height          =   6855
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12091
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
   Begin VB.PictureBox Pic_date 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      ScaleHeight     =   375
      ScaleWidth      =   8295
      TabIndex        =   12
      Top             =   1200
      Width           =   8295
      Begin VB.Label Lbl_dateAu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5160
         TabIndex        =   16
         Top             =   0
         Width           =   2610
      End
      Begin VB.Label Lbl_DateDu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   0
         Width           =   2130
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Du :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Au :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   4560
         TabIndex        =   13
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Image Pic_ShowMenu 
      Height          =   375
      Left            =   8880
      Picture         =   "Frm_StatPLNG.frx":10C1E
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Pic_MaskMenu 
      Height          =   375
      Left            =   8880
      Picture         =   "Frm_StatPLNG.frx":11574
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label LBL_Titre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statistiques PLANNING par conducteur"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6345
   End
   Begin VB.Image PicBox_Header 
      Height          =   1095
      Left            =   -120
      Picture         =   "Frm_StatPLNG.frx":1229A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "Frm_StatPLNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Cmd_FindPLNG_Click()
    Dim Condt As String
On Error GoTo Err
    Condt = Cbo_Conducteur.FirstValue
    
    If cda_Db.Value > cda_fin.Value Then
        MsgBox "Vérifier date de recherche!...", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    Call initgrid_Detail
    Call AfficheStatPLNG_Cond(cda_Db.Value, cda_fin.Value, SBC_Dest.FirstValue, Condt)
    Call Pic_MaskMenu_Click
    Lbl_dateAu.Caption = cda_fin.Value
    Lbl_DateDu.Caption = cda_Db.Value
    
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub AfficheStatPLNG_Cond(ByVal Date_db As Date, ByVal Date_f As Date, ByVal DESTINATION As String, ByVal cond As String)

Dim LOBJ_statPLNG As New PLANNING
Dim rs As New Recordset
Dim Lrs As New Recordset
Dim JourDate As Integer
Dim NbrPlng As Integer

On Error GoTo Err

Set rs = LOBJ_statPLNG.Get_DetailPLNG(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Date_db, Date_f, DESTINATION, cond)
If ErrNumber <> 0 Then
    ErrNumber = 0
    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        Set Lrs = LOBJ_statPLNG.Get_CountPLNG(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Date_db, Date_f, rs("TOURNEE"), rs("CONDUCTEUR"))
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        If Not Lrs.EOF Then
            NbrPlng = Lrs("nbrPLNG")
        End If
        Set Lrs = Nothing
        With grid_detail
            .AddRow
            If cond <> "0000" And cond <> "" Then
                .CellDetails .Rows, .ColumnIndex("Destination"), rs("TOURNEE") & " -- " & rs("CONDUCTEUR") & " -- " & NbrPlng
            Else
                .CellDetails .Rows, .ColumnIndex("Conducteur"), rs("CONDUCTEUR") & " -- " & rs("TOURNEE") & " -- " & NbrPlng
            End If
            
            '.CellDetails .Rows, .ColumnIndex("Destination"),rs("TOURNEE")
            If rs("JOUR") = "lundi" Then JourDate = 0
            If rs("JOUR") = "mardi" Then JourDate = 1
            If rs("JOUR") = "mercredi" Then JourDate = 2
            If rs("JOUR") = "jeudi" Then JourDate = 3
            If rs("JOUR") = "vendredi" Then JourDate = 4
            If rs("JOUR") = "samedi" Then JourDate = 5
            If rs("JOUR") = "dimanche" Then JourDate = 6

            .CellDetails .Rows, .ColumnIndex("Date"), Format(DateCell(rs("DATEDU"), JourDate), " dddd - dd/mm/yyyy")
        End With
    rs.MoveNext
    Wend
    
    If cond <> "0000" And cond <> "" Then
        grid_detail.ColumnWidth("Conducteur") = 0
    Else
        grid_detail.ColumnWidth("Conducteur") = 200
    End If
   grid_detail.Redraw = True
End If
rs.Close
With grid_detail
         .GroupRowBackColor = RGB(251, 246, 206)
         .GroupRowForeColor = QBColor(12)
         If cond <> "0000" And cond <> "" Then
            .ColumnIsGrouped(2) = True
        Else
            .ColumnIsGrouped(1) = True
        End If
         .GroupRowForeColor = QBColor(10)
         .HideGroupingBox = True
         .AllowGrouping = True
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Private Function DateCell(ByVal DateSearch As Date, ByVal jour As Integer) As Date
    Dim LOBJ_Find As New PLANNING, Lrs_Date As New Recordset
On Error GoTo Err
    Set Lrs_Date = LOBJ_Find.GetDate_NewPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, DateSearch, jour, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Function
    End If
    Set LOBJ_Find = Nothing
    If Not Lrs_Date.EOF Then DateCell = Lrs_Date.Fields("datedebut")
    Set Lrs_Date = Nothing
Exit Function
Err:
    MsgBox Err.Description, vbExclamation
End Function

Private Sub Form_Load()

On Error GoTo Err
    cda_Db.Value = "01/01/" & Year(Date)
    cda_fin.Value = Format(Date, "DD/MM/YYYY")
    Cbo_Conducteur.AddItem "0000", "Tous"
    SBC_Dest.AddItem "0000", "Tous"
    Call Affiche_Personnel_SBCombo(Cbo_Conducteur)
    Call Affiche_DestPLNG
    Call initgrid_Detail
    Cbo_Conducteur.ListIndex = 0
    SBC_Dest.ListIndex = 0
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Public Sub Affiche_DestPLNG()

Dim Lobj_Destination As DESTINATION
Dim rs As New Recordset

Set Lobj_Destination = New DESTINATION
Set rs = Lobj_Destination.Get_toutDestPLNG(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

If Not rs.EOF Then
    While Not rs.EOF
        With SBC_Dest
            .AddItem rs("Numero"), rs("Libelle")
        End With
        rs.MoveNext
    Wend
End If
End Sub

Private Sub initgrid_Detail()
    With grid_detail
        .Redraw = False
        .ClearSelection
        .ClearRows
        .Clear
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
        .AddColumn "Conducteur", "Conducteur", , , 200, , , , , , , CCLSortString
        .AddColumn "Destination", "Destination", , , 160
        .AddColumn "Date", "Date", , , 150, , , , , , , CCLSortDate
        .AddColumn "NULL", ""
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub

Private Sub grid_detail_ColumnClick(ByVal lCol As Long)
    Dim sTag As String
    Dim i As Long
   
        With grid_detail.SortObject
           .Clear
           .SortColumn(1) = lCol
        
           sTag = grid_detail.ColumnTag(lCol)
           If (sTag = "") Then
              sTag = "DESC"
              .SortOrder(1) = CCLOrderAscending
           Else
              sTag = ""
              .SortOrder(1) = CCLOrderDescending
           End If
           grid_detail.ColumnTag(lCol) = sTag
        
            Select Case grid_detail.ColumnKey(lCol)
                Case "Conducteur"
                     .SortType(1) = CCLSortString
'                Case "Destination"
'                     .SortType(1) = CCLSortString
                Case "Date"
                     .SortType(1) = CCLSortDate

           End Select
        End With
        Screen.MousePointer = vbHourglass
        grid_detail.Sort
        Screen.MousePointer = vbDefault
End Sub

Private Sub Pic_MaskMenu_Click()
    Pic_MaskMenu.Visible = False
    Pic_ShowMenu.Visible = True
    Tab_FindView.Visible = False
End Sub

Private Sub Pic_ShowMenu_Click()
    If Tab_FindView.Visible = False Then
        Pic_MaskMenu.Visible = True
        Pic_ShowMenu.Visible = False
        Tab_FindView.Visible = True
        Pic_ConsltConge.Visible = True
        
    Else
        Pic_MaskMenu.Visible = False
        Pic_ShowMenu.Visible = True
        Tab_FindView.Visible = False
    End If
End Sub

