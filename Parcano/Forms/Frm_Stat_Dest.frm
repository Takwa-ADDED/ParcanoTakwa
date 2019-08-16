VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Stat_Dest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statistiques Destination"
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   15900
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab Tab_Satistiques 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Statistiques Destinations"
      TabPicture(0)   =   "Frm_Stat_Dest.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Pic_ControlStatR"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox Pic_ControlStatR 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   120
         ScaleHeight     =   7815
         ScaleWidth      =   14415
         TabIndex        =   2
         Top             =   240
         Width           =   14415
         Begin VB.PictureBox Pic_Details_Dest 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   6735
            Left            =   120
            ScaleHeight     =   6705
            ScaleWidth      =   11865
            TabIndex        =   12
            Top             =   720
            Width           =   11895
            Begin SToolBox.SGrid SGrid_Details 
               Height          =   5295
               Left            =   120
               TabIndex        =   13
               Top             =   1080
               Width           =   11655
               _ExtentX        =   20558
               _ExtentY        =   9340
               RowMode         =   -1  'True
               BackgroundPictureHeight=   0
               BackgroundPictureWidth=   0
               BackColor       =   16777152
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   2
               DisableIcons    =   -1  'True
               SelectionAlphaBlend=   -1  'True
               SelectionOutline=   -1  'True
               MaxVisibleRows  =   0
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Retour"
               BeginProperty Font 
                  Name            =   "Sitka Text"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   10920
               TabIndex        =   16
               Top             =   6360
               Width           =   735
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Height          =   375
               Left            =   2280
               TabIndex        =   15
               Top             =   480
               Width           =   4335
            End
            Begin VB.Label Label2 
               Caption         =   "Destination:"
               BeginProperty Font 
                  Name            =   "Sitka Small"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   360
               TabIndex        =   14
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.ComboBox Cbo_Destination 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "Frm_Stat_Dest.frx":001C
            Left            =   1800
            List            =   "Frm_Stat_Dest.frx":001E
            TabIndex        =   4
            Top             =   240
            Width           =   3735
         End
         Begin MSComctlLib.ListView List_Details1 
            Height          =   6735
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   11880
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Numero"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Destination"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Nombre Tournées"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Key             =   "valeur"
               Text            =   "Total Distance/KM"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Total Durée/H:M"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Moyenne Distance/KM"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Moyenne Durée/H"
               Object.Width           =   2822
            EndProperty
         End
         Begin SToolBox.SCommand cmd_FindDest 
            Height          =   345
            Left            =   5760
            TabIndex        =   5
            Top             =   240
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
            Picture         =   "Frm_Stat_Dest.frx":0020
            ButtonType      =   1
         End
         Begin MSComCtl2.DTPicker Dta_Fin 
            Height          =   375
            Left            =   9600
            TabIndex        =   6
            Top             =   240
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
            Format          =   303038465
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker Dta_Debut 
            Height          =   375
            Left            =   7080
            TabIndex        =   7
            Top             =   240
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
            Format          =   303038465
            CurrentDate     =   42860
         End
         Begin VB.Image Cmd_Find 
            Height          =   495
            Left            =   11880
            Picture         =   "Frm_Stat_Dest.frx":035A
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destination"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   6480
            TabIndex        =   9
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   9000
            TabIndex        =   8
            Top             =   240
            Width           =   600
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statistiques Destinations"
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
      Left            =   960
      TabIndex        =   11
      Top             =   360
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   0
      Picture         =   "Frm_Stat_Dest.frx":10F5C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image PicBox_Header 
      Height          =   1005
      Left            =   0
      Picture         =   "Frm_Stat_Dest.frx":2E6B6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15855
   End
   Begin VB.Label m 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statistiques"
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
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "Frm_Stat_Dest.frx":6A1D0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Frm_Stat_Dest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Dim VCodeDestination  As String         'Code Destination
Dim itmX As ListItem



Private Sub Cmd_Find_Click()
Pic_Details_Dest.Visible = False

  Dim VCode As String, LObj_V As New DESTINATION, Lrs_V As New Recordset
On Error GoTo Err
    If Dta_Debut.Value > Dta_Fin.Value Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    List_Details1.ListItems.Clear
  
'    VCode = Cbo_Destination.Text
'     If VCode <> "Tous" And VCode <> "" Then
'        Set Lrs_V = LObj_V.GetDestByLibDest(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
'        If ErrNumber <> 0 Then
'            ErrNumber = 0
'            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
'            Exit Sub
'        End If
'        Set LObj_V = Nothing
'        If Not Lrs_V.EOF Then
'            VCode = Lrs_V("Numero")
'            Set Lrs_V = Nothing
'        Else
'            MsgBox "Destination introuvable!..."
'            Set Lrs_V = Nothing
'            Exit Sub
'        End If
'    End If
    If Cbo_Destination.Text = "Tous" Then
        Call AfficheDetails_Tous_Dest(Dta_Debut.Value, Dta_Fin.Value)
    Else
        Call AfficheDetails_ParDestination(Cbo_Destination.Text, CDate(Dta_Debut.Value), CDate(Dta_Fin.Value))
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub

Public Sub AfficheDetails_Tous_Dest(ByVal VdateD As Date, ByVal vDateF As Date)
    'variables globales
    Dim LOBJ_StatRp As DESTINATION
    Dim rs As New Recordset
    Dim i
    'variables Details
    Dim nbVoy As Double
    Dim Valeur As Double
    Dim duree As Double
    Dim dureeM As Long
    Dim heurmin As Long
    Dim heur As Long
    Dim min As Long
    Dim Moy_Duree As Long
    Dim Moy_DureeM As Long
    Dim Calcul As Long
    Dim HeurMoy As Long
    Dim MinMoy As Long
    Dim Moy_Dist As Double
On Error GoTo Err

    Set itmX = List_Details1.ListItems.Add()
    
    Set LOBJ_StatRp = New DESTINATION
    'nombre des Voyages
    Set rs = LOBJ_StatRp.Get_SumNbrVoyage(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
    
    itmX.SubItems(1) = "Tous"
        nbVoy = 0
        nbVoy = nbVoy + rs("nbVoy")
        
        itmX.SubItems(2) = CStr(rs("nbVoy"))
        'Total Distance
        If Not IsNull(rs("valeur")) Then
            Valeur = 0
            Valeur = Valeur + rs("valeur")
            Moy_Dist = Valeur / nbVoy
            itmX.SubItems(3) = CStr(Format(rs("valeur"), "#,##0.0"))
            itmX.SubItems(5) = Format(Moy_Dist, "#,##0.0")
        Else
            itmX.SubItems(3) = "Valeur Null"
            itmX.SubItems(5) = "Valeur Null"
        End If
        'Total Durée
        If Not IsNull(rs("totduree")) And Not IsNull(rs("totdureeM")) Then
            duree = 0
            dureeM = 0
'            duree = duree + rs("totduree")
            dureeM = dureeM + rs("totdureeM")
            heur = rs("totdureeM") \ 60
            heurmin = 60 * heur
            min = dureeM - heurmin
'            Moy_Duree = duree / nbVoy
            Moy_DureeM = dureeM \ nbVoy
            HeurMoy = Moy_DureeM \ 60
            Calcul = HeurMoy * 60
            MinMoy = Moy_DureeM - Calcul
'            itmX.SubItems(4) = CStr(Format(rs("totduree"), "#,##0.0"))
            itmX.SubItems(4) = CStr(heur) & " : " & CStr(Format(min, "##00"))
            itmX.SubItems(6) = CStr(HeurMoy) & " : " & CStr(Format(MinMoy, "##00"))
        Else
            itmX.SubItems(4) = "Duree Null"
            itmX.SubItems(6) = "Moyenne Duree Null"
        End If
    End If
    rs.Close
'    'Destination + nbr voyage par destination
    Set rs = LOBJ_StatRp.Get_NbrRepStatistGrpDest(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
'    Numero DESTINATION
        
    If Not rs.EOF Then
        While Not rs.EOF
        
            
            If Not IsNull(rs("nbVoy")) Then
            nbVoy = 0
            nbVoy = nbVoy + rs("nbVoy")
            Set itmX = List_Details1.ListItems.Add(, , rs("Numero"))
            itmX.SubItems(1) = CStr(rs("Libelle"))
            itmX.SubItems(2) = CStr(rs("nbVoy"))
            End If
      'Total Distance
            If Not IsNull(rs("valeur")) Then
                Valeur = 0
                Valeur = Valeur + rs("valeur")
                Moy_Dist = Valeur / nbVoy
                itmX.SubItems(3) = CStr(Format(rs("valeur"), "#,##0.0"))
                itmX.SubItems(5) = Format(Moy_Dist, "#,##0.0")
            Else
                itmX.SubItems(3) = "Valeur Null"
                itmX.SubItems(5) = "Valeur Null"
            End If
      'Total Durée
        If Not IsNull(rs("totduree")) Then
           duree = 0
            dureeM = 0
'            duree = duree + rs("totduree")
            dureeM = dureeM + rs("totdureeM")
            heur = rs("totdureeM") \ 60
            heurmin = 60 * heur
            min = dureeM - heurmin
'            Moy_Duree = duree / nbVoy
           Moy_DureeM = dureeM \ nbVoy
            HeurMoy = Moy_DureeM \ 60
            Calcul = HeurMoy * 60
            MinMoy = Moy_DureeM - Calcul
           itmX.SubItems(4) = CStr(heur) & " : " & CStr(Format(min, "##00"))
             itmX.SubItems(6) = CStr(HeurMoy) & " : " & CStr(Format(MinMoy, "##00"))
        Else
            itmX.SubItems(4) = "Duree Null"
            itmX.SubItems(6) = "Moyenne Duree Null"
        End If
            
            
        rs.MoveNext
        Wend
    End If
    rs.Close
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheDetails_ParDestination(ByVal Libelle As String, ByVal VdateD As Date, ByVal vDateF As Date)
    Dim LOBJ_StatRp As New DESTINATION
    Dim rs As New Recordset
   
    'variables Details
   Dim nbVoy As Double
    Dim Valeur As Double
'    Dim duree As Double
    Dim dureeM As Long
    Dim heurmin As Long
    Dim heur As Long
    Dim min As Long
'    Dim Moy_Duree As Long
    Dim Moy_DureeM As Long
    Dim Calcul As Long
    Dim HeurMoy As Long
    Dim MinMoy As Long
    Dim Moy_Dist As Double
On Error GoTo Err

    'nombre des Voyages
    Set rs = LOBJ_StatRp.Get_ValRepStatistDest(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Libelle, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        Set itmX = List_Details1.ListItems.Add(, , rs("Numero"))
        itmX.SubItems(1) = CStr(rs("Libelle"))
    
        If Not IsNull(rs("nbVoy")) Then
            nbVoy = 0
            nbVoy = nbVoy + rs("nbVoy")
            
            itmX.SubItems(2) = CStr(rs("nbVoy"))
            Else
            itmX.SubItems(2) = "Valeur Null"
        End If
        'Total Distance
        If Not IsNull(rs("valeur")) Then
            Valeur = 0
            Valeur = Valeur + rs("valeur")
            Moy_Dist = Valeur / nbVoy
            itmX.SubItems(3) = CStr(Format(rs("valeur"), "#,##0.0"))
            itmX.SubItems(5) = Format(Moy_Dist, "#,##0.0")
        Else
            itmX.SubItems(3) = "Valeur Null"
            itmX.SubItems(5) = "Valeur Null"
        End If
        'Total Durée
        If Not IsNull(rs("totdureeM")) Then
            
            dureeM = 0
            dureeM = dureeM + rs("totdureeM")
            heur = rs("totdureeM") \ 60
            heurmin = 60 * heur
            min = dureeM - heurmin
            Moy_DureeM = dureeM \ nbVoy
            HeurMoy = Moy_DureeM \ 60
            Calcul = HeurMoy * 60
            MinMoy = Moy_DureeM - Calcul
            itmX.SubItems(4) = CStr(heur) & " : " & CStr(Format(min, "##00"))
            itmX.SubItems(6) = CStr(HeurMoy) & " : " & CStr(Format(MinMoy, "##00"))
        Else
            itmX.SubItems(4) = "Duree Null"
            itmX.SubItems(6) = "Moyenne Duree Null"
        End If
    End If
    rs.Close
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub cmd_FindDest_Click()
Pic_Details_Dest.Visible = False
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "STATDestinations"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub



Private Sub Dta_Debut_Click()
Pic_Details_Dest.Visible = False
End Sub



Private Sub Dta_Fin_Click()
Pic_Details_Dest.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Pic_Details_Dest.Visible = False
End Sub

Private Sub Form_Load()
    Pic_Details_Dest.Visible = False
    Call Initgrid_Details_Dest
    Dta_Debut.Value = "01/" & Month(Date) & "/" & Year(Date)
    Dta_Fin.Value = Date

  
    Cbo_Destination.AddItem ("Tous"), 0
   
    Call Affiche_Destination_Combo(Cbo_Destination)
  
    Cbo_Destination.ListIndex = 0
   
    VCodeDestination = "  -  Tous"
End Sub


'Private Sub Cbo_Destination_Change()
'Dim i As Integer, start As Integer
'    Dim ShiftDown As Boolean
'    Dim CtrlDown As Boolean
'    Dim AltDown As Boolean
'    ShiftDown = (theshift And vbShiftMask) > 0
'    CtrlDown = (theshift And vbCtrlMask) > 0
'    AltDown = (theshift And vbAltMask) > 0
'    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
'        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
'    Else
'        start = Len(Cbo_Destination.Text)
'        For i = 0 To Cbo_Destination.ListCount - 1
'            If Left(Cbo_Destination.List(i), start) = Cbo_Destination.Text Then
'                Cbo_Destination.Text = Cbo_Destination.List(i)
'            End If
'        Next
'        Cbo_Destination.SelStart = start
'        Cbo_Destination.SelLength = Len(Cbo_Destination.Text)
'    End If
'End Sub

Private Sub Cbo_Destination_Click()
List_Details1.ListItems.Clear
Pic_Details_Dest.Visible = False
Dim LOBJ_Dest As New DESTINATION
    Dim Lrs_Dest As Recordset
    If Cbo_Destination.ListIndex = 0 Then
        VCodeDestination = "  -  Tous"
    Else
        Set Lrs_Dest = LOBJ_Dest.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, Cbo_Destination.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LOBJ_Dest = Nothing
        
        If Not Lrs_Dest.EOF Then VCodeDestination = Lrs_Dest("Numero")
    End If
End Sub


Public Sub AfficheRow_Destination(ByVal VCode As String)
    Dim LObj_Find As New DESTINATION
    Dim Lrs_Find As New Recordset
    Dim cboDest As ComboBox
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        'Charge
        If Frm_FindView.StrSource = "STATDestinations" Then Set cboDest = Cbo_Destination
        If Not IsNull(Lrs_Find("Libelle")) Then
            cboDest.Text = Lrs_Find("Libelle")
            VCodeDestination = Lrs_Find("Numero")
        End If
    End If
    Lrs_Find.Close
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub Label4_Click()
Pic_Details_Dest.Visible = False
End Sub

Private Sub List_Details1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Pic_Details_Dest.Visible = False
 Call ListView_Header_Click(List_Details1, ColumnHeader)
End Sub
Public Sub ListView_Header_Click(LView As ListView, ColumnHeader As MSComctlLib.ColumnHeader)
            LView.Sorted = True
            LView.SortOrder = Abs(Not -LView.SortOrder)
            LView.SortKey = ColumnHeader.Index - 1
            LView.Sorted = False
    End Sub

Private Sub List_Details1_DblClick()
Dim VCode
Dim DateD As Date
Dim DateF As Date
Dim i As Integer
On Error GoTo Err
    i = List_Details1.SelectedItem.Index
    VCode = List_Details1.ListItems.Item(i)
    DateD = Dta_Debut.Value
    DateF = Dta_Fin.Value
    Call AfficheDetailsDes(VCode, DateD, DateF)
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub

Public Sub AfficheDetailsDes(ByVal VCode As String, ByVal DateD As Date, ByVal DateF As Date)
    Dim LOBJ_Dest   As New DESTINATION
    Dim Lrs_Dest    As New Recordset
On Error GoTo Err
    
    SGrid_Details.ClearRows
    
    Set Lrs_Dest = LOBJ_Dest.GetRow_Details_Dest(ErrNumber, ErrDescription, ErrSourceDetail, VCode, DateD, DateF, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    
    Set LOBJ_Dest = Nothing

If Not Lrs_Dest.EOF Then
Label3.Caption = Lrs_Dest("Destination")
        While Not Lrs_Dest.EOF
        SGrid_Details.Redraw = False
            With SGrid_Details
            .AddRow
            .CellDetails .Rows, .ColumnIndex("Vehicule"), Lrs_Dest.Fields("Vehicule"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("Conducteur"), Lrs_Dest.Fields("Conducteur"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("CompteurSortie"), Lrs_Dest.Fields("CompteurSortie"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("CompteurEntre"), Lrs_Dest.Fields("CompteurEntre"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("HeureSortie"), Lrs_Dest.Fields("HeureSortie"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("HeureEntre"), Lrs_Dest.Fields("HeureEntre"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("OperateurSortie"), Lrs_Dest.Fields("OperateurSortie"), DT_LEFT
            .CellDetails .Rows, .ColumnIndex("OperateurEntre"), Lrs_Dest.Fields("OperateurEntre"), DT_LEFT
                    
            End With
        Lrs_Dest.MoveNext
        Wend
        
        SGrid_Details.Redraw = True
        Pic_Details_Dest.Visible = True
        cmd_FindDest.SetFocus
        
        If SGrid_Details.Rows > 0 Then
            SGrid_Details.SelectedRow = 1
            SGrid_Details.SetFocus
        End If

    
    Else
        MsgBox "Veuillez Choisir une Destination!", vbInformation
     
    End If
On Error Resume Next
    Set Lrs_Dest = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Sub

Private Sub Initgrid_Details_Dest()
With SGrid_Details
.Redraw = False

    

    .HideGroupingBox = True
    .AllowGrouping = True
 
    .GroupRowBackColor = vbWindowBackground
    .GroupRowForeColor = vbWindowText
    
    .GridLineColor = &HE0E0E0
    .GridFillLineColor = vbBlue
    .GridLines = True
    
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .DrawFocusRectangle = False
    
 
    .AddColumn "Numero", "Numero", , , , False
    .AddColumn "Vehicule", "Vehicule", , , 100, , , , , , , CCLSortString
    .AddColumn "Conducteur", "Conducteur", , , 100
    .AddColumn "CompteurSortie", "CPT Sortie", , , 60
    .AddColumn "CompteurEntre", "CPT Entrée", , , 60
    .AddColumn "HeureSortie", "HeureSortie", , , 100
    .AddColumn "HeureEntre", "HeureEntre", , , 100
    .AddColumn "OperateurSortie", "OperateurSortie", , , 100
    .AddColumn "OperateurEntre", "OperateurEntre", , , 100
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
.Redraw = True
End With

End Sub

Private Sub List_Details1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Pic_Details_Dest.Visible = False
End Sub

Private Sub Pic_Details_Dest_LostFocus()
Pic_Details_Dest.Visible = False

End Sub

Private Sub SGrid_Details_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
If KeyCode = vbKeyEscape Then Pic_Details_Dest.Visible = False
End Sub

Private Sub SGrid_Details_ColumnClick(ByVal lCol As Long)

Dim sTag As String
    Dim i As Long
    With SGrid_Details.SortObject
        .Clear
        .SortColumn(1) = lCol
        sTag = SGrid_Details.ColumnTag(lCol)
        If (sTag = "") Then
            sTag = "DESC"
            .SortOrder(1) = CCLOrderAscending
        Else
            sTag = ""
            .SortOrder(1) = CCLOrderDescending
        End If
        SGrid_Details.ColumnTag(lCol) = sTag
        Select Case SGrid_Details.ColumnKey(lCol)
            Case "Numero"
                 .SortType(1) = CCLSortNumeric
            Case "Vehicule"
                 .SortType(1) = CCLSortString
            Case "Conducteur"
                 .SortType(1) = CCLSortString
            Case "CompteurSortie"
                 .SortType(1) = CCLSortNumeric
            Case "CompteurEntre"
                 .SortType(1) = CCLSortNumeric
            Case "HeureSortie"
                 .SortType(1) = CCLSortDate
            Case "HeureEntre"
                 .SortType(1) = CCLSortDate
            Case "OperateurSortie"
                 .SortType(1) = CCLSortString
            Case "OperateurEntre"
                 .SortType(1) = CCLSortString
        End Select
    End With
    Screen.MousePointer = vbHourglass
    SGrid_Details.Sort
    Screen.MousePointer = vbDefault
End Sub
