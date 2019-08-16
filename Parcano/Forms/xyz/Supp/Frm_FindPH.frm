VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_FindPH 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "."
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8100
   ClipControls    =   0   'False
   Icon            =   "Frm_FindPH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PixBox_Lister 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   720
      ScaleHeight     =   2955
      ScaleWidth      =   6795
      TabIndex        =   3
      Top             =   1080
      Width           =   6855
      Begin SToolBox.SCheckBox ChBox_Supprimer 
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         Caption         =   "SCheckBox1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
      End
      Begin SToolBox.SOptionButton Obtn_Tous 
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "SOptionButton1"
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
      Begin VB.PictureBox PicBox_Date 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2640
         ScaleHeight     =   375
         ScaleWidth      =   3495
         TabIndex        =   8
         Top             =   1440
         Width           =   3495
         Begin MSComCtl2.DTPicker DBox_DDebut 
            Height          =   375
            Left            =   480
            TabIndex        =   20
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632256
            Format          =   142409729
            CurrentDate     =   42859
         End
         Begin MSComCtl2.DTPicker DBox_DFin 
            Height          =   375
            Left            =   2160
            TabIndex        =   21
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632256
            Format          =   142409729
            CurrentDate     =   42859
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "De :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "à :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   9
            Top             =   0
            Width           =   255
         End
      End
      Begin SToolBox.SCommand Cmd_Search 
         Height          =   375
         Left            =   4200
         TabIndex        =   6
         Top             =   2400
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Afficher"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Frm_FindPH.frx":000C
         BackColor       =   8421504
         ForeColor       =   16777215
         ButtonType      =   1
      End
      Begin SToolBox.SCheckBox ChBox_BetweenDate 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Enabled         =   0   'False
         Value           =   1
         Caption         =   "SCheckBox1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
      End
      Begin SToolBox.SCommand Cmd_Masque 
         Height          =   255
         Left            =   5880
         TabIndex        =   7
         Top             =   120
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
      Begin SToolBox.SOptionButton Obtn_AFaire 
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   450
         BackStyle       =   0
         Value           =   1
         Caption         =   "SOptionButton3"
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
      Begin SToolBox.SBiCombo Cbo_Conducteur 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   1920
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.Label Lbl_cond 
         BackStyle       =   0  'Transparent
         Caption         =   "Conducteur :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Line Line9 
         X1              =   120
         X2              =   0
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label LblAFaire 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   4935
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   120
         Y1              =   840
         Y2              =   1320
      End
      Begin VB.Line Line7 
         X1              =   1560
         X2              =   2880
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   1560
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line5 
         X1              =   1560
         X2              =   1560
         Y1              =   840
         Y2              =   1320
      End
      Begin VB.Line Line4 
         X1              =   2880
         X2              =   2880
         Y1              =   840
         Y2              =   1320
      End
      Begin VB.Line Line3 
         X1              =   1560
         X2              =   2880
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line2 
         X1              =   2880
         X2              =   6840
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Supprime..."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   1560
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "A faire..."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tous..."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Afficher Liste des programmes"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   4815
      End
      Begin VB.Image Image2 
         Height          =   135
         Left            =   -120
         Picture         =   "Frm_FindPH.frx":0527
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   6975
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   0
         Picture         =   "Frm_FindPH.frx":22AD
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date de programme:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   2055
      End
   End
   Begin SToolBox.SCommand Cmd_MasqueListe 
      Height          =   255
      Left            =   7560
      TabIndex        =   0
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
   Begin SToolBox.SGrid grid1 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
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
   Begin VB.Label Lbl_Msg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   720
      Width           =   3375
   End
   Begin VB.Image PicBox_Menu 
      Height          =   255
      Left            =   6960
      Picture         =   "Frm_FindPH.frx":35647
      Stretch         =   -1  'True
      Top             =   840
      Width           =   615
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
      TabIndex        =   2
      Top             =   240
      Width           =   3330
   End
   Begin VB.Image PicBox_Header 
      Height          =   975
      Left            =   0
      Picture         =   "Frm_FindPH.frx":35F99
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "Frm_FindPH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------------
'Déclaration des Variables Global***
'----------------------------------------------------------------------------------------------------------------------------------
    Public StrSource As String
    Public ViewSupp As String
    Public RETOUR As Integer

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Form_Load()
On Error GoTo Err
    Call InitialiseControlBox
    LBL_Titre.Caption = "Liste des " & StrSource & "s"
    Select Case StrSource
'        Case "ProgChauffeurs"
'            PicBox_Menu.Visible = True
'            Cbo_Conducteur.AddItem "0000", "Tous"
'            Call Affiche_Personnel_SBCombo(Cbo_Conducteur)
'            Call Initgrid1_ProgChauffeur
'            Call Affiche_ProgChauffeursAvecDetails(Date, Date, 4, ViewSupp, Cbo_Conducteur.FirstValue)
'            LBL_Titre.Caption = "Liste des Programme"
            
'        Case "FournisseurPH"
'            Call Initgrid1_FournisseurAchat
'            Call Affiche_FournisseurAchat
            
'        Case "VehiculePH"
'            Call Initgrid1_Vehicule
'            Call Affiche_Vehicule
            
'        Case "VehiculeSup"
'            Call Initgrid1_Vehicule
'            Call Affiche_Vehicule
            
'        Case "ConducteurPing"
'            LBL_Titre.Caption = "Liste des Conducteurs Actif"
'            Me.Caption = "Liste des Conducteurs | " & App.ProductName
'            Call Initgrid1_PersonnelActifDisp
'            Call Affiche_PersonnelActifDisp
        
'        Case "VehiculePing"
'            LBL_Titre.Caption = "Liste des Vehicules Actif"
'            Me.Caption = "Liste des Véhicules | " & App.ProductName
'            Call Initgrid1_Vehicule
'            Call Affiche_Vehicule
            
        End Select
         
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub Form_Activate()
    If grid1.Rows = 0 Then MsgBox "Pas de données à visualiser", vbInformation
End Sub

'----------------------------------------------------------------------------------------------------------------------------------
'Subilations Chargement Detail par SGrid***
'----------------------------------------------------------------------------------------------------------------------------------
Private Sub grid1_ColumnClick(ByVal lCol As Long)
    Dim sTag As String
    Dim i As Long

    PixBox_Lister.Visible = False
   
    If (StrSource = "BLPieceReparation" Or StrSource = "FacturePieceReparation") Then
        With grid1.SortObject
           .Clear
           .SortColumn(1) = lCol
        
           sTag = grid1.ColumnTag(lCol)
           If (sTag = "") Then
              sTag = "DESC"
              .SortOrder(1) = CCLOrderAscending
           Else
              sTag = ""
              .SortOrder(1) = CCLOrderDescending
           End If
           grid1.ColumnTag(lCol) = sTag
        
            Select Case grid1.ColumnKey(lCol)
                Case "Numero"
                     .SortType(1) = CCLSortNumeric
                Case "DatePiece"
                     .SortType(1) = CCLSortDate
                Case "DateOperation"
                     .SortType(1) = CCLSortDate
                Case "Fournisseur"
                     .SortType(1) = CCLSortString
           End Select
        End With
        Screen.MousePointer = vbHourglass
        grid1.Sort
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub grid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    Dim VCode

    PixBox_Lister.Visible = False
    
On Error GoTo Err

    VCode = grid1.CellText(lRow, 1)
    Select Case StrSource
        
        Case "FournisseurAchat"
            Unload Me
            Frm_PrgChauf.AfficheRowProgrammeCH (VCode)
            
        Case "ProgChauffeurs"
            VCode = grid1.CellText(lRow, 2)
            If VCode <> "" Then
                Unload Me
                Frm_PrgChauf.AfficheRowProgrammeCH (VCode)
            End If
            
        Case "FournisseurPH"
            Unload Me
            Frm_PrgChauf.AfficheRowFournisseurPH (VCode)
            
        Case "VehiculePH"
            Unload Me
            Frm_PrgChauf.AfficheRowVehiculePH (VCode)
            
        Case "VehiculeSup"
            Unload Me
            Frm_Supervision.AfficheRowVehiculeSup (VCode)
            
        Case "VehiculePing"
            Unload Me
            Frm_PLANNING.AfficheRowVehiculePing (VCode)
            
        Case "ConducteurPing"
            Unload Me
            Frm_PLANNING.AfficheRowConducteurPing (VCode)
            
    End Select

Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    Dim VCode
    
On Error GoTo Err

    VCode = grid1.CellText(grid1.SelectedRow, 1)
    Select Case KeyCode
        Case vbKeyF2, vbKeyReturn
            Select Case StrSource
            
                Case "ProgChauffeurs"
                    VCode = grid1.CellText(grid1.SelectedRow, 2)
                    If VCode <> "" Then
                        Unload Me
                        Frm_PrgChauf.AfficheRowProgrammeCH (VCode)
                    End If
                    
                Case "VehiculePH"
                    Unload Me
                    Frm_PrgChauf.AfficheRowVehiculePH (VCode)
                    
                Case "FournisseurPH"
                    Unload Me
                    Frm_PrgChauf.AfficheRowFournisseurPH (VCode)
                    
                Case "VehiculePing"
                    Unload Me
                    Frm_PLANNING.AfficheRowVehiculePing (VCode)
                    
                Case "ConducteurPing"
                    Unload Me
                    Frm_PLANNING.AfficheRowConducteurPing (VCode)
                    
                Case vbKeyEscape
                    Unload Me
            End Select
    End Select

Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'----------------------------------------------------------------------------------------------------------------------------------
'Subilations Affiche Details dans SGrid***
'----------------------------------------------------------------------------------------------------------------------------------
Public Sub Affiche_Vehicule()
    Dim Lobj_Vehicule As New VEHICULE
    Dim rs As New Recordset
    
    Set rs = Lobj_Vehicule.GetAll_VehiculeActifNonSupp(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
        
    If Not rs.EOF Then
        grid1.Redraw = False
        While Not rs.EOF
            With grid1
                .AddRow
                .CellDetails .Rows, 1, rs("Code")
                .CellDetails .Rows, .ColumnIndex("Libelle"), rs("Matricule")
                .CellDetails .Rows, .ColumnIndex("Marque"), rs("Marque")
                .CellDetails .Rows, .ColumnIndex("Type"), rs("Type")
                .CellDetails .Rows, .ColumnIndex("Energie"), rs("Energie")
                .CellDetails .Rows, .ColumnIndex("Puissance"), rs("Puissance")
                If StrSource <> "VehiculePH" And StrSource <> "VehiculePing" Then .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")
            End With
            rs.MoveNext
        Wend
        grid1.Redraw = True
    End If
End Sub
Public Sub Affiche_FournisseurAchat()
    Dim LOBJ_Stat As Station
    Dim rs As New Recordset
    
    Set LOBJ_Stat = New Station
    
    Set rs = LOBJ_Stat.Get_FournisAchat(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    
    If Not rs.EOF Then
        grid1.Redraw = False
        While Not rs.EOF
            With grid1
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
        grid1.Redraw = True
    End If
End Sub
Public Sub Affiche_ProgChauffeursAvecDetails(ByVal Ddebut As String, ByVal Dfin As String, ByVal Param As Integer, ByVal ViewSupp As String, ByVal cond As String)
    Dim LObj_Find As ProgChauf
    Dim Lrs As Recordset

    Set LObj_Find = New ProgChauf
    Set Lrs = LObj_Find.GetRow_ProgramChauffeur(ErrNumber, ErrDescription, ErrSourceDetail, Ddebut, Dfin, Param, ViewSupp, cond, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    
    If Not Lrs.EOF Then
        grid1.Redraw = False
        While Not Lrs.EOF
        Dim AssProg As String
        AssProg = Lrs.Fields("DateProgramme") & "    Conducteur :  " & Lrs.Fields("conducteur") & "    Véhicule :  " & Lrs.Fields("Matricule")
        If Lrs.Fields("Supp") = "O" Then AssProg = Lrs.Fields("DateProgramme") & "    Conducteur :  " & Lrs.Fields("conducteur") & "    Véhicule :  " & Lrs.Fields("Matricule") & "            || ** Programme Supprimer **"
            With grid1
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Code"), AssProg, , , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("CodeProg"), Lrs("Code")
                .CellDetails .Rows, .ColumnIndex("Order"), Lrs.Fields("ProgOrder"), DT_CENTER, , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("Fournisseur"), Lrs.Fields("Fournisseur"), , , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("TxtCommande"), Lrs.Fields("TxtCommande"), , , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("TxtPaiement"), Lrs.Fields("TxtPaiement"), , , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("TxtObservation"), Lrs.Fields("TxtObservation"), , , RGB(225, 237, 226)
                .CellDetails .Rows, .ColumnIndex("NULL"), "", , , RGB(225, 237, 226)
            End With
            Lrs.MoveNext
        Wend
        grid1.Redraw = True
        Set Lrs = Nothing
    
        With grid1
            .GroupRowBackColor = RGB(251, 246, 206)
            .GroupRowForeColor = QBColor(12)
            .ColumnIsGrouped(1) = True
            .GroupRowForeColor = QBColor(10)
            .HideGroupingBox = True
            .AllowGrouping = True
'            Call expandAllGroups(grid1)
       End With
    End If
    If grid1.Rows > 0 Then grid1.SelectedRow = 1
End Sub
Public Sub Affiche_PersonnelActifDisp()
    Dim LOBJ_Pers As New CONDUCTEUR
    Dim rs As New Recordset
    
    Set rs = LOBJ_Pers.GetAll_ConducteursActifNonSupp(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    
    If Not rs.EOF Then
        grid1.Redraw = False
        While Not rs.EOF
            With grid1
                .AddRow
                .CellDetails .Rows, 1, rs("Code")
                .CellDetails .Rows, .ColumnIndex("Libelle"), rs("libelle")
                If StrSource <> "ConducteurPH" And StrSource <> "ConducteurPing" Then .CellDetails .Rows, .ColumnIndex("Actif"), rs("Actif")
            End With
            rs.MoveNext
        Wend
        grid1.Redraw = True
    End If
End Sub
'----------------------------------------------------------------------------------------------------------------------------------
'Subilations Initialise SGrid***
'----------------------------------------------------------------------------------------------------------------------------------
Public Sub Initgrid1_Vehicule()
    With grid1
        .Redraw = False
        ' Allow the grid1 to be grouped, but
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
        
        .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
        .AddColumn "Libelle", "Matricule", , , 140
        .AddColumn "Marque", "Marque", , , 40
        .AddColumn "Type", "Type", eSortType:=CCLSortStringNoCase
        .AddColumn "Energie", "Energie", eSortType:=CCLSortStringNoCase
        .AddColumn "Puissance", "Puissance", sFmtString:="short date", eSortType:=CCLSortDateDayAccuracy
        If StrSource <> "VehiculePH" And StrSource <> "VehiculePing" Then .AddColumn "Actif", "Acif", , , 40
        
        .AddColumn "Q", ""
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Private Sub Initgrid1_PersonnelActifDisp()
    With grid1
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
        If StrSource <> "ConducteurPH" And StrSource <> "ConducteurPing" Then .AddColumn "Actif", "Actif", , , 40
    
        
        .AddColumn "Q", ""
        .StretchLastColumnToFit = True
    End With
End Sub
Private Sub Initgrid1_FournisseurAchat()
    With grid1
        .Redraw = False
        ' Allow the grid1 to be grouped, but
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
Public Sub Initgrid1_ProgChauffeur()
    With grid1
        .Redraw = False
        .ClearSelection
        .ClearRows
        .Clear

        ' Allow the grid1 to be grouped, but
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
        
        .AddColumn "Code", ""
        .AddColumn "CodeProg", "Code", , , , False
        
        .AddColumn "Order", "Order", , , 60
        .AddColumn "Fournisseur", "Fournisseur", , , 150
        .AddColumn "TxtCommande", "Commande", , , 150
        .AddColumn "TxtPaiement", "Paiement", , , 150
        .AddColumn "TxtObservation", "Observation", , , 150
        .AddColumn "NULL", ""

        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub

'----------------------------------------------------------------------------------------------------------------------------------
'Subilations Liste de Rechecrhe***
'----------------------------------------------------------------------------------------------------------------------------------
Rem Afficher Liste de Recherche***
Private Sub PicBox_Menu_Click()
    If PixBox_Lister.Visible = True Then
        PixBox_Lister.Visible = False
    Else
        PixBox_Lister.Visible = True
'        Call ChBox_BetweenDate_Click
        DBox_DDebut.Value = Date
        DBox_DFin.Value = Date
    End If
End Sub

Rem Masque Liste de Recherche***
Private Sub Cmd_Masque_Click()
    PixBox_Lister.Visible = False
End Sub
Private Sub PicBox_Header_Click()
    PixBox_Lister.Visible = False
End Sub
Private Sub Cmd_MasqueListe_Click()
    PixBox_Lister.Visible = False
End Sub
Rem Option de Recherche (Tous, Fait, A faire, Supprimer)***
Private Sub Obtn_Tous_Click()
    Line1.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    Line7.Visible = False
    Line6.Visible = False
    Line8.Visible = False
    LblAFaire.Visible = False
    PicBox_Date.Visible = True
    ChBox_BetweenDate.Visible = True
    Label2.Visible = True
    ChBox_Supprimer.Visible = True
    Label8.Visible = True
    ChBox_Supprimer.Value = vbUnchecked
    ChBox_BetweenDate.Value = vbUnchecked
'    ChBox_BetweenDate_Click
    DBox_DDebut.Value = Date
    DBox_DFin.Value = Date
End Sub
Private Sub Obtn_AFaire_Click()
    Line1.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    Line7.Visible = True
    Line6.Visible = True
    Line8.Visible = True
    LblAFaire.Caption = "Tous les programmes a faire et non supprimer"
    LblAFaire.Visible = True
    PicBox_Date.Visible = False
    ChBox_BetweenDate.Visible = False
    Label2.Visible = False
    ChBox_Supprimer.Visible = False
    Label8.Visible = False
    ChBox_Supprimer.Value = vbUnchecked
    ChBox_BetweenDate.Value = vbUnchecked
End Sub
Rem Serach Program***
Private Sub Cmd_Search_Click()
    Dim ToDayDate As String
    Dim Ddebut As String
    Dim Dfin As String
    Dim Msg As VbMsgBoxResult
    Dim ViewSupp As String
    Dim cond As String
        ToDayDate = Date
        Ddebut = DBox_DDebut.Value
        Dfin = DBox_DFin.Value
    
    Call Initgrid1_ProgChauffeur
    cond = Cbo_Conducteur.FirstValue
    'A Faire...
    If Obtn_AFaire.Value = vbChecked Then
        grid1.ClearRows
        Call Affiche_ProgChauffeursAvecDetails(Ddebut, Dfin, 4, ViewSupp, cond)
        PixBox_Lister.Visible = False
        Lbl_Msg.Visible = False
        If grid1.Rows = 0 Then
            Msg = MsgBox("Aucun programme en attend!..." & vbCr & " Voulez-vous afficher Tous", vbOKCancel + vbInformation, "Information!...")
            If Msg = vbCancel Then Exit Sub
            Call Affiche_ProgChauffeursAvecDetails(Ddebut, Dfin, 0, ViewSupp, cond)
            PixBox_Lister.Visible = False
            Lbl_Msg.Visible = False
        End If
    End If
    'Tous...
    If Obtn_Tous.Value = vbChecked Then
        If ChBox_Supprimer.Value = vbChecked Then
            ViewSupp = "O"
        Else
            ViewSupp = "N"
        End If
         
        If DBox_DDebut.Value < DBox_DFin.Value Then
            Call Affiche_ProgChauffeursAvecDetails(Ddebut, Dfin, 0, ViewSupp, cond)
            PixBox_Lister.Visible = False
            Lbl_Msg.Visible = False
        Else
            MsgBox "Verifier date de recherche"
        End If
    End If
End Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub InitialiseControlBox()
    PicBox_Menu.Visible = False
    PixBox_Lister.Visible = False
    Lbl_Msg.Visible = False
    PicBox_Date.Visible = False
    ChBox_BetweenDate.Visible = False
    Label2.Visible = False
    ChBox_Supprimer.Visible = False
    Label8.Visible = False
    ViewSupp = "N"
    Line1.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    Line7.Visible = True
    Line6.Visible = True
    Line8.Visible = True
    LblAFaire.Visible = True
    LblAFaire.Caption = "-> Tous les programmes a faire et non supprimer"
    End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
