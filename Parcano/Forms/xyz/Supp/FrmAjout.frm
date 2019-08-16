VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAjout 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gestion de Traffic Auto"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
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
   Icon            =   "FrmAjout.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   11400
   Begin SToolBox.SCommand cmd_r 
      Height          =   615
      Left            =   10200
      TabIndex        =   20
      Top             =   720
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
   Begin SToolBox.SCommand SCommand1 
      Height          =   615
      Left            =   7200
      TabIndex        =   19
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      Caption         =   "Alerte"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   255
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   240
   End
   Begin VB.TextBox txt_Observation 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8400
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Txt_KM 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   8400
      MaxLength       =   6
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   3855
      Left            =   11280
      TabIndex        =   6
      Top             =   1200
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   6800
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
      Picture         =   "FrmAjout.frx":000C
   End
   Begin MSComctlLib.ListView LSV_Exterieur 
      Height          =   3015
      Left            =   0
      TabIndex        =   16
      Top             =   5160
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   5318
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
         Text            =   "Matricule"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Conducteur"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Destination"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Heure.S"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Heure.E"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "CPT.S"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "CPT.E"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Distance (KM)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Durée"
         Object.Width           =   2540
      EndProperty
   End
   Begin SToolBox.SGrid grid_vehicule 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5530
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
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin SToolBox.SGrid grid_Conducteur 
      Height          =   3135
      Left            =   2880
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5530
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
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin SToolBox.SGrid grid_destination 
      Height          =   3135
      Left            =   5640
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5530
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
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin MSComctlLib.ListView Lsv_Depot 
      Height          =   3855
      Left            =   -14400
      TabIndex        =   17
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
   Begin SToolBox.SCommand CmdPrint 
      Height          =   615
      Left            =   8640
      TabIndex        =   18
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BackStyle       =   0
      Caption         =   "Compteurs"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmAjout.frx":018E
   End
   Begin VB.Label lbl_Compteur 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Compteur Actuelle !!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   15
      Top             =   1440
      Width           =   2655
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
      TabIndex        =   14
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
      Left            =   5400
      TabIndex        =   13
      Top             =   840
      Width           =   1575
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
      TabIndex        =   12
      Top             =   840
      Width           =   4215
   End
   Begin VB.Image Img_alarme 
      Height          =   600
      Left            =   0
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Observation"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8760
      TabIndex        =   11
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Compteur"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8400
      TabIndex        =   10
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Destination"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Conducteur"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vehicule"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
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
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "FrmAjout"
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


Public Sub AfficheExterieur()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim Couleur As String
Dim datesys As Date

Dim min As Long
Dim heur As Long
Dim Dur As Long
Dim temp As String

datesys = Date
LSV_Exterieur.ListItems.Clear
SQL = " Select * from fichetraffic where  HeureSortie<" & SQLText(datesys) & " And heureEntre is Null Order by HeureEntre ,heureSortie"

rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then

    While Not rs.EOF

            Set itmX = LSV_Exterieur.ListItems.Add(, , "")
            itmX.SubItems(1) = CStr(rs("Numero"))
            itmX.SubItems(2) = rs("vehicule")
            itmX.SubItems(3) = rs("Conducteur")
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
SQL = " Select * from fichetraffic where CONVERT(VARCHAR,HeureSortie,103)=" & SQLText(datesys) & " OR CONVERT(VARCHAR,HeureEntre,103)=" & SQLText(datesys) & " Order by HeureEntre, heureSortie"
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then

    While Not rs.EOF
         Dur = 0
        heur = 0
        min = 0
        temp = ""

            Set itmX = LSV_Exterieur.ListItems.Add(, , "")
            itmX.SubItems(1) = CStr(rs("Numero"))
            itmX.SubItems(2) = rs("vehicule")
            itmX.SubItems(3) = rs("Conducteur")
            itmX.SubItems(4) = rs("Destination")
            itmX.SubItems(5) = Format(rs("heureSortie"), "hh:mm")
            If Not IsNull(rs("HeureENtre")) Then itmX.SubItems(6) = Format(rs("HeureENtre"), "hh:mm")
            If Not IsNull(rs("CompteurSortie")) Then itmX.SubItems(7) = rs("CompteurSortie")
            If Not IsNull(rs("CompteurEntre")) Then itmX.SubItems(8) = rs("CompteurEntre")
            If Not IsNull(rs("CompteurEntre")) Then itmX.SubItems(9) = Val(rs("CompteurEntre")) - Val(rs("CompteurSortie")) & " KM"
            If Not IsNull(rs("HeureENtre")) Then
                'Calcule de durée
                Dur = DateDiff("n", rs("HeureSortie"), rs("HeureEntre"))
                heur = Dur \ 60
                min = Dur - (heur * 60)
                temp = CStr(heur) & ":" & CStr(min)
            itmX.SubItems(10) = temp
            Else
                'Calcule de durée
                Dur = DateDiff("n", rs("HeureSortie"), Now)
                heur = Dur \ 60
                min = Dur - (heur * 60)
                temp = CStr(heur) & ":" & CStr(min)
            itmX.SubItems(10) = temp
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
   
           If (Lsv_Depot.ListItems(i).SubItems(1) = LSV_Exterieur.ListItems(j).SubItems(2)) _
                And (Len(LSV_Exterieur.ListItems(j).SubItems(6)) = 0) Then
           Lsv_Depot.ListItems.Remove (i)
           End If
    Next
Next


rs.Close
End Sub

Private Sub cmd_r_Click()

Call Affiche_Vehicule
Call Affiche_Conducteur
Call Affiche_Destination
Call AfficheDepot
Call AfficheExterieur

Txt_KM.Enabled = True
Txt_KM.Text = ""
txt_observation.Text = ""

Call grid_vehicule_ColumnClick(1)
Call grid_Conducteur_ColumnClick(1)


End Sub

Private Sub CmdPrint_Click()
Dim i As Integer

Unload FrmMajFT
 With FrmControlePwd
        .Sible = "Consult_Compteurs"
        .Show
End With
Exit Sub

'FrmCompteurs.Show
End Sub

Private Sub CmdSave_Click()
Dim EnMission As Boolean
Dim j As Integer
Dim ii As Integer

Dim LInt_NumCompteur As Long
Dim NumeroTxt As String
Dim SelectedV As String
Dim SelectedC As String
Dim SelectedD As String
Dim Kilometrage As Long
Dim Observation As String
Dim heure As Date
Dim NumeroFiche As String
Dim intLoopIndex As Integer
Dim compteur As Long

Dim SQL As String
Dim rs As New ADODB.Recordset

Dim Vehicule As String
Dim Conducteur As String

'Controle des droits

        Dim rQ As New ADODB.Recordset
        SQL = "Select * from utilisateur where Ins_FT = 1 and code= " & LInt_UserId
        rQ.Open SQL, CNB, adOpenDynamic
        If rQ.EOF Then
            rQ.Close
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
'Controle dispinibilité
Vehicule = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
Conducteur = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)

SQL = "Select * from vehicule  where Matricule=" & SQLText(Vehicule)
rs.Open SQL, CNB, adOpenKeyset

If Not (rs.EOF) Then
    If (rs("Disponible") = "N") Then
          MsgBox "Voiture Hors Service    ", vbInformation
        grid_vehicule.SetFocus
        rs.Close
        Exit Sub
    End If
End If
rs.Close

SQL = "Select * from Personnel  where Libelle=" & SQLText(Conducteur)
rs.Open SQL, CNB, adOpenKeyset

If Not (rs.EOF) Then
    If (rs("Disponible") = "N") Then
          MsgBox "Conducteur Hors Service    ", vbInformation
        grid_vehicule.SetFocus
        rs.Close
        Exit Sub
    End If
End If
rs.Close



'Vérifier Véhicule
If grid_vehicule.SelectionCount = 0 Then
   MsgBox "Sélectionner Vehicule      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If
'Verifier Conducteur
If grid_Conducteur.Enabled = True Then
    EnMission = False
For j = 1 To LSV_Exterieur.ListItems.Count
                If ((grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)) _
                = LSV_Exterieur.ListItems(j).SubItems(3) And LSV_Exterieur.ListItems(j).SubItems(6) = "" And LSV_Exterieur.ListItems(j).SubItems(4) <> "REPARATION") Then
                EnMission = True
            End If
        Next
If EnMission = True Then
     MsgBox "Conducteur déja en Mission     ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
   End If
   
'Sélection Conducteur
If grid_Conducteur.SelectionCount = 0 Then
   MsgBox "Sélectionner Conducteur      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If
End If

'Sélection déstination
If grid_destination.Enabled = True Then
If grid_destination.SelectionCount = 0 Then
   MsgBox "Sélectionner la destination     ", vbInformation
   grid_destination.SetFocus
   Exit Sub
End If
End If



'Sélection de véhicule
SelectedV = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
'sélection de l'heure
heure = Now
        
'get Operation & Numero de fiche
NumeroFiche = "Auto"
ReDim Operation(1)
Operation = ReturnOperation(SelectedV)
If (Operation(0) = "E") Then
    NumeroFiche = Operation(1)
End If
     Select Case MsgBox("Confirmez vous l'enregistrement", vbYesNoCancel + vbDefaultButton2 + vbInformation)
            Case vbYes
        
    If Operation(0) = "S" Then
        SelectedC = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)
        SelectedD = grid_destination.CellText(grid_destination.SelectedRow, grid_destination.SelectedCol)
        Kilometrage = CompteurVehicule(SelectedV)
        Observation = txt_observation.Text
               LInt_NumCompteur = return_Compteur() + 1
        
          NumeroTxt = Format(LInt_NumCompteur, "00000")
            CNB.BeginTrans
    
             'Insertion enregistrement
        SQL = "Insert into ficheTraffic  (Numero,Vehicule,CompteurSortie,Conducteur,Destination,HeureSortie, Observation, OperateurSortie) values ("
        SQL = SQL & SQLText(NumeroTxt)
        SQL = SQL & "," & SQLText(SelectedV)
        SQL = SQL & "," & SQLText(Kilometrage)
        SQL = SQL & "," & SQLText(SelectedC)
        SQL = SQL & "," & SQLText(SelectedD)
        SQL = SQL & "," & SQLText(heure)
        SQL = SQL & "," & SQLText(Observation)
        SQL = SQL & "," & SQLText(LStr_NameUser)
        SQL = SQL & ")"
        CNB.Execute SQL
        CNB.CommitTrans
        MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
        
        Call AfficheExterieur
        Call AfficheDepot
        
    ElseIf Operation(0) = "E" Then
        NumeroTxt = Format(CStr(NumeroFiche), "00000")
        
        'Verification de Compteur
        If Txt_KM.Text = "" Then
             MsgBox "Compteur Vide !!!     ", vbInformation
           Txt_KM.SetFocus
           Exit Sub
        End If
        
        
        If grid_vehicule.SelectionCount <> 0 Then
        If Txt_KM <> "" Then
            SelectedV = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
            compteur = CompteurVehicule(SelectedV)
            lbl_Compteur.Caption = compteur
        'Controle de compteur
        If Val(Txt_KM.Text) < Val(compteur) Then
            MsgBox ("Nouveau compteur invalid")
            Txt_KM.SetFocus
            Exit Sub
            
        End If
        
        If Val(Txt_KM.Text) - Val(compteur) > 1200 Then
            MsgBox ("Nouveau compteur invalid : Plus que 1200 klm")
            Txt_KM.SetFocus
            Exit Sub
            
        End If
        End If
        End If

        
        CNB.BeginTrans
        SQL = "Update fichetraffic Set "
        SQL = SQL & " HeureEntre = " & SQLText(heure)
        SQL = SQL & " , CompteurEntre = " & SQLText(Txt_KM.Text)
        SQL = SQL & " , OperateurEntre = " & SQLText(LStr_NameUser)
        SQL = SQL & " , ObservationEntre = " & SQLText(Observation)
        SQL = SQL & " where fichetraffic.Numero = " & NumeroTxt
        CNB.Execute SQL
        
        SQL = "Update Vehicule Set CompteurFT = " & SQLText(Txt_KM.Text) & " where Matricule=" & SQLText(SelectedV)
        CNB.Execute SQL
        
        CNB.CommitTrans
        MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
    
    Call AfficheExterieur
    Call AfficheDepot

    End If
    
    Call Affiche_Vehicule
    Call Affiche_Conducteur
    Call Affiche_Destination
    Txt_KM.Text = ""
    txt_observation.Text = ""
    grid_vehicule.SetFocus
    Call grid_vehicule_ColumnClick(1)
    Call grid_Conducteur_ColumnClick(1)
    
Case vbCancel
    Call Affiche_Vehicule
    Call Affiche_Conducteur
    Call Affiche_Destination
    Txt_KM.Text = ""
    txt_observation.Text = ""
    grid_vehicule.SetFocus
    Call grid_vehicule_ColumnClick(1)
    Call grid_Conducteur_ColumnClick(1)
End Select
    
   
Exit Sub


End Sub

Private Sub CmdSave_GotFocus()

If grid_vehicule.SelectionCount = 0 Then
   MsgBox "Sélectionner Vehicule      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If

If grid_Conducteur.Enabled = True Then
If grid_Conducteur.SelectionCount = 0 Then
   MsgBox "Sélectionner un Conducteur     ", vbInformation
   grid_Conducteur.SetFocus
   Exit Sub
End If
End If

If grid_destination.Enabled = True Then
If grid_destination.SelectionCount = 0 Then
   MsgBox "Sélectionner la destination     ", vbInformation
   grid_destination.SetFocus
   Exit Sub
End If
End If

If Len(Trim(txt_observation.Text)) = 0 Then
    txt_observation.Text = "SANS OBSERVATION"
    Exit Sub
End If

End Sub

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
Call Initgrid_Destination
Call Affiche_Destination

Call grid_vehicule_ColumnClick(1)
Call grid_Conducteur_ColumnClick(1)

Img_alarme.Picture = LoadPicture(App.Path & "\Images\button-green.bmp")
dat = Date
Lbl_date.Caption = UCase(Format(Now, "dddd-dd-mm-yyyy"))
dat = Time
Lbl_heure.Caption = dat

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
         .SortType(1) = CCLSortBackColor
      End Select
   End With
   Screen.MousePointer = vbHourglass
   grid_Conducteur.Sort
   Screen.MousePointer = vbDefault

Me.WindowState = 2

End Sub



Private Sub grid_Conducteur_DblClick(ByVal lRow As Long, ByVal lCol As Long)

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim Conducteur As String
    
    On Error GoTo Err

'Controle d'acces
SQL = "Select * from utilisateur where MAJ_Disp = 1 and code= " & LInt_UserId
rs.Open SQL, CNB, adOpenDynamic
If rs.EOF Then
    rs.Close
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
Else
    rs.Close
End If
    
Conducteur = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)

SQL = "Select * from Personnel  where Libelle=" & SQLText(Conducteur)
rs.Open SQL, CNB, adOpenKeyset


If (Not rs.EOF) Then
    If rs("disponible") = "O" Then
        If MsgBox("EN-Service -> " & Conducteur & " comme HORS-Service?", vbYesNo + vbDefaultButton2 + vbCritical) = vbYes Then
            
            Call SaveDispo(Conducteur, "HS")
            SQL = "Update Personnel Set disponible='N'  where Libelle=" & SQLText(Conducteur)
            CNB.Execute SQL
        End If
    Else
        If MsgBox("HORS-Service -> " & Conducteur & " -> EN_Service ?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
            Call SaveDispo(Conducteur, "ES")
            SQL = "Update Personnel Set disponible='O'  where Libelle=" & SQLText(Conducteur)
            CNB.Execute SQL
        End If
    End If
End If

Call cmd_r_Click

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub grid_Conducteur_GotFocus()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim Vehicule As String

Vehicule = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
SQL = "Select * from vehicule  where Matricule=" & SQLText(Vehicule)
rs.Open SQL, CNB, adOpenKeyset

If Not (rs.EOF) Then
    If (rs("Disponible") = "N") Then
          MsgBox "Voiture Hors Service    ", vbInformation
        grid_vehicule.SetFocus
        rs.Close
        Exit Sub
    End If
End If
rs.Close

End Sub

Private Sub grid_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"

End Sub



Private Sub grid_destination_GotFocus()
Dim EnMission As Boolean
Dim j As Integer

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim Vehicule As String
Dim Conducteur As String


Vehicule = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
Conducteur = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)

SQL = "Select * from vehicule  where Matricule=" & SQLText(Vehicule)
rs.Open SQL, CNB, adOpenKeyset

If Not (rs.EOF) Then
    If (rs("Disponible") = "N") Then
          MsgBox "Voiture Hors Service    ", vbInformation
        grid_vehicule.SetFocus
        rs.Close
        Exit Sub
    End If
End If
rs.Close

SQL = "Select * from Personnel  where Libelle=" & SQLText(Conducteur)
rs.Open SQL, CNB, adOpenKeyset

If Not (rs.EOF) Then
    If (rs("Disponible") = "N") Then
          MsgBox "Conducteur Hors Service    ", vbInformation
        grid_vehicule.SetFocus
        rs.Close
        Exit Sub
    End If
End If
rs.Close


Conducteur = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)
Vehicule = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)

If grid_Conducteur.Enabled = True Then
    EnMission = False
    For j = 1 To LSV_Exterieur.ListItems.Count
        If (Conducteur = LSV_Exterieur.ListItems(j).SubItems(3)) And (LSV_Exterieur.ListItems(j).SubItems(6) = "" And LSV_Exterieur.ListItems(j).SubItems(4) <> "REPARATION") Then
            EnMission = True
        End If
    Next
    If EnMission = True Then
        MsgBox "Conducteur déja en Mission     ", vbInformation
        grid_vehicule.SetFocus
        Exit Sub
   End If
End If

End Sub

Private Sub grid_destination_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
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


Private Sub grid_vehicule_DblClick(ByVal lRow As Long, ByVal lCol As Long)

 With FrmControlePwd
        .Sible = "Disponibilité"
        .Show
End With
Exit Sub
End Sub


Private Sub grid_vehicule_GotFocus()
If grid_Conducteur.Enabled = False Then grid_Conducteur.Enabled = True
If grid_destination.Enabled = False Then grid_destination.Enabled = True
lbl_Compteur.Caption = "Compteur Actuelle !!"

End Sub

Private Sub grid_vehicule_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub grid_vehicule_LostFocus()
Dim i As Integer
Dim SelectedV As String
Dim compteur As Long

'SelectedV=Matricule de voiture sélectionner
If grid_vehicule.SelectionCount <> 0 Then
    SelectedV = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
    compteur = CompteurVehicule(SelectedV)
End If

'get Operation
ReDim Operation(1)
Operation = ReturnOperation(SelectedV)
    If (Operation(0) = "E") Then
        grid_Conducteur.Enabled = False
        grid_destination.Enabled = False
        Txt_KM.Enabled = True
        Txt_KM.TabIndex = 4
        Txt_KM.SetFocus
    Else
     Txt_KM.Text = compteur
     Txt_KM.Enabled = False
     Txt_KM.TabIndex = 60
    End If
 If grid_Conducteur.Enabled = True Then
 Img_alarme.Picture = LoadPicture(App.Path & "\Images\button-green.bmp")
 Else
 Img_alarme.Picture = LoadPicture(App.Path & "\Images\button-red.bmp")
 End If
 
 
End Sub

Private Sub LSV_Exterieur_DblClick()
Dim i As Integer

Unload FrmMajFT
 With FrmControlePwd
        i = LSV_Exterieur.SelectedItem.Index
        .vCode = LSV_Exterieur.ListItems(i).SubItems(1)
        .Sible = "MajTraffic"
        .Show
End With
Exit Sub

End Sub



Private Sub SCommand1_Click()
Frm_Main.Alerte
End Sub



Private Sub Timer1_Timer()
Lbl_heure = Time
End Sub



Private Sub Txt_KM_GotFocus()
Dim EnMission As Boolean
Dim j As Integer
Dim compteur As Long
Dim SelectedV As String

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim Vehicule As String
Dim Conducteur As String

Vehicule = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
Conducteur = grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)
SQL = "Select * from vehicule  where Matricule=" & SQLText(Vehicule)
rs.Open SQL, CNB, adOpenKeyset

If Not (rs.EOF) Then
    If (rs("Disponible") = "N") Then
          MsgBox "Voiture Hors Service    ", vbInformation
        grid_vehicule.SetFocus
        rs.Close
        Exit Sub
    End If
End If
rs.Close

SQL = "Select * from Personnel  where Libelle=" & SQLText(Conducteur)
rs.Open SQL, CNB, adOpenKeyset

If Not (rs.EOF) Then
    If (rs("Disponible") = "N") Then
          MsgBox "Conducteur Hors Service    ", vbInformation
        grid_vehicule.SetFocus
        rs.Close
        Exit Sub
    End If
End If
rs.Close


If grid_vehicule.SelectionCount = 0 Then
   MsgBox "Sélectionner Vehicule      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If

If grid_vehicule.SelectionCount <> 0 Then
    SelectedV = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
    compteur = CompteurVehicule(SelectedV)
    If Len(CStr(compteur)) = 6 Then
    Txt_KM.Text = Left(CStr(compteur), 2)
    ElseIf Len(CStr(compteur)) = 5 Then
    Txt_KM.Text = Left(CStr(compteur), 1)
    Else
    Txt_KM.Text = 0
    End If
End If

If grid_Conducteur.Enabled = True Then
    EnMission = False
For j = 1 To LSV_Exterieur.ListItems.Count
                If ((grid_Conducteur.CellText(grid_Conducteur.SelectedRow, grid_Conducteur.SelectedCol)) _
                = LSV_Exterieur.ListItems(j).SubItems(3) And LSV_Exterieur.ListItems(j).SubItems(6) = "" And LSV_Exterieur.ListItems(j).SubItems(4) <> "REPARATION") Then
                EnMission = True
            End If
        Next
If EnMission = True Then
     MsgBox "Conducteur déja en Mission     ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
   End If

If grid_Conducteur.SelectionCount = 0 Then
   MsgBox "Sélectionner Conducteur      ", vbInformation
   grid_vehicule.SetFocus
   Exit Sub
End If
End If

If grid_destination.Enabled = True Then
If grid_destination.SelectionCount = 0 Then
   MsgBox "Sélectionner la destination     ", vbInformation
   grid_destination.SetFocus
   Exit Sub
End If
End If

Txt_KM.SelStart = Len(Txt_KM.Text)
        


End Sub

Private Sub Txt_KM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Txt_KM_KeyPress(KeyAscii As Integer)
If Not (Chr(KeyAscii) Like "[0123456789]") And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub



Private Sub Txt_KM_LostFocus()
Dim SelectedV As String
Dim compteur As Long

 'Verification de Compteur
        
        
        If grid_Conducteur.Enabled = True Then
        If grid_vehicule.SelectionCount <> 0 Then
        
'        If Txt_KM.Text = "" Then
'             MsgBox "Compteur Vide !!!     ", vbInformation
'           Txt_KM.SetFocus
'           Exit Sub
'        End If
        If Txt_KM <> "" Then
            SelectedV = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)
            compteur = CompteurVehicule(SelectedV)
            lbl_Compteur.Caption = compteur
        'Controle de compteur
        If Val(Txt_KM.Text) < Val(compteur) Then
            MsgBox ("Nouveau compteur invalid")
            Txt_KM.SetFocus
            Exit Sub
            
        End If
        
        If Val(Txt_KM.Text) - Val(compteur) > 1200 Then
            MsgBox ("Nouveau compteur invalid : Plus que 1200 klm")
            Txt_KM.SetFocus
            Exit Sub
            
        End If
        
        End If
        End If
        End If
End Sub

Private Sub Txt_Observation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
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
    
    .AddColumn "Matricule", "", , , 140
    
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
grid_vehicule.SelectedRow = 1
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
    
    .AddColumn "Libelle", "", , , 140
    
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub

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
                If (rs("Libelle") = LSV_Exterieur.ListItems(j).SubItems(3) And LSV_Exterieur.ListItems(j).SubItems(6) = "" And LSV_Exterieur.ListItems(j).SubItems(4) <> "REPARATION") Then
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
grid_Conducteur.SelectedRow = 1



End Sub

Public Sub Initgrid_Destination()
With grid_destination
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
    
    .AddColumn "LibelleD", "", , , 140
    
    .StretchLastColumnToFit = True
    .Redraw = True
End With

End Sub

Public Sub Affiche_Destination()
Dim SQL As String
Dim rs As New ADODB.Recordset

grid_destination.ClearRows
SQL = "Select * from Destination where actif = 1  order by Libelle"
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    grid_destination.Redraw = False
    While Not rs.EOF
        With grid_destination
            .AddRow
            
            .CellDetails .Rows, .ColumnIndex("LibelleD"), rs("Libelle")
        End With
        rs.MoveNext
    Wend
    grid_destination.Redraw = True
End If
grid_destination.SelectedRow = 1
End Sub

Private Function return_Compteur() As Long
Dim rD As New ADODB.Recordset
Dim SQL As String
return_Compteur = 0
SQL = "select Max(Numero) from fichetraffic "
rD.Open SQL, CNB, adOpenKeyset
If Not rD.EOF Then
return_Compteur = rD(0)
End If

rD.Close
End Function

Public Function CompteurVehicule(ByVal vCode As String) As Long
    Dim rD As New ADODB.Recordset
    Dim SQL As String
    SQL = "Select * from vehicule where Matricule = " & SQLText(vCode)
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


Public Sub SaveDispo(ByVal Conducteur As String, ByVal Operation As String)

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim numero As String
Dim LInt_NumCompteur As Long

On Error GoTo Err


    
        'Si Operation = En-Service
        
    If Operation = "ES" Then
        CNB.BeginTrans
        
        'Update Ancien Ligne
        SQL = "Update DispoPerso Set HFin=" & SQLText(Now) & " "
        SQL = SQL & " where Numero = (Select Max(Numero) from DispoPerso where Conducteur=" & SQLText(Conducteur) & " And Etat = 'Hors-Service')"
        CNB.Execute SQL
        
        'Créer Nouvelle Ligne
        LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_DispoPerso")
        If ErrNumber <> 0 Then
           MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
           ErrNumber = 0
           Exit Sub
        End If
        'Insertion enregistrement assiette
        numero = Format(LInt_NumCompteur, "00000")
        'Insertion enregistrement
        SQL = "Insert into DispoPerso (Numero,Conducteur,Etat,HDebut) values ("
        SQL = SQL & SQLText(numero)
        SQL = SQL & "," & SQLText(Conducteur)
        SQL = SQL & "," & SQLText("En-Service")
        SQL = SQL & "," & SQLText(Now)
        SQL = SQL & ")"
        CNB.Execute SQL
        CNB.CommitTrans
        
        'Si Operation = Hors-Service
    ElseIf Operation = "HS" Then
        CNB.BeginTrans
        
        'Update Ancien Ligne
        SQL = "Update DispoPerso Set HFin=" & SQLText(Now)
        SQL = SQL & " where Numero = (Select Max(Numero) from DispoPerso where Conducteur=" & SQLText(Conducteur) & " And Etat= 'En-Service') "
        CNB.Execute SQL
        
        'Créer Nouvelle Ligne
        LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_DispoPerso")
        If ErrNumber <> 0 Then
           MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
           ErrNumber = 0
           Exit Sub
        End If
        'Insertion enregistrement assiette
        numero = Format(LInt_NumCompteur, "00000")
        'Insertion enregistrement
        SQL = "Insert into DispoPerso (Numero,Conducteur,Etat,HDebut) values ("
        SQL = SQL & SQLText(numero)
        SQL = SQL & "," & SQLText(Conducteur)
        SQL = SQL & "," & SQLText("Hors-Service")
        SQL = SQL & "," & SQLText(Now)
        SQL = SQL & ")"
        CNB.Execute SQL
        CNB.CommitTrans
    End If
Exit Sub
Err:
CNB.RollbackTrans
MsgBox Err.Description, vbInformation


End Sub

Public Sub MajDisp()
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim Vehicule As String
    
    On Error GoTo Err
Vehicule = grid_vehicule.CellText(grid_vehicule.SelectedRow, grid_vehicule.SelectedCol)

SQL = "Select * from vehicule  where Matricule=" & SQLText(Vehicule)
rs.Open SQL, CNB, adOpenKeyset


If (Not rs.EOF) Then
    If rs("disponible") = "O" Then
        If MsgBox("EN-Service -> " & Vehicule & " -> HORS-Service ?", vbYesNo + vbDefaultButton2 + vbCritical) = vbYes Then
            
            SQL = "Update Vehicule Set disponible='N'  where Matricule=" & SQLText(Vehicule)
            CNB.Execute SQL
        End If
    Else
        If MsgBox("HORS-Service -> " & Vehicule & " -> EN-SERVICE ?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
            
            SQL = "Update Vehicule Set disponible='O'  where Matricule=" & SQLText(Vehicule)
            CNB.Execute SQL
        End If
    End If
End If

Call cmd_r_Click

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub
