VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmStatCarburant 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Statistiques Carburant"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   11175
   Begin VB.ComboBox cbo_Matricule 
      Height          =   315
      ItemData        =   "FrmStatCarburant.frx":0000
      Left            =   2160
      List            =   "FrmStatCarburant.frx":0002
      TabIndex        =   7
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   300
      Left            =   4080
      TabIndex        =   0
      Top             =   2520
      Width           =   375
   End
   Begin SToolBox.SDateBox cda_Debut 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Tag             =   "M"
      Top             =   2520
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
   End
   Begin SToolBox.SDateBox cda_Fin 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Tag             =   "M"
      Top             =   2520
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
   End
   Begin MSComctlLib.ListView Lsv_Details 
      Height          =   1215
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   2143
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Vehicule"
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "valeur Carburant"
         Object.Width           =   3177
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Prix Litre"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "NB.Litre Carburant"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Km.Parcouru"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Consommation Par 100 km"
         Object.Width           =   2118
      EndProperty
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   345
      Left            =   5400
      TabIndex        =   8
      Top             =   1680
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
      Picture         =   "FrmStatCarburant.frx":0004
      ButtonType      =   1
   End
   Begin MSComctlLib.ListView Lsv_detailP 
      Height          =   4215
      Left            =   0
      TabIndex        =   10
      Top             =   4320
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7435
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pièce"
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vehicule"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Nb.Litres"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Montant"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "KM-Parcourue"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Date "
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Compteur"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cons/100KM"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Véhicule:_ _ _ _ _ _ :"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   2040
   End
   Begin VB.Label m 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statistiques Carburant"
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
      Caption         =   "Période du :"
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
      Left            =   0
      TabIndex        =   4
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Au :"
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
      Left            =   2295
      TabIndex        =   3
      Top             =   2520
      Width           =   405
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmStatCarburant.frx":0357
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20895
   End
End
Attribute VB_Name = "FrmStatCarburant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim thekey As Integer
Dim theshift As Integer
Dim frm As New FrmAllBonCarburant

Private Sub cbo_Matricule_Change()
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
        start = Len(cbo_Matricule.Text)
        For i = 0 To cbo_Matricule.ListCount - 1
            If Left(cbo_Matricule.List(i), start) = cbo_Matricule.Text Then
                cbo_Matricule.Text = cbo_Matricule.List(i)
            End If
        Next
        cbo_Matricule.SelStart = start
        cbo_Matricule.SelLength = Len(cbo_Matricule.Text)
    End If
End Sub

Private Sub cbo_Matricule_GotFocus()

'Call Affiche_Matricule_Combo(cbo_Matricule)
'cbo_Matricule.AddItem ("Tous"), 0
End Sub

Private Sub cbo_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cbo_Matricule_KeyUp(KeyCode As Integer, Shift As Integer)
 thekey = KeyCode
    theshift = Shift
End Sub

Private Sub cda_Fin_Change()
'On Error GoTo Err
If KeyCode = vbKeyReturn Then
'    Call AfficheDetails_PourCreation(txt_MatriculeStation.Text, cda_Debut.Text, cda_Fin.Text)
End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation


End Sub

Private Sub cmdFindMatricule_Click()
On Error GoTo Err

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Véhicule Stat carburant"
    .Show
End With
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub Command1_Click()

On Error GoTo Err

Lsv_Details.ListItems.Clear
Lsv_detailP.ListItems.Clear

If cbo_Matricule.Text = "Tous" Then
Call AfficheDetails_Tous(cda_Debut.Text, cda_Fin.Text)
Else
Call AfficheDetails_ParVehicule(cbo_Matricule.Text, cda_Debut.Text, cda_Fin.Text)
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Public Sub AfficheRow_Vehicule(ByVal vcode As String)

Dim SQL As String
Dim rs As New ADODB.Recordset

SQL = "Select * from vehicule where code = " & SQLText(vcode) & " OR Matricule= " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Matricule")) Then
    cbo_Matricule.Text = rs("Matricule")
    End If
Else
    MsgBox "Code introuvable", vbInformation
    cbo_Matricule.SetFocus
    Exit Sub
End If
rs.Close

End Sub

Public Function Return_CodVehicule(ByVal Matricule As String) As String
    Dim SQL As String
    Dim rs As New ADODB.Recordset

SQL = "Select Code from vehicule where Matricule= " & SQLText(Matricule)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    If Not IsNull(rs("Code")) Then
    Return_CodVehicule = CStr(rs("Code"))
    End If
End If

rs.Close
End Function



Public Sub AfficheDetails_Tous(ByVal VdateD As Date, ByVal vDateF As Date)
Dim SQL As String
Dim rs As New ADODB.Recordset


Dim TLitre As Double
Dim Valeur As Double

Dim MaxC As Long
Dim MinC As Long
Dim NbKM As Long

Dim KmCarburant As Double

Dim Name_Table As String
Name_Table = "DetBonCarburant"


'detail P
SQL = "SELECT Numero,DetBonCarburant.CompteurCarburant As Compteur, Vehicule.Matricule,"
SQL = SQL & " DetBonCarburant.Litre As Litre, DetBonCarburant.prixLitre As prixLitre,"
SQL = SQL & " DetBonCarburant.dateDoc As DateDoc , (DetBonCarburant.CompteurCarburant - DetBonCarburant.AnCompteur) As Kilometrage"
SQL = SQL & " From DetBonCarburant"
SQL = SQL & " inner join Vehicule"
SQL = SQL & " on DetBonCarburant.vehicule = vehicule.code"
SQL = SQL & " where DetBonCarburant.datedoc Between" & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00"))

rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    While Not rs.EOF
            Set itmX = Lsv_detailP.ListItems.Add(, , rs("Numero"))
            itmX.SubItems(1) = CStr(rs("Matricule"))
            itmX.SubItems(2) = CStr(rs("Litre"))
            itmX.SubItems(3) = CStr(rs("Litre") * rs("prixLitre"))
            itmX.SubItems(4) = CStr(rs("Kilometrage"))
            itmX.SubItems(5) = CStr(rs("DateDoc"))
            itmX.SubItems(6) = CStr(rs("Compteur"))
            If rs("Kilometrage") <> 0 Then
                itmX.SubItems(7) = CStr(Format((rs("Litre") * 100) / rs("Kilometrage"), "#,##0.000"))
            Else
                itmX.SubItems(7) = 0
            End If
        rs.MoveNext
    Wend
End If

rs.Close

'totox Lsv Details
'Parcourir Lsv_detailP
Dim Valc As Double
Dim NBL As Long

Valc = 0
NBL = 0
  
For i = 1 To Lsv_detailP.ListItems.Count
Valc = Valc + Lsv_detailP.ListItems(i).SubItems(3)
NBL = NBL + Lsv_detailP.ListItems(i).SubItems(2)
Next
Set itm = Lsv_Details.ListItems.Add(, , "Tous")
    itm.SubItems(1) = CStr(Format(Valc, "#,##0.000"))
    itm.SubItems(3) = CStr(Format(NBL, "#,##0"))
    

'Details Lsv_details
SQL = "Select SUM(D.Litre) As Litre, D.prixLitre As Prix,V.Matricule As Vehicule" _
& " , Max(D.CompteurCarburant) AS MaxC, Min(D.CompteurCarburant) As MinC" _
& " From DetBonCarburant D , Vehicule V" _
& " Where D.Vehicule = V.Code" _
& " AND D.datedoc Between" & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00")) _
& " group by  D.prixLitre, V.Matricule"
rs.Open SQL, CNB, adOpenKeyset


If Not rs.EOF Then

    While Not rs.EOF
    'Lire et prix Litre
        TLitre = 0
        Valeur = 0

        TLitre = TLitre + rs("Litre")
        Valeur = Valeur + rs("Litre") * rs("Prix")

        Set itmX = Lsv_Details.ListItems.Add(, , CStr(rs("Vehicule")))
                itmX.SubItems(1) = CStr(Format(Valeur, "#,##0.000"))
                itmX.SubItems(2) = CStr(Format(rs("Prix"), "#,##0.000"))
                itmX.SubItems(3) = CStr(Format(TLitre, "#,##0"))

        'Km Parcourus
        MaxC = 0
        MinC = 0
        NbKM = 0
        If Not IsNull(rs("MaxC")) Then
        If Not IsNull(rs("MinC")) Then
            MaxC = MaxC + rs("MaxC")
            MinC = MinC + rs("MinC")

            NbKM = MaxC - MinC
            itmX.SubItems(4) = CStr(NbKM)

        End If
        Else
            itmX(1).SubItems(4) = "Valeur Null"
        Exit Sub
        End If
        'Consommation par 100 KM
        If Not (itmX.SubItems(4) = 0) Then
            KmCarburant = ((itmX.SubItems(3) * 100) / itmX.SubItems(4))
            itmX.SubItems(5) = CStr(Format(KmCarburant, "#,##0.000"))
        Else
            itmX.SubItems(5) = "zéro km!!"
        End If

            rs.MoveNext
         Wend
End If


rs.Close


    

End Sub
Public Sub AfficheDetails_ParVehicule(ByVal Matricule As String, ByVal VdateD As Date, ByVal vDateF As Date)
Dim SQL As String
Dim rs As New ADODB.Recordset

Dim i As Integer

Dim TLitre As Double
Dim Valeur As Double

Dim MaxC As Long
Dim MinC As Long
Dim NbKM As Long

Dim KmCarburant As Double

''selection de code de Vehicule
Dim CodV As String
CodV = Return_CodVehicule(Matricule)


'Remplissage de Grid Lsv_detailP
SQL = "SELECT  Distinct Numero,  Litre, prixLitre, dateDoc, CompteurCarburant,"
SQL = SQL & " (CompteurCarburant - AnCompteur) As Kilometrage"
SQL = SQL & " From DetBonCarburant"
SQL = SQL & " Where DetBonCarburant.datedoc Between" & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00")) & " And "
SQL = SQL & " DetBonCarburant.Vehicule=" & SQLText(CodV) & " Order by CompteurCarburant"
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    While Not rs.EOF
            Set itmX = Lsv_detailP.ListItems.Add(, , rs("Numero"))
            itmX.SubItems(1) = CStr(Matricule)
            itmX.SubItems(2) = CStr(rs("Litre"))
            itmX.SubItems(3) = CStr(rs("Litre") * rs("prixLitre"))
            itmX.SubItems(4) = CStr(rs("Kilometrage"))
            itmX.SubItems(5) = CStr(rs("DateDoc"))
            itmX.SubItems(6) = CStr(rs("CompteurCarburant"))
            If rs("Kilometrage") <> 0 Then
                itmX.SubItems(7) = CStr(Format((rs("Litre") * 100) / rs("Kilometrage"), "#,##0.000"))
            Else
                itmX.SubItems(7) = 0
            End If
        rs.MoveNext
    Wend
End If

rs.Close



'Parcourir Lsv_detailP
Dim Valc As Double
Dim NBL As Long

Valc = 0
NBL = 0
  
For i = 1 To Lsv_detailP.ListItems.Count
Valc = Valc + Lsv_detailP.ListItems(i).SubItems(3)
NBL = NBL + Lsv_detailP.ListItems(i).SubItems(2)
Next
Set itm = Lsv_Details.ListItems.Add(, , CStr(Matricule))
    itm.SubItems(1) = CStr(Format(Valc, "#,##0.000"))
    itm.SubItems(3) = CStr(Format(NBL, "#,##0"))
    
    'Nombre de kilomètres
   If Not (Lsv_detailP.ListItems.Count = 0) Then
    itm.SubItems(4) = Lsv_detailP.ListItems(i - 1).SubItems(6) - Lsv_detailP.ListItems(1).SubItems(6)
    Else
    itm.SubItems(4) = 0
    End If
    
    'consommation par 100 km
      If Not (itm.SubItems(4) = 0) Then
            KmCarburant = ((itm.SubItems(3) * 100) / itm.SubItems(4))
            itm.SubItems(5) = CStr(Format(KmCarburant, "#,##0.000"))
        Else
            itm.SubItems(5) = "zéro km!!"
        End If
End Sub



Private Sub Form_Load()
cbo_Matricule.AddItem "Tous", 0
Call Affiche_Matricule_Combo(cbo_Matricule)
cbo_Matricule.ListIndex = 0

cda_Debut.Text = "01/" & Month(Date) & "/" & Year(Date)
cda_Fin.Text = Date

Me.WindowState = 2
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

Private Sub Lsv_detailP_DblClick()
Dim vcode
Dim i As Integer
On Error GoTo Err
i = Lsv_detailP.SelectedItem.Index
vcode = Lsv_detailP.ListItems.Item(i)
With FrmAllBonCarburant
    .AfficheRow (vcode)
    .Show
    End With
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub


