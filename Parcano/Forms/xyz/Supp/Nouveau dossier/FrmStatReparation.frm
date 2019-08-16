VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStatReparation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Stat Réparations"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   13995
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   300
      Left            =   4560
      TabIndex        =   1
      Top             =   2520
      Width           =   375
   End
   Begin VB.ComboBox cbo_Matricule 
      Height          =   315
      ItemData        =   "FrmStatReparation.frx":0000
      Left            =   1440
      List            =   "FrmStatReparation.frx":0007
      TabIndex        =   0
      Top             =   1800
      Width           =   3015
   End
   Begin SToolBox.SDateBox cda_Debut 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Tag             =   "M"
      Top             =   2520
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
   End
   Begin SToolBox.SDateBox cda_Fin 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
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
      TabIndex        =   4
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Vehicule"
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre réparations"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "valeur"
         Text            =   "valeur Reparation "
         Object.Width           =   3177
      EndProperty
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   345
      Left            =   4560
      TabIndex        =   5
      Top             =   1800
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
      Picture         =   "FrmStatReparation.frx":0011
      ButtonType      =   1
   End
   Begin MSComctlLib.ListView Lsv_detailP 
      Height          =   4215
      Left            =   0
      TabIndex        =   6
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pièce"
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vehicule"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Désignation"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Qte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "TOT.TTC.NET"
         Object.Width           =   2540
      EndProperty
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
      Left            =   2760
      TabIndex        =   10
      Top             =   2520
      Width           =   405
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
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Label m 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statistiques Reparation"
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
      TabIndex        =   8
      Top             =   240
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Véhicule:"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   885
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmStatReparation.frx":0364
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20895
   End
End
Attribute VB_Name = "FrmStatReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim thekey As Integer
Dim theshift As Integer
Public Sub AfficheDetails_Tous(ByVal VdateD As Date, ByVal vDateF As Date)
'variables globales
Dim SQL As String
Dim rs As New ADODB.Recordset

'variables DetailsP
Dim TotHTBrut As Double
Dim TotTTC As Double
Dim Fcode As String
Dim Qte As Double
Dim PUHT As Double
Dim Remise As Double
Dim tva As Double
Dim RP As Double
Dim TotalG As Double

'variables Details
Dim nbRep As Double
Dim Valeur As Double

Set itmX = Lsv_Details.ListItems.Add(, , "Tous")

'nombre des reparations
SQL = " Select count(*) As nbrRep from AsspieceReparation, detailpieceReparation where" _
& " AssPieceReparation.Numero = detailPieceReparation.Numero" _
& " And AssPieceReparation.datePiece" _
& " between " & SQLText(VdateD) & " and " & SQLText(vDateF)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    nbRep = 0
     nbRep = nbRep + rs("nbrRep")
   itmX.SubItems(1) = CStr(rs("nbrRep"))
       End If
    
rs.Close

'TTC
SQL = "select  Sum (totTTC)As valeur from AssPieceReparation where " _
& "Datepiece between" & SQLText(VdateD) & " and " & SQLText(vDateF)
rs.Open SQL, CNB, adOpenKeyset

If Not rs.EOF Then
    If Not IsNull(rs("Valeur")) Then
        Valeur = 0
        Valeur = Valeur + rs("valeur")
                itmX.SubItems(2) = CStr(Format(rs("valeur"), "#,##0.000"))
    Else
    itmX.SubItems(2) = "Valeur Null"
    End If
End If
    
rs.Close

'Vehicule + nbr reparations par vehicule

SQL = " Select vehicule , count(*) As nbrRep  from AsspieceReparation, detailpieceReparation where" _
& " AssPieceReparation.Numero = detailPieceReparation.Numero" _
& " And AssPieceReparation.datePiece" _
& " between " & SQLText(VdateD) & " and " & SQLText(vDateF) & " group by Vehicule"
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
While Not rs.EOF
    nbRep = 0
     nbRep = nbRep + rs("nbrRep")
    Set itmX = Lsv_Details.ListItems.Add(, , CStr(rs("vehicule")))
   itmX.SubItems(1) = CStr(rs("nbrRep"))
   rs.MoveNext
   Wend
       End If
    
rs.Close

'Valeur Reparation par vehicule
For i = 2 To Lsv_Details.ListItems.Count
SQL = "select Qte, PUHT, Remise, TVA, remisePiece from AssPieceReparation,detailPiecereparation where " _
 & " AssPieceReparation.Numero= detailPiecereparation.Numero And " _
& " vehicule =" & SQLText(Lsv_Details.ListItems(i)) & " " _
& " And Datepiece between " & SQLText(VdateD) & " and " & SQLText(vDateF)
rs.Open SQL, CNB, adOpenKeyset

    TotalG = 0
    If Not rs.EOF Then
    While Not rs.EOF
    TotTTC = 0
    TotHTBrut = 0
    Qte = 0
    PUHT = 0
    Remise = 0
    tva = 0
    RP = 0
    Qte = rs("Qte")
    PUHT = rs("PUHT")
    Remise = rs("Remise")
    tva = rs("tva")
    RP = rs("remisePiece")
    TotHTBrut = FrmSaisiePieceReparation.Return_TotHT(Qte, PUHT, Remise)
    TotTTC = TotTTC + (TotHTBrut + (TotHTBrut * (tva / 100)))
    TotTTC = TotTTC - (TotTTC * RP / 100)
    
    TotalG = TotalG + TotTTC
     Lsv_Details.ListItems(i).SubItems(2) = Format(TotalG, "#,##0.000")
    
     rs.MoveNext
    Wend
End If
rs.Close
Next


'detailP

SQL = "Select * from DetailPieceReparation , AssPieceReparation " _
& " Where DetailPieceReparation.Numero = AssPieceReparation.Numero " _
& " And AssPieceReparation.datePiece" _
& " between " & SQLText(VdateD) & " and " & SQLText(vDateF) & " order by DetailPieceReparation.vehicule, AssPieceReparation.datePiece"

rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    While Not rs.EOF
    TotTTC = 0
    TotHTBrut = 0
    Qte = 0
    PUHT = 0
    Remise = 0
    tva = 0
    
    Qte = rs("Qte")
    PUHT = rs("PUHT")
    Remise = rs("Remise")
    tva = rs("tva")
    
    TotHTBrut = FrmSaisiePieceReparation.Return_TotHT(Qte, PUHT, Remise)
    TotTTC = TotHTBrut + (TotHTBrut * (tva / 100))
    
    
           
            
            Set itmX = Lsv_detailP.ListItems.Add(, , rs("Numero"))
            itmX.SubItems(1) = rs("datePiece")
            itmX.SubItems(2) = rs("Vehicule")
            itmX.SubItems(3) = rs("Designation")
            itmX.SubItems(4) = rs("Qte")
            itmX.SubItems(5) = Format(TotTTC, "#,##0.000")
        rs.MoveNext
    Wend
End If


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
Call Affiche_Matricule_Combo(cbo_Matricule)
cbo_Matricule.AddItem ("Tous"), 0
End Sub


Private Sub cbo_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cbo_Matricule_KeyUp(KeyCode As Integer, Shift As Integer)
thekey = KeyCode
    theshift = Shift
End Sub

Private Sub cmdFindMatricule_Click()
On Error GoTo Err

Unload FrmFind_Fils
With FrmFind_Fils
    .StrSource = "Véhicule Stat reparation"
    .Show
End With
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

Public Sub AfficheDetails_ParVehicule(ByVal Matricule As String, ByVal VdateD As Date, ByVal vDateF As Date)
'variables globales
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim i As Integer
'variables DetailsP
Dim TotHTBrut As Double
Dim TotTTC As Double
Dim Fcode As String
Dim Qte As Double
Dim PUHT As Double
Dim Remise As Double
Dim tva As Double
Dim RemiseP As Double

'variables Details
Dim nbRep As Double
Dim Valeur As Double

Set itmX = Lsv_Details.ListItems.Add(, , "Tous")

'nombre des reparations
SQL = " Select count(*) As nbrRep from AsspieceReparation, detailpieceReparation where" _
& " AssPieceReparation.Numero = detailPieceReparation.Numero" _
& " And detailPieceReparation.vehicule= " & SQLText(Matricule) & " And " _
& "  AssPieceReparation.datePiece" _
& " between " & SQLText(VdateD) & " and " & SQLText(vDateF)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    nbRep = 0
    nbRep = nbRep + rs("nbrRep")
   itmX.SubItems(1) = CStr(rs("nbrRep"))
       End If
    
rs.Close

'detail P

SQL = "Select * from DetailPieceReparation , AssPieceReparation "
SQL = SQL & " Where DetailPieceReparation.Numero = AssPieceReparation.Numero "
SQL = SQL & " And detailPieceReparation.vehicule= " & SQLText(Matricule) & " And "
SQL = SQL & " AssPieceReparation.datePiece"
SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    While Not rs.EOF
    TotTTC = 0
    TotHTBrut = 0
    Qte = 0
    PUHT = 0
    Remise = 0
    tva = 0
    RemiseP = 0
    
    Qte = rs("Qte")
    PUHT = rs("PUHT")
    Remise = rs("Remise")
    tva = rs("tva")
    RemiseP = rs("RemisePiece")
    
    TotHTBrut = FrmSaisiePieceReparation.Return_TotHT(Qte, PUHT, Remise)
    TotHTBrut = TotHTBrut - (RemiseP * (tva / 100))
    TotTTC = TotHTBrut + (TotHTBrut * (tva / 100))
    
   
            Set itmX = Lsv_detailP.ListItems.Add(, , rs("Numero"))
            itmX.SubItems(1) = rs("datePiece")
            itmX.SubItems(2) = rs("Vehicule")
            itmX.SubItems(3) = rs("Designation")
            itmX.SubItems(4) = rs("Qte")
            itmX.SubItems(5) = Format(TotTTC, "#,##0.000")
        rs.MoveNext
    Wend
End If

'Totale réparation
Valeur = 0
If Lsv_detailP.ListItems.Count > 0 Then
    For i = 1 To Lsv_detailP.ListItems.Count
        Valeur = Valeur + Lsv_detailP.ListItems(i).SubItems(5)
    Next
End If


Set itmX = Lsv_Details.ListItems.Item(1)
    
         itmX.SubItems(2) = CStr(Valeur)

End Sub

Private Sub Command1_Click()

On Error GoTo Err

Lsv_Details.ListItems.Clear
Lsv_detailP.ListItems.Clear

If cbo_Matricule.Text = "Tous" Then
Call AfficheDetails_Tous(cda_Debut.Text, cda_fin.Text)
Else
Call AfficheDetails_ParVehicule(cbo_Matricule.Text, cda_Debut.Text, cda_fin.Text)
End If

Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub


Private Sub Form_Load()

cbo_Matricule.AddItem ("Tous"), 0
cbo_Matricule.ListIndex = 0

cda_Debut.Text = "01/" & Month(Date) & "/" & Year(Date)
cda_fin.Text = Date


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
    'ViderZone (frm)
    With FrmPieceReparation
        .AfficheRow (vcode)
        .Show
        End With
    Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
