VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmMajStatFT 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maj.Stat.FT"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SToolBox.STimeBox txt_HeureSortie 
      Height          =   285
      Left            =   3000
      TabIndex        =   18
      Top             =   3000
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   503
   End
   Begin SToolBox.SCommand SCommand1 
      Height          =   495
      Left            =   6480
      TabIndex        =   15
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Enregistrer "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8454016
   End
   Begin SToolBox.SDateBox cbo_DateSortie 
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Top             =   3000
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
   End
   Begin VB.TextBox txt_observation 
      Height          =   2085
      Left            =   6480
      TabIndex        =   13
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txt_cptEntre 
      Height          =   405
      Left            =   6480
      TabIndex        =   9
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox txt_cptSortie 
      Height          =   405
      Left            =   6480
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
   Begin VB.ComboBox cbo_destination 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   2400
      Width           =   2775
   End
   Begin VB.ComboBox cbo_Conducteur 
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   1800
      Width           =   2775
   End
   Begin VB.ComboBox cbo_vehicle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
   Begin SToolBox.SCommand SCommand2 
      Height          =   495
      Left            =   7920
      TabIndex        =   16
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Annuler"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421631
   End
   Begin SToolBox.SDateBox cbo_DateEntre 
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      Top             =   3600
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Text            =   ""
   End
   Begin SToolBox.STimeBox txt_heureEntre 
      Height          =   285
      Left            =   3000
      TabIndex        =   19
      Top             =   3600
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   503
   End
   Begin VB.Label lbl_NumFiche 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NumFiche"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mise a jour fiche traffic"
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
      Left            =   720
      TabIndex        =   20
      Top             =   360
      Width           =   3915
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cpt.Sortie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lbl14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cpt.Entre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lbl447 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Observation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "H.Entrée"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "H.Sortie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Conducteur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicule"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "FrmMajStatFT.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "FrmMajStatFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim thekey As Integer
Dim theshift As Integer


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
On Error GoTo Err
    Call Affiche_Personnel_Combo(cbo_Conducteur)
Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Private Sub Cbo_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Cbo_Conducteur_KeyUp(KeyCode As Integer, Shift As Integer)
 thekey = KeyCode
    theshift = Shift
End Sub

Public Sub selectFT(ByVal vcode As String)
Dim NumFiche As String
Dim Matricule As String
Dim Conducteur As String
Dim Destination As String
Dim CompteurEntre As Long
Dim CompteurSortie As Long
Dim HeureSortie As String
Dim HeureEntre As String
Dim Observation As String

Dim SQL As String
Dim rs As New ADODB.Recordset
 SQL = "Select * from Fichetraffic where Numero = " & SQLText(vcode)
 rs.Open SQL, CNB, adOpenDynamic
    If Not rs.EOF Then
        If (Not (IsNull(rs("Numero")))) Then NumFiche = rs("Numero")
      If (Not (IsNull(rs("Vehicule")))) Then Matricule = rs("Vehicule")
      If (Not (IsNull(rs("Conducteur")))) Then Conducteur = rs("Conducteur")
      If (Not (IsNull(rs("Destination")))) Then Destination = rs("Destination")
      If (Not (IsNull(rs("CompteurEntre")))) Then CompteurEntre = rs("CompteurEntre")
      If (Not (IsNull(rs("CompteurSortie")))) Then CompteurSortie = rs("CompteurSortie")
      If (Not (IsNull(rs("HeureEntre")))) Then HeureEntre = rs("HeureEntre")
      If (Not (IsNull(rs("HeureSortie")))) Then HeureSortie = rs("HeureSortie")
      If (Not (IsNull(rs("Observation")))) Then Observation = rs("Observation")
    End If
    
    If Not (IsNull(NumFiche)) Then
    lbl_NumFiche.Caption = NumFiche
    End If
    
    If Not (IsNull(Matricule)) Then
    cbo_vehicle.Text = Matricule
    End If
    
    If Not (IsNull(Conducteur)) Then
    cbo_Conducteur.Text = Conducteur
    End If
    
    If Not (IsNull(Destination)) Then
    cbo_destination.Text = Destination
    End If
    
     If Not (IsNull(HeureSortie)) Then
        cbo_DateSortie.Text = Format(HeureSortie, "dd/mm/yyyy")
    End If
    
     If Not (IsNull(HeureSortie)) Then
        txt_HeureSortie.Text = Format(HeureSortie, "hh:mm:ss")
    End If
    
    If Not (IsNull(HeureEntre)) Then
        cbo_DateEntre.Text = Format(HeureEntre, "dd/mm/yyyy")
    End If
    
    If Not (IsNull(HeureEntre)) Then
        txt_heureEntre.Text = Format(HeureEntre, "hh:mm:ss")
    End If
    
    If Not (IsNull(CompteurSortie)) Then
        txt_cptSortie.Text = CompteurSortie
    End If
    
    If Not (IsNull(CompteurEntre)) Then
        txt_cptEntre.Text = CompteurEntre
    End If
    
    If Not (IsNull(Observation)) Then
        txt_observation.Text = Observation
    End If
    
   
    
   
 


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
        start = Len(cbo_destination.Text)
        For i = 0 To cbo_destination.ListCount - 1
            If Left(cbo_destination.List(i), start) = cbo_destination.Text Then
                cbo_destination.Text = cbo_destination.List(i)
            End If
        Next
        cbo_destination.SelStart = start
        cbo_destination.SelLength = Len(cbo_destination.Text)
    End If
End Sub

Private Sub cbo_destination_GotFocus()
Call Affiche_Destination_Combo(cbo_destination)
End Sub


Private Sub cbo_destination_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cbo_destination_KeyUp(KeyCode As Integer, Shift As Integer)
  thekey = KeyCode
    theshift = Shift
End Sub

Private Sub cbo_vehicle_Change()
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
        start = Len(cbo_vehicle.Text)
        For i = 0 To cbo_vehicle.ListCount - 1
            If Left(cbo_vehicle.List(i), start) = cbo_vehicle.Text Then
                cbo_vehicle.Text = cbo_vehicle.List(i)
            End If
        Next
        cbo_vehicle.SelStart = start
        cbo_vehicle.SelLength = Len(cbo_vehicle.Text)
    End If
End Sub


Private Sub cbo_vehicle_GotFocus()

    Call Affiche_Matricule_Combo(cbo_vehicle)

End Sub


Private Sub cbo_vehicle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub cbo_vehicle_KeyUp(KeyCode As Integer, Shift As Integer)
     thekey = KeyCode
    theshift = Shift
End Sub


Private Sub SCommand1_Click()
Dim SelectedV As String
Dim SelectedC As String
Dim SelectedD As String
Dim SelectedCE As Long
Dim SelectedCS
Dim Observation As String
Dim HeureSortie As Date
Dim HeureEntre As Date
Dim NumFiche As String
Dim SQL As String
Dim rs As New ADODB.Recordset

On Error GoTo Err

        NumFiche = lbl_NumFiche.Caption
        SelectedV = cbo_vehicle.Text
        SelectedC = cbo_Conducteur.Text
        SelectedD = cbo_destination.Text
        SelectedCE = CStr(txt_cptEntre.Text)
        SelectedCS = CStr(txt_cptSortie.Text)
        Observation = txt_observation.Text
        
        HeureSortie = Format(cbo_DateSortie.Text, "d/m/yyyy") & " " & Format(txt_HeureSortie.Text, "hh:mm:ss")
         If Not (IsNull(SelectedCE)) And (SelectedCE > 0) Then
        HeureEntre = Format(cbo_DateEntre.Text, "d/m/yyyy") & " " & Format(txt_heureEntre.Text, "hh:mm:ss")
        End If
        
    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    'Modification
        
        CNB.BeginTrans
        SQL = "Update fichetraffic Set "
        SQL = SQL & " Vehicule = " & SQLText(SelectedV)
        SQL = SQL & ", CompteurSortie = " & SelectedCS
        SQL = SQL & ", Conducteur = " & SQLText(SelectedC)
        SQL = SQL & ", Destination = " & SQLText(SelectedD)
        SQL = SQL & ", Observation = " & SQLText(Observation)
        SQL = SQL & ",  HeureSortie = " & SQLText(HeureSortie)
        If Not (IsNull(SelectedCE)) And (SelectedCE > 0) Then
        SQL = SQL & ",  CompteurEntre = " & SelectedCE
        SQL = SQL & ",  HeureEntre = " & SQLText(HeureEntre)
        End If
        SQL = SQL & " where fichetraffic.Numero = " & SQLText(NumFiche)
     CNB.Execute SQL
    CNB.CommitTrans
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
End If
    Call Frm_Statistiques.Cmd_Searchft_Click
   ' Call Frm_Statistiques.Affiche_FT(Frm_Statistiques.cbo_VehiculeFT, Frm_Statistiques.cbo_ConducteurFT, Frm_Statistiques.cbo_DestinationFT, Frm_Statistiques.cda_Debutft, Frm_Statistiques.cda_FinFT)
    Frm_Statistiques.Tab_Satistiques.Tab = 2
    Unload Me
Exit Sub
Err:
CNB.RollbackTrans
Exit Sub
MsgBox Err.Description, vbInformation
End Sub

Private Sub SCommand2_Click()
Unload Me
End Sub


