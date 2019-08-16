VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form Frm_MajStatFT 
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
      Picture         =   "Frm_MajStatFT.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "Frm_MajStatFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim thekey As Integer
    Dim theshift As Integer

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
Private Sub Cbo_Conducteur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub Cbo_Conducteur_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
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

Private Sub SCommand2_Click()
    Unload Me
End Sub
Private Sub Cbo_Conducteur_Change()
     Dim I As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_Conducteur.Text)
        For I = 0 To cbo_Conducteur.ListCount - 1
            If Left(cbo_Conducteur.List(I), start) = cbo_Conducteur.Text Then
                cbo_Conducteur.Text = cbo_Conducteur.List(I)
            End If
        Next
        cbo_Conducteur.SelStart = start
        cbo_Conducteur.SelLength = Len(cbo_Conducteur.Text)
    End If
End Sub
Private Sub cbo_destination_Change()
    Dim I As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_destination.Text)
        For I = 0 To cbo_destination.ListCount - 1
            If Left(cbo_destination.List(I), start) = cbo_destination.Text Then
                cbo_destination.Text = cbo_destination.List(I)
            End If
        Next
        cbo_destination.SelStart = start
        cbo_destination.SelLength = Len(cbo_destination.Text)
    End If
End Sub
Private Sub cbo_vehicle_Change()
    Dim I As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_vehicle.Text)
        For I = 0 To cbo_vehicle.ListCount - 1
            If Left(cbo_vehicle.List(I), start) = cbo_vehicle.Text Then
                cbo_vehicle.Text = cbo_vehicle.List(I)
            End If
        Next
        cbo_vehicle.SelStart = start
        cbo_vehicle.SelLength = Len(cbo_vehicle.Text)
    End If
End Sub
Private Sub Cbo_Conducteur_GotFocus()
    On Error GoTo Err
    Call Affiche_Personnel_Combo(cbo_Conducteur)
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Public Sub selectFT(ByVal VCode As String)
    Dim LObj_Find As New Traffic
    Dim Lrs_Trafic As New Recordset

On Error GoTo Err
    
    Set Lrs_Trafic = LObj_Find.GetRow_Traffic_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    
    If Not Lrs_Trafic.EOF Then
        If (Not (IsNull(Lrs_Trafic("Numero")))) Then lbl_NumFiche.Caption = Lrs_Trafic("Numero")
        If (Not (IsNull(Lrs_Trafic("Vehicule")))) Then cbo_vehicle.Text = Lrs_Trafic("MatVehi")
        If (Not (IsNull(Lrs_Trafic("Conducteur")))) Then cbo_Conducteur.Text = Lrs_Trafic("LibCond")
        If (Not (IsNull(Lrs_Trafic("Destination")))) Then cbo_destination.Text = Lrs_Trafic("LibDest")
        If (Not (IsNull(Lrs_Trafic("CompteurEntre")))) Then txt_cptEntre.Text = Lrs_Trafic("CompteurEntre")
        If (Not (IsNull(Lrs_Trafic("CompteurSortie")))) Then txt_cptSortie.Text = Lrs_Trafic("CompteurSortie")
        If (Not (IsNull(Lrs_Trafic("HeureEntre")))) Then
            cbo_DateEntre.Text = Format(Lrs_Trafic("HeureEntre"), "dd/mm/yyyy")
            txt_heureEntre.Text = Format(Lrs_Trafic("HeureEntre"), "hh:mm:ss")
        End If
        If (Not (IsNull(Lrs_Trafic("HeureSortie")))) Then
            cbo_DateSortie.Text = Format(Lrs_Trafic("HeureSortie"), "dd/mm/yyyy")
            txt_HeureSortie.Text = Format(Lrs_Trafic("HeureSortie"), "hh:mm:ss")
        End If
        If (Not (IsNull(Lrs_Trafic("ObservationEntre")))) Then txt_observation.Text = Lrs_Trafic("ObservationEntre")
    End If
    Set Lrs_Trafic = Nothing

Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub SCommand1_Click()
    Dim LObj_FindUser As Utilisateur
    Dim LObj_Find As New Traffic
    Dim Lrs_Trafic As New Recordset, Lrs_User As Recordset
    Dim Lobj_Vehicule As VEHICULE
    Dim Lobj_Conducteur As CONDUCTEUR
    Dim Lobj_Destination As DESTINATION
    Dim Lrs_Conducteur As Recordset
    Dim Lrs_Vehicule As Recordset
    Dim Lrs_Destination As Recordset
    Dim Code_Conducteur As String, Code_vehicule As String, Code_Destination As String
    Dim SelectedCE As Long, SelectedCS As Long
    Dim HeureSortie As Date, HeureEntre As Date

On Error GoTo Err
    
    Set LObj_FindUser = New Utilisateur
    Set Lrs_User = LObj_FindUser.GetRow_User_Maj_FT(ErrNumber, ErrDescription, ErrSourceDetail, LInt_UserId, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_FindUser = Nothing

    If Lrs_User.EOF Then
        Set Lrs_User = Nothing
        MsgBox "Modification n'est pas accessible!.." & vbNewLine & "Vous ne disposez peut-etre pas des autorisations nécessaires pour modifier un Traffic", vbExclamation, "Parcano..."
        Exit Sub
    End If
    Set Lrs_User = Nothing


    If cbo_Conducteur.Text = "" Or cbo_vehicle.Text = "" Or cbo_destination.Text = "" Then
        MsgBox "Vérifier votre insertion!...", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    If CStr(txt_cptEntre.Text) < CStr(txt_cptSortie.Text) Then
        MsgBox "Vérifier Compteur d'entré", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    '-- Code Conducteur***
        Set Lobj_Conducteur = New CONDUCTEUR
        Set Lrs_Conducteur = Lobj_Conducteur.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, cbo_Conducteur.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set Lobj_Conducteur = Nothing
        If Not Lrs_Conducteur.EOF Then Code_Conducteur = Lrs_Conducteur("Code")
        Set Lrs_Conducteur = Nothing
    
    '-- Code Vehicule***
        Set Lobj_Vehicule = New VEHICULE
        Set Lrs_Vehicule = Lobj_Vehicule.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, cbo_vehicle.Text)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set Lobj_Vehicule = Nothing
        If Not Lrs_Vehicule.EOF Then Code_vehicule = Lrs_Vehicule("Code")
        Set Lrs_Vehicule = Nothing
    
    '-- Code Destination***
        Set Lobj_Destination = New DESTINATION
        Set Lrs_Destination = Lobj_Destination.GetRow_Destination_Bycode(ErrNumber, ErrDescription, ErrSourceDetail, cbo_destination.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set Lobj_Destination = Nothing
        If Not Lrs_Destination.EOF Then Code_Destination = Lrs_Destination("Numero")
        Set Lrs_Destination = Nothing
        
    SelectedCE = CStr(txt_cptEntre.Text)
    SelectedCS = CStr(txt_cptSortie.Text)
    HeureSortie = Format(cbo_DateSortie.Text, "d/m/yyyy") & " " & Format(txt_HeureSortie.Text, "hh:mm:ss")
    If Not (IsNull(SelectedCE)) And (SelectedCE > 0) Then HeureEntre = Format(cbo_DateEntre.Text, "d/m/yyyy") & " " & Format(txt_heureEntre.Text, "hh:mm:ss")

    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
    If Trim(txt_observation.Text) = "" Then txt_observation.Text = "Sans Observation"
    Set Lrs_Trafic = CreateEmptyRS_Traffic()
    With Lrs_Trafic
        .AddNew
        .Fields("Vehicule") = Code_vehicule
        .Fields("CompteurSortie") = SelectedCS
        .Fields("Conducteur") = Code_Conducteur
        .Fields("Destination") = Code_Destination
        .Fields("CompteurEntre") = SelectedCE
        .Fields("Observation") = txt_observation.Text
        .Fields("HeureSortie") = Format(HeureSortie, "dd/mm/yyyy hh:mm:ss")
        .Fields("HeureEntre") = Format(HeureEntre, "dd/mm/yyyy hh:mm:ss")
        .Fields("UserUpdate") = LInt_UserId
    End With
    Call LObj_Find.UpDate_Traffic(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Lrs_Trafic, lbl_NumFiche.Caption, SelectedCE, cbo_destination.Text)

    MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
    
    Call Frm_Statistiques.Cmd_Searchft_Click
    Frm_Statistiques.Tab_Satistiques.Tab = 2
    Unload Me
   
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
