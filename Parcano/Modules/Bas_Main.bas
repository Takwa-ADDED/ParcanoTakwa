Attribute VB_Name = "Bas_Main"
Option Explicit
Public N_XLSNG As Long
Public CNB As ADODB.Connection
Public CNR As ADODB.Connection
Public ErrNumber As Long
Public ErrSourceDetail As String
Public ErrDescription As String
'Vriable connection
Public LOBJ_CON As New DBConnexion
'Code utilisateur
Public LInt_UserId As String
Public LInt_UserIdMaj As String
'Nom et Prénom utilisateur
Public LStr_NameUser As String
Private Sub Main()
    'si pas de nouvelle version est l'exe commence par 0000000 copier son contenu dans le vrai est lancer le vrai sans 000000
    If App.EXEName = "000000_Parcano" Then
        CopyFileAny App.Path & "\" & App.EXEName & ".exe", App.Path & "\" & Replace(App.EXEName, "000000_", "", , , vbTextCompare) & ".exe"
        Shell App.Path & "\" & Replace(App.EXEName, "000000_", "", , , vbTextCompare) & ".exe", vbNormalFocus
        Exit Sub
    Else
        If TEST_NEW_VERSION = True Then
            Call CHECK_NEW_VERSION
        Else
            On Error Resume Next
            If ExisteFile(App.Path & "\000000_Parcano.exe") Then Kill (App.Path & "\000000_Parcano.exe")
            On Error GoTo 0
            frmconnexionx.Show
        End If
        Shell ("regsvr32 " & App.Path & "\GestionParc.dll /s")
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub
Public Function ExisteFile(Xfile As String) As Boolean
    
    On Error Resume Next
    
    Dim exist
    exist = Dir(Xfile)
    If exist <> "" And Xfile <> "" Then
        ExisteFile = True
    Else
        ExisteFile = False
    End If
    
End Function
Public Sub N_Ligne(GridS As SGrid)
    With GridS
        If .Rows > 0 Then
            Dim i As Integer
            For i = 1 To .Rows
                .CellDetails i, .ColumnIndex("N"), i, DT_RIGHT, , &H80000015, &H80000005  '&H8000000D, &H80000009
            Next i
        End If
    End With
End Sub
Public Function Get_NameUserByCode(ByVal VCode As String) As String

Dim Lobj_User As Utilisateur
Dim rs As New Recordset

Set Lobj_User = New Utilisateur
Set rs = Lobj_User.GetRow_UserByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Function
End If
If Not rs.EOF Then
   Get_NameUserByCode = rs("NomPrn")
End If

End Function

Public Sub ExistDonnee(ByVal cbo As SBiCombo)
    Dim RCount As Integer, i As Integer, exist As Boolean, MConducteur As String, TConducteur As String
On Error GoTo Err
    RCount = cbo.ListCount
    TConducteur = cbo.Text
    For i = 0 To RCount - 1
        cbo.ListIndex = i
        MConducteur = cbo.Text
        If TConducteur = MConducteur Then
            exist = True
            Exit For
        Else
            exist = False
        End If
    Next i
    If i = RCount Then
        If exist = False Then
            MsgBox "Saisie invalide!...    " & vbCr & "Vérifier donnée saisie...    ", vbExclamation, App.ProductName
            cbo.Text = ""
        End If
    Else
        cbo.ListIndex = i
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub
Public Sub ExistDonneeCbo(ByVal cbo As ComboBox)
    Dim RCount As Integer, i As Integer, exist As Boolean, MConducteur As String, TConducteur As String
On Error GoTo Err
    RCount = cbo.ListCount
    TConducteur = cbo.Text
    For i = 0 To RCount - 1
        cbo.ListIndex = i
        MConducteur = cbo.Text
        If TConducteur = MConducteur Then
            exist = True
            Exit For
        Else
            exist = False
        End If
    Next i
    If i = RCount Then
        If exist = False Then
            MsgBox "Saisie invalide!...    " & vbCr & "Vérifier donnée saisie...    ", vbExclamation, App.ProductName
            cbo.Text = ""
        End If
    Else
        cbo.ListIndex = i
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub ViderZone(ByVal POBJ_FORM As Form)
    Dim ctl
    For Each ctl In POBJ_FORM.Controls

        If TypeOf ctl Is TextBox Or TypeOf ctl Is STextBox Or _
            TypeOf ctl Is SCombo Or TypeOf ctl Is SDataCombo Or TypeOf ctl Is SBiCombo Or TypeOf ctl Is SBiCombo2 Then
                ctl.Text = ""

        ElseIf TypeOf ctl Is SMTextBox Then
                ctl.Text = ""

        ElseIf TypeOf ctl Is SDateBox Then
                ctl.Text = "__/__/____"

        ElseIf TypeOf ctl Is CheckBox Then
            ctl.Value = 0
        End If
        
        

    Next ctl

End Sub
Function SQLText(txt As Variant) As String

    Dim N As Integer
    N = InStr(1, txt, "'")
    If N > 0 Then
        SQLText = "'" & Left$(txt, N) & SQLText(Right$(txt, Len(txt) - N))
    Else
        SQLText = "'" & txt & "'"
    End If
    
End Function
'Cette fonction exige la saisie au champs obligatoire
Public Function CheckMandatory(ByVal POBJ_FORM As Form) As Integer

Dim ctl
For Each ctl In POBJ_FORM.Controls

    If TypeOf ctl Is TextBox Or TypeOf ctl Is STextBox Or TypeOf ctl Is ComboBox Or _
        TypeOf ctl Is SCombo Or TypeOf ctl Is SDataCombo Or TypeOf ctl Is SBiCombo Or TypeOf ctl Is SBiCombo2 Then

        If ctl.Enabled And Trim(ctl.Text) = "" And Left(ctl.Tag, 1) = "M" Then
            MsgBox "Le champ Indiqué ne doit pas être vide", vbQuestion
            ctl.Text = Trim(ctl.Text)
            ctl.BackColor = &HC0FFFF
            ctl.SetFocus
            CheckMandatory = "1"
            CheckMandatory = CheckMandatory & Mid(ctl.Tag, 2, 1)
            
            Exit Function
        End If

    ElseIf TypeOf ctl Is SMTextBox Then

        If ctl.Enabled And Trim(Replace(ctl.Text, vbNewLine, "")) = "" And Left(ctl.Tag, 1) = "M" Then
    
            ctl.Text = Trim(Replace(ctl.Text, vbNewLine, ""))
            MsgBox "Le champ indiqué ne doit pas être vide", vbQuestion
            ctl.SetFocus
            CheckMandatory = "1"
            CheckMandatory = CheckMandatory & Mid(ctl.Tag, 2, 1)

            Exit Function
        End If

    ElseIf TypeOf ctl Is SDateBox Then

        If ctl.Enabled And ctl.Text = "__/__/____" And Left(ctl.Tag, 1) = "M" Or _
           ctl.Enabled And ctl.Text = "" And Left(ctl.Tag, 1) = "M" Then
            
            MsgBox "Le champ indiqué ne doit pas être vide", vbQuestion
            'ctl.SetFocus
            CheckMandatory = "1"
            CheckMandatory = CheckMandatory & Mid(ctl.Tag, 2, 1)

            Exit Function
        End If

    ElseIf TypeOf ctl Is STimeBox Then
        If ctl.Enabled And ctl.Text = "__ : __ : __" And Left(ctl.Tag, 1) = "M" Or _
           ctl.Enabled And ctl.Text = "" And Left(ctl.Tag, 1) = "M" Then
            
            MsgBox "Le champ indiqué ne doit pas être vide", vbQuestion
            ctl.SetFocus
            CheckMandatory = "1"
            CheckMandatory = CheckMandatory & Mid(ctl.Tag, 2, 1)
            Exit Function
        End If
    
    End If

Next ctl

End Function

Public Function Crement_Compteur(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal Cn As ADODB.Connection, _
              ByVal NameProcStock As String, _
              ByVal LStr_VCompteur As String) As Long
              
    
    Dim datcmd As ADODB.Command
    
    On Error GoTo ErrHandler
    
    Set datcmd = New ADODB.Command
    Set datcmd.ActiveConnection = Cn
    
    datcmd.CommandText = NameProcStock
    datcmd.CommandType = adCmdStoredProc
    datcmd.Parameters.Append datcmd.CreateParameter("@Str", adVarChar, adParamInput, 50, LStr_VCompteur)
    datcmd.Parameters.Append datcmd.CreateParameter("@ValCompt", adDouble, adParamOutput, 0)
    datcmd.Execute , , adExecuteNoRecords
    Crement_Compteur = datcmd.Parameters("@ValCompt").Value

    Set datcmd = Nothing
    
    Exit Function
ErrHandler:

    Set datcmd = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description

End Function

Public Sub MouseOn()
    Screen.MousePointer = vbHourglass
End Sub
Public Sub MouseOff()
    Screen.MousePointer = vbDefault
End Sub

'Afficher Liste de tout les personnels actifs
Public Sub Affiche_Personnel_Combo(cbo As ComboBox)

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_AllActifPers(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        With cbo
            .AddItem rs("libelle")
        End With
        rs.MoveNext
    Wend
End If
End Sub
'Afficher Liste de tout les personnels actifs
Public Sub Affiche_Personnel_SBCombo(cbo As SBiCombo)

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_AllActifPers(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        With cbo
            .AddItem rs("Code"), rs("libelle")
        End With
        rs.MoveNext
    Wend
End If
End Sub

'Afficher tout les stations de tout les types
Public Sub Affiche_Station_SBCombo(cbo As SBiCombo)
cbo.Clear
Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.Get_AllStat(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    cbo.AddItem "0000", "Tous"
    While Not rs.EOF
        With cbo
            .AddItem rs("Code"), rs("libelle")
        End With
        rs.MoveNext
    Wend
    cbo.ListIndex = 0
End If

End Sub
Public Sub Affiche_StatRep_SBCombo(cbo As SBiCombo)
cbo.Clear
Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.Get_StatRep(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    cbo.AddItem "0000", "Tous"
    While Not rs.EOF
        With cbo
            .AddItem rs("Code"), rs("libelle")
        End With
        rs.MoveNext
    Wend
    cbo.ListIndex = 0
End If

End Sub
Public Sub Affiche_StatCarb_SBCombo(cbo As SBiCombo)

cbo.Clear
Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.Get_StationCarb(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    cbo.AddItem "0000", "Tous"
    While Not rs.EOF
        With cbo
            .AddItem rs("Code"), rs("libelle")
        End With
        rs.MoveNext
    Wend
    
End If

End Sub
'Afficher tout les stations de tout les types
Public Sub Affiche_Station_Combo(cbo As ComboBox)

Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.Get_AllStat(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        With cbo
            .AddItem rs("libelle")
        End With
        rs.MoveNext
    Wend
End If
End Sub
Public Sub Affiche_StatRep_Combo(cbo As ComboBox)

Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.Get_StatRep(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        With cbo
            .AddItem rs("libelle")
        End With
        rs.MoveNext
    Wend
End If
End Sub
'Afficher tout les stations de tout les types
Public Sub Affiche_StatCarb_Combo(cbo As ComboBox)

Dim LOBJ_Station As Station
Dim rs As New Recordset

Set LOBJ_Station = New Station
Set rs = LOBJ_Station.Get_StationCarb(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        With cbo
            .AddItem rs("libelle")
        End With
        rs.MoveNext
    Wend
End If
End Sub


'Afficher liste des matricules de tout les véhicules
Public Sub Affiche_Matricule_Combo(cbo As ComboBox)

Dim Lobj_Vehicule As VEHICULE
Dim rs As New Recordset

Set Lobj_Vehicule = New VEHICULE
Set rs = Lobj_Vehicule.Get_AllMatrVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        With cbo
            .AddItem rs("matricule")
        End With
        rs.MoveNext
    Wend

End If
End Sub

Public Sub Affiche_Matricule_SBCombo(cbo As SBiCombo)

Dim Lobj_Vehicule As VEHICULE
Dim rs As New Recordset

Set Lobj_Vehicule = New VEHICULE
Set rs = Lobj_Vehicule.Get_AllMatrVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        With cbo
            .AddItem rs("Code"), rs("matricule")
        End With
        rs.MoveNext
    Wend

End If
End Sub


'afficher liste des stations de type fournisseur
Public Sub Affiche_Fournisseur_Combo(cbo As ComboBox)

Dim LOBJ_Station As Station
Dim rs As New Recordset

cbo.Clear
Set LOBJ_Station = New Station
Set rs = LOBJ_Station.GetAllStat_Fournis(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        With cbo
            .AddItem rs("libelle")
        End With
        rs.MoveNext
    Wend
End If
End Sub

Public Sub Affiche_Type_Combo(cbo As ComboBox)

Dim Lobj_Destination As DESTINATION
Dim rs As New Recordset

cbo.Clear
Set Lobj_Destination = New DESTINATION
Set rs = Lobj_Destination.Get_toutTypeDest(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

If Not rs.EOF Then
    While Not rs.EOF
        With cbo
            .AddItem rs("type")
        End With
        rs.MoveNext
    Wend
End If
End Sub

'Affficher tout les libelles des destinations
Public Sub Affiche_Destination_Combo(cbo As ComboBox)

Dim Lobj_Destination As DESTINATION
Dim rs As New Recordset

Set Lobj_Destination = New DESTINATION
Set rs = Lobj_Destination.Get_toutLibDest(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

If Not rs.EOF Then
    While Not rs.EOF
        With cbo
            .AddItem rs("Libelle")
        End With
        rs.MoveNext
    Wend
End If
End Sub

Public Sub Affiche_Personnel_ListBox(LB As SListBox)

Dim LOBJ_Personnel As personnel
Dim rs As New Recordset
LB.Clear

Set LOBJ_Personnel = New personnel
Set rs = LOBJ_Personnel.Get_AllPers(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

If Not rs.EOF Then
    While Not rs.EOF
        With LB
            .AddItem rs("libelle")
        End With
        rs.MoveNext
    Wend
End If
End Sub

Public Sub Affiche_Destination_ListBox(LB As SListBox)

Dim Lobj_Destination As DESTINATION
Dim rs As New Recordset
LB.Clear
Set Lobj_Destination = New DESTINATION
Set rs = Lobj_Destination.Get_DestByType(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
If Not rs.EOF Then
    While Not rs.EOF
        With LB
            .AddItem rs("libelle")
        End With
        rs.MoveNext
    Wend
End If
End Sub

'Affiche liste des véhicules order by Matricule
Public Sub Affiche_Matricule_ListBox(LB As SListBox)

Dim Lobj_Vehicule As VEHICULE
Dim rs As New Recordset
LB.Clear
Set Lobj_Vehicule = New VEHICULE
Set rs = Lobj_Vehicule.Get_AllVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

If Not rs.EOF Then
    While Not rs.EOF
        With LB
            .AddItem rs("matricule")
        End With
        rs.MoveNext
    Wend
End If
End Sub

Public Function CHECK_ACCES(ByVal acces As String, ByVal User As String) As Boolean
    Dim LObj_Find As New Utilisateur
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.USER_ACCESS(ErrNumber, ErrDescription, ErrSourceDetail, CNB, acces, User)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Function
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        CHECK_ACCES = True
        Set Lrs_Find = Nothing
        Exit Function
    Else
        CHECK_ACCES = False
        Set Lrs_Find = Nothing
    End If
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Function
'=================================================
'Fonctions Retourne Compteur Traffic, Compteur Véhicule, Opertaion***
'=================================================
Public Function return_Compteur() As Long
    Dim LObj_Find   As New Traffic
    Dim Lrs_Traffic As Recordset
On Error GoTo Err
    return_Compteur = 0
    Set Lrs_Traffic = LObj_Find.GetMAx_NumeroTraffic(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Function
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Traffic.EOF Then return_Compteur = Lrs_Traffic(0)
    Set Lrs_Traffic = Nothing
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Function
Public Function return_CDVehi(ByVal matric As String) As String
    Dim LObj_Find   As New VEHICULE
    Dim Lrs As Recordset
On Error GoTo Err
    return_CDVehi = ""
    Set Lrs = LObj_Find.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, matric)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Function
    End If
    Set LObj_Find = Nothing
    If Not Lrs.EOF Then return_CDVehi = Lrs.Fields("Code")
    Set Lrs = Nothing
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Function
Public Function ReturnOperation(ByVal Matricule As String) As String()
    Dim LObj_Find   As New Traffic
    Dim Lrs_Traffic As Recordset
    Dim Tableau()   As String
    ReDim Tableau(1)
    Dim Name_Tab    As String
On Error GoTo Err
    Name_Tab = "FicheTraffic"
    Set Lrs_Traffic = LObj_Find.GetTraffic_ByMarticuleVehicule(ErrNumber, ErrDescription, ErrSourceDetail, Name_Tab, Matricule, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Function
    End If
    Set LObj_Find = Nothing
    Tableau(0) = ("S")
    While Not Lrs_Traffic.EOF
        If (Lrs_Traffic("Numero")) = "36560" Then
        
        Dim x As Integer
        x = 1
        End If
        If IsNull(Lrs_Traffic("HeureEntre")) Then
            Tableau(0) = ("E")
            Tableau("1") = (Lrs_Traffic("Numero"))
        End If
     Lrs_Traffic.MoveNext
     Wend
    ReturnOperation = Tableau()
    Set Lrs_Traffic = Nothing
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Function
Public Function CompteurVehicule(ByVal Matricule As String) As Long
    Dim LObj_Find       As New VEHICULE
    Dim Lrs_Vehicule    As Recordset
On Error GoTo Err
    Set Lrs_Vehicule = LObj_Find.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Matricule)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Function
    End If
    Set LObj_Find = Nothing
    If Not (IsNull(Lrs_Vehicule("CompteurFT"))) Then CompteurVehicule = Lrs_Vehicule("CompteurFT")
    Set Lrs_Vehicule = Nothing
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Function
'# Code de Lubrifiant...
Public Function GetCode_Lubrif(ByVal VCode As String) As Integer
    Dim LOBJ_Lub    As New Produit_Lubrifiant
    Dim rs          As New Recordset
On Error GoTo Err
    Set rs = LOBJ_Lub.Get_ProdLubByLib(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Function
    End If
    Set LOBJ_Lub = Nothing
    If Not rs.EOF Then
        GetCode_Lubrif = Val(rs("Numero"))
    End If
    Set rs = Nothing
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Function
'# Dernier compteur du véhicule en entrant (ficheTraffic)
Public Function MaxCompteurVehicule(ByVal VCode As String) As String
    Dim Lobj_Vehicule   As New VEHICULE
    Dim rs1             As New Recordset
    Dim Name_Tab        As String
On Error GoTo Err
    MaxCompteurVehicule = "0"
    Name_Tab = "FicheTraffic"
    Set rs1 = Lobj_Vehicule.Get_DerCompt(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Name_Tab, VCode)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Function
    End If
    Set Lobj_Vehicule = Nothing
    If Not rs1.EOF Then
        If Not IsNull(rs1("maxCpt")) Then MaxCompteurVehicule = rs1("maxCpt")
    End If
    Set rs1 = Nothing
Exit Function
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName
End Function
