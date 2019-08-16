Attribute VB_Name = "MainBase"

Option Explicit

'Initialise SBiComboBox (Conducteur, Vehicule, Destination, Fournisseur Achat,...)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Sub Initialise_SBICombo_Cond(ByRef SBiCbo As SBiCombo)
     Dim LObj_Find As New Conducteur
     Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetAll_ConducteursActif(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    SBiCbo.AddItem "0000", "Tous"
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
            SBiCbo.AddItem Lrs_Find("Code"), Lrs_Find("Libelle")
        Lrs_Find.MoveNext
        Wend
    End If
    SBiCbo.ListIndex = 0
    Set LObj_Find = Nothing
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Initialise_SBICombo_PngDest(ByRef SBiCbo As SBiCombo)
    Dim LObj_Find As New DESTINATION
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.Get_toutDestPLNG(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    SBiCbo.AddItem "0000", "Tous"
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
                SBiCbo.AddItem Lrs_Find("Numero"), Lrs_Find("Libelle")
            Lrs_Find.MoveNext
        Wend
    End If
    SBiCbo.ListIndex = 0
    Set LObj_Find = Nothing
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Initialise_SBICombo_AllDest(ByRef SBiCbo As SBiCombo)
    Dim LObj_Find As New DESTINATION
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.Get_ActifDest(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    SBiCbo.AddItem "0000", "Tous"
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
                SBiCbo.AddItem Lrs_Find("Numero"), Lrs_Find("Libelle")
            Lrs_Find.MoveNext
        Wend
    End If
    SBiCbo.ListIndex = 0
    Set LObj_Find = Nothing
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Initialise_SBICombo_Vehic(ByRef SBiCbo As SBiCombo)
    Dim LObj_Find As New VEHICULE
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetAllActifVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    SBiCbo.AddItem "0000", "Tous"
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
            SBiCbo.AddItem Lrs_Find("Code"), Lrs_Find("Matricule")
        Lrs_Find.MoveNext
        Wend
    End If
    SBiCbo.ListIndex = 0
    Set LObj_Find = Nothing
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
Public Sub Initialise_SBICombo_FseurPH(ByRef SBiCbo As SBiCombo)
    Dim LObj_Find As New Fournisseur
    Dim Lrs_Find As New Recordset
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetRow_Fournisseurs_ByType(ErrNumber, ErrDescription, ErrSourceDetail, "Fournisseur Achat", CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    SBiCbo.AddItem "0000", "Tous"
    If Not Lrs_Find.EOF Then
        While Not Lrs_Find.EOF
            SBiCbo.AddItem Lrs_Find("Code"), Lrs_Find("Libelle")
        Lrs_Find.MoveNext
        Wend
    End If
    Set LObj_Find = Nothing
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub

