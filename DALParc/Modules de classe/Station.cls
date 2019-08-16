VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Station"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "Station"

'Incrementation ou décrementation du nombre des BonsCarb de la station en question
Public Sub UpdateNBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal un As Integer, _
              ByVal txt As String)

    Const sPROC_NAME As String = OBJ_NAME & ".UpdateNBC"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update station set numbc = numbc + " & un & " where code =  " & SQLText(txt)

    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Sub UpdateNBV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal un As Integer, _
              ByVal txt As String)

    Const sPROC_NAME As String = OBJ_NAME & ".UpdateNBV"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update station set NUMBV = NUMBV + " & un & " where code =  " & SQLText(txt)

    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Sub Update_NUMFCT(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal un As Integer, _
              ByVal txt As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_NUMFCT"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update station set NUMFCT = NUMFCT + " & un & " where code =  " & SQLText(txt)

    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Retourne une station active par son code ou libelle
Public Function GetStatByCodeLib(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetStatByCodeLib"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "Select * from station where code = " & SQLText(vcode) & " OR libelle like '" & vcode & "%' "

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set GetStatByCodeLib = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne une station active par son code
Public Function GetStationByCode(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetStationByCode"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "Select * from station where Actif=1 And code = " & SQLText(vcode)

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set GetStationByCode = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


'Retourne tout les stations active de type Fournisseur
Public Function GetStat_Fournisseur(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetStat_Fournisseur"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Station  where  actif=1 And type='Fournisseur' order by libelle"

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set GetStat_Fournisseur = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function



'Retourne liste des stations de type fournisseur
Public Function GetAllStat_Fournis(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetAllStat_Fournis"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Station  where type='Fournisseur' order by libelle"

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set GetAllStat_Fournis = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne tout les stations active de tout les types
Public Function Get_ActifStat(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ActifStat"

    Dim StrSql As String
    Dim LRs_Record As New ADODB.Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Station  where  actif=1 order by libelle"
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_ActifStat = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function Get_ActifStatRep(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ActifStatRep"

    Dim StrSql As String
    Dim LRs_Record As New ADODB.Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Station  where  actif=1 and type<>'Fournisseur Achat' order by libelle"
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_ActifStatRep = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne tout les stations de tout les types
Public Function Get_AllStat(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AllStat"

    Dim StrSql As String
    Dim LRs_Record As New ADODB.Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Station  order by libelle"
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_AllStat = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function Get_StatRep(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_StatRep"

    Dim StrSql As String
    Dim LRs_Record As New ADODB.Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Station where type <>'Fournisseur Achat' order by libelle"
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_StatRep = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne tout les stations active de type StationCarburant

Public Function Get_StationCarb(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_StationCarb"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Station where Actif=1 And Type ='Station carburant' order by libelle"

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_StationCarb = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne le nombre des bonCarburant de la station en question
Public Function Get_NBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_NBC"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select Count(Numero) as numbc from AssBoncarburant where station = " & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_NBC = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_NumFact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_NumFact"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select numFCT from station where code= " & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_NumFact = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne le nombre des bonCarburant de la station en question
Public Function Get_NBV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_NBV"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select Count(Numero) as numbvdg from Ass_BonVidange where Station = " & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_NBV = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Sub Delete_Stat(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vSupp As String, _
              ByVal CodeUser As String, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Stat"

    Dim SQL As String

    On Error GoTo ErrHandler
    If vSupp <> "N" And vSupp <> "O" Then
        MsgBox "Variable Invalide", vbCritical
        Exit Sub
    End If
    
    SQL = "Update Station set"
    If vSupp = "O" Then
        SQL = SQL & " supp=" & SQLText(vSupp)
        SQL = SQL & ", Userdelete=" & SQLText(CodeUser)
    ElseIf vSupp = "N" Then
        SQL = SQL & " supp=" & SQLText(vSupp)
        SQL = SQL & ", UserAdd=" & SQLText(CodeUser)
    End If
    SQL = SQL & "where code = " & Val(vcode)
    CNB.Execute SQL

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Sub Insert_Stat(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_Stat"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
    StrSql = "Insert into station (Code,Libelle,Type,Adresse,Ville,CPOSTAL,Activite,telephone,mobile,fax,email, Actif,UserInsert) values ("

    If IsNull(PRs_Record!code) Then
        StrSql = StrSql & "Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!code)
    End If
    
     If IsNull(PRs_Record!Libelle) Then
        StrSql = StrSql & ", Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Libelle)
    End If
    
    If IsNull(PRs_Record!Type) Then
        StrSql = StrSql & "," & "NULL"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Type)
    End If
    
    If IsNull(PRs_Record!Adresse) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Adresse)
    End If
    
    If IsNull(PRs_Record!Ville) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Ville)
    End If
    
    If IsNull(PRs_Record!CPOSTAL) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Val(PRs_Record!CPOSTAL)
    End If
    
    If IsNull(PRs_Record!Activite) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Activite)
    End If
    
    If IsNull(PRs_Record!Telephone) Then
        StrSql = StrSql & "," & "NULL"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Telephone)
    End If
    
    If IsNull(PRs_Record!Mobile) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Mobile)
    End If
    
    If IsNull(PRs_Record!fax) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!fax)
    End If
    
    If IsNull(PRs_Record!email) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!email)
    End If
    
    If IsNull(PRs_Record!Actif) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Val(PRs_Record!Actif)
    End If
    
    If IsNull(PRs_Record!USERINSERT) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!USERINSERT)
    End If

    StrSql = StrSql & ")"

    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Sub Update_Stat(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_Stat"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update Station set"
   
    If IsNull(PRs_Record!Libelle) Then
        StrSql = StrSql & " Libelle = Null"
    Else
        StrSql = StrSql & " Libelle = " & SQLText(PRs_Record!Libelle)
    End If
    
    If IsNull(PRs_Record!Type) Then
        StrSql = StrSql & ", Type = Null"
    Else
        StrSql = StrSql & ", Type = " & SQLText(PRs_Record!Type)
    End If
    
    If IsNull(PRs_Record!Adresse) Then
        StrSql = StrSql & ", Adresse = Null"
    Else
        StrSql = StrSql & ", Adresse = " & SQLText(PRs_Record!Adresse)
    End If
    
    If IsNull(PRs_Record!Ville) Then
        StrSql = StrSql & " , Ville = Null"
    Else
        StrSql = StrSql & " , Ville = " & SQLText(PRs_Record!Ville)
    End If
    
    If IsNull(PRs_Record!CPOSTAL) Then
        StrSql = StrSql & " , CPOSTAL = 0"
    Else
        StrSql = StrSql & " , CPOSTAL = " & Val(PRs_Record!CPOSTAL)
    End If
    
    If IsNull(PRs_Record!Activite) Then
        StrSql = StrSql & " , Activite = Null"
    Else
        StrSql = StrSql & " , Activite = " & SQLText(PRs_Record!Activite)
    End If
    
    If IsNull(PRs_Record!Telephone) Then
        StrSql = StrSql & " , telephone = Null"
    Else
        StrSql = StrSql & " , telephone = " & SQLText(PRs_Record!Telephone)
    End If
    
    If IsNull(PRs_Record!Mobile) Then
        StrSql = StrSql & " , mobile = Null"
    Else
        StrSql = StrSql & " , mobile = " & SQLText(PRs_Record!Mobile)
    End If
    
    If IsNull(PRs_Record!fax) Then
        StrSql = StrSql & " , fax = Null"
    Else
        StrSql = StrSql & " , fax = " & SQLText(PRs_Record!fax)
    End If


    If IsNull(PRs_Record!email) Then
        StrSql = StrSql & " , email = Null"
    Else
        StrSql = StrSql & " , email = " & SQLText(PRs_Record!email)
    End If
    
    If IsNull(PRs_Record!Actif) Then
        StrSql = StrSql & " , Actif = 0"
    Else
        StrSql = StrSql & " , Actif = " & Val(PRs_Record!Actif)
    End If
    
    If IsNull(PRs_Record!userupdate) Then
        StrSql = StrSql & " , UserUpdate = Null"
    Else
        StrSql = StrSql & " , UserUpdate = " & SQLText(PRs_Record!userupdate)
    End If


    StrSql = StrSql & " Where Code = " & SQLText(PRs_Record!code)
    
 CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

