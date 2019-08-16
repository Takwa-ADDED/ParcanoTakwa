VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conducteur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "Conducteur"

Public Function GetRow_Conducteur_ByLibelle(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal Libelle As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL         As String
    Dim LRs_Record  As New Recordset
    
On Error GoTo ErrHandler
    SQL = "SELECT * FROM PERSONNEL WHERE ACTIF=1 and LIBELLE like " & SQLText(Libelle & "%")
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_Conducteur_ByLibelle = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetAll_ConducteursActif(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetAll_ConducteursActif"
    Dim SQL As String, LRs_Record As Recordset

On Error GoTo ErrHandler

    SQL = "SELECT * FROM PERSONNEL WHERE ACTIF=1 ORDER BY LIBELLE"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetAll_ConducteursActif = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function











Public Function GetAll_Abreviation(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal ABr As String, _
                                        ByVal code As String, _
                                        ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL         As String
    Dim LRs_Record  As Recordset
On Error GoTo ErrHandler
    SQL = "SELECT ABREVIATION FROM PERSONNEL WHERE ABREVIATION=" & SQLText(ABr) & " and Code != " & SQLText(code)
        
        Set LRs_Record = New Recordset
        LRs_Record.CursorLocation = adUseClient
        LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
        Set GetAll_Abreviation = LRs_Record
        Set LRs_Record = Nothing
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_ABrByLibelle(ByRef ErrNumber As Long, _
                                                    ByRef ErrDescription As String, _
                                                    ByRef ErrSourceDetail As String, _
                                                    ByVal Lib As String, _
                                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL         As String
    Dim LRs_Record  As Recordset
On Error GoTo ErrHandler
    SQL = "SELECT Abreviation FROM PERSONNEL WHERE Libelle=" & SQLText(Lib)
        Set LRs_Record = New Recordset
        LRs_Record.CursorLocation = adUseClient
        LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
        Set GetRow_ABrByLibelle = LRs_Record
        Set LRs_Record = Nothing
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_Conducteur_ByCode(ByRef ErrNumber As Long, _
                                            ByRef ErrDescription As String, _
                                            ByRef ErrSourceDetail As String, _
                                            ByVal CodeConducteur As String, _
                                            ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL         As String
    Dim LRs_Record  As Recordset
On Error GoTo ErrHandler
    SQL = "SELECT * FROM PERSONNEL WHERE CODE = " & SQLText(CodeConducteur) & " OR Libelle Like '" & CodeConducteur & "%'"
        Set LRs_Record = New Recordset
        LRs_Record.CursorLocation = adUseClient
        LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
        Set GetRow_Conducteur_ByCode = LRs_Record
        Set LRs_Record = Nothing
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Sub Update_Desp_Conducteur(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal vDesp As String, _
                                ByVal Libelle As String, _
                                ByVal CNB As ADODB.Connection)
On Error GoTo ErrHandler
    If vDesp <> "N" And vDesp <> "O" Then
        MsgBox "variable vDesp invalide", vbExclamation
        Exit Sub
    End If
    CNB.Execute "UPDATE PERSONNEL SET DISPONIBLE= " & SQLText(vDesp) & " WHERE LIBELLE=" & SQLText(Libelle)
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub



Public Sub Update_DispPers_In_DispoPerso(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal Libelle As String, _
                                        ByVal ETAT As String, _
                                        ByVal CNB As ADODB.Connection)
    Dim SQL As String
    
On Error GoTo ErrHandler
    
    SQL = "Update DispoPerso Set HFin=" & SQLText(Now)
    SQL = SQL & " where Numero = (Select Max(Numero) from DispoPerso where Conducteur=" & SQLText(Libelle) & " And Etat = " & SQLText(ETAT) & ")"
    CNB.Execute SQL
    
Exit Sub
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub
Public Sub Save_DispoPerso(ByRef ErrNumber As Long, _
                    ByRef ErrDescription As String, _
                    ByRef ErrSourceDetail As String, _
                    ByVal CNB As ADODB.Connection, _
                    ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Save_DispoPerso"

    Dim SQL As String
    
    On Error GoTo ErrHandler

    SQL = "Insert into DispoPerso (Numero,Conducteur,Etat,HDebut) values ("
    SQL = SQL & SQLText(PRs_Record!numero)
    SQL = SQL & "," & SQLText(PRs_Record!Conducteur)
    SQL = SQL & "," & SQLText(PRs_Record!ETAT)
    SQL = SQL & "," & SQLText(PRs_Record!HDebut)
    SQL = SQL & ")"
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub Delete_Restaurer(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CodeCond As String, _
                            ByVal CodeUser As String, _
                            ByVal vSupp As String, _
                            ByVal CNB As ADODB.Connection)
    Dim SQL As String
    
On Error GoTo ErrHandler
    
    If vSupp <> "N" And vSupp <> "O" Then
        MsgBox "variable supp invalide", vbExclamation
        Exit Sub
    End If
    
    SQL = "Update Personnel set"
    SQL = SQL & " Supp=" & SQLText(vSupp)
    If vSupp = "O" Then
        SQL = SQL & ", userdelete=" & SQLText(CodeUser)
    ElseIf vSupp = "N" Then
        SQL = SQL & ", UserAdd=" & SQLText(CodeUser)
    End If
    SQL = SQL & " where code = " & SQLText(CodeCond)
    CNB.Execute SQL
    
Exit Sub
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub
Public Sub Update_Conducteur(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CodeCond As String, _
                            ByVal CNB As ADODB.Connection, _
                            ByVal PRs_Record As ADODB.Recordset)
    Dim SQL As String
    
On Error GoTo ErrHandler
    
    SQL = "Update Personnel set"
    SQL = SQL & " Libelle= " & SQLText(PRs_Record!Libelle) & ","
    SQL = SQL & "CIN= " & SQLText(PRs_Record!CIN) & ","
    SQL = SQL & " Fonction= " & SQLText(PRs_Record!Fonction) & ","
    SQL = SQL & " Telephone= " & SQLText(PRs_Record!Telephone) & ","
    SQL = SQL & " Mobile= " & SQLText(PRs_Record!Mobile) & ","
    SQL = SQL & " Permie= " & SQLText(PRs_Record!Permie) & ","
    SQL = SQL & " Datlivr= " & SQLText(PRs_Record!Datelivr) & ","
    SQL = SQL & " lieulivr= " & SQLText(PRs_Record!lieulivr) & ","
    SQL = SQL & " Abreviation= " & SQLText(PRs_Record!ABr) & ","
    SQL = SQL & " Actif= " & PRs_Record!Actif & ","
    If Not IsNull(PRs_Record!PicBox) Then SQL = SQL & " PicBox= " & SQLText(PRs_Record!PicBox) & ","
    SQL = SQL & " userupdate= " & SQLText(PRs_Record!userupdate)
    SQL = SQL & " where code = " & SQLText(CodeCond)
    CNB.Execute SQL
    
Exit Sub
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub
Public Sub Insert_Conducteur(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CNB As ADODB.Connection, _
                            ByVal PRs_Record As ADODB.Recordset)
    Dim SQL As String
    
On Error GoTo ErrHandler
    
    SQL = "Insert into personnel(Code, Libelle,Abreviation, CIN, Fonction, telephone, mobile, permie, datlivr, lieulivr, Actif, disponible, picbox, userinsert) values ("
    SQL = SQL & SQLText(PRs_Record!numero)
    SQL = SQL & "," & SQLText(PRs_Record!Libelle)
    SQL = SQL & "," & SQLText(PRs_Record!ABr)
    If Not IsNull(PRs_Record!CIN) And (PRs_Record!CIN) <> "" Then
        SQL = SQL & "," & PRs_Record!CIN
    Else
        SQL = SQL & "," & 0
    End If
    SQL = SQL & "," & SQLText(PRs_Record!Fonction)
    If Not IsNull(PRs_Record!Telephone) And (PRs_Record!Telephone) <> "" Then
        SQL = SQL & "," & SQLText(PRs_Record!Telephone)
    Else
        SQL = SQL & "," & 0
    End If
    SQL = SQL & "," & SQLText(PRs_Record!Mobile)
    SQL = SQL & "," & SQLText(PRs_Record!Permie)
    SQL = SQL & "," & SQLText(PRs_Record!Datelivr)
    SQL = SQL & "," & SQLText(PRs_Record!lieulivr)
    SQL = SQL & "," & PRs_Record!Actif
    SQL = SQL & "," & SQLText(PRs_Record!Disponible)
    If Not IsNull(PRs_Record!PicBox) Then SQL = SQL & "," & SQLText(PRs_Record!PicBox)
    SQL = SQL & "," & SQLText(PRs_Record!USERINSERT)
    SQL = SQL & ")"
    CNB.Execute SQL
    
Exit Sub
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub

'=======================================================
'=======================Gestion des congés =============
'Liste des congés
Public Function Get_AllCongeConduc(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal date_deb As Date, _
              ByVal date_fin As Date, _
              ByVal Conduc As String, _
              ByVal sup As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AllCongeConduc"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select G.* , P.Libelle from GestionConge G INNER JOIN Personnel P On G.Conducteur=P.Code "
   StrSql = StrSql & " Where Type ='Congé' And "
   If sup = "O" Then
        StrSql = StrSql & " G.Supp = 'O' "
    ElseIf sup = "N" Then
        StrSql = StrSql & " G.Supp = 'N' "
    End If
   StrSql = StrSql & " And ((g.datedu >= " & SQLText(date_deb) & " And g.datedu <= " & SQLText(date_fin)
   StrSql = StrSql & ") Or (" & SQLText(date_deb) & " > = g.datedu And " & SQLText(date_deb) & " <= g.dateau ))"
   
    If Conduc = "0000" Or Conduc = "" Then
        StrSql = StrSql & " Order By G.Conducteur"
    Else
        StrSql = StrSql & "And G.Conducteur = " & SQLText(Conduc)
    End If
    

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_AllCongeConduc = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_CongeCondByCode(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As Integer) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_CongeCondByCode"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select G.* , P.Libelle from GestionConge G INNER JOIN Personnel P On G.Conducteur=P.Code Where G.Numero = " & Val(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_CongeCondByCode = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Sub Insert_Conge(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_Conge"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
    StrSql = "Insert into GestionConge (Conducteur, DateDu, DateAu,Type, Observation , UserInsert) Values ("
    
     If IsNull(PRs_Record!Conducteur) Then
        StrSql = StrSql & "Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!Conducteur)
    End If
    
    If IsNull(PRs_Record!DateDu) Then
        StrSql = StrSql & ", Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!DateDu)
    End If
    
     If IsNull(PRs_Record!DateAu) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!DateAu)
    End If
    
    If IsNull(PRs_Record!Type) Then
        StrSql = StrSql & ", Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Type)
    End If
    
    If IsNull(PRs_Record!Observation) Then
        StrSql = StrSql & ", Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Observation)
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

Public Sub Update_Conge(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CNB As ADODB.Connection, _
                            ByVal PRs_Record As ADODB.Recordset)
     
     Const sPROC_NAME As String = OBJ_NAME & ".Update_Conge"
     
    Dim SQL As String
    
On Error GoTo ErrHandler
    
    SQL = "Update GestionConge set"
    SQL = SQL & " DateDu = " & SQLText(PRs_Record!DateDu)
    SQL = SQL & ", DateAu = " & SQLText(PRs_Record!DateAu)
    
    If IsNull(PRs_Record!Observation) Then
        SQL = SQL & ", Observation = Null"
    Else
        SQL = SQL & ", Observation = " & SQLText(PRs_Record!Observation)
    End If
    
    SQL = SQL & ", UserUpdate = " & SQLText(PRs_Record!userupdate)
    SQL = SQL & " where Numero = " & SQLText(PRs_Record!numero)
    CNB.Execute SQL
    
Exit Sub
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub

Public Sub Delete_Conge(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CodeCong As String, _
                            ByVal CodeUser As String, _
                            ByVal CNB As ADODB.Connection)
    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Conge"
    
    Dim SQL As String
    
On Error GoTo ErrHandler
    
    SQL = "Update GestionConge set Supp= 'O' , userdelete = " & SQLText(CodeUser)
    SQL = SQL & " where Numero = " & SQLText(CodeCong)
    
    CNB.Execute SQL
    
Exit Sub
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub

Public Function Get_MaxNumConge(ByRef ErrNumber As Long, _
                                                ByRef ErrDescription As String, _
                                                ByRef ErrSourceDetail As String, _
                                                ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT Max(Numero) as maxNum from GESTIONCONGE "
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set Get_MaxNumConge = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_CongeByCodeCondAndDate(ByRef ErrNumber As Long, _
                                                ByRef ErrDescription As String, _
                                                ByRef ErrSourceDetail As String, _
                                                ByVal CodeConducteur As String, _
                                                ByVal DateTraffic As Date, _
                                                ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT G.NUMERO, C.LIBELLE"
    SQL = SQL & " FROM GESTIONCONGE G"
    SQL = SQL & " INNER JOIN PERSONNEL C ON G.CONDUCTEUR=C.CODE"
    SQL = SQL & " WHERE Type = 'Congé' And CONDUCTEUR=" & SQLText(CodeConducteur) & " AND G.SUPP='N' AND "
    SQL = SQL & " ((DATEDU <= (cast(" & SQLText(DateTraffic) & " AS DATETIME))) AND (DATEAU >= (CAST(" & SQLText(DateTraffic) & " AS DATETIME))))"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_CongeByCodeCondAndDate = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function GetRow_CongeByDate(ByRef ErrNumber As Long, _
                                                ByRef ErrDescription As String, _
                                                ByRef ErrSourceDetail As String, _
                                                ByVal DateTraffic As Date, _
                                                ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT G.NUMERO, C.LIBELLE"
    SQL = SQL & " FROM GESTIONCONGE G"
    SQL = SQL & " INNER JOIN PERSONNEL C ON G.CONDUCTEUR=C.CODE"
    SQL = SQL & " WHERE Type = 'Congé' AND G.SUPP='N' AND "
    SQL = SQL & " ((DATEDU <= (cast(" & SQLText(DateTraffic) & " AS DATETIME))) AND (DATEAU >= (CAST(" & SQLText(DateTraffic) & " AS DATETIME))))"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_CongeByDate = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'========================= Repos =======================
'=======================================================
Public Function Get_repos(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal date_repos As Date, _
              ByVal date_fin As Date) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_repos"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "SELECT G.* , P.LIBELLE "
    StrSql = StrSql & " FROM GESTIONCONGE G"
    StrSql = StrSql & " INNER JOIN PERSONNEL P ON G.CONDUCTEUR=P.CODE"
    StrSql = StrSql & " WHERE TYPE = 'REPOS' AND G.SUPP = 'N' AND "
    StrSql = StrSql & " G.DATEDU >= CAST(" & SQLText(date_repos) & " AS DATETIME) AND G.DATEDU <= CAST(" & SQLText(date_fin) & " AS DATETIME) ORDER BY P.LIBELLE"


    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_repos = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


Public Function Get_ReposByCodeCondAndDate(ByRef ErrNumber As Long, _
                                                ByRef ErrDescription As String, _
                                                ByRef ErrSourceDetail As String, _
                                                ByVal CodeConducteur As String, _
                                                ByVal DateRepos As Date, _
                                                ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT G.NUMERO, C.LIBELLE"
    SQL = SQL & " FROM GESTIONCONGE G"
    SQL = SQL & " INNER JOIN PERSONNEL C ON G.CONDUCTEUR=C.CODE"
    SQL = SQL & " WHERE Type = 'Repos' And CONDUCTEUR=" & SQLText(CodeConducteur)
    SQL = SQL & " AND DATEDU = " & SQLText(DateRepos)
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set Get_ReposByCodeCondAndDate = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function Get_ReposByDate(ByRef ErrNumber As Long, _
                                                ByRef ErrDescription As String, _
                                                ByRef ErrSourceDetail As String, _
                                                ByVal DateRepos As Date, _
                                                ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT G.NUMERO, C.LIBELLE AS CONDUCTEUR , G.Observation,DATEDU "
    SQL = SQL & " FROM GESTIONCONGE G"
    SQL = SQL & " INNER JOIN PERSONNEL C ON G.CONDUCTEUR=C.CODE"
    SQL = SQL & " WHERE Type = 'Repos' "
    SQL = SQL & " AND DATEDU = " & SQLText(DateRepos)
    SQL = SQL & "  AND g.Supp ='N'"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set Get_ReposByDate = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'GestionConducteurs***

'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'PLANNING***
                                                
    Public Function GetAll_ConducteursActifNonSupp(ByRef ErrNumber As Long, _
                                                ByRef ErrDescription As String, _
                                                ByRef ErrSourceDetail As String, _
                                                ByVal Actif As String, _
                                                ByVal Supp As String, _
                                                ByVal CNB As ADODB.Connection) As Recordset
                                                
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

'    SQL = "SELECT * FROM personnel  "
'    If StrSource = "Personnel_actif" Then SQL = SQL & " WHERE ACTIF=1 AND SUPP='N'"
'    SQL = SQL & "ORDER BY Libelle"

    SQL = "Select * from personnel "
    If Actif = "O" And Supp = "N" Then
        SQL = SQL & " where Actif = 1 and supp='N' "

    End If
    SQL = SQL & " order by Libelle "

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetAll_ConducteursActifNonSupp = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'GestionConge***
Public Function Get_AllCongeByCond_Date(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal CNB As ADODB.Connection, _
                                        ByVal DateDu As String, _
                                        ByVal DateFin As String, _
                                        ByVal Code_Cond As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AllCongeByCond_Date"
    Dim SQL As String, LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT G.NUMERO,C.LIBELLE AS CONDUCTEUR,G.TYPE,G.OBSERVATION,G.DATEDU,G.DATEAU,G.SUPP "
    SQL = SQL & " FROM GESTIONCONGE G"
    SQL = SQL & " INNER JOIN PERSONNEL C ON G.CONDUCTEUR=C.CODE"
    SQL = SQL & " WHERE Type ='Congé' And G.Supp = 'N' And ((G.DATEDU >= CAST(" & SQLText(DateDu) & " AS DATETIME)) AND (G.DATEDU <= CAST(" & SQLText(DateFin) & " AS DATETIME))) "
    If Code_Cond <> "0000" Then SQL = SQL & " AND G.CONDUCTEUR=" & SQLText(Code_Cond)
    SQL = SQL & " ORDER BY G.DATEDU"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_AllCongeByCond_Date = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Sub Update_Repos(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CNB As ADODB.Connection, _
                            ByVal PRs_Record As ADODB.Recordset)
     
    Const sPROC_NAME As String = OBJ_NAME & ".Update_Repos"
    Dim SQL As String
    
On Error GoTo ErrHandler
    
    SQL = "Update GestionConge set"
    SQL = SQL & " Conducteur = " & SQLText(PRs_Record!Conducteur)
    SQL = SQL & ", Observation= " & SQLText(PRs_Record!Observation)
    SQL = SQL & ", UserUpdate = " & SQLText(PRs_Record!userupdate)
    SQL = SQL & " where Numero = " & PRs_Record!numero
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub
Public Sub Delete_Repos(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal Cond As String, _
                            ByVal DateRepos As String, _
                            ByVal CodeUser As String, _
                            ByVal CNB As ADODB.Connection)
    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Repos"
    
    Dim SQL As String
    
On Error GoTo ErrHandler
    
    SQL = "Update GestionConge set Supp= 'O' , userdelete = " & SQLText(CodeUser)
    SQL = SQL & " where Type = 'Repos' And Conducteur = " & SQLText(Cond) & " And DateDu = " & SQLText(DateRepos)
    
    CNB.Execute SQL
    
Exit Sub
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub


