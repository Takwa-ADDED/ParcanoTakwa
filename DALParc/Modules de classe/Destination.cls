VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Destination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "Destination"

Public Function Get_Destination(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_Destination"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Destination order by libelle "

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_Destination = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_DestTrafic(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_DestTrafic"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Destination where type <> 'Planning' order by libelle "

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_DestTrafic = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


'Retourne tout les types des destinations
Public Function Get_toutTypeDest(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_toutTypeDest"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select Distinct(type) from Destination"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_toutTypeDest = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne tout les libelles des destinations
Public Function Get_toutLibDest(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_toutLibDest"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select Distinct(Libelle) from Destination where Type <> 'Planning'"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_toutLibDest = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne tout les libelles des destinations
Public Function Get_toutDestPLNG(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_toutDestPLNG"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select Numero , Libelle from Destination where Type = 'Planning' and Actif = 1 "

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_toutDestPLNG = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne toute les destinations actives
Public Function Get_ActifDest(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ActifDest"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Destination where Actif = 1  order by libelle"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_ActifDest = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_Tournee(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_Tournee"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Tournee order by libelle"
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_Tournee = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Sub UpDate_Destination(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CNB As ADODB.Connection, _
                            ByVal PRs_Record As ADODB.Recordset, _
                            ByVal NumeroTxt As String)

    Const sPROC_NAME As String = OBJ_NAME & ".UpDate_Destination"

    Dim SQL As String
    
    On Error GoTo ErrHandler

    SQL = "Update destination Set "
    SQL = SQL & " Type = " & SQLText(PRs_Record!Type)
    SQL = SQL & " , Libelle = " & SQLText(PRs_Record!Libelle)
    SQL = SQL & " , Actif = " & PRs_Record!Actif
    SQL = SQL & " ,MaxDuree = " & SQLText(Format((PRs_Record!MaxDuree), "hh:mm:ss"))
    SQL = SQL & " , MaxCompteur = " & SQLText(PRs_Record!MaxCompteur)
    SQL = SQL & " , MinCompteur = " & SQLText(PRs_Record!MinCompteur)
    SQL = SQL & " , UserUpdate = " & SQLText(PRs_Record!USERUPDATE)
    If Not IsNull(PRs_Record!Temps) Then
        SQL = SQL & " , Temps = " & SQLText(PRs_Record!Temps)
    Else
        SQL = SQL & " , Temps = NULL "
    End If
    SQL = SQL & " where Numero = " & NumeroTxt
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub Save_Destination(ByRef ErrNumber As Long, _
                        ByRef ErrDescription As String, _
                        ByRef ErrSourceDetail As String, _
                        ByVal CNB As ADODB.Connection, _
                        ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Save_Destination"

    Dim SQL As String
    
On Error GoTo ErrHandler

    SQL = "Insert into Destination (Numero,Type,Libelle, Actif,MaxDuree,MaxCompteur,MinCompteur,UserInsert,Temps) values ("
    SQL = SQL & SQLText(PRs_Record!numero)
    SQL = SQL & "," & SQLText(PRs_Record!Type)
    SQL = SQL & "," & SQLText(PRs_Record!Libelle)
    SQL = SQL & "," & PRs_Record!Actif
    SQL = SQL & "," & SQLText(Format((PRs_Record!MaxDuree), "hh:mm:ss"))
    SQL = SQL & "," & SQLText(PRs_Record!MaxCompteur)
    SQL = SQL & "," & SQLText(PRs_Record!MinCompteur)
    SQL = SQL & "," & SQLText(PRs_Record!USERINSERT)
    If Not IsNull(PRs_Record!Temps) Then SQL = SQL & "," & SQLText(PRs_Record!Temps)
    SQL = SQL & ")"
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub

Public Function GetNumero_MaxCompteur(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "select Max(Numero) from Destination"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetNumero_MaxCompteur = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function Get_DestByType(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_DestByType"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "Select * from Destination order by Type "

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_DestByType = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetAll_DestinationActif(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from Destination where actif = 1 and Type <> 'Planning'  order by Type, Libelle"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetAll_DestinationActif = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_Destination_ByLibelle(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal Libelle As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * From destination where libelle=" & SQLText(Libelle) & " Or Numero= " & SQLText(Libelle)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_Destination_ByLibelle = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_Details_Dest(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal Vcode As String, _
                                    ByVal DateD As Date, _
                                    ByVal DateF As Date, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

SQL = "select F.Numero, D.Libelle as Destination, V.Matricule as Vehicule, P.Libelle as Conducteur,  F.Compteursortie, F.CompteurEntre, F.HeureSortie, F.HeureEntre, F.OperateurSortie, F.OperateurEntre"
SQL = SQL & " from FicheTraffic F "
SQL = SQL & " left join Vehicule V on F.Vehicule = V.Code"
SQL = SQL & " left join Personnel P on F.Conducteur = P.Code"
SQL = SQL & " left join Destination D on F.Destination = D.Numero"

SQL = SQL & " where F.destination =" & SQLText(Vcode) & "  and Cast(F.HeureSortie as date) between " & SQLText(DateD) & " and " & SQLText(DateF)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_Details_Dest = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function GetRow_Destination_ByLib(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal code As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * From destination where libelle like '" & code & "%' "
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_Destination_ByLib = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_Destination_ByCode(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal code As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * From destination where Numero = " & SQLText(code) & " Or libelle= " & SQLText(code)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_Destination_ByCode = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Sub Delete_Add_Destination(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal code As String, _
                                ByVal vSupp As String, _
                                ByVal CodeUser As String, _
                                ByVal CNB As ADODB.Connection)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Add_Destination"

    Dim SQL As String
    
    If vSupp <> "N" And vSupp <> "O" Then
        MsgBox "variable supp invalide", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    SQL = "Update Destination set "
    SQL = SQL & " Supp=" & SQLText(vSupp)
    If vSupp = "O" Then
        SQL = SQL & " , UserDelete=" & SQLText(CodeUser)
    ElseIf vSupp = "N" Then
        SQL = SQL & " , UserAdd=" & SQLText(CodeUser)
    End If
    SQL = SQL & " where Numero =" & SQLText(code)
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
'=====================================================================
'PLANNING
Public Function GetAll_DestinationActifNonSupp(ByRef ErrNumber As Long, _
                                                        ByRef ErrDescription As String, _
                                                        ByRef ErrSourceDetail As String, _
                                                        ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetAll_DestinationActifDisponibleExist"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Destination where Actif = 1  and supp='N'order by libelle"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set GetAll_DestinationActifNonSupp = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetAll_DestinationActifTourneeDisponibleExist(ByRef ErrNumber As Long, _
                                                            ByRef ErrDescription As String, _
                                                            ByRef ErrSourceDetail As String, _
                                                            ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from Destination where actif = 1 and type='Planning' and supp='N' order by Ordre"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetAll_DestinationActifTourneeDisponibleExist = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


Public Function Get_SumNbrVoyage(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_SumNbrVoyage"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    
    SQL = " Select count(*) As nbVoy, (Sum (F.CompteurEntre - F.CompteurSortie)) as valeur,  sum(DATEDIFF(HOUR, F.HeureSortie, F.HeureEntre)) as totduree, sum(DATEDIFF(Minute, F.HeureSortie, F.HeureEntre)) as totdureeM from ficheTraffic F, Destination D where"
    SQL = SQL & " F.Destination = D.Numero And F.Supp ='N' and Cast(F.HeureSortie as date)"
    SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF)
'   BETWEEN CAST('" & Ddebut & "' AS DATETIME) AND CAST('" & Dfin & "' AS DATETIME)"
'    SQL = " Select count(*) As nbrVoy ,Sum (Distinct F.Destination)As valeur from ficheTraffic F, Destination D where"
'    SQL = SQL & " F.Numero = D.Numero And F.Supp ='N' and F.HeureSortie"
'    SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF)
 
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_SumNbrVoyage = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_ValRepStatistDest(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal Vcode As String, _
              ByVal VdateD As Date, _
              ByVal vDateF As Date) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ValRepStatistDest"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler


 SQL = " Select D.Numero, libelle, count(*) As nbVoy, (Sum (F.CompteurEntre - F.CompteurSortie)) as valeur,  sum(DATEDIFF(HOUR, F.HeureSortie, F.HeureEntre)) as totduree, sum(DATEDIFF(Minute, F.HeureSortie, F.HeureEntre)) as totdureeM from ficheTraffic F, Destination D where"
    SQL = SQL & " F.Destination = D.Numero And F.Supp ='N' and D.libelle=" & SQLText(Vcode) & " and Cast(F.HeureSortie as date)"
    SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF) & " Group by Libelle, D.Numero"

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_ValRepStatistDest = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function Get_NbrRepStatistGrpDest(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_NbrRepStatistGrpDest"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

'SQL = " Select vehicule , count(*) As nbrRep  from AsspieceReparation A, detailpieceReparation D where"
'SQL = SQL & " A.Numero = D.Numero and A.Supp ='N' "
'SQL = SQL & " And A.datePiece"
'SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF) & " group by Vehicule"
 SQL = " Select D.Numero, Libelle, count(*) As nbVoy, (Sum (F.CompteurEntre - F.CompteurSortie)) as valeur,  sum(DATEDIFF(HOUR, F.HeureSortie, F.HeureEntre)) as totduree, sum(DATEDIFF(Minute, F.HeureSortie, F.HeureEntre)) as totdureeM from ficheTraffic F, Destination D where"
    SQL = SQL & " F.Destination = D.Numero And F.Supp ='N' and Cast(F.HeureSortie as date)"
    SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF) & "  group by Libelle, D.Numero"
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_NbrRepStatistGrpDest = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function GetDestByLibDest(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal Vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetDestByLibDest"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "Select * from destination libelle = " & SQLText(Vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetDestByLibDest = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

