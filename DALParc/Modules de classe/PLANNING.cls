VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PLANNING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Private Const OBJ_NAME As String = "PLANNING"

Public Function GetRow_PLANNINGByDateDu(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal DateDu As String, _
                                        ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT P.CODE, P.DATEDU, P.DATEAU, P.JOUR, P.SUPP, P.DATECREATION, P.HeureEntre, "
    SQL = SQL & " C.LIBELLE AS CONDUCTEUR,"
    SQL = SQL & " V.MATRICULE AS VEHICULE,"
    SQL = SQL & " T.LIBELLE AS TOURNEE,"
    SQL = SQL & " U.NOMPRN AS USERINSERT"
    SQL = SQL & " FROM PLANNING P"
    SQL = SQL & " INNER JOIN PERSONNEL C ON C.CODE=P.CONDUCTEUR"
    SQL = SQL & " FULL JOIN VEHICULE V ON V.CODE=P.VEHICULE"
    SQL = SQL & " INNER JOIN DESTINATION T ON T.NUMERO=P.TOURNEE"
    SQL = SQL & " INNER JOIN UTILISATEUR U ON U.CODE=P.USERINSERT"
    SQL = SQL & " WHERE P.Supp='N' AND CONVERT(VARCHAR(10),P.DATEDU,103)=" & SQLText(DateDu)
    SQL = SQL & " ORDER BY P.CODE"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_PLANNINGByDateDu = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetCode_PLANNING(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal DateDu As String, _
                                ByVal jour As String, _
                                ByVal Tournee As String, _
                                ByVal Conducteur As String, _
                                ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT CODE FROM PLANNING"
    SQL = SQL & " WHERE CONVERT(VARCHAR(10),DATEDU,103)=" & SQLText(DateDu) & " AND"
    SQL = SQL & " CONDUCTEUR=" & SQLText(Conducteur) & " AND"
    SQL = SQL & " TOURNEE =" & SQLText(Tournee) & " AND"
    SQL = SQL & " JOUR = " & SQLText(jour)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetCode_PLANNING = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetPLANNING_ByCode(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal CodePLANNING As String, _
                                ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT * FROM PLANNING"
    SQL = SQL & " WHERE CODE=" & SQLText(CodePLANNING)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetPLANNING_ByCode = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Sub Save_PLANNING(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CNB As ADODB.Connection, _
                            ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Save_PLANNING"
    Dim SQL As String
    
On Error GoTo ErrHandler

    SQL = "INSERT INTO PLANNING (DATECREATION, DATEDU, DATEAU, TOURNEE, CONDUCTEUR, VEHICULE,DATEJOUR, "
    SQL = SQL & "JOUR, HEUREENTRE, USERINSERT) VALUES ("
    SQL = SQL & SQLText(PRs_Record!DATECREAT)
    SQL = SQL & "," & SQLText(PRs_Record!DateDu)
    SQL = SQL & "," & SQLText(PRs_Record!DateAu)
    SQL = SQL & "," & SQLText(PRs_Record!Tournee)
    SQL = SQL & "," & SQLText(PRs_Record!Conducteur)
    If Not (IsNull(PRs_Record!Vehicule)) Then SQL = SQL & "," & SQLText(PRs_Record!Vehicule) Else SQL = SQL & ",NULL"
    SQL = SQL & "," & SQLText(PRs_Record!DATEJOUR)
    SQL = SQL & "," & SQLText(PRs_Record!jour)
    SQL = SQL & "," & SQLText(PRs_Record!HeureEntre)
    SQL = SQL & "," & SQLText(PRs_Record!USERINSERT)
    SQL = SQL & ")"
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub Update_PLANNING(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CodePLANNING As Integer, _
                            ByVal CNB As ADODB.Connection, _
                            ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_PLANNING"
    Dim SQL As String
    
    On Error GoTo ErrHandler

        SQL = "Update PLANNING"
        SQL = SQL & " SET TOURNEE=" & SQLText(PRs_Record!Tournee)
        SQL = SQL & "," & "CONDUCTEUR=" & SQLText(PRs_Record!Conducteur)
        SQL = SQL & "," & "DateJour=" & SQLText(PRs_Record!DATEJOUR)
        If Not (IsNull(PRs_Record!Vehicule)) Then SQL = SQL & "," & "VEHICULE=" & SQLText(PRs_Record!Vehicule) Else SQL = SQL & ", VEHICULE= NULL"
        SQL = SQL & "," & "DATEEDIT=" & SQLText(PRs_Record!DateEdit)
        SQL = SQL & "," & "HeureEntre=" & SQLText(PRs_Record!HeureEntre)
        SQL = SQL & "WHERE CODE=" & SQLText(CodePLANNING)
        CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Function GetDate_WEnd(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal DateDu As String, _
                            ByVal Interval As Integer, _
                            ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset

On Error GoTo ErrHandler

    SQL = "SELECT DATEADD(DAY," & Interval & "," & SQLText(DateDu) & ") as DateWEnd" ' FROM ASS_PLANNING"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetDate_WEnd = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetDate_NewPLANNING(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal DateDu As String, _
                            ByVal Interval As Integer, _
                            ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset

On Error GoTo ErrHandler

    SQL = "SELECT DateAdd(DAY," & Interval & "," & SQLText(DateDu) & ") as DateDebut" ' FROM ASS_PLANNING"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetDate_NewPLANNING = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetHeureEntre_TourneeByDest(ByRef ErrNumber As Long, _
                                            ByRef ErrDescription As String, _
                                            ByRef ErrSourceDetail As String, _
                                            ByVal DateSortie As String, _
                                            ByVal CodeDest As String, _
                                            ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset

On Error GoTo ErrHandler

    
    SQL = "SELECT CONVERT(VARCHAR(8),CAST(" & SQLText(DateSortie) & " AS DATETIME) + D.MAXDUREE,108) AS DUREE "
    SQL = SQL & " From Destination d"
'    SQL = SQL & " FROM PLANNING P"
'    SQL = SQL & " INNER JOIN DESTINATION D ON P.Tournee=D.NUMERO"
    SQL = SQL & " Where d.Numero = " & SQLText(CodeDest)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetHeureEntre_TourneeByDest = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Sub Save_TmpPLANNING1(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal DateDu As Date, _
                            ByVal DateAu As Date, _
                            ByVal jour As String, _
                            ByVal Detail As String, _
                            ByVal Tournee As String, _
                            ByVal CNB As ADODB.Connection)

    Const sPROC_NAME As String = OBJ_NAME & ".Save_TmpPLANNING"
    Dim SQL As String
    
    On Error GoTo ErrHandler


        SQL = "Insert Into TMP_PLANNING (DATEDU,DATEAU, TOURNEE," & jour & ") Values ("
        SQL = SQL & Format(SQLText(DateDu), " dd/mm/yyyy hh:mm:ss")
        SQL = SQL & "," & Format(SQLText(DateAu), " dd/mm/yyyy hh:mm:ss")
        SQL = SQL & "," & SQLText(Tournee)
        SQL = SQL & "," & SQLText(Detail)
        SQL = SQL & ")"
        CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub Save_TmpPLANNING(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal DateDu As Date, _
                            ByVal DateAu As Date, _
                            ByVal Tournee As String, _
                            ByVal LUNDI As String, _
                            ByVal MARDI As String, _
                            ByVal MERCREDI As String, _
                            ByVal JEUDI As String, _
                            ByVal VENDREDI As String, _
                            ByVal SAMDI As String, _
                            ByVal DIMANCHE As String, _
                            ByVal CNB As ADODB.Connection)

    Const sPROC_NAME As String = OBJ_NAME & ".Save_TmpPLANNING"
    Dim SQL As String
    
    On Error GoTo ErrHandler


        SQL = "Insert Into TMP_PLANNING (DATEDU,DATEAU, TOURNEE,LUNDI, MARDI, MERCREDI, JEUDI, VENDREDI, SAMDI, DIMANCHE) Values ("
        SQL = SQL & Format(SQLText(DateDu), " dd/mm/yyyy hh:mm:ss")
        SQL = SQL & "," & Format(SQLText(DateAu), " dd/mm/yyyy hh:mm:ss")
        SQL = SQL & "," & SQLText(Tournee)
        SQL = SQL & "," & SQLText(LUNDI)
        SQL = SQL & "," & SQLText(MARDI)
        SQL = SQL & "," & SQLText(MERCREDI)
        SQL = SQL & "," & SQLText(JEUDI)
        SQL = SQL & "," & SQLText(VENDREDI)
        SQL = SQL & "," & SQLText(SAMDI)
        SQL = SQL & "," & SQLText(DIMANCHE)
        SQL = SQL & ")"
        CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub Delete_TmpPLANNING(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CNB As ADODB.Connection)

    Const sPROC_NAME As String = OBJ_NAME & ".Save_TmpPLANNING"
    Dim SQL As String
    
    On Error GoTo ErrHandler

        SQL = "DELETE FROM TMP_PLANNING"
        CNB.Execute SQL
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Function GetCode_PLANNINGByJourTourneeDate(ByRef ErrNumber As Long, _
                                                    ByRef ErrDescription As String, _
                                                    ByRef ErrSourceDetail As String, _
                                                    ByVal DateDu As String, _
                                                    ByVal jour As String, _
                                                    ByVal Tournee As String, _
                                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT CODE FROM PLANNING"
    SQL = SQL & " WHERE CONVERT(VARCHAR(10),DATEDU,103)=" & SQLText(DateDu) & " AND"
    SQL = SQL & " TOURNEE =" & SQLText(Tournee) & " AND"
    SQL = SQL & " Supp ='N' AND"
    SQL = SQL & " JOUR = " & SQLText(jour)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetCode_PLANNINGByJourTourneeDate = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Sub Delete_RestaurerPLANNING(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal CodeProg As String, _
                                ByVal vSupp As String, _
                                ByVal CodeUser As String, _
                                ByVal DATESupp As Date, _
                                ByVal CNB As ADODB.Connection)
    
    Dim SQL As String
    
On Error GoTo ErrHandler
    
    If vSupp <> "N" And vSupp <> "O" Then
        MsgBox "variable supp invalide", vbExclamation
        Exit Sub
    End If
    
    SQL = "Update PLANNING set "
    SQL = SQL & " Supp=" & SQLText(vSupp)
    If vSupp = "O" Then
        SQL = SQL & " , UserDelete=" & SQLText(CodeUser)
        SQL = SQL & " , Datedelete=" & SQLText(DATESupp)
    ElseIf vSupp = "N" Then
        SQL = SQL & " , UserAdd=" & SQLText(CodeUser)
    End If
    SQL = SQL & " where code =" & SQLText(CodeProg)
    CNB.Execute SQL
    
    
Exit Sub
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub
Public Function Get_CountPLNG(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As String, _
              ByVal vDateF As String, _
              ByVal Dest As String, _
              ByVal Cond As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_CountPLNG"
    Dim SQL As String
    Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
    
    SQL = "SELECT DISTINCT COUNT(*) AS NBRPLNG , "
    SQL = SQL & " C.LIBELLE AS CONDUCTEUR,"
    SQL = SQL & " T.LIBELLE AS TOURNEE "
    SQL = SQL & " FROM PLANNING P"
    SQL = SQL & " INNER JOIN PERSONNEL C ON C.CODE=P.CONDUCTEUR"
    SQL = SQL & " INNER JOIN DESTINATION T ON T.NUMERO=P.TOURNEE"
    SQL = SQL & " WHERE P.DATEJOUR IS NOT NULL AND P.SUPP='N' AND P.DATEDU BETWEEN " & SQLText(VdateD) & " AND " & SQLText(vDateF)
    If Dest <> "0000" And Dest <> "" Then SQL = SQL & " AND T.LIBELLE = " & SQLText(Dest)
    If Cond <> "0000" And Cond <> "" Then SQL = SQL & " AND C.LIBELLE = " & SQLText(Cond)
    SQL = SQL & " GROUP BY C.LIBELLE ,T.LIBELLE"
    CNB.Execute SQL
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_CountPLNG = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function Get_CondPLNG(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal semPlng As Date, _
              ByVal jour As String, _
              ByVal Cond As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_CondPLNG"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    
    SQL = "SELECT C.LIBELLE AS CONDUCTEUR,P.DATEDU, P.JOUR,"
    SQL = SQL & " T.LIBELLE AS TOURNEE "
    SQL = SQL & " FROM PLANNING P"
    SQL = SQL & " INNER JOIN PERSONNEL C ON C.CODE=P.CONDUCTEUR"
    SQL = SQL & " INNER JOIN DESTINATION T ON T.NUMERO=P.TOURNEE WHERE P.Supp='N' AND "
    SQL = SQL & "  P.DATEDU = " & SQLText(semPlng) & " And P.JOUR = " & SQLText(jour)
    SQL = SQL & " And P.CONDUCTEUR = " & SQLText(Cond)
  
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_CondPLNG = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function Get_DetailPLNG(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As String, _
              ByVal vDateF As String, _
              ByVal Dest As String, _
              ByVal Cond As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_DetailPLNG"
    Dim SQL As String
    Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
    
    SQL = "SELECT P.CODE AS NUMERO,C.LIBELLE AS CONDUCTEUR,P.DATEDU, P.JOUR,P.DATEJOUR,"
    SQL = SQL & " T.LIBELLE AS TOURNEE "
    SQL = SQL & " FROM PLANNING P"
    SQL = SQL & " INNER JOIN PERSONNEL C ON C.CODE=P.CONDUCTEUR"
    SQL = SQL & " INNER JOIN DESTINATION T ON T.NUMERO=P.TOURNEE"
    SQL = SQL & " WHERE P.DATEJOUR IS NOT NULL AND P.SUPP='N' AND P.DATEDU BETWEEN " & SQLText(VdateD) & " AND " & SQLText(vDateF)
    If Dest <> "0000" And Dest <> "" Then SQL = SQL & " AND P.TOURNEE = " & SQLText(Dest)
    If Cond <> "0000" And Cond <> "" Then SQL = SQL & " AND P.CONDUCTEUR = " & SQLText(Cond)
    SQL = SQL & "ORDER BY P.DATEJOUR , C.LIBELLE"

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_DetailPLNG = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


Public Function Get_DetailPLNG_TG(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As String, _
              ByVal vDateF As String, _
              ByVal Dest As String, _
              ByVal Cond As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_DetailPLNG_TG"
    Dim SQL As String
    Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
    
    SQL = "SELECT P.CODE AS NUMERO, C.LIBELLE AS CONDUCTEUR, V.MATRICULE AS VEHICULE, P.DATEDU, P.JOUR,P.DATEJOUR,"
    SQL = SQL & " T.LIBELLE AS TOURNEE"
    SQL = SQL & " FROM PLANNING P"
    SQL = SQL & " INNER JOIN PERSONNEL C ON C.CODE=P.CONDUCTEUR"
    SQL = SQL & " INNER JOIN VEHICULE V ON V.CODE=P.VEHICULE"
    SQL = SQL & " INNER JOIN DESTINATION T ON T.NUMERO=P.TOURNEE"
    SQL = SQL & " WHERE P.DATEJOUR IS NOT NULL AND P.SUPP='N' AND P.DATEDU BETWEEN " & SQLText(VdateD) & " AND " & SQLText(vDateF)
    SQL = SQL & " AND (P.JOUR='SAMEDI' OR P.JOUR ='DIMANCHE')"
    If Dest <> "0000" And Dest <> "" Then SQL = SQL & " AND P.TOURNEE = " & SQLText(Dest)
    If Cond <> "0000" And Cond <> "" Then SQL = SQL & " AND P.CONDUCTEUR = " & SQLText(Cond)
    SQL = SQL & "ORDER BY P.DATEJOUR, P.JOUR, C.LIBELLE "

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_DetailPLNG_TG = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

