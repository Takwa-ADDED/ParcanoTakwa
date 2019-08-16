VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Traffic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
    Private Const OBJ_NAME As String = "Traffic"
Public Function GETALL_SUPERVISIONTRAFFICBYDATE(ByRef ErrNumber As Long, _
                                            ByRef ErrDescription As String, _
                                            ByRef ErrSourceDetail As String, _
                                            ByVal TableTrafic As String, _
                                            ByVal DateDebut As Date, _
                                            ByVal DateFin As Date, _
                                            ByVal Conducteur As String, _
                                            ByVal Vehicule As String, _
                                            ByVal Distination As String, _
                                            ByVal VType As String, _
                                            ByVal YearTrafic As Integer, _
                                            ByVal CNB As ADODB.Connection) As Recordset
    
    Const sPROC_NAME As String = OBJ_NAME & ".GETALL_SUPERVISIONTRAFFICBYDATE"
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT "
        SQL = SQL & " FT.Numero, FT.CompteurSortie, FT.CompteurEntre , FT.OperateurEntre, FT.OperateurSortie, "
        SQL = SQL & " CONVERT(VARCHAR(8),FT.HeureEntre,108) AS HeureEntre,"
        SQL = SQL & " CONVERT(VARCHAR(8),FT.HeureSortie,108) AS HeureSortie,"
        SQL = SQL & " Cn.Libelle AS LibelleCond,"
        SQL = SQL & " Vc.Matricule AS MatriculeVehic,"
        SQL = SQL & " Ds.Libelle AS LibelleDest, Ds.MaxCompteur,"
        SQL = SQL & " CONVERT(VARCHAR(8),Ds.MaxDuree,108) AS MaxDuree,"
        SQL = SQL & " ((FT.CompteurEntre)-(FT.CompteurSortie))AS Kmt,"
        SQL = SQL & " CONVERT(VARCHAR(8), (FT.heureentre) - (FT.heuresortie), 108) As Duree,"
        SQL = SQL & " CONVERT(VARCHAR,FT.Heuresortie,103) as DateSortie, "
        SQL = SQL & " DATEDIFF (S ,FT.heuresortie,FT.heureentre) as NSecond, "
        SQL = SQL & " (((FT.CompteurEntre)-(FT.CompteurSortie))-Ds.MaxCompteur)as Difk, "
        SQL = SQL & " CONVERT(VarChar(8), (((FT.heureentre) - (FT.heuresortie)) - Ds.MaxDuree), 108) As DifD, "
        SQL = SQL & " CONVERT(VarChar(8), (Ds.MaxDuree - ((FT.heureentre) - (FT.heuresortie))), 108) As DifDm"
    SQL = SQL & " FROM "
        SQL = SQL & " " & TableTrafic & " FT"
        SQL = SQL & " Inner Join Personnel Cn ON FT.Conducteur = Cn.Code"
        SQL = SQL & " INNER JOIN Vehicule Vc ON FT.Vehicule = Vc.Code"
        SQL = SQL & " INNER JOIN Destination DS ON FT.Destination = DS.Numero"
    SQL = SQL & " WHERE "
        SQL = SQL & "(( ((CONVERT(VARCHAR,FT.Heureentre,103) >= (cast(" & SQLText(DateDebut) & " as datetime))) AND (CONVERT(VARCHAR,FT.HeureEntre,103) <= (cast(" & SQLText(DateFin) & " as datetime)))) )  or FT.Heureentre is null ) "
        SQL = SQL & " AND (FT.Supp= 'N') "
'        SQL = SQL & " AND (((FT.HeureEntre) IS NOT NULL) AND ((FT.OperateurEntre) IS NOT NULL)) "
        SQL = SQL & " AND (((FT.conducteur) IS NOT NULL) AND ((FT.vehicule) IS NOT NULL) AND ((FT.Destination) IS NOT NULL)) "
'        SQL = SQL & " AND (((Ds.MaxDuree) IS NOT NULL) AND ((Ds.MaxCompteur) IS NOT NULL)) "
        If Conducteur <> "0000" And Conducteur <> "Tous" Then SQL = SQL & " AND (FT.Conducteur = " & SQLText(Conducteur) & ") "
        If Vehicule <> "0000" And Vehicule <> "Tous" Then SQL = SQL & " AND (FT.Vehicule = " & SQLText(Vehicule) & ") "
        If Distination <> "0000" And Distination <> "Tous" Then SQL = SQL & " AND (FT.Destination = " & SQLText(Distination) & ") "
        If VType = "Anomalie" Then
            SQL = SQL & " And ((((CAST(FT.CompteurEntre AS INT)) - (CAST(FT.CompteurSortie AS INT))) >= (CAST(Ds.MaxCompteur AS INT))) OR "
            SQL = SQL & " (((FT.HeureEntre)-(FT.HeureSortie)) >= (CAST(Ds.MaxDuree AS DATETIME))))"
        End If
    SQL = SQL & " ORDER BY "
        SQL = SQL & "cast(CONVERT(VARCHAR,FT.Heuresortie,103) As datetime), FT.heureSortie "
        
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GETALL_SUPERVISIONTRAFFICBYDATE = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GETALL_STATISTIQUESSERVICES(ByRef ErrNumber As Long, _
                                                    ByRef ErrDescription As String, _
                                                    ByRef ErrSourceDetail As String, _
                                                    ByVal TableTrafic As String, _
                                                    ByVal DateDebut As Date, _
                                                    ByVal DateFin As Date, _
                                                    ByVal Conducteur As String, _
                                                    ByVal CNB As ADODB.Connection) As Recordset
    
    Const sPROC_NAME As String = OBJ_NAME & ".GETALL_STATISTIQUESSERVICES"
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT "
        SQL = SQL & "FT.NUMERO AS NTRAFIC,"
        SQL = SQL & "CN.LIBELLE AS CONDUCTEUR, "
        SQL = SQL & "VH.MATRICULE AS VEHICULE, "
        SQL = SQL & "DS.LIBELLE AS DESTINATION, "
        SQL = SQL & "CONVERT(VARCHAR(8),FT.HEURESORTIE,108) AS HEURESORTIE, "
        SQL = SQL & "CONVERT(VARCHAR(8),FT.HEUREENTRE,108) AS HEUREENTRE, "
        SQL = SQL & "DP.NUMERO AS NDISP, "
        SQL = SQL & "DP.ETAT, "
        SQL = SQL & "CONVERT(VARCHAR(10),DP.HDEBUT,103) AS DATEDEBUT, "
        SQL = SQL & "CONVERT(VARCHAR(8),DP.HDEBUT,108) AS HDEBUT, "
        SQL = SQL & "CONVERT(VARCHAR(8),DP.HFIN,108) AS HFIN, "
        SQL = SQL & "CONVERT(VARCHAR(8), (DP.HFIN) - (DP.HDEBUT), 108) As DUREE, "
        SQL = SQL & "CONVERT(VARCHAR(8), (FT.HEUREENTRE) - (FT.HEURESORTIE), 108) As DUREETRAFIC "
    SQL = SQL & "FROM " & TableTrafic & " FT "
    SQL = SQL & "INNER JOIN PERSONNEL CN ON FT.CONDUCTEUR=CN.CODE "
    SQL = SQL & "INNER JOIN VEHICULE VH ON FT.VEHICULE=VH.CODE "
    SQL = SQL & "INNER JOIN DESTINATION DS ON FT.DESTINATION=DS.NUMERO "
    SQL = SQL & "INNER JOIN DISPOPERSO DP ON CN.CODE=DP.CONDUCTEUR "
    SQL = SQL & "Where "
        SQL = SQL & "(CN.LIBELLE=" & SQLText(Conducteur) & ") AND "
        SQL = SQL & "(DP.HDEBUT BETWEEN " & SQLText(DateDebut) & " AND " & SQLText(DateFin) & ") AND "
        SQL = SQL & "(DP.HFIN IS NOT NULL) AND "
        SQL = SQL & "(FT.HEUREENTRE BETWEEN  DP.HDEBUT AND DP.HFIN) AND"
        SQL = SQL & "(FT.HEURESORTIE BETWEEN  DP.HDEBUT AND DP.HFIN)"
    SQL = SQL & "ORDER BY "
        SQL = SQL & "CN.LIBELLE , DP.HDEBUT "
        
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GETALL_STATISTIQUESSERVICES = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GET_COUNTVEHICULE(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CNB As ADODB.Connection, _
                                    ByVal VEHIC As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GET_COUNTVEHICULE"
    Dim SQL As String, LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT COUNT(*) AS COUNT FROM FICHETRAFFIC WHERE VEHICULE=" & SQLText(VEHIC)
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GET_COUNTVEHICULE = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GET_MAXCOMPTRAFFIC(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CNB As ADODB.Connection, _
                                    ByVal VEHIC As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GET_MAXCOMPTRAFFIC"
    Dim SQL As String, LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT COMPTEURENTRE FROM FICHETRAFFIC WHERE NUMERO = (SELECT MAX(CAST(NUMERO AS INT)) AS NUM FROM FICHETRAFFIC WHERE VEHICULE=" & SQLText(VEHIC) & ")"
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GET_MAXCOMPTRAFFIC = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function



Public Function GetRow_Traffic_ByCode(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal CodeTraffic As String, _
                                        ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "select F.*, C.Libelle as LibCond, V.Matricule as MatVehi, D.Libelle as LibDest "
    SQL = SQL & " from Fichetraffic F"
    SQL = SQL & " INNER JOIN Personnel C ON F.CONDUCTEUR=C.CODE"
    SQL = SQL & " INNER JOIN VEHICULE V ON F.VEHICULE=V.CODE"
    SQL = SQL & " INNER JOIN DESTINATION D ON F.DESTINATION=D.NUMERO"
    SQL = SQL & " Where"
    SQL = SQL & " F.numero = " & SQLText(CodeTraffic)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_Traffic_ByCode = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetAll_TrafficVehiculeExterieur(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal Name_Table As String, _
                                        ByVal DateSys As Date, _
                                        ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT"
    SQL = SQL & " FT.*, CN.CODE AS CODECOND, CN.ABREVIATION AS ABREVCOND, CN.LIBELLE AS LIBELLECOND,"
    SQL = SQL & " VC.CODE AS CODEVEHIC, VC.ABREVIATION AS ABREVVEH , VC.MATRICULE AS MATRICULEVEHIC,"
    SQL = SQL & " DS.NUMERO AS CODEDEST, DS.LIBELLE AS LIBELLEDEST,DS.MAXCOMPTEUR, DS.MAXDUREE "
    SQL = SQL & " FROM " & Name_Table & " FT"
    SQL = SQL & " INNER JOIN PERSONNEL CN ON FT.CONDUCTEUR = CN.CODE"
    SQL = SQL & " INNER JOIN VEHICULE VC ON FT.VEHICULE = VC.CODE"
    SQL = SQL & " INNER JOIN DESTINATION DS ON FT.DESTINATION = DS.NUMERO"
'    SQL = SQL & " WHERE (CONVERT(VARCHAR,FT.HEURESORTIE,103)) <=" & SQLTEXT(DATESYS) & " AND FT.HEUREENTRE IS NULL"
    SQL = SQL & " WHERE CAST((CONVERT(VARCHAR,FT.HEURESORTIE,103)) AS DATETIME) <=CAST(" & SQLText(DateSys) & "  AS DATETIME) AND FT.HEUREENTRE IS NULL"
    SQL = SQL & " ORDER BY FT.HEUREENTRE ,FT.HEURESORTIE"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetAll_TrafficVehiculeExterieur = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetAll_TrafficByDateSys(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal Name_Table As String, _
                                        ByVal DateSys As Date, _
                                        ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler
    
    SQL = "select FT.*, Cn.Code as codeCond, Cn.ABREVIATION as abrevCond, Cn.Libelle as LibelleCond, Vc.Code as CodeVehic, "
    SQL = SQL & " Vc.ABREVIATION as abrevVeh , Vc.Matricule as MatriculeVehic, Ds.Numero as CodeDest, Ds.Libelle as LibelleDest, DS.MaxCompteur , DS.MaxDuree,DS.MinCompteur"
    SQL = SQL & " from " & Name_Table & " FT"
    SQL = SQL & " Inner Join Personnel Cn on FT.Conducteur = Cn.Code"
    SQL = SQL & " INNER JOIN Vehicule Vc ON FT.Vehicule = Vc.Code"
    SQL = SQL & " INNER JOIN Destination DS ON FT.Destination = DS.Numero"
    SQL = SQL & " where CONVERT(VARCHAR,FT.HeureEntre,103)= " & SQLText(DateSys) & " Order by FT.HeureEntre, FT.heureSortie"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetAll_TrafficByDateSys = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetTraffic_ByMarticuleVehicule(ByRef ErrNumber As Long, _
                                                ByRef ErrDescription As String, _
                                                ByRef ErrSourceDetail As String, _
                                                ByVal Name_Table As String, _
                                                ByVal Matricule As String, _
                                                ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from " & Name_Table & " where Vehicule = " & SQLText(Matricule)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetTraffic_ByMarticuleVehicule = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetMAx_NumeroTraffic(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "select Max(Numero) from fichetraffic "

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetMAx_NumeroTraffic = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Sub UpDate_Traffic(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CNB As ADODB.Connection, _
                            ByVal PRs_Record As ADODB.Recordset, _
                            ByVal NumeroFich As String, _
                            ByVal SelectedCE As Long, _
                            ByVal Selectedd As String)

    Const sPROC_NAME As String = OBJ_NAME & ".UpDate_Traffic"

    Dim SQL As String
    
    On Error GoTo ErrHandler

    SQL = "Update fichetraffic Set "
    SQL = SQL & " Vehicule = " & SQLText(PRs_Record!Vehicule)
    SQL = SQL & ", CompteurSortie = " & PRs_Record!CompteurSortie
    SQL = SQL & ", Conducteur = " & SQLText(PRs_Record!Conducteur)
    SQL = SQL & ", Destination = " & SQLText(PRs_Record!Destination)
    SQL = SQL & ",  HeureSortie = " & SQLText(PRs_Record!HEURESORTIE)
    SQL = SQL & ", ObservationEntre=" & SQLText(PRs_Record!Observation)
    SQL = SQL & ", userupdate=" & SQLText(PRs_Record!userupdate)
    If Not (IsNull(SelectedCE)) And (SelectedCE > 0) Then
        SQL = SQL & ",  CompteurEntre = " & PRs_Record!CompteurEntre
        SQL = SQL & ",  HeureEntre = " & SQLText(PRs_Record!HeureEntre)
    End If
    If Selectedd = "REPARATION" Then
        SQL = SQL & ",  OperateurEntre = " & SQLText(PRs_Record!OperateurEntre)
    End If
    SQL = SQL & " where fichetraffic.Numero = " & SQLText(NumeroFich)
     CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub UpDate_Traffic_VE(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CNB As ADODB.Connection, _
                            ByVal PRs_Record As ADODB.Recordset, _
                            ByVal NumeroTxt As String)

    Const sPROC_NAME As String = OBJ_NAME & ".UpDate_Traffic_VE"

    Dim SQL As String
    
    On Error GoTo ErrHandler

    SQL = "Update fichetraffic Set "
    SQL = SQL & " HeureEntre = " & SQLText(PRs_Record!HeureEntre)
    SQL = SQL & " , CompteurEntre = " & SQLText(PRs_Record!CompteurEntre)
    SQL = SQL & " , OperateurEntre = " & SQLText(PRs_Record!OperateurEntre)
    SQL = SQL & " , ObservationEntre = " & SQLText(PRs_Record!Observation)
    SQL = SQL & " where fichetraffic.Numero = " & NumeroTxt
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub Save_Traffic(ByRef ErrNumber As Long, _
                        ByRef ErrDescription As String, _
                        ByRef ErrSourceDetail As String, _
                        ByVal CNB As ADODB.Connection, _
                        ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Save_Traffic"

    Dim SQL As String
    
    On Error GoTo ErrHandler

    SQL = "Insert into ficheTraffic (Numero,Vehicule,CompteurSortie,Conducteur,Destination,HeureSortie, OperateurSortie,UserInsert) values ("
    SQL = SQL & SQLText(PRs_Record!numero)
    SQL = SQL & "," & SQLText(PRs_Record!Vehicule)
    SQL = SQL & "," & SQLText(PRs_Record!CompteurSortie)
    SQL = SQL & "," & SQLText(PRs_Record!Conducteur)
    SQL = SQL & "," & SQLText(PRs_Record!Destination)
    SQL = SQL & "," & SQLText(PRs_Record!HEURESORTIE)
    SQL = SQL & "," & SQLText(PRs_Record!OperateurSortie)
    SQL = SQL & " ," & SQLText(PRs_Record!USERINSERT)
    SQL = SQL & ")"
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub Delete_Add_Traffic(ByRef ErrNumber As Long, _
                        ByRef ErrDescription As String, _
                        ByRef ErrSourceDetail As String, _
                        ByVal code As String, _
                        ByVal vSupp As String, _
                        ByVal CodeUser As String, _
                        ByVal CNB As ADODB.Connection)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Add_Traffic"

    Dim SQL As String
    
    If vSupp <> "N" And vSupp <> "O" Then
        MsgBox "variable supp invalide", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    SQL = "Update Fichetraffic set "
    SQL = SQL & " Supp = " & SQLText(vSupp)
    If vSupp = "O" Then
        SQL = SQL & " , UserDelete = " & SQLText(CodeUser)
    ElseIf vSupp = "N" Then
        SQL = SQL & " , UserAdd = " & SQLText(CodeUser)
    End If
    SQL = SQL & " where Numero = " & SQLText(code)
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub










