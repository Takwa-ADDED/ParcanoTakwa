VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Imprimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Private Const OBJ_NAME As String = "Imprimer"

Public Function PrintProgChauf(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal CNB As ADODB.Connection, _
                                ByVal vcode As String) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".PrintProgChauf"
    Dim SQL As String
    Dim LRs_Record As Command

    'Query SQL "Select INNER JOIN"***
    SQL = "select A.Code as CodeProg, A.DateCreation, A.DateProgramme, D.TxtCommande, D.TxtPaiement,D.TxtObservation, D.ProgOrder, S.Libelle as Fournisseur, P.Libelle as Conducteur, V.Matricule, U.NomPrn"
    SQL = SQL & " From Ass_ProgChauf A"
    SQL = SQL & " inner join Det_ProgChauf D on D.CodeProgChauf=A.Code"
    SQL = SQL & " inner join Station S on D.codefournisseur=S.code"
    SQL = SQL & " Inner join Personnel P on A.CodeConducteur = P.Code"
    SQL = SQL & " Inner join Vehicule V on A.codeVehicule=V.Code"
    SQL = SQL & " Inner Join Utilisateur U on A.UserInsert=U.Code"
    SQL = SQL & " where A.Supp='N' And A.Code=" & SQLText(vcode) & " Order By D.ProgOrder"

    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set PrintProgChauf = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Function
Public Function PrintCompteur(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal CNB As ADODB.Connection) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".PrintCompteur"
    Dim SQL As String
    Dim LRs_Record As Command

    SQL = "SELECT Matricule, CompteurFT, CompteurCarburant, CompteurVidange from Vehicule where actif = 1 order by Matricule"

    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set PrintCompteur = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Function
Public Function PrintAnomalieTrafic(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal Name_Table As String, _
                                ByVal DateDebut As Date, _
                                ByVal DateFin As Date, _
                                ByVal Conducteur As String, _
                                ByVal Vehicule As String, _
                                ByVal Distination As String, _
                                ByVal anomali As Boolean, _
                                ByVal CNB As ADODB.Connection) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".PrintAnomalieTrafic"
    Dim SQL As String
    Dim LRs_Record As Command

    'Query SQL "Select INNER JOIN"***
    SQL = "SELECT"
        SQL = SQL & " FT.operateurEntre, FT.OperateurSortie,"
        SQL = SQL & " CONVERT(VARCHAR(8),FT.HeureEntre,108) AS HeureEntre,"
        SQL = SQL & " CONVERT(VARCHAR(8),FT.HeureSortie,108) AS HeureSortie,"
        SQL = SQL & " Cn.Libelle AS LibelleCond,"
        SQL = SQL & " Vc.Matricule AS MatriculeVehic,"
        SQL = SQL & " Ds.Libelle AS LibelleDest,Ds.MaxCompteur,"
        SQL = SQL & " convert(varchar(8),Ds.MaxDuree,108) as MaxDuree,"
        SQL = SQL & " ((FT.compteurentre)-(FT.compteursortie))as Kmt,"
        SQL = SQL & " convert(VarChar(8), (FT.HeureEntre) - (FT.heureSortie), 108) As duree,"
        SQL = SQL & " CONVERT(VARCHAR,FT.Heuresortie,103) as DateSortie"
    SQL = SQL & " From"
        SQL = SQL & " " & Name_Table & " FT"
        SQL = SQL & " Inner Join Personnel Cn ON FT.Conducteur = Cn.Code"
        SQL = SQL & " INNER JOIN Vehicule Vc ON FT.Vehicule = Vc.Code"
        SQL = SQL & " INNER JOIN Destination DS ON FT.Destination = DS.Numero"
    SQL = SQL & " Where"
        SQL = SQL & " ((CONVERT(VARCHAR,FT.Heureentre,103) >= (cast(" & SQLText(DateDebut) & " as datetime))) AND (CONVERT(VARCHAR,FT.HeureEntre,103) <= (cast(" & SQLText(DateFin) & " as datetime)))) AND "
        SQL = SQL & " (FT.Supp= 'N') AND"
        SQL = SQL & " (((FT.HeureEntre) IS NOT NULL) AND ((FT.OperateurEntre) IS NOT NULL)) AND"
        SQL = SQL & " (((FT.conducteur) IS NOT NULL) AND ((FT.vehicule) IS NOT NULL) AND ((FT.Destination) IS NOT NULL)) AND"
        SQL = SQL & " (((Ds.MaxDuree) IS NOT NULL) AND ((Ds.MaxCompteur) IS NOT NULL)) "
        If Conducteur <> "0000" And Conducteur <> "  -  Tous" Then SQL = SQL & " AND (FT.Conducteur = " & SQLText(Conducteur) & ") "
        If Vehicule <> "0000" And Vehicule <> "  -  Tous" Then SQL = SQL & " AND (FT.Vehicule = " & SQLText(Vehicule) & ") "
        If Distination <> "0000" And Distination <> "  -  Tous" Then SQL = SQL & "AND (FT.Destination = " & SQLText(Distination) & ") "
        If anomali = True Then
        SQL = SQL & " And ((((CAST(FT.CompteurEntre AS INT)) - (CAST(FT.CompteurSortie AS INT))) >= (CAST(Ds.MaxCompteur AS INT))) OR"
        SQL = SQL & " (((FT.HeureEntre)-(FT.HeureSortie)) >= (CAST(Ds.MaxDuree AS DATETIME))))"
        End If
    SQL = SQL & " Order By"
        SQL = SQL & " FT.HeureEntre , FT.heureSortie"
    

    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set PrintAnomalieTrafic = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Function
Public Function PrintPLANNING(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal CNB As ADODB.Connection, _
                                ByVal DateDu As Date) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".PrintPLANNING"
    Dim SQL As String
    Dim LRs_Record As Command

    SQL = "SELECT * FROM TMP_PLANNING WHERE TOURNEE <> 'REPOS' AND CONVERT(VARCHAR(10),DateDu,103) = (cast(" & SQLText(DateDu) & " as datetime)) order by code"

    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set PrintPLANNING = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Function
Public Function PrintREPOS(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal CNB As ADODB.Connection, _
                                ByVal DateDu As Date) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".PrintREPOS"
    Dim SQL As String
    Dim LRs_Record As Command

    SQL = "SELECT * FROM TMP_PLANNING WHERE TOURNEE = 'REPOS' AND CONVERT(VARCHAR(10),DateDu,103) = (cast(" & SQLText(DateDu) & " as datetime)) order by code"

    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set PrintREPOS = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Function

Public Function Print_StatServices(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal TableTrafic As String, _
                                    ByVal DateDebut As Date, _
                                    ByVal DateFin As Date, _
                                    ByVal Conducteur As String, _
                                    ByVal CNB As ADODB.Connection) As Command
    
    Const sPROC_NAME As String = OBJ_NAME & ".Print_StatServices"
    Dim SQL As String
    Dim LRs_Record As Command
    
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

    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set Print_StatServices = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

