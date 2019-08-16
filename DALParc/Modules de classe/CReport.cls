VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "CReport"

'===================== Bon Carburant ====================
'========================================================
Public Function PrintOutBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal numero As String) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".PrintOutBC"

    Dim StrSql As String
    Dim LRs_Record As Command

    On Error GoTo ErrHandler
    
    StrSql = "SELECT A.Numero,A.Heure, A.DateDoc, A.Station, A.Conducteur, "
    StrSql = StrSql & "  A.Valeur,A.dateop, D.Vehicule, D.Litre AS DetLitre,prixLitre as prix,tva,"
    StrSql = StrSql & " D.CompteurCarburant As Compteur, S.Libelle AS NStation, V.Matricule,V.marque, V.TYPE, V.Energie,"
    StrSql = StrSql & " P.Libelle AS PERS, S.Adresse, S.Ville, S.TELEPHONE, S.FAX, A.NBC , D.AnomalieConsom "
    StrSql = StrSql & " FROM Station S INNER JOIN AssBoncarburant A INNER JOIN"
    StrSql = StrSql & " DetBonCarburant D ON A.Numero = D.Numero INNER JOIN"
    StrSql = StrSql & " Vehicule V ON D.Vehicule = V.Code INNER JOIN"
    StrSql = StrSql & " Personnel P ON A.Conducteur = P.Code ON S.Code = A.Station"
    StrSql = StrSql & " where A.numero= " & SQLText(numero)

    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = StrSql
    LRs_Record.CommandType = adCmdText
    
    Set PrintOutBC = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function PrintOutBC2(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal numero As String) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".PrintOutBC2"

    Dim SQL As String
    Dim LRs_Record As Command

    On Error GoTo ErrHandler
    
    SQL = "Select SUM(Litre) As Volume, SUM(Litre * PrixLitre) As Valeur, Energie"
    SQL = SQL & " from DetBoncarburant Where  numero = " & SQLText(numero)
    SQL = SQL & " Group By Energie"
    
    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    
    Set PrintOutBC2 = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'===================== Bon vidange =====================
'========================================================
Public Function PrintOutBV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal numero As String) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".PrintOutBV"

    Dim SQL As String
    Dim LRs_Record As Command

    On Error GoTo ErrHandler
    
    SQL = "SELECT Ass_BonVidange.Numero, Ass_BonVidange.DateDoc,Ass_BonVidange.heure,Ass_BonVidange.DateOp,Ass_BonVidange.valeur,"
    SQL = SQL & "Ass_BonVidange.NBKLMvid AS NBKLM, Vehicule.Matricule, Vehicule.Energie, Personnel.Libelle AS Pers,"
    SQL = SQL & " Station.Libelle AS Nstation, Ass_BonVidange.NBC, Det_BonVidange.Libelle AS NLubrifiant,"
    SQL = SQL & " Det_BonVidange.Libelle AS NProduit, Det_BonVidange.THT as prix,Det_BonVidange.tva,Det_BonVidange.prixTTC as prixttc , Vehicule.CompteurFT As compteur,"
    SQL = SQL & " Ass_BonVidange.CompteurVidange as dercompt,Vehicule.marque"
    SQL = SQL & " From Ass_BonVidange"
    SQL = SQL & " Inner Join Vehicule"
    SQL = SQL & " On vehicule.Code=Ass_BonVidange.Vehicule"
    SQL = SQL & " INNER JOIN Personnel"
    SQL = SQL & " ON Personnel.Code =Ass_BonVidange.Conducteur"
    SQL = SQL & " INNER JOIN Det_BonVidange"
    SQL = SQL & " on  Det_BonVidange.Numero =Ass_BonVidange.Numero"
    SQL = SQL & " INNER JOIN Station"
    SQL = SQL & " ON Station.Code=Ass_BonVidange.Station"
    SQL = SQL & " Where Ass_BonVidange.Numero = " & SQLText(numero)
    
    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    
    Set PrintOutBV = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'============== Bon Commande Réception ====================
'==========================================================
Public Function PrintOutBRepar(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal numero As String) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".PrintOutBRepar"

    Dim StrSql As String
    Dim LRs_Record As Command

    On Error GoTo ErrHandler
    
    StrSql = "SELECT A.Numero As Field1, A.DateCreation As dateCreation,S.Libelle As Fournisseur,"
    StrSql = StrSql & "D.Désignation As Designation , ROUND(D.Qté,0) As Qte ,D.Observation As Observation "
    StrSql = StrSql & "FROM  AssBCReparation A Inner Join DetailBCReparation D "
    StrSql = StrSql & "ON A.Numero = D.Numero "
    StrSql = StrSql & " INNER JOIN Station S ON A.fournisseur = S.Code "
    StrSql = StrSql & "where A.Numero=" & SQLText(numero)
    
    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = StrSql
    LRs_Record.CommandType = adCmdText
    
    Set PrintOutBRepar = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'=================== Imprimer Piece Réparation ==============
'============================================================
Public Function PrintOutPieceRepar(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal numero As String, _
              ByVal MainOeuvre As Double) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".PrintOutPieceRepar"

    Dim StrSql As String
    Dim LRs_Record As Command

    On Error GoTo ErrHandler
    
    StrSql = "Select A.Numero As Number ,A.Type, DatePiece ,refPiece, dateOperation, "
    StrSql = StrSql & " S.Libelle As Fournisseur , remisePiece, Designation, Qte, PUHT, remise, tva"
    StrSql = StrSql & " from  AssPieceReparation A INNER JOIN DetailPieceReparation D ON"
    StrSql = StrSql & " A.Numero = D.Numero"
    StrSql = StrSql & " INNER JOIN Station S ON A.fournisseur = S.Code "
    StrSql = StrSql & " where A.Numero =" & SQLText(numero)
    If MainOeuvre <> 0 Then
        StrSql = StrSql & " Union All select distinct A.Numero As Number ,A.Type, DatePiece ,refPiece, dateOperation ,  S.Libelle As Fournisseur , remisePiece, 'Main doeuvre' as Designation,"
        StrSql = StrSql & " 1 as Qte, PrixMOeuvre as PUHT,remisePiece as  remise, TVA_MOeuvre as tva from  AssPieceReparation A "
        StrSql = StrSql & " INNER JOIN DetailPieceReparation D ON A.Numero = D.Numero INNER JOIN Station S ON A.fournisseur = S.Code "
        StrSql = StrSql & " Where A.Numero = " & SQLText(numero)
    End If
    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = StrSql
    LRs_Record.CommandType = adCmdText
    
    Set PrintOutPieceRepar = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
'==========================Imprimer Facture ==========================
'============================================================
Public Function PrintOutFact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal numero As String) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".PrintOutFact"

    Dim SQL As String
    Dim LRs_Record As Command

    On Error GoTo ErrHandler

'Carburant
    SQL = "select * from ("
    SQL = SQL & " Select AssBonCarburant.numfact,energie as designation,Sum(detBonCarburant.Litre) as Qte,prixht,tva"
    SQL = SQL & " From assBonCarburant, detBonCarburant"
    SQL = SQL & " where assBonCarburant.numero= detBonCarburant.numero and transf='O' and numfact= " & SQLText(numero)
    SQL = SQL & " GROUP BY assBonCarburant.numfact,energie,prixht,tva,prixlitre) as Q,FactureCarburant,station"
    SQL = SQL & " Where q.numfact = FactureCarburant.numero and station.code=FactureCarburant.station"
    'Vidange
    SQL = SQL & " Union All"
    SQL = SQL & " select * from ( Select Ass_BonVidange.numfact,Det_BonVidange.Libelle, Sum(1)as Qte, "
    SQL = SQL & " MIN(Det_BonVidange.tht) AS Prixht,MIN(Det_BonVidange.tva) AS tva"
    SQL = SQL & " From Det_BonVidange"
    SQL = SQL & " INNER JOIN Ass_BonVidange ON Det_BonVidange.Numero = Ass_BonVidange.Numero"
    SQL = SQL & " where transf='O' and numfact=" & SQLText(numero)
    SQL = SQL & " GROUP BY Ass_BonVidange.numfact,Det_BonVidange.Libelle ) as Q,FactureCarburant,station"
    SQL = SQL & " Where q.numfact = FactureCarburant.numero And station.code = FactureCarburant.station"
    'Rèception
    SQL = SQL & " Union All"
    SQL = SQL & " select * from ( Select AssPieceReparation.numfact,DetailPieceReparation.Designation, SUM(DetailPieceReparation.Qte) AS Qte,"
    SQL = SQL & " DetailPieceReparation.PUHT AS Prixht,DetailPieceReparation.TVA AS tva"
    SQL = SQL & " From DetailPieceReparation"
    SQL = SQL & " INNER JOIN AssPieceReparation ON DetailPieceReparation.Numero = AssPieceReparation.Numero"
    SQL = SQL & " where transf='O' and numfact=" & SQLText(numero)
    SQL = SQL & " GROUP BY AssPieceReparation.numfact,Designation,PUHT,tva ) as Q,FactureCarburant,station"
    SQL = SQL & " Where q.numfact = FactureCarburant.numero And station.code = FactureCarburant.station"
    'Moeuvre
    SQL = SQL & " Union All"
    SQL = SQL & " select * from ( Select numfact,'Main doeuvre' as Designation, 1 AS Qte, "
    SQL = SQL & " PrixMOeuvre AS Prixht,TVA_MOeuvre AS tva "
    SQL = SQL & " From AssPieceReparation "
    SQL = SQL & " where transf='O' and numfact=" & SQLText(numero)
    SQL = SQL & " GROUP BY numfact,PrixMOeuvre,TVA_MOeuvre ) as Q,FactureCarburant,station"
    SQL = SQL & " Where q.numfact = FactureCarburant.numero And station.code = FactureCarburant.station"


    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    
    Set PrintOutFact = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'========================== Congé ==========================
'===========================================================
Public Function PrintOut_CongeConduc(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal date_deb As Date, _
              ByVal date_fin As Date) As Command


    Const sPROC_NAME As String = OBJ_NAME & ".PrintOut_CongeConduc"

    Dim StrSql As String
    Dim LRs_Record As Command

    On Error GoTo ErrHandler

    StrSql = "select G.DateDu as DateDu ,G.DateAu as DateAu, G.Observation as Observation , G.Conducteur , P.Libelle as Libelle from GestionConge G INNER JOIN Personnel P On G.Conducteur=P.Code "
    StrSql = StrSql & "Where G.type <>'Repos' and convert(varchar(10),g.datedu,103) >= " & SQLText(date_deb) & " And convert(varchar(10),g.datedu,103) <=" & SQLText(date_fin)
'    If Conduc = "0000" Or Conduc = "" Then
        StrSql = StrSql & " Order By G.Conducteur"
'    Else
'        StrSql = StrSql & "And G.Conducteur = " & SQLText(Conduc)
'    End If
    

    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = StrSql
    LRs_Record.CommandType = adCmdText
    
    Set PrintOut_CongeConduc = LRs_Record
    Set LRs_Record = Nothing


    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'========================== Statistiques Carburant ==========================
'=============================================================================
Public Function Get_StatistBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As Date, _
              ByVal vDateF As Date) As Command

Const sPROC_NAME As String = OBJ_NAME & ".Get_StatistBC"

Dim SQL As String
Dim LRs_Record As Command

On Error GoTo ErrHandler

SQL = "select distinct X.Numero,X.AnomalieConsom , x.prixLitre As prix, x.dateDoc , x.matricule,"
SQL = SQL & " isnull(Y.CompteurCarburant,0) as AnCmpt, X.CompteurCarburant As Compteur, X.Litre ,(x.prixLitre * X.Litre) as montant, "
SQL = SQL & " isnull((X.CompteurCarburant - Y.CompteurCarburant),0) as KmParc ,"
SQL = SQL & " isnull((X.Litre*100/(X.CompteurCarburant - Y.CompteurCarburant)),0) as Consom "
SQL = SQL & " from("
SQL = SQL & " select rank() Over(Order by d.CompteurCarburant) as rank, D.Numero , Vehicule , d.CompteurCarburant , D.Litre ,a.dateDoc ,prixLitre ,AnomalieConsom ,"
SQL = SQL & " v.matricule  from  DetBonCarburant D , AssBonCarburant a , "
SQL = SQL & " Vehicule v where D.Numero=a.Numero and d.vehicule=v.code and a.Supp = 'N' and A.datedoc Between " & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00"))
SQL = SQL & " ) AS Y Right Join("
SQL = SQL & " select rank()  Over(Order by d.CompteurCarburant) as rank , d.Numero , Vehicule , d.CompteurCarburant , D.Litre , a.dateDoc ,prixLitre ,AnomalieConsom,"
SQL = SQL & " v.matricule from  DetBonCarburant D , AssBonCarburant a ,"
SQL = SQL & " Vehicule v where D.Numero=a.Numero and d.vehicule=v.code and a.Supp = 'N' and A.datedoc Between " & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00"))
SQL = SQL & " ) AS X ON Y.rank+1 = X.rank and y.vehicule = X.vehicule Order by x.matricule"


    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    
Set Get_StatistBC = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_StatistBCVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As Date, _
              ByVal vDateF As Date, _
            ByVal CodV As String) As Command

Const sPROC_NAME As String = OBJ_NAME & ".Get_StatistBCVeh"

Dim SQL As String
Dim LRs_Record As Command

On Error GoTo ErrHandler
SQL = "select distinct X.Numero,X.AnomalieConsom , x.prixLitre As prix, x.datedoc , x.matricule,"
SQL = SQL & " isnull(Y.CompteurCarburant,0) as AnCmpt, X.CompteurCarburant As Compteur, X.Litre ,(x.prixLitre * X.Litre) as montant,"
SQL = SQL & " isnull((X.CompteurCarburant - Y.CompteurCarburant),0) as KmParc ,"
SQL = SQL & " isnull((X.Litre*100/(X.CompteurCarburant - Y.CompteurCarburant)),0) as Consom"
SQL = SQL & "  from("
SQL = SQL & " select rank() Over(Order by d.CompteurCarburant) as rank, D.Numero , Vehicule , d.CompteurCarburant , D.Litre ,a.dateDoc ,prixLitre ,AnomalieConsom ,"
SQL = SQL & " v.matricule  from  DetBonCarburant D , AssBonCarburant a , "
SQL = SQL & " Vehicule v where D.Numero=a.Numero and d.vehicule=v.code and a.Supp = 'N' And v.matricule=" & SQLText(CodV) & " and A.datedoc Between " & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00"))
SQL = SQL & " ) AS Y Right Join("
SQL = SQL & " select rank()  Over(Order by d.CompteurCarburant) as rank , d.Numero , Vehicule , d.CompteurCarburant , D.Litre , a.dateDoc ,prixLitre ,AnomalieConsom,"
SQL = SQL & " v.matricule from  DetBonCarburant D , AssBonCarburant a ,"
SQL = SQL & " Vehicule v where D.Numero=a.Numero and d.vehicule=v.code And a.Supp = 'N' And v.matricule = " & SQLText(CodV) & " and A.datedoc Between " & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00"))
SQL = SQL & ") AS X ON Y.rank+1 = X.rank and y.vehicule = X.vehicule Order by x.matricule"


    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set Get_StatistBCVeh = LRs_Record
    Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Print_StatistBCTot(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As Date, _
              ByVal vDateF As Date) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".Print_StatistBCTot"

    Dim SQL As String
    Dim LRs_Record As New Command

    On Error GoTo ErrHandler

SQL = "Select Max(D.Numero) as Numero , SUM(D.Litre) As Litre, D.prixLitre As Prix,V.Matricule As Vehicule"
SQL = SQL & " ,(SUM(D.Litre) * D.prixLitre ) as TotalPx "
SQL = SQL & " ,(Max(D.CompteurCarburant) - Min(D.CompteurCarburant)) as KmPr From AssBonCarburant A , DetBonCarburant D , Vehicule V"
SQL = SQL & " Where A.Numero = D.Numero and D.Vehicule = V.Code"
SQL = SQL & " AND A.Supp = 'N' And A.datedoc Between" & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00"))
SQL = SQL & " group by V.Matricule , D.prixLitre"


    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set Print_StatistBCTot = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Print_StatistBCTotVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As Date, _
              ByVal vDateF As Date, _
              ByVal CodV As String) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".Print_StatistBCTotVeh"

    Dim SQL As String
    Dim LRs_Record As New Command

    On Error GoTo ErrHandler

SQL = "Select Max(D.Numero) as Numero , SUM(D.Litre) As Litre, D.prixLitre As Prix,V.Matricule As Vehicule "
SQL = SQL & " , (SUM(D.Litre) * D.prixLitre ) as TotalPx "
SQL = SQL & ",(Max(D.CompteurCarburant) - Min(D.CompteurCarburant)) as KmPr From AssBonCarburant A , DetBonCarburant D , Vehicule V"
SQL = SQL & " Where A.Numero = D.Numero and D.Vehicule = V.Code And v.matricule=" & SQLText(CodV)
SQL = SQL & " AND A.Supp = 'N' And A.datedoc Between" & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00"))
SQL = SQL & " group by V.Matricule , D.prixLitre"


    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set Print_StatistBCTotVeh = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'========================== Statistiques Réparation ==========================
'=============================================================================

Public Function Print_DetRepStatist(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As Date, _
              ByVal vDateF As Date) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".Print_DetRepStatist"

    Dim SQL As String
    Dim LRs_Record As New Command

    On Error GoTo ErrHandler

SQL = "Select D.* , A.DatePiece , (((((Qte * PUHT) - (Qte * PUHT * Remise / 100)) * tva / 100) + ((Qte * PUHT) - (Qte * PUHT * Remise / 100))) - (((((Qte * PUHT) - (Qte * PUHT * Remise / 100)) * tva / 100) + ((Qte * PUHT) - (Qte * PUHT * Remise / 100))) * A.RemisePiece/100)) as montant"
SQL = SQL & " from DetailPieceReparation D , AssPieceReparation A Where D.Numero = A.Numero "
SQL = SQL & " And A.Supp = 'N' And A.datePiece between " & SQLText(VdateD) & " and " & SQLText(vDateF)
SQL = SQL & " order by D.vehicule, A.datePiece"

    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set Print_DetRepStatist = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Print_RepStat(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As Date, _
              ByVal vDateF As Date) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".Print_RepStat"

    Dim SQL As String
    Dim LRs_Record As New Command

    On Error GoTo ErrHandler

    SQL = " Select distinct D.vehicule  ,count(*) As nbrRep ,  "
    SQL = SQL & " Sum (((((Qte * PUHT) - (Qte * PUHT * Remise / 100)) * tva / 100) + ((Qte * PUHT) - (Qte * PUHT * Remise / 100))) - (((((Qte * PUHT) - (Qte * PUHT * Remise / 100)) * tva / 100) + ((Qte * PUHT) - (Qte * PUHT * Remise / 100))) * A.RemisePiece/100)) As ttc "
    SQL = SQL & " from AsspieceReparation A, detailpieceReparation D where A.Numero = D.Numero And A.Supp = 'N' And A.datePiece"
    SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF)
    SQL = SQL & " group by D.Vehicule"


    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set Print_RepStat = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Print_RepStatVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal Matricule As String, _
              ByVal VdateD As Date, _
              ByVal vDateF As Date) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".Print_RepStatVeh"

    Dim SQL As String
    Dim LRs_Record As New Command

    On Error GoTo ErrHandler

   SQL = " Select D.vehicule, count(*) As nbrRep , "
    SQL = SQL & "Sum(((((Qte * PUHT) - (Qte * PUHT * Remise / 100)) * tva / 100) + ((Qte * PUHT) - (Qte * PUHT * Remise / 100))) - (((((Qte * PUHT) - (Qte * PUHT * Remise / 100)) * tva / 100) + ((Qte * PUHT) - (Qte * PUHT * Remise / 100))) * A.RemisePiece/100)) As ttc "
    SQL = SQL & " from AsspieceReparation A, detailpieceReparation D where A.Numero = D.Numero"
    SQL = SQL & " And A.Supp = 'N' And D.vehicule= " & SQLText(Matricule) & "  And A.datePiece "
    SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF) & " group by d.vehicule"

    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set Print_RepStatVeh = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne details des pieces de repar pour statistiques
Public Function Print_DetRepStatVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal Matricule As String, _
              ByVal VdateD As Date, _
              ByVal vDateF As Date) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".Print_DetRepStatVeh"

    Dim SQL As String
    Dim LRs_Record As New Command

    On Error GoTo ErrHandler

    SQL = " Select D.* , A.DatePiece , "
    SQL = SQL & "(((((Qte * PUHT) - (Qte * PUHT * Remise / 100)) * tva / 100) + ((Qte * PUHT) - (Qte * PUHT * Remise / 100))) - (((((Qte * PUHT) - (Qte * PUHT * Remise / 100)) * tva / 100) + ((Qte * PUHT) - (Qte * PUHT * Remise / 100))) * A.RemisePiece/100)) As montant "
    SQL = SQL & "from AsspieceReparation A , detailpieceReparation D where A.Numero = D.Numero  And A.Supp = 'N' And D.vehicule= " & SQLText(Matricule) & " And "
    SQL = SQL & " A.datePiece between " & SQLText(VdateD) & " and " & SQLText(vDateF)
    
    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set Print_DetRepStatVeh = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Print_Alertes(ByRef ErrNumber As Long, ByRef ErrDescription As String, ByRef ErrSourceDetail As String, ByVal CNB As ADODB.Connection) As Command

    Const sPROC_NAME As String = OBJ_NAME & ".Print_Alertes"

    Dim SQL As String
    Dim LRs_Record As New Command

    On Error GoTo ErrHandler

    SQL = "Select code,Marque,Matricule,Compteurvidange,compteurcarburant,CompteurEntre,"
    SQL = SQL & "DATEDIFF(Day, GETDATE(), DateFinAssur) as Assur,"
    SQL = SQL & "DATEDIFF(Day, GETDATE(), DateFinVisite) as Visite,"
    SQL = SQL & "DATEDIFF(Day, GETDATE(), DateFinTax) as Taxe,"
    SQL = SQL & "IIF(Compteurvidange>0 and (CompteurEntre-Compteurvidange) >= 10000, 'VidangeOblig', '') as Vidange"
    SQL = SQL & " from Vehicule V, Fichetraffic F where V.code=F.vehicule"
    SQL = SQL & " and V.supp='N' and V.actif=1"
    SQL = SQL & " and CompteurEntre in (select Max(compteurEntre) from FicheTraffic group by Vehicule)"
    SQL = SQL & " and  (DATEDIFF(Day, GETDATE(), DateFinAssur) < 0  or DATEDIFF(Day, GETDATE(), DateFinVisite) < 0 or DATEDIFF(Day, GETDATE(), DateFinTax) < 0 or (Compteurvidange>0 and (CompteurEntre-Compteurvidange)>= 10000))"
    SQL = SQL & " group by DateFinAssur,Matricule,DateFinVisite,DateFinTax,Compteurvidange,compteurcarburant, code,Marque, vehicule,CompteurEntre"

    Set LRs_Record = New Command
    Set LRs_Record.ActiveConnection = CNB
    LRs_Record.CommandText = SQL
    LRs_Record.CommandType = adCmdText
    Set Print_Alertes = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
