VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BonCarburant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "BonCarburant"

'Retourne l'assiette du BC à partir du numéroBC
Public Function Get_AssBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AssBC"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "Select * from AssBonCarburant where Numero =" & SQLText(vcode)

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_AssBC = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne les assiettes des BC pour l'afficher dans Forme FrmFind_BC
Public Function Get_BonCarburant(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal DateDu As Date, _
              ByVal DateAu As Date, _
              ByVal Station As String, _
              ByVal TYP As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_BonCarburant"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler


SQL = "SELECT A.Numero,A.Supp, A.DateDoc,A.NumFact,S.Libelle,"
SQL = SQL & " P.Libelle AS Pers,VALEUR"
SQL = SQL & " From AssBonCarburant A"
SQL = SQL & " INNER JOIN Station S ON A.Station =S.Code"
SQL = SQL & " INNER JOIN Personnel P ON A.Conducteur =P.Code"
SQL = SQL & " Where A.DateDoc Between " & SQLText(DateDu) & " and " & SQLText(DateAu)

If Station <> "0000" Then SQL = SQL & " and A.Station = " & SQLText(Station)

'If  supprimer
If TYP = "BC Supprimé" Then
    SQL = SQL & " and A.Supp = 'O'"
Else 'non supprimé
    SQL = SQL & " and A.SUPP = 'N'"
    
    Select Case TYP
        Case "BC Facturé"
            SQL = SQL & " And A.Numfact is not NULL"
        Case "BC Non Facturé"
            SQL = SQL & " And A.Numfact is NULL Or A.Numfact =''"
    End Select
    
End If

SQL = SQL & " Order by A.Numero"

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_BonCarburant = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne la valeur du compteur de carburant dans l'avant dernier bon insérer pour un véhicule
Public Function Get_AnComptCar(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal numbon As String, _
              ByVal vcode As String) As Recordset

 Const sPROC_NAME As String = OBJ_NAME & ".Get_AnComptCar"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "Select CompteurCarburant as maxCpt from DetBonCarburant where Vehicule = " & SQLText(vcode) & "And cast (Numero as int) = ( "
    StrSql = StrSql & " select max(cast(numero As int))From DetBonCarburant where cast (Numero as int) < " & Val(numbon)
    StrSql = StrSql & " and vehicule = " & SQLText(vcode) & ")"
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_AnComptCar = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne les details d'un BC
Public Function Get_DetailBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_DetailBC"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
'DetBonCarburant.AnCompteur,
SQL = "SELECT D.*, A.DateDoc, V.Matricule"
SQL = SQL & " FROM AssBoncarburant A INNER JOIN DetBonCarburant D ON A.Numero=D.Numero INNER JOIN Vehicule V ON D.Vehicule = V.Code"
SQL = SQL & " WHERE D.Numero = " & SQLText(vcode)

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_DetailBC = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne les details d'un BC pour création de la facture
Public Function Get_DetBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_DetBC"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
'DetBonCarburant.AnCompteur,
SQL = "SELECT * from DetBonCarburant Where Numero = " & SQLText(vcode)

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_DetBC = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Produits carburant pour création d'une facture
Public Function Get_ProdCarb(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vCodeSation As String, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_ProdCarb"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler

SQL = "Select Sum(D.Litre) as Qte,energie,prixht,0 As Remise, tva"
SQL = SQL & " From assBonCarburant A, detBonCarburant D"
SQL = SQL & " where A.numero= D.numero and transf='N' and A.Supp='N' and A.dateop between " & SQLText(VdateD) & " and " & SQLText(vDateF) & " and Station=" & SQLText(vCodeSation)
SQL = SQL & " GROUP BY energie,prixht,tva,prixlitre"

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_ProdCarb = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Chercher les bon carburant d'une facture précise
Public Function Get_BCByNumFact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vNumero) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_BCByNumFact"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler

SQL = "Select Sum(D.Litre) as Qte,energie,PrixHT,tva,0 As Remise"
SQL = SQL & " From assBonCarburant A, detBonCarburant D"
SQL = SQL & " where A.numero= D.numero and transf='O' and numfact=" & SQLText(vNumero)
SQL = SQL & " GROUP BY energie,prixht,tva,prixlitre"


LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_BCByNumFact = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Suppression de l'assiette d'un bonCarburant
Public Sub Delete_AssBON(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal CodeUser As String, _
              ByVal id As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_AssBON"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update AssBonCarburant Set Supp = 'O' , UserDelete = " & SQLText(CodeUser) & " where Numero =" & SQLText(id)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Suppression des details d'un bonCarburant
Public Sub Delete_DetailBON(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal id As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_DetailBON"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Delete from DetBonCarburant where Numero = " & SQLText(id)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Suppression des details d'un bonCarburant
Public Sub Delete_DetBCBySupp(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal id As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_DetailBON"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update DetBonCarburant Set Supp='O' where Numero = " & SQLText(id)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Insertion des nouveaux details pour un nouveau bonC ou pour un ancien BC (MAJ : ajout des details)
Public Sub Insert_DetailBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_DetailBC"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
PRs_Record.MoveFirst
While Not PRs_Record.EOF
    StrSql = "Insert into DetBonCarburant (Numero,Vehicule,Energie,CompteurCarburant,litre,PrixHT,prixLitre,Observation,AnomalieConsom) values ("

    If IsNull(PRs_Record!numero) Then
        StrSql = StrSql & "Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!numero)
    End If

    If IsNull(PRs_Record!Vehicule) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Vehicule)
    End If
     
    If IsNull(PRs_Record!Energie) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Energie)
    End If
    
    If IsNull(PRs_Record!CompteurCarburant) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & PRs_Record!CompteurCarburant
    End If
    
    If IsNull(PRs_Record!Litre) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace(PRs_Record!Litre, ",", ".")
    End If
    
    If IsNull(PRs_Record!prixht) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace(PRs_Record!prixht, ",", ".")
    End If

    If IsNull(PRs_Record!prixLitre) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace(PRs_Record!prixLitre, ",", ".")
    End If
    
'    If IsNull(PRs_Record!TVA) Then
'        StrSql = StrSql & "," & "0"
'    Else
'        StrSql = StrSql & "," & Replace(PRs_Record!TVA, ",", ".")
'    End If

    If IsNull(PRs_Record!Observation) Then
        StrSql = StrSql & ", Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Observation)
    End If
    
    If IsNull(PRs_Record!AnomalieConsom) Then
        StrSql = StrSql & ", 0 "
    Else
        StrSql = StrSql & "," & Replace(PRs_Record!AnomalieConsom, ",", ".")
    End If


    StrSql = StrSql & ")"

    CNB.Execute StrSql
    
PRs_Record.MoveNext
Wend

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Insertion de l'assiette d'un nouveau BC
Public Sub Insert_AssBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_AssBC"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Insert into AssBonCarburant (Numero,DateDoc,Heure,Station,Conducteur,valeur,nbc,dateop ,UserInsert) values ("

    If IsNull(PRs_Record!numero) Then
        StrSql = StrSql & "Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!numero)
    End If

    If IsNull(PRs_Record!datedoc) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!datedoc)
    End If
    
    If IsNull(PRs_Record!Heure) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Heure)
    End If
     
    If IsNull(PRs_Record!Station) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Station)
    End If
    
    If IsNull(PRs_Record!Conducteur) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Conducteur)
    End If
    
    If IsNull(PRs_Record!valeur) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace(PRs_Record!valeur, ",", ".")
    End If
    
    If IsNull(PRs_Record!NBC) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Val(PRs_Record!NBC)
    End If
    
    If IsNull(PRs_Record!dateOP) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!dateOP)
    End If
    
    If IsNull(PRs_Record!UserInsert) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!UserInsert)
    End If
    
   
    StrSql = StrSql & ")"

    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'MAJ de l'assiette d'un BC
Public Sub Update_AssBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_AssBC"

    Dim StrSql As String

    On Error GoTo ErrHandler
'Les champs à modifier Conducteur,valeur,nbc,dateop
    StrSql = "Update AssBoncarburant set Conducteur = "

    If IsNull(PRs_Record!Conducteur) Then
        StrSql = StrSql & "  Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!Conducteur)
    End If
    
    If IsNull(PRs_Record!valeur) Then
        StrSql = StrSql & ", valeur = 0"
    Else
        StrSql = StrSql & ", valeur = " & Replace(PRs_Record!valeur, ",", ".")
    End If
    
    If IsNull(PRs_Record!NBC) Then
        StrSql = StrSql & ", nbc = 0"
    Else
        StrSql = StrSql & ", nbc = " & Val(PRs_Record!NBC)
    End If
    
    If IsNull(PRs_Record!dateOP) Then
        StrSql = StrSql & ", dateop = Null"
    Else
        StrSql = StrSql & ", dateop = " & SQLText(PRs_Record!dateOP)
    End If
    
    If IsNull(PRs_Record!userupdate) Then
        StrSql = StrSql & ", UserUpdate = Null"
    Else
        StrSql = StrSql & ", UserUpdate = " & SQLText(PRs_Record!userupdate)
    End If
    
    StrSql = StrSql & " Where Numero = " & SQLText(PRs_Record!numero)

    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Retourne le numero max des bons pour un vehicule
Public Function Get_MaxNumBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_MaxNumBC"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "Select max(cast(numero As int)) as maxNum from detBonCarburant where vehicule = " & SQLText(vcode)

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_MaxNumBC = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Lors de la suppression d'une facture
Public Sub Update_NumFact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_NumFact"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update AssBoncarburant set transf='N' , numfact='' where Numfact=" & SQLText(vcode)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Facture transferer
Public Sub Insert_NumFact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String, _
              ByVal N As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_NumFact"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update AssBoncarburant set transf='O'  , numfact= " & SQLText(vcode) & " where Numero=" & SQLText(N)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'========================Statistique==========================================
'=============================================================================

'Retourne les details d'un BC pour statistiques de tout les véhicules durant une période
Public Function Get_StatistDetBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_StatistDetBC"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler

SQL = "SELECT D.Numero,D.vehicule , D.CompteurCarburant As Compteur,D.AnomalieConsom , V.Matricule,"
SQL = SQL & " D.Litre As Litre, D.prixLitre As prixLitre, A.dateDoc As DateDoc"
SQL = SQL & " From AssBonCarburant A inner join DetBonCarburant D "
SQL = SQL & " on A.Numero=D.Numero inner join Vehicule V "
SQL = SQL & " on D.vehicule = V.code"
SQL = SQL & " where A.Supp = 'N' and A.datedoc Between" & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00"))

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_StatistDetBC = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne les details d'un BC
Public Function Get_StatistBC(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_StatistBC"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
'DetBonCarburant.AnCompteur,
SQL = "Select Max(D.Numero) as Numero , SUM(D.Litre) As Litre, D.prixLitre As Prix,V.Matricule As Vehicule"
SQL = SQL & " , Max(D.CompteurCarburant) AS MaxC, Min(D.CompteurCarburant) As MinC"
SQL = SQL & " From AssBonCarburant A , DetBonCarburant D , Vehicule V"
SQL = SQL & " Where A.Numero = D.Numero and D.Vehicule = V.Code"
SQL = SQL & " AND A.Supp ='N' and A.datedoc Between" & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00"))
SQL = SQL & " group by V.Matricule , D.prixLitre"

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_StatistBC = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne les details d'un BC pour statistiques de tout les véhicules durant une période
Public Function Get_StatistBCVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As String, _
              ByVal vDateF As String, _
              ByVal CodV As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_StatistBCVeh"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
'DetBonCarburant.AnCompteur,
SQL = "SELECT Distinct D.Numero, D.Litre, D.prixLitre, A.DateDoc, D.CompteurCarburant as compteur , D.AnomalieConsom , V.Matricule as Matricule"
SQL = SQL & " From DetBonCarburant D , AssBonCarburant A , Vehicule v Where A.Numero=D.Numero and D.Vehicule=V.code And A.Supp ='N' and A.datedoc Between "
SQL = SQL & SQLText(Format(VdateD, "dd/mm/yyyy 00:00:00:00")) & " and " & SQLText(Format(vDateF, "dd/mm/yyyy 23:59:59:00"))
SQL = SQL & " And D.Vehicule=" & SQLText(CodV) & " Order by D.CompteurCarburant"

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_StatistBCVeh = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_MoyConsom(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal Vdate As Date, _
              ByVal CodV As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_MoyConsom"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
'DetBonCarburant.AnCompteur,
SQL = "select Y.CompteurCarburant as AnCmpt, X.CompteurCarburant , X.Litre ,"
SQL = SQL & " X.CompteurCarburant - Y.CompteurCarburant as KmParc , "
SQL = SQL & " (X.Litre*100/(X.CompteurCarburant - Y.CompteurCarburant)) as Consom "
SQL = SQL & " from( "
SQL = SQL & " select rank() Over(Order by CompteurCarburant) as rank, D.Numero , Vehicule , CompteurCarburant , D.Litre ,a.dateop "
SQL = SQL & " from  DetBonCarburant D, AssBonCarburant a "
SQL = SQL & " where D.Numero=a.Numero and vehicule  = " & SQLText(CodV)
SQL = SQL & " and datediff(month , a.dateop," & SQLText(Vdate) & ") <= 6 "
SQL = SQL & " ) AS Y Inner Join ("
SQL = SQL & " select rank()  Over(Order by CompteurCarburant) as rank , d.Numero , Vehicule , CompteurCarburant , D.Litre , a.dateop "
SQL = SQL & " from  DetBonCarburant D , AssBonCarburant a "
SQL = SQL & " where D.Numero=a.Numero and vehicule  = " & SQLText(CodV)
SQL = SQL & " and datediff(month , a.dateop, " & SQLText(Vdate) & ") <= 6 "
SQL = SQL & " ) AS X ON Y.rank+1 = X.rank Order by Y.CompteurCarburant"

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_MoyConsom = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

