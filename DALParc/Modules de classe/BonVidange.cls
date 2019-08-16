VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BonVidange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "BonVidange"

Public Sub Insert_LubBV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_LubBV"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
PRs_Record.MoveFirst
While Not PRs_Record.EOF
    StrSql = "Insert into Det_BonVidange (Numero, Libelle,THT,TVA, prixTTC) values ("

    If IsNull(PRs_Record!numero) Then
        StrSql = StrSql & "Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!numero)
    End If

    If IsNull(PRs_Record!Libelle) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Libelle)
    End If
    
    
    If IsNull(PRs_Record!THT) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace(PRs_Record!THT, ",", ".")
    End If
    
    If IsNull(PRs_Record!tva) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace(PRs_Record!tva, ",", ".")
    End If
    
    
    If IsNull(PRs_Record!prixTTC) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace(PRs_Record!prixTTC, ",", ".")
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

Public Sub Insert_BV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_BV"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
    StrSql = "Insert into Ass_BonVidange(Numero,DateDoc,Vehicule,Station,Conducteur,valeur,heure,NBC,dateop,CompteurVidange,NBKLMvid , UserInsert) values ("

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
    
    If IsNull(PRs_Record!Vehicule) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Vehicule)
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
    
    If IsNull(PRs_Record!Heure) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Heure)
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
    
    If IsNull(PRs_Record!CompteurVidange) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Val(PRs_Record!CompteurVidange)
    End If
    
     If IsNull(PRs_Record!NBKLMvid) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Val(PRs_Record!NBKLMvid)
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

Public Sub Update_BV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_BV"

    Dim StrSql As String

    On Error GoTo ErrHandler
'Les champs à modifier Conducteur,valeur,NBC,dateop,CompteurVidange,NBKLMvid

    StrSql = "Update Ass_BonVidange "

    If IsNull(PRs_Record!Conducteur) Then
        StrSql = StrSql & " set Conducteur = 0"
    Else
        StrSql = StrSql & " set Conducteur = " & SQLText(PRs_Record!Conducteur)
    End If
    
    If IsNull(PRs_Record!valeur) Then
        StrSql = StrSql & ", valeur = 0"
    Else
        StrSql = StrSql & ", valeur = " & Replace(PRs_Record!valeur, ",", ".")
    End If
    
    If IsNull(PRs_Record!NBC) Then
        StrSql = StrSql & ", NBC = 0"
    Else
        StrSql = StrSql & ", NBC = " & Val(PRs_Record!NBC)
    End If
    
    If IsNull(PRs_Record!dateOP) Then
        StrSql = StrSql & ", dateop = 0"
    Else
        StrSql = StrSql & ", dateop = " & SQLText(PRs_Record!dateOP)
    End If
    
    If IsNull(PRs_Record!Heure) Then
        StrSql = StrSql & ", Heure = 0"
    Else
        StrSql = StrSql & ", Heure = " & SQLText(PRs_Record!Heure)
    End If
    
    If IsNull(PRs_Record!CompteurVidange) Then
        StrSql = StrSql & ", CompteurVidange = 0"
    Else
        StrSql = StrSql & ", CompteurVidange = " & Val(PRs_Record!CompteurVidange)
    End If
    
    If IsNull(PRs_Record!NBKLMvid) Then
        StrSql = StrSql & ", NBKLMvid = 0"
    Else
        StrSql = StrSql & ", NBKLMvid = " & Val(PRs_Record!NBKLMvid)
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

'Suppresion d'un bonVidange non facturé par Update : mettre la valeur de Supp ='O'
Public Sub DeleteBV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal id As String, _
              ByVal idUser As String)

    Const sPROC_NAME As String = OBJ_NAME & ".DeleteBV"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update Ass_BonVidange Set Supp ='O' , UserDelete = " & SQLText(idUser) & " Where Numero = " & SQLText(id)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Suppression des détails d'un BV
Public Sub DeleteBV_Lub(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal id As String)

    Const sPROC_NAME As String = OBJ_NAME & ".DeleteBV_Lub"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update Det_BonVidange set Supp='O' where Numero = " & SQLText(id)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub


'Suppression des détails d'un BV
Public Sub Delete_DetBV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal id As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_DetBV"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Delete from Det_BonVidange where Numero =" & SQLText(id)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'retourne un bon de vidange
Public Function Get_BV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset
              
Const sPROC_NAME As String = OBJ_NAME & ".Get_BV"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    StrSql = "Select * from Ass_BonVidange where Numero = " & SQLText(vcode)
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_BV = LRs_Record
    Set LRs_Record = Nothing

Exit Function

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Function

Public Function Get_DerBV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset
              
Const sPROC_NAME As String = OBJ_NAME & ".Get_DerBV"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    ' dernier bonV ajouté pour ce véhicule
    StrSql = "Select Datedoc,DateOp , CompteurVidange, NBKLMvid from Ass_BonVidange where Vehicule = " & SQLText(vcode) & " And Numero = (select Max(Numero)from Ass_BonVidange where Vehicule = " & SQLText(vcode) & ")"
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_DerBV = LRs_Record
    Set LRs_Record = Nothing

Exit Function

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Function
Public Function Get_LastCompeteurV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset
              
Const sPROC_NAME As String = OBJ_NAME & ".Get_LastCompeteurV"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    
    SQL = "select CompteurVidange, dateop from Ass_BonVidange where dateop = (SELECT MAX (dateop) FROM Ass_BonVidange)"

    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_LastCompeteurV = LRs_Record
    Set LRs_Record = Nothing

Exit Function

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Function

Public Function Get_BVAfich(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal DateDu As Date, _
              ByVal DateAu As Date, _
              ByVal Station As String, _
              ByVal TYP As String) As Recordset
              
Const sPROC_NAME As String = OBJ_NAME & ".Get_BVAfich"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    
    SQL = " Select Numero, dateDoc,A.Supp, S.Libelle As station, V.Matricule, P.Libelle As Conducteur, Valeur"
    SQL = SQL & " FROM Ass_BonVidange A INNER JOIN"
    SQL = SQL & " Vehicule V ON A.Vehicule = V.Code INNER JOIN"
    SQL = SQL & " Station S ON A.Station = S.Code INNER JOIN"
    SQL = SQL & " Personnel P ON A.Conducteur = P.Code"
    SQL = SQL & " Where A.DateDoc Between " & SQLText(DateDu) & " and " & SQLText(DateAu)
    
    If Station <> "0000" Then SQL = SQL & " and A.Station = " & SQLText(Station)

    'If  supprimer
    If TYP = "BV Supprimé" Then
        SQL = SQL & " and A.Supp = 'O'"
    Else 'non supprimé
        SQL = SQL & " and A.SUPP = 'N'"
        
        Select Case TYP
            Case "BV Facturé"
                SQL = SQL & " And A.Numfact is not NULL"
            Case "BV Non Facturé"
                SQL = SQL & " And A.Numfact is NULL Or A.Numfact =''"
        End Select
        
    End If
    
    SQL = SQL & " Order by Numero"
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_BVAfich = LRs_Record
    Set LRs_Record = Nothing

Exit Function

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Function

'Retourne la valeur du maxNumero
Public Function Get_MaxNumBV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset
              
Const sPROC_NAME As String = OBJ_NAME & ".Get_MaxNumBV"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    StrSql = "select Max(Numero) as maxNum from Ass_BonVidange  "
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_MaxNumBV = LRs_Record
    Set LRs_Record = Nothing

Exit Function

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Function

'Retourne le numero max des bons pour un vehicule
Public Function Get_NumMaxBV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_NumMaxBV"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "Select max(cast(numero As int)) as maxNum from Ass_BonVidange where vehicule =" & SQLText(vcode)

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_NumMaxBV = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne la valeur du compteur de vidange dans l'avant dernier bon insérer pour un véhicule
Public Function Get_AnComptVdg(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal numbon As String, _
              ByVal vcode As String) As Recordset

 Const sPROC_NAME As String = OBJ_NAME & ".Get_AnComptVdg"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "Select CompteurVidange as maxCpt from Ass_BonVidange where Vehicule = " & SQLText(vcode) & " And cast (Numero as int) = ( "
    StrSql = StrSql & " select max(cast(numero As int)) From Ass_BonVidange where cast (Numero as int) < " & Val(numbon)
    StrSql = StrSql & " And vehicule = " & SQLText(vcode) & ")"
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_AnComptVdg = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'retourne les details de la table Det_BonVidange selon son numero
Public Function Get_Lub_BV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_Lub_BV"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Det_BonVidange Where Det_BonVidange.Numero=" & SQLText(vcode)
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_Lub_BV = LRs_Record
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

    StrSql = "Update Ass_BonVidange set transf= 'N' , NumFact='' where NumFact = " & SQLText(vcode)
   
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

    StrSql = "Update Ass_BonVidange set transf='O'  , numfact= " & SQLText(vcode) & " where Numero=" & SQLText(N)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Retourne les details d'un Bv pour création de la facture
Public Function Get_DetBV(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_DetBV"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler

SQL = "SELECT Distinct Det_BonVidange.Libelle, 1 as Qte , SUM(THT) as THT From Det_BonVidange"
SQL = SQL & " INNER JOIN Ass_BonVidange ON Det_BonVidange.Numero = Ass_BonVidange.Numero"
SQL = SQL & " WHERE Ass_BonVidange.Numero =" & SQLText(vcode) & " Group by Libelle "
    

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_DetBV = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Produits vdg pour création d'une facture
Public Function Get_ProdVdg(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vCodeSation As String, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_ProdVdg"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
'SUM(Det_BonVidange.Qte)
SQL = "Select D.Libelle, SUM(1) AS Qte, MIN(D.THT) AS Prixht,MIN(D.prixTTC) AS prixtt,MIN(D.tva) AS tva, SUM(prixTTC) AS Valeur,0 As Remise"
SQL = SQL & " From Det_BonVidange D"
SQL = SQL & " INNER JOIN Ass_BonVidange A ON D.Numero  = A.Numero"
SQL = SQL & " where transf='N' and A.Supp='N' and dateop between  " & SQLText(VdateD) & " and " & SQLText(vDateF) & " and Station=" & SQLText(vCodeSation)
SQL = SQL & " GROUP BY D.Libelle, D.prixTTC"

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_ProdVdg = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Chercher les bon vidange d'une facture précise
Public Function Get_BVByNumFact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vNumero) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_BVByNumFact"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
'SUM(Det_BonVidange.Qte)
SQL = "Select Det_BonVidange.Libelle, SUM(1) AS Qte, MIN(Det_BonVidange.THT) AS Prixht,"
SQL = SQL & " MIN(Det_BonVidange.PrixTTC) AS prixtt,MIN(Det_BonVidange.tva) AS tva,"
SQL = SQL & " SUM(PrixTTC) AS Valeur, 0 As Remise"
SQL = SQL & " From Det_BonVidange"
SQL = SQL & " INNER JOIN Ass_BonVidange ON Ass_BonVidange.Numero = Det_BonVidange.Numero"
SQL = SQL & " where transf='O' and numfact=" & SQLText(vNumero)
SQL = SQL & " GROUP BY Det_BonVidange.Libelle"


LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_BVByNumFact = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function



'Public Function Get_BVAfich(ByRef ErrNumber As Long, _
'              ByRef ErrDescription As String, _
'              ByRef ErrSourceDetail As String, _
'              ByVal CNB As ADODB.Connection) As Recordset
'
'Const sPROC_NAME As String = OBJ_NAME & ".Get_BVAfich"
'
'    Dim SQL As String
'    Dim LRs_Record As New Recordset
'
'    On Error GoTo ErrHandler
'    SQL = "SELECT   Numero,  Ass_BonVidange.DateDoc, Station.Libelle AS Station, Vehicule.Matricule, "
'    SQL = SQL & " Personnel.Libelle AS Conducteur,Valeur"
'    SQL = SQL & " FROM Ass_BonVidange INNER JOIN"
'    SQL = SQL & " Vehicule ON Ass_BonVidange.Vehicule = Vehicule.Code INNER JOIN"
'    SQL = SQL & " Station ON Ass_BonVidange.Station = Station.Code INNER JOIN"
'    SQL = SQL & " Personnel ON Ass_BonVidange.Conducteur = Personnel.Code"
'    SQL = SQL & " Order by Numero "
'    LRs_Record.CursorLocation = adUseClient
'    LRs_Record.Open SQL, CNB, adOpenKeyset
'    Set Get_BVAfich = LRs_Record
'    Set LRs_Record = Nothing
'
'Exit Function
'
'ErrHandler:
'
'    ErrNumber = Err.Number
'    ErrDescription = Err.Description
'    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
'
'End Function
