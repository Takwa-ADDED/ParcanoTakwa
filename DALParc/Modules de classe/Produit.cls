VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Produit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "Produit"

Public Function Get_ProduitRepa(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ProduitRepa"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "Select 'Lub' As type, Numero ,Libelle , tva  , prixht "
    SQL = SQL & " from Produit where Produit.Actif='O'"
    SQL = SQL & " Union All"
    SQL = SQL & " Select 'Prod' As type, Code As Numero,Libelle , TVA , tht As prixht"
    SQL = SQL & " from Lubrifiant where Lubrifiant.Actif='O'"
    SQL = SQL & " Order by type, Libelle"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_ProduitRepa = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_ProdRepaByInitial(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal vlibelle As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ProdRepaByInitial"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "Select 'Lub' As type, Numero ,Libelle , tva  , prixht "
    SQL = SQL & " from Produit where Produit.Actif='O' And Produit.Libelle Like '" & vlibelle & "%'"
    SQL = SQL & " Union All"
    SQL = SQL & " Select 'Prod' As type, Code As Numero,Libelle , TVA , tht As prixht"
    SQL = SQL & " from Lubrifiant where Lubrifiant.Actif='O'And Lubrifiant.Libelle Like '" & vlibelle & "%'"
    SQL = SQL & " Order by type, Libelle"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_ProdRepaByInitial = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_ProdBycode(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ProdBycode"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "Select * from Produit where Numero = " & SQLText(vcode)
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_ProdBycode = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_Produits(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_Produits"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "select * from Produit where Actif='O' order by libelle"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_Produits = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_AllProduits(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AllProduits"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "select * from Produit order by libelle"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_AllProduits = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


'TypeReparation
Public Function Get_ProdTypRepa(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ProdTypRepa"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = "select * from Produit where Libelle=" & SQLText(vcode)
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_ProdTypRepa = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Sub Update_Produit(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_Produit"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update Produit set"
   
    If IsNull(PRs_Record!tva) Then
        StrSql = StrSql & " tva = 0"
    Else
        StrSql = StrSql & " tva = " & Replace((PRs_Record!tva), ",", ".")
    End If
    
    If IsNull(PRs_Record!DatePrix) Then
        StrSql = StrSql & ", DatePrix = Null"
    Else
        StrSql = StrSql & ", DatePrix = " & SQLText(PRs_Record!DatePrix)
    End If
    
    If IsNull(PRs_Record!prixht) Then
        StrSql = StrSql & ", prixht = 0"
    Else
        StrSql = StrSql & ", prixht = " & Replace((PRs_Record!prixht), ",", ".")
    End If
    
    If IsNull(PRs_Record!actif) Then
        StrSql = StrSql & " , Actif = 'O'"
    Else
        StrSql = StrSql & " , Actif = " & SQLText(PRs_Record!actif)
    End If

    StrSql = StrSql & " Where Libelle = " & SQLText(PRs_Record!Libelle)
    
 CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Sub Insert_Produit(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_Produit"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
    StrSql = "Insert into Produit (Numero, Libelle, tva, prixht, DatePrix, Actif) Values ("

    If IsNull(PRs_Record!numero) Then
        StrSql = StrSql & "Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!numero)
    End If
    
     If IsNull(PRs_Record!Libelle) Then
        StrSql = StrSql & "Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!Libelle)
    End If
    
    If IsNull(PRs_Record!tva) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!tva), ",", ".")
    End If
    
    If IsNull(PRs_Record!prixht) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!prixht), ",", ".")
    End If
    
    If IsNull(PRs_Record!DatePrix) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!DatePrix)
    End If
    
    If IsNull(PRs_Record!actif) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!actif)
    End If
    

    StrSql = StrSql & ")"

    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub


'Numero Max dans la table Produit( type reparation)
Public Function Get_MaxNum(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_MaxNum"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = "select Max(Numero) as maxnum from Produit "
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_MaxNum = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Sub Delete_Prod(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Prod"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update Produit Set Supp = 'O' where Numero =" & SQLText(vcode)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub
