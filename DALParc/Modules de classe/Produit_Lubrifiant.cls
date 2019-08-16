VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Produit_Lubrifiant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "Produit_Lubrifiant"

Public Function Get_ProduitRepa(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ProduitRepa"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "Select Numero ,Libelle , prixht,tva , Type_PL as type "
    SQL = SQL & " from Produit_Lubrifiant where Actif='O' Order by Type_PL DESC , Libelle ASC "
    
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

Public Function Get_ProdLubByInit(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal vLibelle As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ProdLubByInit"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "Select Numero ,Libelle , tva  , prixht ,Actif, Type_PL as type "
    SQL = SQL & " from Produit_Lubrifiant where Actif='O' And Libelle Like '" & vLibelle & "%'"
    SQL = SQL & " Or Numero = " & Val(vLibelle) & " Order by Libelle"
   
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_ProdLubByInit = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Recherche par code d'un prodLub
Public Function Get_ProdLubBycode(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ProdLubBycode"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "Select * from Produit_Lubrifiant where Numero = " & Val(vcode)
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_ProdLubBycode = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Tout les produits actifs
Public Function Get_Produits(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_Produits"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "select * from Produit_Lubrifiant where Type_PL = 'Produit' And Actif='O' order by libelle"
    
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

    SQL = "select * from Produit_Lubrifiant where Type_PL = 'Produit' order by libelle"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_AllProduits = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Tout les articles (Lubrifiants et produits)
Public Function Get_AllArticles(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AllArticles"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "select * from Produit_Lubrifiant Order by Type_PL DESC , Libelle ASC "
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_AllArticles = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Tout les articles actifs
Public Function Get_ActifArticles(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ActifArticles"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "select * from Produit_Lubrifiant where Actif='O' Order by Type_PL DESC , Libelle ASC "
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_ActifArticles = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'TypeReparation
Public Function Get_ProdLubByLib(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ProdLubByLib"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = "select * from Produit_Lubrifiant where Libelle = " & SQLText(vcode)
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_ProdLubByLib = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Sub Update_ProdLub(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_ProdLub"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update Produit_Lubrifiant set "
    
    If IsNull(PRs_Record!Libelle) Then
        StrSql = StrSql & "Libelle = Null"
    Else
        StrSql = StrSql & "Libelle = " & SQLText(PRs_Record!Libelle)
    End If
   
    If IsNull(PRs_Record!tva) Then
        StrSql = StrSql & ", tva = 0"
    Else
        StrSql = StrSql & ", tva = " & Replace((PRs_Record!tva), ",", ".")
    End If
    
    If IsNull(PRs_Record!Type_PL) Then
        StrSql = StrSql & ", Type_PL = Null"
    Else
        StrSql = StrSql & ", Type_PL = " & SQLText(PRs_Record!Type_PL)
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
    
    If IsNull(PRs_Record!Actif) Then
        StrSql = StrSql & " , Actif = 'O'"
    Else
        StrSql = StrSql & " , Actif = " & SQLText(PRs_Record!Actif)
    End If
    
    If IsNull(PRs_Record!OperateurSaisi) Then
        StrSql = StrSql & ", OperateurSaisi = Null"
    Else
        StrSql = StrSql & ", OperateurSaisi = " & SQLText(PRs_Record!OperateurSaisi)
    End If

    StrSql = StrSql & " Where Numero = " & PRs_Record!numero
    
 CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Sub Insert_ProdLub(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_ProdLub"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
    StrSql = "Insert into Produit_Lubrifiant ( Libelle, tva, prixht, DatePrix,Type_PL, Actif, OperateurSaisi) Values ("
    
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
    
     If IsNull(PRs_Record!Type_PL) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Type_PL)
    End If
    
    If IsNull(PRs_Record!Actif) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Actif)
    End If
    
    If IsNull(PRs_Record!OperateurSaisi) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!OperateurSaisi)
    End If

    StrSql = StrSql & ")"

    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub


'Numero Max dans la table Produit_Lubrifiant(type reparation)
Public Function Get_MaxNum(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_MaxNum"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = "select Max(Numero) as maxnum from Produit_Lubrifiant "
    
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
Public Sub Delete_Add_ProdLub(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal code As String, _
                                ByVal vSupp As String, _
                                ByVal CNB As ADODB.Connection)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Add_ProdLub"

    Dim SQL As String
    
    If vSupp <> "N" And vSupp <> "O" Then
        MsgBox "variable supp invalide", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    SQL = "Update Produit_Lubrifiant set Supp =" & SQLText(vSupp)

'    If vSupp = "O" Then
'        Sql = Sql & " , UserDelete=" & SQLText(CodeUser)
'    ElseIf vSupp = "N" Then
'        Sql = Sql & " , UserAdd=" & SQLText(CodeUser)
'    End If

    SQL = SQL & " where Numero =" & Val(code)
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub

'***************************************************************
'******************************Lubrifiant*****************************

'Retourne la liste de libelles de tout les Lubrifiants actifs
Public Function Get_LibLubActif(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_LibLubActif"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select libelle from Produit_Lubrifiant where Type_PL ='Lubrifiant' And actif='O' and supp='N' order by libelle "

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_LibLubActif = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne la liste des tout les Lubrifiants actifs
Public Function Get_LubActif(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_LubActif"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Produit_Lubrifiant where Type_PL ='Lubrifiant' And  actif='O' order by libelle "

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_LubActif = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_Lubrifiant(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_Lubrifiant"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Produit_Lubrifiant where Type_PL ='Lubrifiant' order by libelle"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_Lubrifiant = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_LibelLubrif(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_LibelLubrif"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select libelle from Produit_Lubrifiant where Type_PL ='Lubrifiant' order by libelle"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_LibelLubrif = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'retourne le code d'un lubrifiant selon son libelle
Public Function Get_CodeLub(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_CodeLub"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select Numero from Produit_Lubrifiant where libelle =" & SQLText(vcode)
    

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_CodeLub = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'retourne libelle d'un lubrifiant selon son code
Public Function Get_LibelleLub(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_LibelleLub"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select Libelle from Produit_Lubrifiant where Numero = " & Val(vcode)
    

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_LibelleLub = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

