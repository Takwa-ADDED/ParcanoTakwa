VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Energie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "Energie"

'Retourne les libellés des différentes type d'énergie
Public Function Get_Energie(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_Energie"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "Select distinct libelle from energie "

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_Energie = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function Get_EnergByLibelle(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal txt As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_EnergByLibelle"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from energie where libelle like '" & txt & "%' "

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_EnergByLibelle = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_EnergByLiborCod(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal txt As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_EnergByLiborCod"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from energie where libelle = " & SQLText(txt) & " or Code = " & "'" & txt & "'"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_EnergByLiborCod = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_Energ(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_Energ"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Energie order by libelle "

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_Energ = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne le prix de l'energie par son libelle
Public Function Get_PrixEnergie(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal libele As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_PrixEnergie"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "select Prix, tht, tva from energie where libelle = " & SQLText(libele)

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_PrixEnergie = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Sub Delete_Add_Energie(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal vcode As String, _
                                ByVal vSupp As String, _
                                ByVal CNB As ADODB.Connection)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Add_Energie"

    Dim SQL As String
    
    If vSupp <> "N" And vSupp <> "O" Then
        MsgBox "variable supp invalide", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    SQL = "Update Energie set Supp = " & SQLText(vSupp)

'    If vSupp = "O" Then
'        Sql = Sql & " , UserDelete=" & SQLText(CodeUser)
'    ElseIf vSupp = "N" Then
'        Sql = Sql & " , UserAdd=" & SQLText(CodeUser)
'    End If
    SQL = SQL & " where Code =" & SQLText(vcode)
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub

Public Sub Insert_Energie(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_Energie"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
    StrSql = "Insert into energie  (Code,libelle,tht,tva,Prix,UserInsert) values ("
    
    StrSql = StrSql & SQLText(PRs_Record!code)
    
    If IsNull(PRs_Record!Libelle) Then
        StrSql = StrSql & ", Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Libelle)
    End If
    
    If IsNull(PRs_Record!THT) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!THT), ",", ".")
    End If
    
    If IsNull(PRs_Record!tva) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!tva), ",", ".")
    End If
    
    If IsNull(PRs_Record!prix) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!prix), ",", ".")
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

Public Sub Update_Energie(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_Energie"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update Energie set "
   
   If IsNull(PRs_Record!THT) Then
        StrSql = StrSql & " tht = 0"
    Else
        StrSql = StrSql & " tht = " & Replace((PRs_Record!THT), ",", ".")
    End If
   
    If IsNull(PRs_Record!tva) Then
        StrSql = StrSql & ", tva = 0"
    Else
        StrSql = StrSql & ", tva = " & Replace((PRs_Record!tva), ",", ".")
    End If
    

    If IsNull(PRs_Record!prix) Then
        StrSql = StrSql & " , prix =0"
    Else
        StrSql = StrSql & " , prix = " & Replace((PRs_Record!prix), ",", ".")
    End If
    
    If IsNull(PRs_Record!USERINSERT) Then
        StrSql = StrSql & ", UserInsert = Null"
    Else
        StrSql = StrSql & ", UserInsert = " & SQLText(PRs_Record!USERINSERT)
    End If

    StrSql = StrSql & " Where Libelle = " & SQLText(PRs_Record!Libelle)
    
 CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

