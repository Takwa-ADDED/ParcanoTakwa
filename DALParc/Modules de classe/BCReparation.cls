VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "BCReparation"

Public Sub Delete_BRepa(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vuser As String, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_BRepa"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update AssBCReparation Set Supp ='O' , UserDelete = " & SQLText(vuser) & " Where Numero = " & SQLText(vcode)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Suppression des détails d'un BV
Public Sub Delete_DetRep(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_DetRep"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Delete from detailBCReparation where Numero = " & SQLText(vcode)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub
'Suppression des détails d'un BV
Public Sub Delete_DetRepBySup(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_DetRepBySup"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update detailBCReparation Set Supp ='O' where Numero = " & SQLText(vcode)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Sub Update_BRepar(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_BRepar"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update AssBCReparation set"
    
    If IsNull(PRs_Record!Fournisseur) Then
        StrSql = StrSql & " Fournisseur = Null"
    Else
        StrSql = StrSql & " Fournisseur = " & SQLText(PRs_Record!Fournisseur)
    End If

    If IsNull(PRs_Record!Conducteur) Then
        StrSql = StrSql & " , Conducteur = Null"
    Else
        StrSql = StrSql & " , Conducteur = " & SQLText(PRs_Record!Conducteur)
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

Public Sub Update_DetBRepar(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_DetBRepar"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update detailBCreparation set désignation = "
    
    If IsNull(PRs_Record!désignation) Then
        StrSql = StrSql & "  Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!désignation)
    End If

    If IsNull(PRs_Record!Qté) Then
        StrSql = StrSql & " , Qté = 0"
    Else
        StrSql = StrSql & " , Qté = " & Val(PRs_Record!Qté)
    End If

    If IsNull(PRs_Record!Vehicule) Then
        StrSql = StrSql & ", Vehicule = Null"
    Else
        StrSql = StrSql & ", Vehicule = " & SQLText(PRs_Record!Vehicule)
    End If
    
    StrSql = StrSql & " Where Numero = " & SQLText(PRs_Record!numero)

    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Sub Insert_BCRepar(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_BCRepar"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
    StrSql = "Insert into AssBCReparation (Numero,DateCreation, Fournisseur,Conducteur,UserInsert) values ("

    If IsNull(PRs_Record!numero) Then
        StrSql = StrSql & " Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!numero)
    End If

    StrSql = StrSql & "," & SQLText(PRs_Record!DateCreation)
  
    If IsNull(PRs_Record!Fournisseur) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Fournisseur)
    End If
    
    If IsNull(PRs_Record!Conducteur) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Conducteur)
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

Public Sub Insert_DetBCRepar(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_DetBCRepar"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
PRs_Record.MoveFirst
While Not PRs_Record.EOF
    StrSql = "Insert into detailBCreparation (Numero,désignation, Qté,Vehicule,Observation) values ("

    If IsNull(PRs_Record!numero) Then
        StrSql = StrSql & " Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!numero)
    End If

    If IsNull(PRs_Record!désignation) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!désignation)
    End If
    
    If IsNull(PRs_Record!Qté) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Val(PRs_Record!Qté)
    End If
    
    If IsNull(PRs_Record!Vehicule) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Vehicule)
    End If
    
     If IsNull(PRs_Record!Observation) Then
        StrSql = StrSql & "," & " Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Observation)
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

'Retourne Max des numeros des BCReparation
Public Function Get_MaxNum(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_MaxNum"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "select Max(Numero)as maxNum from AssBCReparation"

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_MaxNum = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne les détails d'un bon de réparation par le numero de ce bon
Public Function Get_DetBRepar(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal vcode As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_DetBRepar"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "SELECT * from detailBCReparation WHERE detailBCReparation.Numero = " & SQLText(vcode)

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_DetBRepar = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_AssBRepar(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal vcode As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AssBRepar"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    StrSql = "SELECT * from AssBCReparation WHERE Numero = " & SQLText(vcode)

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_AssBRepar = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_Reparation(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_Reparation"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = "SELECT Numero,DateCreation,S.Code As CoDFournissseur, S.Libelle As Fournisseur from AssBCReparation A"
    SQL = SQL & " INNER JOIN Station S ON A.fournisseur=S.code"
    SQL = SQL & " Order by Numero "
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_Reparation = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_ReparNTrans(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal DateDu As Date, _
              ByVal DateAu As Date, _
              ByVal Station As String, _
              ByVal vtransf As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ReparNTrans"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = "SELECT Numero,DateCreation,S.Code As CoDFournissseur, S.Libelle As Fournisseur from AssBCReparation A"
    SQL = SQL & " INNER JOIN Station S ON A.fournisseur=S.code Where "
    SQL = SQL & "DateCreation Between " & SQLText(DateDu) & " and " & SQLText(DateAu)
    If Station <> "0000" Then SQL = SQL & " and A.fournisseur = " & SQLText(Station)
    
    'BC transféré
    If vtransf = "BC transféré" Then SQL = SQL & " And A.Transf = 'O'  "
    'BC non transféré
    If vtransf = "BC Non transféré" Then SQL = SQL & "And A.Transf = 'N' "
    
    SQL = SQL & " Order by Numero "
 
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_ReparNTrans = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

