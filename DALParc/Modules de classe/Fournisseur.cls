VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fournisseur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Private Const OBJ_NAME As String = "Fournisseur"
Public Function GetRow_Fournisseur_ByCode(ByRef ErrNumber As Long, _
                                                ByRef ErrDescription As String, _
                                                ByRef ErrSourceDetail As String, _
                                                ByVal vcode As String, _
                                                ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * From Station where ACTIF=1 and Code =" & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_Fournisseur_ByCode = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_Fournisseur_ACHAT_ByCode(ByRef ErrNumber As Long, _
                                                ByRef ErrDescription As String, _
                                                ByRef ErrSourceDetail As String, _
                                                ByVal vcode As String, _
                                                ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * From Station where ACTIF=1 and TYPE='FOURNISSEUR ACHAT' and Code =" & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_Fournisseur_ACHAT_ByCode = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_Fournisseur_ACHAT_ByLibelle(ByRef ErrNumber As Long, _
                                                    ByRef ErrDescription As String, _
                                                    ByRef ErrSourceDetail As String, _
                                                    ByVal CNB As ADODB.Connection, _
                                                    ByVal vLibelle As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetRow_Fournisseur_ACHAT_ByLibelle"
    Dim SQL As String
    Dim LRs_Record As New Recordset
On Error GoTo ErrHandler

    SQL = "SELECT * FROM STATION  WHERE ACTIF=1 and TYPE='FOURNISSEUR ACHAT' and Libelle like " & SQLText(vLibelle & "%") & " ORDER BY LIBELLE"
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetRow_Fournisseur_ACHAT_ByLibelle = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetAll_Fournisseur_ACHAT(ByRef ErrNumber As Long, _
                                                    ByRef ErrDescription As String, _
                                                    ByRef ErrSourceDetail As String, _
                                                    ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetAll_Fournisseur_ACHAT"
    Dim SQL As String
    Dim LRs_Record As New Recordset
On Error GoTo ErrHandler

    SQL = "SELECT * FROM STATION  WHERE ACTIF=1 and TYPE='FOURNISSEUR ACHAT' ORDER BY LIBELLE"
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetAll_Fournisseur_ACHAT = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


















Public Function Get_FournisAchat(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_FournisAchat"
    Dim SQL As String
    Dim LRs_Record As New Recordset
On Error GoTo ErrHandler
    SQL = "SELECT * FROM STATION  WHERE TYPE='FOURNISSEUR ACHAT' ORDER BY LIBELLE"
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_FournisAchat = LRs_Record
    Set LRs_Record = Nothing
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function GetRow_Fournisseurs_ByType(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal TYP As String, _
                                        ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from station where actif =1 and  Type = " & SQLText(TYP) & " Order by cast(Code as int)"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_Fournisseurs_ByType = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function





