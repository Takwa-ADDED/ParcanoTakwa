VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Personnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "Personnel"

'Verifier les droits d'accès de l'utilisateur
Public Function Verif_USER_Access(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal access As String, _
              ByVal UserId As Long) As Boolean

    Const sPROC_NAME As String = OBJ_NAME & ".Verif_USER_Access"

    Dim StrSql As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
' access = MAJ ou Supp ou INS_BV
    StrSql = "Select * from utilisateur where " & access & " = 1 and code= " & UserId
 
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    
    If Not LRs_Record.EOF Then
         Verif_USER_Access = True
    Else
         Verif_USER_Access = False
         LRs_Record.Close
    End If
    
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne le code du conducteur (de la table Personnel) à partir de son nom(libelle)
Public Function GetCODE_CONDUCTEUR(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal txt As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetCODE_CONDUCTEUR"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "Select Code from personnel where libelle = " & SQLText(txt)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set GetCODE_CONDUCTEUR = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_CONDUCTEUR(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal txt As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_CONDUCTEUR"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "Select * from personnel where code = " & SQLText(txt) & " Or Libelle = " & SQLText(txt)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_CONDUCTEUR = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne la liste de tout les personnels actifs
Public Function Get_AllActifPers(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AllActifPers"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Personnel where  actif=1 and Supp='N' order by libelle"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_AllActifPers = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Liste de tout les personnels
Public Function Get_AllPers(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AllPers"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Personnel order by libelle"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_AllPers = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Liste de tout les personnels
Public Function Get_AllUsers(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AllUsers"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Utilisateur order by NOMPRN"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_AllUsers = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function



'Liste de tout les utilisateur actif
Public Function Get_AllActifUsers(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AllActifUsers"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Utilisateur where Actif =1 order by NOMPRN"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_AllActifUsers = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne l'abreviation du nom(Libelle) du personnel (de la table Personnel)
Public Function Get_Abrev(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal txt As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_Abrev"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "Select ABREVIATION from personnel where libelle = " & SQLText(txt)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_Abrev = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne nom du personnel par son abreviation
Public Function Get_LibAbrev(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal txt As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_LibAbrev"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "Select libelle from personnel where ABREVIATION  = " & SQLText(txt)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_LibAbrev = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

