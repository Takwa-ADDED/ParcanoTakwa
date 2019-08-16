VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBConnexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Private Const OBJ_NAME As String = "DBConnexion"
Public Sub Disconnect(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByRef Cn As ADODB.Connection)

    Const sPROC_NAME As String = OBJ_NAME & ".Disconnect"
On Error GoTo ErrHandler
    Cn.Close
    Set Cn = Nothing
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub Connect(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByRef CNB As ADODB.Connection, _
              ByVal P_StrServer As String, _
              ByVal P_StrDataBase As String, _
              ByVal P_StrUserId As String, _
              ByVal P_StrPassword As String, _
              ByVal P_StrServer2 As String, _
              ByVal P_StrPassword2 As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Connect"
On Error GoTo ErrHandler
    Set CNB = New Connection
    Set CNR = New Connection
    With CNB
    On Error Resume Next
        .ConnectionString = "Provider = SQLOLEDB;user id =" & P_StrUserId & ";Password= " & P_StrPassword & "; Server=" & P_StrServer & "; Database=" & P_StrDataBase & "; Network=dbmssocn"
        CNR.ConnectionString = .ConnectionString
        .CommandTimeout = 500
        .ConnectionTimeout = 10
        .Open
        If (CNB.Errors.Count <> 0) Then
            If MsgBox("Erreur Serveur, voulez vous basculer vers le serveur secondaire ?", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
                On Error GoTo ErrHandler
                .ConnectionString = "Provider = SQLOLEDB;user id =" & P_StrUserId & ";Password= " & P_StrPassword2 & "; Server=" & P_StrServer2 & "; Database=" & P_StrDataBase & "; Network=dbmssocn"
                CNR.ConnectionString = .ConnectionString
                .CommandTimeout = 500
                .ConnectionTimeout = 500
                .Open
            End If
        End If
    End With
    
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
