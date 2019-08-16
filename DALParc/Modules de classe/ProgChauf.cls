VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProgChauf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Private Const OBJ_NAME As String = "MC_ProgChauf"

Public Sub Save_AssProgCH(ByRef ErrNumber As Long, _
                    ByRef ErrDescription As String, _
                    ByRef ErrSourceDetail As String, _
                    ByVal CNB As ADODB.Connection, _
                    ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Save_AssProgCH"
    Dim SQL As String
    
On Error GoTo ErrHandler

    SQL = "INSERT INTO ASS_PROGCHAUF (CODE,CODECONDUCTEUR,CODEVEHICULE,DATECREATION,DATEPROGRAMME,USERINSERT) VALUES ("
    SQL = SQL & SQLText(PRs_Record!code)
    SQL = SQL & "," & SQLText(PRs_Record!CodeConducteur)
    SQL = SQL & "," & SQLText(PRs_Record!CODEVEHICULE)
    SQL = SQL & "," & SQLText(PRs_Record!DATECREATION)
    SQL = SQL & "," & SQLText(PRs_Record!DATEPROGRAMME)
    SQL = SQL & "," & SQLText(PRs_Record!USERINSERT)
    SQL = SQL & ")"
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub Update_Ass_ProgChauf(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal CNB As ADODB.Connection, _
                                ByVal PRs_Record As ADODB.Recordset)
    
    Dim SQL             As String
    
On Error GoTo ErrHandler
    
    SQL = "UPDATE ASS_PROGCHAUF SET "
    SQL = SQL & " DATEPROGRAMME=" & SQLText(PRs_Record!DATEPROGRAMME)
    SQL = SQL & " , USERUPDATE=" & SQLText(PRs_Record!USERUPDATE)
    SQL = SQL & " , USERADD=" & SQLText(PRs_Record!USERUPDATE)
    SQL = SQL & " , CODECONDUCTEUR = " & SQLText(PRs_Record!CodeConducteur)
    SQL = SQL & " , CODEVEHICULE = " & SQLText(PRs_Record!CODEVEHICULE)
    SQL = SQL & " WHERE CODE =" & SQLText(PRs_Record!code)
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub
Public Sub Save_DetProgCH(ByRef ErrNumber As Long, _
                    ByRef ErrDescription As String, _
                    ByRef ErrSourceDetail As String, _
                    ByVal CNB As ADODB.Connection, _
                    ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Save_DetProgCH"

    Dim SQL As String

    On Error GoTo ErrHandler
    
    PRs_Record.MoveFirst
    While Not PRs_Record.EOF
        SQL = "Insert Into Det_ProgChauf (CodeProgChauf,CodeFournisseur,TxtCommande,TxtPaiement,TxtObservation,ProgOrder) Values ("
        SQL = SQL & SQLText(PRs_Record!CodeProgChauf)
        SQL = SQL & "," & SQLText(PRs_Record!CodeFournisseur)
        SQL = SQL & "," & SQLText(PRs_Record!TxtCommande)
        SQL = SQL & "," & SQLText(PRs_Record!TxtPaiement)
        SQL = SQL & "," & SQLText(PRs_Record!TxtObservation)
        SQL = SQL & "," & SQLText(PRs_Record!ProgOrder)
        SQL = SQL & ")"
        CNB.Execute SQL
        PRs_Record.MoveNext
    Wend

Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub Delete_Det_Chauf(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal CodeDetailProg As String, _
                                ByVal CNB As ADODB.Connection)
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "DELETE FROM DET_PROGCHAUF WHERE CODEPROGCHAUF =" & SQLText(CodeDetailProg)
    CNB.Execute SQL

Exit Sub
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub
Public Sub Delete_Restaurer(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal CNB As ADODB.Connection, _
                                ByVal CodeProg As String, _
                                ByVal vSupp As String, _
                                ByVal CodeUser As String)
    
    Dim SQL As String
    
On Error GoTo ErrHandler
    
    If vSupp <> "N" And vSupp <> "O" Then
        MsgBox "VARIABLE SUPP INVALIDE", vbExclamation
        Exit Sub
    End If
    SQL = "UPDATE ASS_PROGCHAUF SET SUPP=" & SQLText(vSupp)
    If vSupp = "O" Then
        SQL = SQL & " , USERDELETE=" & SQLText(CodeUser)
    ElseIf vSupp = "N" Then
        SQL = SQL & " , USERADD=" & SQLText(CodeUser)
    End If
    SQL = SQL & " WHERE CODE =" & SQLText(CodeProg)
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub
Public Function GetRow_Ass_ProgramCH(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal CodeProg As String, _
                                ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT A.CODE, A.DATECREATION, A.DATEPROGRAMME, A.SUPP,P.CODE AS CODEPERSONNE, P.LIBELLE,V.CODE AS CODEVEHICULE, V.MATRICULE"
    SQL = SQL & " FROM ASS_PROGCHAUF A"
    SQL = SQL & " INNER JOIN PERSONNEL P ON A.CODECONDUCTEUR  = P.CODE"
    SQL = SQL & " INNER JOIN VEHICULE V ON A.CODEVEHICULE = V.CODE"
    SQL = SQL & " WHERE A.CODE = " & SQLText(CodeProg)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_Ass_ProgramCH = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_DetailsProgramCH(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal CodeProg As String, _
                                ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT D.CODE, D.PROGORDER, S.LIBELLE, D.TXTCOMMANDE, D.TXTPAIEMENT, D.TXTOBSERVATION, S.CODE AS CODEFR, A.DATEPROGRAMME"
    SQL = SQL & " FROM ASS_PROGCHAUF A"
    SQL = SQL & " INNER JOIN DET_PROGCHAUF D ON A.CODE=D.CODEPROGCHAUF"
    SQL = SQL & " INNER JOIN STATION S ON D.CODEFOURNISSEUR = S.CODE"
    SQL = SQL & " WHERE A.CODE = " & SQLText(CodeProg)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_DetailsProgramCH = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_ProgramChauffeur(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal Ddebut As String, _
                                ByVal Dfin As String, _
                                ByVal Param As Integer, _
                                ByVal ViewSupp As String, _
                                ByVal Cond As String, _
                                ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT  A.CODE, A.DATECREATION, A.DATEPROGRAMME,P.CODE AS CD_P, A.SUPP, P.LIBELLE AS CONDUCTEUR, V.MATRICULE ,"
    SQL = SQL & " D.PROGORDER , S.LIBELLE AS FOURNISSEUR, D.TXTCOMMANDE, D.TXTPAIEMENT, D.TXTOBSERVATION"
    SQL = SQL & " FROM  ASS_PROGCHAUF A"
    SQL = SQL & " INNER JOIN PERSONNEL P ON A.CODECONDUCTEUR  = P.CODE"
    SQL = SQL & " INNER JOIN VEHICULE V ON A.CODEVEHICULE = V.CODE"
    SQL = SQL & " INNER JOIN DET_PROGCHAUF D ON A.CODE = D.CODEPROGCHAUF"
    SQL = SQL & " INNER JOIN STATION S ON D.CODEFOURNISSEUR = S.CODE"
    If Param = 4 Then
        SQL = SQL & " WHERE A.SUPP='N' AND A.DATEPROGRAMME >= " & SQLText(Date)
        If Cond <> "0000" And Cond <> "" Then SQL = SQL & " AND A.CODECONDUCTEUR = " & SQLText(Cond)
    End If
    If Param = 0 Then
        SQL = SQL & " WHERE A.DATEPROGRAMME BETWEEN CAST('" & Ddebut & "' AS DATETIME) AND CAST('" & Dfin & "' AS DATETIME)"
        If ViewSupp = "N" Then SQL = SQL & " AND A.SUPP='N'"
        If Cond <> "0000" And Cond <> "" Then SQL = SQL & " AND A.CODECONDUCTEUR = " & SQLText(Cond)
    End If
    SQL = SQL & " ORDER BY A.CODE , D.PROGORDER"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_ProgramChauffeur = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
