VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Utilisateur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Private Const OBJ_NAME As String = "Utilisateur"
    
'==========================================================================
'Verifier les droits d'accès de l'utilisateur
'==========================================================================
Public Function USER_ACCESS(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal access As String, _
              ByVal UserId As Long) As Recordset

    Const sPROC_NAME    As String = OBJ_NAME & ".USER_ACCESS"
    Dim SQL             As String
    Dim LRs_Record      As New Recordset

On Error GoTo ErrHandler

    SQL = "Select * from utilisateur where actif = 1 and " & access & " = 1 and code= " & UserId
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set USER_ACCESS = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
'==========================================================================
'Liste de tout les utilisateur actif
'==========================================================================
Public Function GetAll_UsersActif(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetAll_UsersActif"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Utilisateur where Actif =1 order by NOMPRN"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set GetAll_UsersActif = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
'==========================================================================
'Connexion
'==========================================================================
Public Function GetRow_UsersByLoginPwd(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal Pwd As String, _
                                    ByVal Login As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetRow_UsersByLoginPwd"
    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Utilisateur where MP like " & SQLText(Pwd) & " and NOMPRN like " & SQLText(Login)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set GetRow_UsersByLoginPwd = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
'Vérifier Mot de passe***
Public Function GetRow_UsersByPwd(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal vcode As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetRow_UsersByPwd"
    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "select * from Utilisateur where MP like " & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set GetRow_UsersByPwd = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_UserByCode(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal CodeUser As String, _
                                ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetRow_UserByCode"
    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler
 '" & CodeConducteur & "%'"

    SQL = "Select * from utilisateur where CODE = " & SQLText(CodeUser) & " OR NomPrn Like '" & CodeUser & "%'"


    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetRow_UserByCode = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Sub Delete_Add_USER(ByRef ErrNumber As Long, _
                        ByRef ErrDescription As String, _
                        ByRef ErrSourceDetail As String, _
                        ByVal code As String, _
                        ByVal vSupp As String, _
                        ByVal CodeUser As String, _
                        ByVal CNB As ADODB.Connection)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Add_USER"

    Dim SQL As String
    
    If vSupp <> "N" And vSupp <> "O" Then
        MsgBox "variable supp invalide", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    SQL = "Update Utilisateur set "
    SQL = SQL & " Supp=" & SQLText(vSupp)
    If vSupp = "O" Then
        SQL = SQL & " , UserDelete=" & SQLText(CodeUser)
    ElseIf vSupp = "N" Then
        SQL = SQL & " , UserAdd=" & SQLText(CodeUser)
    End If
    SQL = SQL & " where Code =" & SQLText(code)
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub UpDate_USER(ByRef ErrNumber As Long, _
                            ByRef ErrDescription As String, _
                            ByRef ErrSourceDetail As String, _
                            ByVal CNB As ADODB.Connection, _
                            ByVal PRs_Record As ADODB.Recordset, _
                            ByVal code As String)

    Const sPROC_NAME As String = OBJ_NAME & ".UpDate_USER"
    Dim SQL As String
    
On Error GoTo ErrHandler

    SQL = "update Utilisateur set"
    SQL = SQL & " MP =" & SQLText(PRs_Record!MP)
    SQL = SQL & ", NomPrn =" & SQLText(PRs_Record!NomPrn)
    SQL = SQL & ", Ins_BC=" & PRs_Record!Ins_BC
    SQL = SQL & ", Maj_BC=" & PRs_Record!Maj_BC
    SQL = SQL & ", Supp_BC=" & PRs_Record!Supp_BC
    SQL = SQL & ", Consult_BC=" & PRs_Record!Consult_BC
    SQL = SQL & ", Ins_BV=" & PRs_Record!Ins_BV
    SQL = SQL & ", Maj_BV=" & PRs_Record!Maj_BV
    SQL = SQL & ", Supp_BV=" & PRs_Record!Supp_BV
    SQL = SQL & ", Consult_BV=" & PRs_Record!Consult_BV
    SQL = SQL & ", Consult_Alerte=" & PRs_Record!Consult_Alerte
    SQL = SQL & ", Ins_BCR=" & PRs_Record!Ins_BCR
    SQL = SQL & ", Maj_BCR=" & PRs_Record!Maj_BCR
    SQL = SQL & ", Supp_BCR=" & PRs_Record!Supp_BCR
    SQL = SQL & ", Consult_BCR=" & PRs_Record!Consult_BCR
    SQL = SQL & ", Ins_PR=" & PRs_Record!Ins_PR
    SQL = SQL & ", Maj_PR=" & PRs_Record!Maj_PR
    SQL = SQL & ", Supp_PR=" & PRs_Record!Supp_PR
    SQL = SQL & ", Consult_PR=" & PRs_Record!Consult_PR
    SQL = SQL & ", Ins_FF=" & PRs_Record!Ins_FF
    SQL = SQL & ", Maj_FF=" & PRs_Record!Maj_FF
    SQL = SQL & ", Supp_FF=" & PRs_Record!Supp_FF
    SQL = SQL & ", Consult_FF=" & PRs_Record!Consult_FF
    SQL = SQL & ", Consult_SC=" & PRs_Record!Consult_SC
    SQL = SQL & ", Ins_FT=" & PRs_Record!Ins_FT
    SQL = SQL & ", Maj_FT=" & PRs_Record!Maj_FT
    SQL = SQL & ", Supp_FT=" & PRs_Record!Supp_FT
    SQL = SQL & ", Consult_FT=" & PRs_Record!Consult_FT
    SQL = SQL & ", Consult_Sup=" & PRs_Record!Consult_Sup
    SQL = SQL & ", Ins_Vehicule=" & PRs_Record!Ins_Vehicule
    SQL = SQL & ", Maj_vehicule=" & PRs_Record!Maj_vehicule
    SQL = SQL & ", Supp_vehicule=" & PRs_Record!Supp_vehicule
    SQL = SQL & ", Consult_vehicule=" & PRs_Record!Consult_vehicule
    SQL = SQL & ", Ins_Fournisseur=" & PRs_Record!Ins_Fournisseur
    SQL = SQL & ", Maj_Fournisseur=" & PRs_Record!Maj_Fournisseur
    SQL = SQL & ", Supp_Fournisseur=" & PRs_Record!Supp_Fournisseur
    SQL = SQL & ", Conslt_Fournisseur=" & PRs_Record!Conslt_Fournisseur
    SQL = SQL & ", Ins_TC=" & PRs_Record!Ins_TC
    SQL = SQL & ", Maj_TC=" & PRs_Record!Maj_TC
    SQL = SQL & ", Supp_TC=" & PRs_Record!Supp_TC
    SQL = SQL & ", Consult_TC=" & PRs_Record!Consult_TC
    SQL = SQL & ", Ins_TV=" & PRs_Record!Ins_TV
    SQL = SQL & ", Maj_TV=" & PRs_Record!Maj_TV
    SQL = SQL & ", supp_TV=" & PRs_Record!supp_TV
    SQL = SQL & ", Consult_TV=" & PRs_Record!Consult_TV
    SQL = SQL & ", Ins_Destination=" & PRs_Record!Ins_Destination
    SQL = SQL & ", Maj_Destination=" & PRs_Record!Maj_Destination
    SQL = SQL & ", Supp_Destination=" & PRs_Record!Supp_Destination
    SQL = SQL & ", Consult_Destination=" & PRs_Record!Consult_Destination
    SQL = SQL & ", Ins_Produit=" & PRs_Record!Ins_Produit
    SQL = SQL & ", Maj_produit=" & PRs_Record!Maj_produit
    SQL = SQL & ", Supp_Produit=" & PRs_Record!Supp_Produit
    SQL = SQL & ", Consult_Produit=" & PRs_Record!Consult_Produit
    SQL = SQL & ", Ins_Personnel=" & PRs_Record!Ins_Personnel
    SQL = SQL & ", Maj_Personnel=" & PRs_Record!Maj_Personnel
    SQL = SQL & ", Supp_personnel=" & PRs_Record!Supp_personnel
    SQL = SQL & ", Consult_personnel=" & PRs_Record!Consult_personnel
    SQL = SQL & ", Ins_Utilisateur=" & PRs_Record!Ins_Utilisateur
    SQL = SQL & ", Maj_Utilisateur=" & PRs_Record!Maj_Utilisateur
    SQL = SQL & ", Supp_Utilisateur=" & PRs_Record!Supp_Utilisateur
    SQL = SQL & ", Consult_Utilisateur=" & PRs_Record!Consult_Utilisateur
    SQL = SQL & ", Actif=" & PRs_Record!Actif
    SQL = SQL & ", Maj_Disp=" & PRs_Record!Maj_Disp
    SQL = SQL & ", Maj_Compt=" & PRs_Record!Maj_Compt
    SQL = SQL & ", Consult_Compteurs=" & PRs_Record!Consult_Compteurs
    SQL = SQL & ", Ins_PCH=" & PRs_Record!Ins_PCH
    SQL = SQL & ", Maj_PCH=" & PRs_Record!Maj_PCH
    SQL = SQL & ", Supp_PCH=" & PRs_Record!Supp_PCH
    SQL = SQL & ", Consult_PCH=" & PRs_Record!Consult_PCH
    SQL = SQL & ", Ins_PLING=" & PRs_Record!Ins_PLING
    SQL = SQL & ", Maj_PLING=" & PRs_Record!Maj_PLING
    SQL = SQL & ", Supp_PLING=" & PRs_Record!Supp_PLING
    SQL = SQL & ", Consult_PLING=" & PRs_Record!Consult_PLING
    SQL = SQL & ", Ins_Conge=" & PRs_Record!Ins_Conge
    SQL = SQL & ", Maj_Conge=" & PRs_Record!Maj_Conge
    SQL = SQL & ", Supp_Conge=" & PRs_Record!Supp_Conge
    SQL = SQL & ", Consult_Conge=" & PRs_Record!Consult_Conge
    SQL = SQL & ", UserUpdate = " & SQLText(PRs_Record!USERUPDATE)
    SQL = SQL & " Where Code=" & code
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub
Public Sub Save_USER(ByRef ErrNumber As Long, _
                        ByRef ErrDescription As String, _
                        ByRef ErrSourceDetail As String, _
                        ByVal CNB As ADODB.Connection, _
                        ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Save_USER"
    Dim SQL As String
On Error GoTo ErrHandler
    SQL = "Insert into Utilisateur  ("
    SQL = SQL & " Code,MP,NomPrn,Ins_BC , Maj_BC, Supp_BC, Consult_BC, Ins_BV, Maj_BV, Supp_BV, Consult_BV,"
    SQL = SQL & " Consult_Alerte , Ins_BCR, Maj_BCR, Supp_BCR, Consult_BCR, Ins_PR,"
    SQL = SQL & " Maj_PR , Supp_PR, Consult_PR, Ins_FF, Maj_FF, Supp_FF, Consult_FF, Consult_SC,"
    SQL = SQL & " Ins_FT, Maj_FT, Supp_FT, Consult_FT,"
    SQL = SQL & " Consult_Sup , Ins_Vehicule, Maj_vehicule, Supp_vehicule, Consult_vehicule,"
    SQL = SQL & " Ins_Fournisseur , Maj_Fournisseur, Supp_Fournisseur, Conslt_Fournisseur,"
    SQL = SQL & " Ins_TC , Maj_TC, Supp_TC, Consult_TC, Ins_TV, Maj_TV, supp_TV,"
    SQL = SQL & " Consult_TV , Ins_Destination, Maj_Destination, Supp_Destination,"
    SQL = SQL & " Consult_Destination ,"
    SQL = SQL & " Ins_Produit , Maj_produit, Supp_Produit, Consult_Produit, Ins_Personnel,"
    SQL = SQL & " Maj_Personnel , Supp_personnel, Consult_personnel, Ins_Utilisateur,"
    SQL = SQL & " Maj_Utilisateur , Supp_Utilisateur, Consult_Utilisateur, Actif, Maj_Disp, Maj_Compt, Consult_Compteurs,"
    SQL = SQL & " Ins_PCH,Maj_PCH , Supp_PCH, Consult_PCH, Ins_PLING,Maj_PLING , Supp_PLING, Consult_PLING,"
    SQL = SQL & " Ins_Conge, Maj_Conge , Supp_Conge, Consult_Conge, UserInsert"
    SQL = SQL & " )Values ("
    
    SQL = SQL & SQLText(PRs_Record!code)
    SQL = SQL & "," & SQLText(PRs_Record!MP)
    SQL = SQL & "," & SQLText(PRs_Record!NomPrn)
    SQL = SQL & "," & PRs_Record!Ins_BC
    SQL = SQL & "," & PRs_Record!Maj_BC
    SQL = SQL & "," & PRs_Record!Supp_BC
    SQL = SQL & "," & PRs_Record!Consult_BC
    SQL = SQL & "," & PRs_Record!Ins_BV
    SQL = SQL & "," & PRs_Record!Maj_BV
    SQL = SQL & "," & PRs_Record!Supp_BV
    SQL = SQL & "," & PRs_Record!Consult_BV
    SQL = SQL & "," & PRs_Record!Consult_Alerte
    SQL = SQL & "," & PRs_Record!Ins_BCR
    SQL = SQL & "," & PRs_Record!Maj_BCR
    SQL = SQL & "," & PRs_Record!Supp_BCR
    SQL = SQL & "," & PRs_Record!Consult_BCR
    SQL = SQL & "," & PRs_Record!Ins_PR
    SQL = SQL & "," & PRs_Record!Maj_PR
    SQL = SQL & "," & PRs_Record!Supp_PR
    SQL = SQL & "," & PRs_Record!Consult_PR
    SQL = SQL & "," & PRs_Record!Ins_FF
    SQL = SQL & "," & PRs_Record!Maj_FF
    SQL = SQL & "," & PRs_Record!Supp_FF
    SQL = SQL & "," & PRs_Record!Consult_FF
    SQL = SQL & "," & PRs_Record!Consult_SC
    SQL = SQL & "," & PRs_Record!Ins_FT
    SQL = SQL & "," & PRs_Record!Maj_FT
    SQL = SQL & "," & PRs_Record!Supp_FT
    SQL = SQL & "," & PRs_Record!Consult_FT
    SQL = SQL & "," & PRs_Record!Consult_Sup
    SQL = SQL & "," & PRs_Record!Ins_Vehicule
    SQL = SQL & "," & PRs_Record!Maj_vehicule
    SQL = SQL & "," & PRs_Record!Supp_vehicule
    SQL = SQL & "," & PRs_Record!Consult_vehicule
    SQL = SQL & "," & PRs_Record!Ins_Fournisseur
    SQL = SQL & "," & PRs_Record!Maj_Fournisseur
    SQL = SQL & "," & PRs_Record!Supp_Fournisseur
    SQL = SQL & "," & PRs_Record!Conslt_Fournisseur
    SQL = SQL & "," & PRs_Record!Ins_TC
    SQL = SQL & "," & PRs_Record!Maj_TC
    SQL = SQL & "," & PRs_Record!Supp_TC
    SQL = SQL & "," & PRs_Record!Consult_TC
    SQL = SQL & "," & PRs_Record!Ins_TV
    SQL = SQL & "," & PRs_Record!Maj_TV
    SQL = SQL & "," & PRs_Record!supp_TV
    SQL = SQL & "," & PRs_Record!Consult_TV
    SQL = SQL & "," & PRs_Record!Ins_Destination
    SQL = SQL & "," & PRs_Record!Maj_Destination
    SQL = SQL & "," & PRs_Record!Supp_Destination
    SQL = SQL & "," & PRs_Record!Consult_Destination
    SQL = SQL & "," & PRs_Record!Ins_Produit
    SQL = SQL & "," & PRs_Record!Maj_produit
    SQL = SQL & "," & PRs_Record!Supp_Produit
    SQL = SQL & "," & PRs_Record!Consult_Produit
    SQL = SQL & "," & PRs_Record!Ins_Personnel
    SQL = SQL & "," & PRs_Record!Maj_Personnel
    SQL = SQL & "," & PRs_Record!Supp_personnel
    SQL = SQL & "," & PRs_Record!Consult_personnel
    SQL = SQL & "," & PRs_Record!Ins_Utilisateur
    SQL = SQL & "," & PRs_Record!Maj_Utilisateur
    SQL = SQL & "," & PRs_Record!Supp_Utilisateur
    SQL = SQL & "," & PRs_Record!Consult_Utilisateur
    SQL = SQL & "," & PRs_Record!Actif
    SQL = SQL & "," & PRs_Record!Maj_Disp
    SQL = SQL & "," & PRs_Record!Maj_Compt
    SQL = SQL & "," & PRs_Record!Consult_Compteurs
    SQL = SQL & "," & PRs_Record!Ins_PCH
    SQL = SQL & "," & PRs_Record!Maj_PCH
    SQL = SQL & "," & PRs_Record!Supp_PCH
    SQL = SQL & "," & PRs_Record!Consult_PCH
    SQL = SQL & "," & PRs_Record!Ins_PLING
    SQL = SQL & "," & PRs_Record!Maj_PLING
    SQL = SQL & "," & PRs_Record!Supp_PLING
    SQL = SQL & "," & PRs_Record!Consult_PLING
    SQL = SQL & "," & PRs_Record!Ins_Conge
    SQL = SQL & "," & PRs_Record!Maj_Conge
    SQL = SQL & "," & PRs_Record!Supp_Conge
    SQL = SQL & "," & PRs_Record!Consult_Conge
    SQL = SQL & "," & SQLText(PRs_Record!USERINSERT)
    SQL = SQL & ")"
    CNB.Execute SQL
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub



'=================================================
'Fiche Traffuic***
'=================================================
Public Function GetRow_User_Ins_FT(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CodeUser As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from utilisateur where Ins_FT = 1 and code= " & CodeUser

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_User_Ins_FT = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_User_Maj_FT(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CodeUser As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from utilisateur where supp_FT = 1 and code= " & CodeUser

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_User_Maj_FT = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_User_Consult_FT(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CodeUser As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from utilisateur where Consult_FT = 1 and code= " & CodeUser

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_User_Consult_FT = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_User_Maj_FT_ByPwd(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal vcode As String, _
                                        ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from utilisateur where Maj_FT = 1 and MP like " & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_User_Maj_FT_ByPwd = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_User_Supp_FT(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CodeUser As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from utilisateur where Maj_FT = 1 and code= " & CodeUser

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_User_Supp_FT = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
'Public Function GetRow_User_Consult_Supervision(ByRef ErrNumber As Long, _
'                                    ByRef ErrDescription As String, _
'                                    ByRef ErrSourceDetail As String, _
'                                    ByVal CodeUser As String, _
'                                    ByVal CNB As ADODB.Connection) As Recordset
'    Dim Sql As String
'    Dim LRs_Record As Recordset
'
'On Error GoTo ErrHandler
'
'    Sql = "Select * from utilisateur where Consult_Sup = 1 and code= " & CodeUser
'
'    Set LRs_Record = New Recordset
'    LRs_Record.CursorLocation = adUseClient
'    LRs_Record.Open Sql, CNB, adOpenStatic, adLockPessimistic
'    Set GetRow_User_Consult_Supervision = LRs_Record
'    Set LRs_Record = Nothing
'
'Exit Function
'ErrHandler:
'    Set LRs_Record = Nothing
'    ErrNumber = Err.Number
'    ErrDescription = Err.Description
'End Function
'=================================================
'Personnel***
'=================================================
'Public Function GetRow_User_Ins_Personnel(ByRef ErrNumber As Long, _
'                                    ByRef ErrDescription As String, _
'                                    ByRef ErrSourceDetail As String, _
'                                    ByVal CodeUser As String, _
'                                    ByVal CNB As ADODB.Connection) As Recordset
'    Dim SQL As String
'    Dim LRs_Record As Recordset
'
'On Error GoTo ErrHandler
'
'    SQL = "Select * from utilisateur where Ins_Personnel = 1 and code= " & CodeUser
'
'    Set LRs_Record = New Recordset
'    LRs_Record.CursorLocation = adUseClient
'    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
'    Set GetRow_User_Ins_Personnel = LRs_Record
'    Set LRs_Record = Nothing
'
'Exit Function
'ErrHandler:
'    Set LRs_Record = Nothing
'    ErrNumber = Err.Number
'    ErrDescription = Err.Description
'End Function
'Public Function GetRow_User_Supp_personnel(ByRef ErrNumber As Long, _
'                                    ByRef ErrDescription As String, _
'                                    ByRef ErrSourceDetail As String, _
'                                    ByVal CodeUser As String, _
'                                    ByVal CNB As ADODB.Connection) As Recordset
'    Dim SQL As String
'    Dim LRs_Record As Recordset
'
'On Error GoTo ErrHandler
'
'    SQL = "Select * from utilisateur where Supp_personnel = 1 and code= " & CodeUser
'
'    Set LRs_Record = New Recordset
'    LRs_Record.CursorLocation = adUseClient
'    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
'    Set GetRow_User_Supp_personnel = LRs_Record
'    Set LRs_Record = Nothing
'
'Exit Function
'ErrHandler:
'    Set LRs_Record = Nothing
'    ErrNumber = Err.Number
'    ErrDescription = Err.Description
'End Function
'Public Function GetRow_User_Maj_Personnel(ByRef ErrNumber As Long, _
'                                    ByRef ErrDescription As String, _
'                                    ByRef ErrSourceDetail As String, _
'                                    ByVal CodeUser As String, _
'                                    ByVal CNB As ADODB.Connection) As Recordset
'    Dim SQL As String
'    Dim LRs_Record As Recordset
'
'On Error GoTo ErrHandler
'
'    SQL = "Select * from utilisateur where Maj_Personnel = 1 and code= " & CodeUser
'
'    Set LRs_Record = New Recordset
'    LRs_Record.CursorLocation = adUseClient
'    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
'    Set GetRow_User_Maj_Personnel = LRs_Record
'    Set LRs_Record = Nothing
'
'Exit Function
'ErrHandler:
'    Set LRs_Record = Nothing
'    ErrNumber = Err.Number
'    ErrDescription = Err.Description
'End Function
'=================================================
'Fournisseur***
'=================================================
Public Function GetRow_User_Ins_Fournisseur(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CodeUser As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from utilisateur where Ins_Fournisseur = 1 and code= " & CodeUser

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_User_Ins_Fournisseur = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_User_Supp_Fournisseur(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CodeUser As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from utilisateur where Supp_Fournisseur = 1 and code= " & CodeUser

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_User_Supp_Fournisseur = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_User_Maj_Fournisseur(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CodeUser As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from utilisateur where Maj_Fournisseur = 1 and code= " & CodeUser

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_User_Maj_Fournisseur = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
'=================================================
'Disponibilité***
'=================================================
Public Function GetRow_User_MAJ_Disp(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CodeUser As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from utilisateur where MAJ_Disp = 1 and code= " & CodeUser

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_User_MAJ_Disp = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_User_Maj_Disp_ByPwd(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal vcode As String, _
                                        ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from utilisateur where Maj_Disp = 1 and MP like " & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_User_Maj_Disp_ByPwd = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


