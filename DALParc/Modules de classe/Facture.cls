VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Facture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "Facture"

'Retourne détails d'une facture par son code
Public Function Get_FactByCode(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset
              
Const sPROC_NAME As String = OBJ_NAME & ".Get_FactByCode"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    SQL = "Select * from FactureCarburant where Numero = " & SQLText(vcode)
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_FactByCode = LRs_Record
    Set LRs_Record = Nothing

Exit Function

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Function

Public Sub Delete_Fact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Fact"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Delete from FactureCarburant where Numero =" & SQLText(vcode)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Function Get_FactFind(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset
              
Const sPROC_NAME As String = OBJ_NAME & ".Get_FactFind"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    SQL = "SELECT FactureCarburant.*,station.libelle from FactureCarburant ,station"
    SQL = SQL & " where station.code=FactureCarburant.station"
    SQL = SQL & " Order by Numero "
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_FactFind = LRs_Record
    Set LRs_Record = Nothing

Exit Function

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Function

Public Sub Insert_Fact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_Fact"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
    StrSql = "Insert into FactureCarburant(Numero,datedoc,Station,Periodedu,periodeAu,ttc_bc,ttc_bv,ttc_pr,ttc_BR,ttc,nbc,dateOP,Timbre, UserInsert) values ("
    
    If IsNull(PRs_Record!numero) Then
        StrSql = StrSql & "Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!numero)
    End If
    
     If IsNull(PRs_Record!datedoc) Then
        StrSql = StrSql & ", Null"
    Else
        StrSql = StrSql & ", " & SQLText(PRs_Record!datedoc)
    End If
    
    If IsNull(PRs_Record!Station) Then
        StrSql = StrSql & "," & "NULL"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Station)
    End If
    
    If IsNull(PRs_Record!Periodedu) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Periodedu)
    End If
    
    If IsNull(PRs_Record!periodeAu) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!periodeAu)
    End If
    
    If IsNull(PRs_Record!ttc_bc) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!ttc_bc), ",", ".")
    End If
    
    If IsNull(PRs_Record!ttc_bv) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!ttc_bv), ",", ".")
    End If
    
    If IsNull(PRs_Record!ttc_pr) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!ttc_pr), ",", ".")
    End If
    
    If IsNull(PRs_Record!ttc_BR) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!ttc_BR), ",", ".")
    End If
    
    If IsNull(PRs_Record!ttc) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!ttc), ",", ".")
    End If
    
    If IsNull(PRs_Record!NBC) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Val(PRs_Record!NBC)
    End If
    
    If IsNull(PRs_Record!dateOP) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!dateOP)
    End If
    
    If IsNull(PRs_Record!Timbre) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!Timbre), ",", ".")
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

Public Sub Update_Fact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_Fact"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update FactureCarburant set"
  
    If IsNull(PRs_Record!Periodedu) Then
        StrSql = StrSql & " Periodedu = Null"
    Else
        StrSql = StrSql & " Periodedu = " & SQLText(PRs_Record!Periodedu)
    End If
    
    If IsNull(PRs_Record!periodeAu) Then
        StrSql = StrSql & ", periodeAu = Null"
    Else
        StrSql = StrSql & ", periodeAu = " & SQLText(PRs_Record!periodeAu)
    End If
    
    If IsNull(PRs_Record!ttc_bc) Then
        StrSql = StrSql & ", ttc_bc = 0"
    Else
        StrSql = StrSql & ", ttc_bc = " & Replace((PRs_Record!ttc_bc), ",", ".")
    End If
    
    If IsNull(PRs_Record!ttc_bv) Then
        StrSql = StrSql & " , ttc_bv = 0"
    Else
        StrSql = StrSql & " , ttc_bv = " & Replace((PRs_Record!ttc_bv), ",", ".")
    End If
    
    If IsNull(PRs_Record!ttc_pr) Then
        StrSql = StrSql & " , ttc_pr = 0"
    Else
        StrSql = StrSql & " , ttc_pr = " & Replace((PRs_Record!ttc_pr), ",", ".")
    End If
    
    If IsNull(PRs_Record!ttc_BR) Then
        StrSql = StrSql & " , ttc_BR = 0"
    Else
        StrSql = StrSql & " , ttc_BR = " & Replace((PRs_Record!ttc_BR), ",", ".")
    End If
    
    If IsNull(PRs_Record!ttc) Then
        StrSql = StrSql & " , ttc = 0"
    Else
        StrSql = StrSql & " , ttc = " & Replace((PRs_Record!ttc), ",", ".")
    End If
   
    If IsNull(PRs_Record!NBC) Then
        StrSql = StrSql & " , nbc = 0"
    Else
        StrSql = StrSql & " , nbc = " & Val(PRs_Record!NBC)
    End If
    
    If IsNull(PRs_Record!dateOP) Then
        StrSql = StrSql & " , dateOP = Null"
    Else
        StrSql = StrSql & " , dateOP = " & SQLText(PRs_Record!dateOP)
    End If


    If IsNull(PRs_Record!Timbre) Then
        StrSql = StrSql & " , Timbre = 0"
    Else
        StrSql = StrSql & " , Timbre = " & Replace((PRs_Record!Timbre), ",", ".")
    End If
    
    If IsNull(PRs_Record!userupdate) Then
        StrSql = StrSql & " , UserUpdate = Null"
    Else
        StrSql = StrSql & " , UserUpdate = " & SQLText(PRs_Record!userupdate)
    End If

    StrSql = StrSql & " Where Numero = " & SQLText(PRs_Record!numero)
    
 CNB.Execute StrSql

Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Charger détails des bons pour la création de la facture
Public Function Get_Details_PourCreation(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vCodeSation As String, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset
              
Const sPROC_NAME As String = OBJ_NAME & ".Get_Details_PourCreation"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    
    'Bon carburant
SQL = "Select DISTINCT 'BC' as TYPE ,Numero,dateop, Valeur"
SQL = SQL & " From assBonCarburant"
SQL = SQL & " where transf='N' and Supp='N' and assBonCarburant.dateop between " & SQLText(VdateD)
SQL = SQL & " and " & SQLText(vDateF) & " and Station=" & SQLText(vCodeSation)
SQL = SQL & " Union All"
'Bon Vidange
SQL = SQL & " Select 'BV' as TYPE ,Numero,dateop, Valeur from Ass_BonVidange "
SQL = SQL & " where transf='N' and Supp='N' and Ass_BonVidange.dateop between " & SQLText(VdateD) & " and " & SQLText(vDateF) & " and Station=" & SQLText(vCodeSation)
SQL = SQL & " Union All"
'Piece Reparation
SQL = SQL & " Select  'PR' as TYPE ,Numero,dateOperation,totTTC AS Valeur from AssPieceReparation "
SQL = SQL & " where AssPieceReparation.Type in ('Piece Reception') And transf='N' and Supp='N' and AssPieceReparation.dateOperation between " & SQLText(VdateD) & " and " & SQLText(vDateF) & " and Fournisseur=" & SQLText(vCodeSation)
SQL = SQL & " Union All"
'Bon Retour
SQL = SQL & " Select  'BR' as TYPE ,Numero,dateOperation,totTTC AS Valeur from AssPieceReparation"
SQL = SQL & " where AssPieceReparation.Type in ('Bon Retour') And transf='N' and Supp='N' and AssPieceReparation.dateOperation between " & SQLText(VdateD) & " and " & SQLText(vDateF) & " and Fournisseur=" & SQLText(vCodeSation)
SQL = SQL & " order by TYPE,Numero"
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_Details_PourCreation = LRs_Record
    Set LRs_Record = Nothing

Exit Function

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Function

'Charger détails des bons pour afficher la facture
Public Function Get_Details_PourAffich(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vNumero As String) As Recordset
              
Const sPROC_NAME As String = OBJ_NAME & ".Get_Details_PourAffich"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    
'Bon carburant
SQL = "Select DISTINCT 'BC' as TYPE ,Numero,dateop, Valeur"
SQL = SQL & " From AssBonCarburant"
SQL = SQL & " Where transf='O' and numfact=" & SQLText(vNumero)
SQL = SQL & " Union All"
'Bon Vidange
SQL = SQL & " Select 'BV' as TYPE ,Numero,dateop, Valeur from Ass_BonVidange"
SQL = SQL & " where transf='O'  and numfact=" & SQLText(vNumero)
SQL = SQL & " Union All"
'Pièce Reparation
SQL = SQL & " Select  'PR' as TYPE ,Numero,dateOperation,totTTC AS Valeur from AssPieceReparation"
SQL = SQL & " where AssPieceReparation.Type='Piece Reception' And transf='O' and numfact=" & SQLText(vNumero)

'Bon Retour
SQL = SQL & " Union All"
SQL = SQL & " Select  'BR' as TYPE ,Numero,dateOperation,totTTC AS Valeur from AssPieceReparation"
SQL = SQL & " where AssPieceReparation.Type ='Bon Retour' And transf='O' and numfact=" & SQLText(vNumero)
SQL = SQL & "order by TYPE, Numero"
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_Details_PourAffich = LRs_Record
    Set LRs_Record = Nothing

Exit Function

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Function
