VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PieceReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "PieceReparation"

'Retourne l'assiette d'une piece de reparation par code
Public Function Get_AssPieceReparation(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AssPieceReparation"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = "Select * from AssPieceReparation where Numero = " & SQLText(vcode)
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_AssPieceReparation = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne les détails d'une piece de reparation par code
Public Function Get_DetPieceReparation(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_DetPieceReparation"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = "Select * from DetailPieceReparation where Numero = " & SQLText(vcode)
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_DetPieceReparation = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


'Retourne le nombre max des pièces saissi
Public Function Get_MaxNum(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_MaxNum"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = "select Max(Numero) from AssPieceReparation "
    
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

'Suppression d'une piece par Supp='O'
Public Sub Delete_AssPRep(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String, _
              ByVal vuser As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_AssPRep"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update AssPieceReparation Set Supp = 'O' , UserDelete = " & SQLText(vuser) & " where Numero =" & SQLText(vcode)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Suppression des détails d'une piece par Supp='O'
Public Sub Delete_DetailPRepaBySup(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_DetailPRepaBySup"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update DetailPieceReparation Set Supp ='O' where Numero = " & SQLText(vcode)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

''Suppression complète des détails d'une piece
Public Sub Delete_DetailPRepa(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_DetailPRepa"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Delete from DetailPieceReparation where Numero = " & SQLText(vcode)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Retourne les assiettes des BC pour l'afficher dans Forme FrmFind suivant certains paramètres
Public Function Get_PieceReparation(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal DateDu As Date, _
              ByVal DateAu As Date, _
              ByVal Station As String, _
              ByVal TYP As String, _
              ByVal TypPiece As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_PieceReparation"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

SQL = "select Numero, DatePiece, dateOperation, A.Supp , A.transf , S.Libelle As Fournisseur  from AssPieceReparation A "
SQL = SQL & "INNER JOIN Station S ON S.code = A.Fournisseur Where "
If TypPiece <> "Tout" Then SQL = SQL & " A.type = " & SQLText(TypPiece) & " And "

SQL = SQL & "DatePiece Between " & SQLText(DateDu) & " and " & SQLText(DateAu)

If Station <> "0000" Then SQL = SQL & " and A.Fournisseur = " & SQLText(Station)
'If  supprimé
If TYP = "Pièce Supprimé" Then
    SQL = SQL & " and A.supp = 'O'"
Else 'non supprimé
    SQL = SQL & " and A.supp = 'N'"
    If TYP = "Pièce Non Facturé" Then
            SQL = SQL & " And A.transf ='N'"
    ElseIf TYP = "Pièce Facturé" Then
            SQL = SQL & " And A.transf ='O'"
    End If
    
End If

SQL = SQL & " Order by A.Numero"

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_PieceReparation = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Piece Reception pa rtype : avoir, piece reception , facture ..
Public Function Get_PieceReparByTyp(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal TypPiece As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_PieceReparByTyp"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = "select Numero, DatePiece, dateOperation , A.Supp , A.transf , S.Libelle As Fournisseur  from AssPieceReparation A "
    SQL = SQL & "INNER JOIN Station S ON S.code = A.Fournisseur"
    If TypPiece = "Tout" Then
        SQL = SQL & " Order by Numero "
    Else
        SQL = SQL & " Where A.type = " & SQLText(TypPiece) & " Order by Numero "
    End If
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_PieceReparByTyp = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Sub Insert_AssPieceRepar(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_AssPieceRepar"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
    StrSql = "Insert into AssPieceReparation (Numero,Type, DatePiece,RemisePiece, totTTC, DateOperation, refPiece, timbre, Fournisseur ,PrixMOeuvre,TVA_MOeuvre, UserInsert) values ("

    If IsNull(PRs_Record!numero) Then
        StrSql = StrSql & "Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!numero)
    End If
    
     If IsNull(PRs_Record!Type) Then
        StrSql = StrSql & ", Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Type)
    End If
    
    StrSql = StrSql & "," & SQLText(PRs_Record!DatePiece)
  
    If IsNull(PRs_Record!RemisePiece) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!RemisePiece), ",", ".")
    End If
    
    If IsNull(PRs_Record!totTTC) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!totTTC), ",", ".")
    End If
    
    If IsNull(PRs_Record!dateoperation) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!dateoperation)
    End If
    
    If IsNull(PRs_Record!refPiece) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!refPiece)
    End If
    
    If IsNull(PRs_Record!Timbre) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!Timbre), ",", ".")
    End If
    
    If IsNull(PRs_Record!Fournisseur) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Fournisseur)
    End If
    
    If IsNull(PRs_Record!PrixMOeuvre) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!PrixMOeuvre), ",", ".")
    End If
    
    If IsNull(PRs_Record!Tva_MOeuvre) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!Tva_MOeuvre), ",", ".")
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

Public Sub Insert_DetPieceRepar(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_DetPieceRepar"

    Dim StrSql As String

    On Error GoTo ErrHandler
PRs_Record.MoveFirst
While Not PRs_Record.EOF
    StrSql = "Insert into DetailPieceReparation (Numero,Designation, Qte,Vehicule, PUHT, Remise, TVA) values ("

    If IsNull(PRs_Record!numero) Then
        StrSql = StrSql & "Null"
    Else
        StrSql = StrSql & SQLText(PRs_Record!numero)
    End If
    
     If IsNull(PRs_Record!Designation) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Designation)
    End If
  
    If IsNull(PRs_Record!Qte) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Val(PRs_Record!Qte)
    End If
    
    If IsNull(PRs_Record!Vehicule) Then
        StrSql = StrSql & "," & "Null"
    Else
        StrSql = StrSql & "," & SQLText(PRs_Record!Vehicule)
    End If
    
    
    If IsNull(PRs_Record!PUHT) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!PUHT), ",", ".")
    End If
    
    If IsNull(PRs_Record!Remise) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!Remise), ",", ".")
    End If
    
    If IsNull(PRs_Record!tva) Then
        StrSql = StrSql & "," & "0"
    Else
        StrSql = StrSql & "," & Replace((PRs_Record!tva), ",", ".")
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

Public Sub Update_PieceRep(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_PieceRep"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update AssPieceReparation set"
   
    If IsNull(PRs_Record!Type) Then
        StrSql = StrSql & " Type = Null"
    Else
        StrSql = StrSql & " Type = " & SQLText(PRs_Record!Type)
    End If
    
    If IsNull(PRs_Record!RemisePiece) Then
        StrSql = StrSql & ", RemisePiece = 0"
    Else
        StrSql = StrSql & ", RemisePiece = " & Replace((PRs_Record!RemisePiece), ",", ".")
    End If

    If IsNull(PRs_Record!totTTC) Then
        StrSql = StrSql & " , totTTC = 0"
    Else
        StrSql = StrSql & " , totTTC = " & Replace((PRs_Record!totTTC), ",", ".")
    End If

    If IsNull(PRs_Record!dateoperation) Then
        StrSql = StrSql & ", DateOperation = Null"
    Else
        StrSql = StrSql & ", DateOperation = " & SQLText(PRs_Record!dateoperation)
    End If
    
    If IsNull(PRs_Record!refPiece) Then
        StrSql = StrSql & ", refPiece = Null"
    Else
        StrSql = StrSql & ", refPiece = " & SQLText(PRs_Record!refPiece)
    End If
    
    If IsNull(PRs_Record!Timbre) Then
        StrSql = StrSql & ", timbre = 0"
    Else
        StrSql = StrSql & ", timbre = " & Replace((PRs_Record!Timbre), ",", ".")
    End If
    
    If IsNull(PRs_Record!Fournisseur) Then
        StrSql = StrSql & ", Fournisseur = Null"
    Else
        StrSql = StrSql & ", Fournisseur = " & SQLText(PRs_Record!Fournisseur)
    End If
    
    If IsNull(PRs_Record!PrixMOeuvre) Then
        StrSql = StrSql & ", PrixMOeuvre = 0"
    Else
        StrSql = StrSql & ", PrixMOeuvre = " & Replace((PRs_Record!PrixMOeuvre), ",", ".")
    End If
    
    If IsNull(PRs_Record!Tva_MOeuvre) Then
        StrSql = StrSql & ", TVA_MOeuvre = 0"
    Else
        StrSql = StrSql & ", TVA_MOeuvre = " & Replace((PRs_Record!Tva_MOeuvre), ",", ".")
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

'Lors de la suppression d'une facture MAJ les pièce inclus passé de transféré à non transféré
Public Sub Update_NumFact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_NumFact"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update AssPieceReparation set transf='N'  , numfact=''  where Numfact=" & SQLText(vcode)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub
'MAJ Ass Pièce de réparation transféré
'Facture transferer
Public Sub Insert_NumFact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String, _
              ByVal N As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_NumFact"

    Dim StrSql As String

    On Error GoTo ErrHandler

    StrSql = "Update AssPieceReparation set transf='O'  , numfact= " & SQLText(vcode) & " where Numero=" & SQLText(N)
   
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub
'MAJ Bon de commande transféré
'MAJ AssBCREparation après transfert
Public Sub Update_Trans(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String, _
              ByVal vnum As String)
              
 Const sPROC_NAME As String = OBJ_NAME & ".Update_Trans"

    Dim StrSql As String

    On Error GoTo ErrHandler
    
    StrSql = "Update AssBCREparation set transf='O' , NumPR =" & SQLText(vcode) & " where numero=" & SQLText(vnum)
    
    CNB.Execute StrSql

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

              
End Sub

'Retourne les details d'un PIRep pour création de la facture
Public Function Get_DetPRep(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_DetPRep"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
'DetBonCarburant.AnCompteur,
SQL = "SELECT  A.PrixMOeuvre,A.TVA_MOeuvre,A.RemisePiece ,A.timbre, D.designation, D.Qte,D.PUHT , D.Remise , D.tva From DetailPieceReparation D"
SQL = SQL & " INNER JOIN AssPieceReparation A ON A.Numero = D.Numero"
SQL = SQL & " WHERE A.Type='Piece Reception' and  D.Numero =" & SQLText(vcode)
    
LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_DetPRep = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne les details d'un Bonretour pour création de la facture
Public Function Get_DetBRetour(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_DetBRetour"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler
'DetBonCarburant.AnCompteur,
SQL = "SELECT  DetailPieceReparation.designation, DetailPieceReparation.Qte,DetailPieceReparation.PUHT From DetailPieceReparation"
SQL = SQL & " INNER JOIN AssPieceReparation ON AssPieceReparation.Numero = DetailPieceReparation.Numero"
SQL = SQL & " WHERE AssPieceReparation.Type='Bon Retour' and  DetailPieceReparation.Numero =" & SQLText(vcode)
    
LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_DetBRetour = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Détails Piece Rep pour création d'une facture
Public Function Get_ProdPRep(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vCodeSation As String, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_ProdPRep"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler

SQL = " Select Distinct  D.Designation,A.RemisePiece as RemiseP , Sum(D.Qte) As Qte, D.PUHT, D.Remise, D.TVA from DetailPieceReparation D"
SQL = SQL & " INNER JOIN AssPieceReparation A ON A.Numero  = D.Numero"
SQL = SQL & " where Type in ('Piece Reception') and transf='N' and A.Supp='N'"
SQL = SQL & " and dateoperation between  " & SQLText(VdateD) & " and " & SQLText(vDateF)
SQL = SQL & " and FOurnisseur=" & SQLText(vCodeSation)
SQL = SQL & " GROUP BY Designation, PUHT, remise, tva ,RemisePiece"

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_ProdPRep = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'MainD'oeuvre, remise , timbre Piece Rep pour création d'une facture
Public Function Get_RmsMOTimbPRep(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vCodeSation As String, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_RmsMOTimbPRep"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler

SQL = " Select PrixMOeuvre as MOeuvre,TVA_MOeuvre , RemisePiece as RemiseP , Timbre as timbre from AssPieceReparation"
SQL = SQL & " where Type in ('Piece Reception') and transf='N' and Supp='N'"
SQL = SQL & " and dateoperation between  " & SQLText(VdateD) & " and " & SQLText(vDateF)
SQL = SQL & " and FOurnisseur = " & SQLText(vCodeSation)

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_RmsMOTimbPRep = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'MainD'oeuvre, remise , timbre Piece Rep pour création d'une facture
Public Function Get_MOTimbPRepFact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vNumero As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_MOTimbPRepFact"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler

SQL = " Select PrixMOeuvre as MOeuvre,TVA_MOeuvre as TVA_MOeuvre,RemisePiece, Timbre as timbre from AssPieceReparation"
SQL = SQL & " where Type ='Piece Reception' and transf='O' and numfact = " & SQLText(vNumero)

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_MOTimbPRepFact = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Details Bon de retour pour création d'une facture
Public Function Get_ProdBRetour(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vCodeSation As String, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_ProdBRetour"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler

SQL = " Select Distinct D.Designation, Sum(D.Qte) As Qte, D.PUHT, D.Remise, D.TVA from DetailPieceReparation D"
SQL = SQL & " INNER JOIN AssPieceReparation A ON A.Numero  = D.Numero"
SQL = SQL & " where Type in ('Bon Retour') and transf='N'and A.Supp='N' and dateoperation between  " & SQLText(VdateD) & " and " & SQLText(vDateF) & " and FOurnisseur=" & SQLText(vCodeSation)
SQL = SQL & " GROUP BY Designation, PUHT, remise, tva"

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_ProdBRetour = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Chercher piece de reparation d'une facture précise
Public Function Get_PRByNumFact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vNumero) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_PRByNumFact"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler

SQL = " Select D.Designation , D.PUHT , Sum(D.Qte) as Qte ,D.Remise,D.tva, A.RemisePiece from DetailPieceReparation D"
SQL = SQL & " INNER JOIN AssPieceReparation A ON A.Numero  = D.Numero"
SQL = SQL & " where Type ='Piece Reception' and transf='O' and numfact=" & SQLText(vNumero)
SQL = SQL & " GROUP BY D.Designation ,D.PUHT ,D.Remise,D.tva , A.RemisePiece"

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_PRByNumFact = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Chercher bonde retour d'une facture précise
Public Function Get_BRByNumFact(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vNumero) As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".Get_BRByNumFact"

Dim SQL As String
Dim LRs_Record As New Recordset

On Error GoTo ErrHandler

SQL = " Select * from DetailPieceReparation"
SQL = SQL & " INNER JOIN AssPieceReparation ON AssPieceReparation.Numero  = DetailPieceReparation.Numero"
SQL = SQL & " where Type ='Bon Retour' and transf='O' and numfact=" & SQLText(vNumero)

LRs_Record.CursorLocation = adUseClient
LRs_Record.Open SQL, CNB, adOpenKeyset
Set Get_BRByNumFact = LRs_Record
Set LRs_Record = Nothing

Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'==================Statistiques ===================
'==================================================

'Retourne details des pieces de repar pour statistiques
Public Function Get_PieceRepStatistVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal Matricule As String, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_PieceRepStatistVeh"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = " Select * from AsspieceReparation A, detailpieceReparation D where"
    SQL = SQL & " A.Numero = D.Numero And A.Supp = 'N' "
    SQL = SQL & " And D.vehicule= " & SQLText(Matricule) & " And "
    SQL = SQL & "  A.datePiece"
    SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF)
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_PieceRepStatistVeh = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_NbrRepStatistVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal Matricule As String, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_nbrRepStatistVeh"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = " Select count(*) As nbrRep from AsspieceReparation A, detailpieceReparation D where"
    SQL = SQL & " A.Numero = D.Numero And A.Supp = 'N' "
    SQL = SQL & " And D.vehicule= " & SQLText(Matricule) & " And "
    SQL = SQL & "  A.datePiece"
    SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF)
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_NbrRepStatistVeh = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_SumNbrRepStatist(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_SumNbrRepStatist"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

    SQL = " Select count(*) As nbrRep ,Sum (Distinct A.totTTC)As valeur from AsspieceReparation A, detailpieceReparation D where"
    SQL = SQL & " A.Numero = D.Numero And A.Supp ='N' and A.datePiece"
    SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF)
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_SumNbrRepStatist = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_NbrRepStatistGrpVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_NbrRepStatistGrpVeh"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

SQL = " Select vehicule , count(*) As nbrRep  from AsspieceReparation A, detailpieceReparation D where"
SQL = SQL & " A.Numero = D.Numero and A.Supp ='N' "
SQL = SQL & " And A.datePiece"
SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF) & " group by Vehicule"
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_NbrRepStatistGrpVeh = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_ValRepStatistVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_ValRepStatistVeh"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

SQL = "select Qte, PUHT, Remise, TVA, remisePiece from AssPieceReparation A ,detailPiecereparation D where "
SQL = SQL & " A.Numero= D.Numero And A.Supp='N' and "
SQL = SQL & " vehicule = " & SQLText(vcode)
SQL = SQL & " And Datepiece between " & SQLText(VdateD) & " and " & SQLText(vDateF)
    
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_ValRepStatistVeh = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_DetRepStatist(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal VdateD As String, _
              ByVal vDateF As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_DetRepStatist"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler

SQL = "Select * from DetailPieceReparation D , AssPieceReparation A "
SQL = SQL & " Where D.Numero = A.Numero And A.Supp = 'N' "
SQL = SQL & " And A.datePiece"
SQL = SQL & " between " & SQLText(VdateD) & " and " & SQLText(vDateF) & " order by d.vehicule, a.datePiece"

    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_DetRepStatist = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


