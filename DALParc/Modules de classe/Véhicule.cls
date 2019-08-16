VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vehicule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const OBJ_NAME As String = "Vehicule"

Public Function GetRow_Vehicule_ByCode(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal CNB As ADODB.Connection, _
                                        ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetRow_Vehicule_ByCode"
    Dim SQL As String
    Dim LRs_Record As Recordset

On Error GoTo ErrHandler

    SQL = "Select * from vehicule where  ACTIF=1 and code = " & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetRow_Vehicule_ByCode = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_Vehicule_ByLibelle(ByRef ErrNumber As Long, _
                                            ByRef ErrDescription As String, _
                                            ByRef ErrSourceDetail As String, _
                                            ByVal CNB As ADODB.Connection, _
                                            ByVal vLibelle As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetRow_Vehicule_ByLibelle"
    Dim SQL As String
    Dim LRs_Record As Recordset

On Error GoTo ErrHandler

    SQL = "Select * from vehicule where Actif=1 And  Matricule like " & SQLText(vLibelle & "%")

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetRow_Vehicule_ByLibelle = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetAll_Vehicule_Actif(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetAll_Vehicule_Actif"
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "SELECT * FROM VEHICULE WHERE ACTIF=1 and Supp='N' ORDER BY MATRICULE"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetAll_Vehicule_Actif = LRs_Record
    Set LRs_Record = Nothing
    
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function









'======================================================================================================


















Public Function GetAll_Abreviation(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal ABr As String, _
                                        ByVal code As String, _
                                        ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME    As String = OBJ_NAME & ".GetAll_Abreviation"
    Dim SQL         As String
    Dim LRs_Record  As New Recordset
On Error GoTo ErrHandler
    SQL = "SELECT Abreviation FROM VEHICULE WHERE  Abreviation=" & SQLText(ABr) & " and Code <> " & SQLText(code)
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetAll_Abreviation = LRs_Record
    Set LRs_Record = Nothing
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetRow_AbrByMatricule(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal Mat As String, _
                                        ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME    As String = OBJ_NAME & ".GetRow_AbrByMatricule"
    Dim SQL         As String
    Dim LRs_Record  As New Recordset
On Error GoTo ErrHandler
    SQL = "SELECT Abreviation FROM VEHICULE WHERE  Matricule=" & SQLText(Mat)
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetRow_AbrByMatricule = LRs_Record
    Set LRs_Record = Nothing
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetAllActifVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetAllActifVeh"
    Dim SQL As String
    Dim LRs_Record As Recordset
On Error GoTo ErrHandler
    SQL = "SELECT * FROM VEHICULE WHERE  ACTIF=1 And Supp='N' ORDER BY MATRICULE"
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetAllActifVeh = LRs_Record
    Set LRs_Record = Nothing
Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function GetAll_VehiculeActifNonSupp(ByRef ErrNumber As Long, _
                                                    ByRef ErrDescription As String, _
                                                    ByRef ErrSourceDetail As String, _
                                                    ByVal CNB As ADODB.Connection, _
                                                    ByVal StrSource As String) As Recordset

    Const sPROC_NAME    As String = OBJ_NAME & ".GetAll_VehiculeActifNonSupp"
    Dim SQL             As String
    Dim LRs_Record      As Recordset
    
On Error GoTo ErrHandler
    SQL = "SELECT * FROM VEHICULE  "
    If StrSource <> "VehiculeBase" Then SQL = SQL & " WHERE ACTIF=1 AND SUPP='N'"
    SQL = SQL & "ORDER BY MATRICULE"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetAll_VehiculeActifNonSupp = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
'MAJ du compteur de carburant après avoir remplir gazoil (ou ..) : mettre nouveau compteur
Public Sub Update_Compt(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal cmpt As Long, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_Compt"

    Dim SQL As String

    On Error GoTo ErrHandler
'cmpt : compteur initial pour ce bon
    SQL = "Update vehicule set CompteurCarburant = " & cmpt & "  where Code = " & SQLText(vcode)

    CNB.Execute SQL

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'MAJ du compteur de vidange : mettre nouveau compteur
Public Sub Update_ComptVdg(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal cmpt As Long, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_ComptVdg"

    Dim SQL As String

    On Error GoTo ErrHandler

    SQL = "Update vehicule set CompteurVidange = " & cmpt & "  where Code = " & SQLText(vcode)

    CNB.Execute SQL

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Function GetVehiculeByCode(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetVehiculeByCode"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "Select * from vehicule where code = " & SQLText(vcode) & " Or Matricule = " & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetVehiculeByCode = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Dernier Compteur d'entree du véhicule à partir du FicheTraffic
Public Function Get_DerCompt(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal Name_Tab As String, _
              ByVal vcode As String) As Recordset

 Const sPROC_NAME As String = OBJ_NAME & ".Get_DerCompt"

    Dim SQL As String
    Dim LRs_Record As New Recordset

    On Error GoTo ErrHandler
    
    SQL = "Select Max(CompteurEntre) as maxCpt from " & Name_Tab & " where Vehicule =" & SQLText(vcode)
'    SQL = SQL & "And cast (Numero as int) = ( select max(cast(numero As int))From " & Name_Tab & " where vehicule = " & SQLText(vcode) & ")"
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_DerCompt = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Vehicule pour forme du bonvidange
Public Function GetVehByCodeMat(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetVehByCodeMat"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "Select * from vehicule where Actif=1 And Supp='N' And  code = " & SQLText(vcode) & " OR Matricule = " & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetVehByCodeMat = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function GetVehicule(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".GetVehicule"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "Select * from vehicule where code = " & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set GetVehicule = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function



Public Function Get_AllVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AllVeh"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "select * from vehicule order by Matricule"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_AllVeh = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne liste des matricules de tout les véhicules
Public Function Get_AllMatrVeh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_AllMatrVeh"

    Dim SQL As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    SQL = "Select code , matricule from Vehicule where actif=1 and Supp='N' order by matricule"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenKeyset
    Set Get_AllMatrVeh = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Suppression d'un véhicule
Public Sub Delete_Vehicule(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal id As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Vehicule"

    Dim SQL As String

    On Error GoTo ErrHandler

    SQL = "Delete from vehicule where code = " & SQLText(id)
   
    CNB.Execute SQL

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Sub Insert_Veh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_Veh"

    Dim SQL As String

    On Error GoTo ErrHandler
    
    SQL = "Insert into vehicule (Code ,Matricule,Abreviation ,TYPE ,Marque ,Puissance ,Energie, "
    SQL = SQL & " genre,Carrosserie,PlaceAssis,PlaceDebout,Cylindre,NbrEssieux,PTAC,PTRA,Charge,NbKlmVidange,PoidsVide,TypeComm,"
    SQL = SQL & " NumSerie ,DateCircul ,NumAssur ,FournisAssur ,AgenceAssur ,"
    SQL = SQL & " DateDebAssur , DAteFinAssur, DateDebVisite, DAteFinVisite,DateDebTax,DateFinTax, DateSortie, Obs, CompteurVidange,CompteurCarburant,CompteurFT,actif,Disponible"

    SQL = SQL & ", UserInsert,PicBox) values ("
    
    If IsNull(PRs_Record!code) Then
        SQL = SQL & "Null"
    Else
        SQL = SQL & SQLText(PRs_Record!code)
    End If

    If IsNull(PRs_Record!Matricule) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!Matricule)
    End If
    
    If IsNull(PRs_Record!ABr) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!ABr)
    End If
    
    If IsNull(PRs_Record!Type) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!Type)
    End If
    
    If IsNull(PRs_Record!Marque) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!Marque)
    End If
    
    If IsNull(PRs_Record!Puissance) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!Puissance)
    End If
    
    If IsNull(PRs_Record!Energie) Then
        SQL = SQL & "," & "0"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!Energie)
    End If
     
    If IsNull(PRs_Record!genre) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!genre)
    End If
    
    If IsNull(PRs_Record!Carrosserie) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!Carrosserie)
    End If
    
    If IsNull(PRs_Record!PlaceAssis) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!PlaceAssis)
    End If
    
    If IsNull(PRs_Record!PlaceDebout) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!PlaceDebout)
    End If
    
    If IsNull(PRs_Record!Cylindre) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!Cylindre)
    End If
    
    If IsNull(PRs_Record!NbrEssieux) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!NbrEssieux)
    End If
    
    If IsNull(PRs_Record!PTAC) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!PTAC)
    End If
    
    If IsNull(PRs_Record!PTRA) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!PTRA)
    End If
    
    If IsNull(PRs_Record!Charge) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!Charge)
    End If
    
    If IsNull(PRs_Record!NbKlmVidange) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!NbKlmVidange)
    End If
    
    If IsNull(PRs_Record!PoidsVide) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!PoidsVide)
    End If
    
    If IsNull(PRs_Record!TypeComm) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!TypeComm)
    End If
    
    If IsNull(PRs_Record!NumSerie) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!NumSerie)
    End If

    If IsNull(PRs_Record!DateCircul) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!DateCircul)
    End If

    If IsNull(PRs_Record!NumAssur) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!NumAssur)
    End If
    
    If IsNull(PRs_Record!FournisAssur) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!FournisAssur)
    End If
    
    If IsNull(PRs_Record!AgenceAssur) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!AgenceAssur)
    End If
    
    If IsNull(PRs_Record!DateDebAssur) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!DateDebAssur)
    End If
    
    If IsNull(PRs_Record!DateFinAssur) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!DateFinAssur)
    End If
    
    If IsNull(PRs_Record!DateDebVisite) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!DateDebVisite)
    End If
    
    If IsNull(PRs_Record!DateFinVisite) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!DateFinVisite)
    End If
    
    If IsNull(PRs_Record!DateDebTax) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!DateDebTax)
    End If
    
    If IsNull(PRs_Record!DateFinTax) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!DateFinTax)
    End If
    
    If IsNull(PRs_Record!DateSortie) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!DateSortie)
    End If
    
    If IsNull(PRs_Record!Obs) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!Obs)
    End If
    
    If IsNull(PRs_Record!CompteurVidange) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!CompteurVidange)
    End If
    
    If IsNull(PRs_Record!CompteurCarburant) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!CompteurCarburant)
    End If
    
    If IsNull(PRs_Record!CompteurFT) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!CompteurFT)
    End If
    
    If IsNull(PRs_Record!Actif) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!Actif)
    End If
    
    If IsNull(PRs_Record!Disponible) Then
        SQL = SQL & "," & "Null"
    Else
        SQL = SQL & "," & SQLText(PRs_Record!Disponible)
    End If
    SQL = SQL & "," & SQLText(PRs_Record!USERINSERT)
    If Not IsNull(PRs_Record!PicBox) Then SQL = SQL & "," & SQLText(PRs_Record!PicBox)
    SQL = SQL & ")"

    CNB.Execute SQL

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Sub Update_Veh(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Update_Veh"

    Dim SQL As String

    On Error GoTo ErrHandler

    SQL = "Update Vehicule set "

    If IsNull(PRs_Record!code) Then
        SQL = SQL & " Code = Null"
    Else
        SQL = SQL & " Code = " & SQLText(PRs_Record!code)
    End If
    
    If IsNull(PRs_Record!Matricule) Then
        SQL = SQL & ", Matricule = Null"
    Else
        SQL = SQL & ", Matricule = " & SQLText(PRs_Record!Matricule)
    End If
    
    If IsNull(PRs_Record!ABr) Then
        SQL = SQL & ", Abreviation= Null"
    Else
        SQL = SQL & ", Abreviation = " & SQLText(PRs_Record!ABr)
    End If
    
    If IsNull(PRs_Record!Type) Then
        SQL = SQL & ", TYPE = Null"
    Else
        SQL = SQL & ", TYPE = " & SQLText(PRs_Record!Type)
    End If
    
    If IsNull(PRs_Record!Marque) Then
        SQL = SQL & ", Marque = Null"
    Else
        SQL = SQL & ", Marque = " & SQLText(PRs_Record!Marque)
    End If
    
    If IsNull(PRs_Record!Puissance) Then
        SQL = SQL & ", Puissance = Null"
    Else
        SQL = SQL & ", Puissance = " & SQLText(PRs_Record!Puissance)
    End If
    
    If IsNull(PRs_Record!Energie) Then
        SQL = SQL & ", Energie = Null"
    Else
        SQL = SQL & ", Energie = " & SQLText(PRs_Record!Energie)
    End If
    
    If IsNull(PRs_Record!genre) Then
        SQL = SQL & ", genre = Null"
    Else
        SQL = SQL & ", genre = " & SQLText(PRs_Record!genre)
    End If
    
    If IsNull(PRs_Record!Carrosserie) Then
        SQL = SQL & ", Carrosserie = Null"
    Else
        SQL = SQL & ", Carrosserie = " & SQLText(PRs_Record!Carrosserie)
    End If
    If IsNull(PRs_Record!PlaceAssis) Then
        SQL = SQL & ", PlaceAssis = 0"
    Else
        SQL = SQL & ", PlaceAssis = " & SQLText(PRs_Record!PlaceAssis)
    End If
    
    If IsNull(PRs_Record!PlaceDebout) Then
        SQL = SQL & ", PlaceDebout = 0"
    Else
        SQL = SQL & ", PlaceDebout = " & SQLText(PRs_Record!PlaceDebout)
    End If
    
    If IsNull(PRs_Record!Cylindre) Then
        SQL = SQL & ", Cylindre = 0"
    Else
        SQL = SQL & ", Cylindre = " & SQLText(PRs_Record!Cylindre)
    End If
    If IsNull(PRs_Record!NbrEssieux) Then
        SQL = SQL & ", NbrEssieux = 0"
    Else
        SQL = SQL & ", NbrEssieux = " & SQLText(PRs_Record!NbrEssieux)
    End If
    
    If IsNull(PRs_Record!PTAC) Then
        SQL = SQL & ", PTAC = 0"
    Else
        SQL = SQL & ", PTAC = " & SQLText(PRs_Record!PTAC)
    End If
    
    If IsNull(PRs_Record!PTRA) Then
        SQL = SQL & ", PTRA = 0"
    Else
        SQL = SQL & ", PTRA = " & SQLText(PRs_Record!PTRA)
    End If
    If IsNull(PRs_Record!Charge) Then
        SQL = SQL & ", Charge = 0"
    Else
        SQL = SQL & ", Charge = " & SQLText(PRs_Record!Charge)
    End If
    
    If IsNull(PRs_Record!NbKlmVidange) Then
        SQL = SQL & ", NbKlmVidange = 0"
    Else
        SQL = SQL & ", NbKlmVidange = " & SQLText(PRs_Record!NbKlmVidange)
    End If
    
    If IsNull(PRs_Record!PoidsVide) Then
        SQL = SQL & ", PoidsVide = 0"
    Else
        SQL = SQL & ", PoidsVide = " & SQLText(PRs_Record!PoidsVide)
    End If
    
    If IsNull(PRs_Record!TypeComm) Then
        SQL = SQL & ", TypeComm = Null"
    Else
        SQL = SQL & ", TypeComm = " & SQLText(PRs_Record!TypeComm)
    End If
    If IsNull(PRs_Record!NumSerie) Then
        SQL = SQL & ", NumSerie = Null"
    Else
        SQL = SQL & ", NumSerie = " & SQLText(PRs_Record!NumSerie)
    End If
    
    If IsNull(PRs_Record!DateCircul) Then
        SQL = SQL & ", DateCircul = Null"
    Else
        SQL = SQL & ", DateCircul = " & SQLText(PRs_Record!DateCircul)
    End If
    
    If IsNull(PRs_Record!NumAssur) Then
        SQL = SQL & ", NumAssur = 0"
    Else
        SQL = SQL & ", NumAssur = " & SQLText(PRs_Record!NumAssur)
    End If
    
    
     If IsNull(PRs_Record!FournisAssur) Then
        SQL = SQL & ", FournisAssur = Null"
    Else
        SQL = SQL & ", FournisAssur = " & SQLText(PRs_Record!FournisAssur)
    End If
    
    If IsNull(PRs_Record!AgenceAssur) Then
        SQL = SQL & ", AgenceAssur = Null"
    Else
        SQL = SQL & ", AgenceAssur = " & SQLText(PRs_Record!AgenceAssur)
    End If
    
    If IsNull(PRs_Record!DateDebAssur) Then
        SQL = SQL & ", DateDebAssur = 0"
    Else
        SQL = SQL & ", DateDebAssur = " & SQLText(PRs_Record!DateDebAssur)
    End If
     If IsNull(PRs_Record!DateFinAssur) Then
        SQL = SQL & ", DAteFinAssur = Null"
    Else
        SQL = SQL & ", DAteFinAssur = " & SQLText(PRs_Record!DateFinAssur)
    End If
    
    If IsNull(PRs_Record!DateDebVisite) Then
        SQL = SQL & ", DateDebVisite = Null"
    Else
        SQL = SQL & ", DateDebVisite = " & SQLText(PRs_Record!DateDebVisite)
    End If
    
    If IsNull(PRs_Record!DateFinVisite) Then
        SQL = SQL & ", DAteFinVisite = Null"
    Else
        SQL = SQL & ", DAteFinVisite = " & SQLText(PRs_Record!DateFinVisite)
    End If
    
    If IsNull(PRs_Record!DateDebTax) Then
        SQL = SQL & ", DateDebTax = Null"
    Else
        SQL = SQL & ", DateDebTax = " & SQLText(PRs_Record!DateDebTax)
    End If
    
    If IsNull(PRs_Record!DateFinTax) Then
        SQL = SQL & ", DateFinTax = Null"
    Else
        SQL = SQL & ", DateFinTax = " & SQLText(PRs_Record!DateFinTax)
    End If
    If IsNull(PRs_Record!DateSortie) Then
        SQL = SQL & ", DateSortie = Null"
    Else
        SQL = SQL & ", DateSortie = " & SQLText(PRs_Record!DateSortie)
    End If
    
    If IsNull(PRs_Record!Obs) Then
        SQL = SQL & ", Obs = Null"
    Else
        SQL = SQL & ", Obs = " & SQLText(PRs_Record!Obs)
    End If
    
    If IsNull(PRs_Record!CompteurVidange) Then
        SQL = SQL & ", CompteurVidange = 0"
    Else
        SQL = SQL & ", CompteurVidange = " & Val(PRs_Record!CompteurVidange)
    End If
    
    If IsNull(PRs_Record!CompteurCarburant) Then
        SQL = SQL & ", CompteurCarburant = 0"
    Else
        SQL = SQL & ", CompteurCarburant = " & Val(PRs_Record!CompteurCarburant)
    End If
    
    If IsNull(PRs_Record!CompteurFT) Then
        SQL = SQL & ", CompteurFT = 0"
    Else
        SQL = SQL & ", CompteurFT = " & Val(PRs_Record!CompteurFT)
    End If
    
    If IsNull(PRs_Record!Actif) Then
        SQL = SQL & ", actif = Null"
    Else
        SQL = SQL & ", actif = " & SQLText(PRs_Record!Actif)
    End If
    
    SQL = SQL & ", Disponible = " & SQLText(PRs_Record!Disponible)
    SQL = SQL & ", userupdate = " & SQLText(PRs_Record!userupdate)
    If Not IsNull(PRs_Record!PicBox) Then SQL = SQL & ", PicBox = " & SQLText(PRs_Record!PicBox)
    SQL = SQL & " Where Code = " & SQLText(PRs_Record!code)

    CNB.Execute SQL

    If Not IsNull(PRs_Record!CompteurVidange) Then
        CNB.Execute "update Ass_BonVidange set CompteurVidange = " & Val(PRs_Record!CompteurVidange) & " where Vehicule = " & SQLText(PRs_Record!code) & " And Numero = (select Max(Numero)from Ass_BonVidange where Vehicule = " & SQLText(PRs_Record!code) & ")"
    End If
    
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source
End Sub

Public Function GetRow_Vehicule_ByDateSortie(ByRef ErrNumber As Long, _
                                            ByRef ErrDescription As String, _
                                            ByRef ErrSourceDetail As String, _
                                            ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

'    SQL = "Select * from vehicule where datesortie is null order by matricule"

    SQL = "select Distinct V.* "
    SQL = SQL & " from Vehicule V"
    SQL = SQL & " where datesortie is null and V.supp='N' order by matricule"
    
    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetRow_Vehicule_ByDateSortie = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function GetAlerte_Vehicule(ByRef ErrNumber As Long, _
                                            ByRef ErrDescription As String, _
                                            ByRef ErrSourceDetail As String, _
                                            ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

SQL = "Select code,Marque,Matricule,Compteurvidange,compteurcarburant, DATEDIFF(Day, GETDATE(), DateFinAssur) as Assur, DATEDIFF(Day, GETDATE(), DateFinVisite) as Visite, DATEDIFF(Day, GETDATE(), DateFinTax) as Taxe from Vehicule where supp='N' and actif=1 group by DateFinAssur,Matricule,DateFinVisite,DateFinTax,Compteurvidange,compteurcarburant,code,Marque"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetAlerte_Vehicule = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetLast_VidangeParVehicule(ByRef ErrNumber As Long, _
                                            ByRef ErrDescription As String, _
                                            ByRef ErrSourceDetail As String, _
                                            ByVal Matricule As String, _
                                            ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

SQL = "SELECT vehicule, MAX (CompteurVidange) AS MaxCpt FROM Ass_BonVidange where Vehicule=Matricule group by Vehicule"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetLast_VidangeParVehicule = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function GetLast_FTParVehicule(ByRef ErrNumber As Long, _
                                            ByRef ErrDescription As String, _
                                            ByRef ErrSourceDetail As String, _
                                            ByVal Matricule As String, _
                                            ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

SQL = "select Vehicule, CompteurEntre from FicheTraffic where CompteurEntre=(select Max(compteurEntre) from FicheTraffic where vehicule= " & SQLText(Matricule) & ") and Vehicule=" & SQLText(Matricule)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetLast_FTParVehicule = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Function Get_Alertes(ByRef ErrNumber As Long, _
                                            ByRef ErrDescription As String, _
                                            ByRef ErrSourceDetail As String, _
                                            ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

SQL = "Select code,Marque,Matricule,Compteurvidange,compteurcarburant,CompteurEntre,DATEDIFF(Day, GETDATE(), DateFinAssur) as Assur,DATEDIFF(Day, GETDATE(), DateFinVisite) as Visite, DATEDIFF(Day, GETDATE(), DateFinTax) as Taxe from Vehicule V, Fichetraffic F where V.code=F.vehicule and V.supp='N' and V.actif=1 and CompteurEntre in (select Max(compteurEntre) from FicheTraffic group by Vehicule) group by DateFinAssur,Matricule,DateFinVisite,DateFinTax,Compteurvidange,compteurcarburant, code,Marque, vehicule,CompteurEntre"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set Get_Alertes = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function



Public Function GetMatricule_Vehicules(ByRef ErrNumber As Long, _
                                    ByRef ErrDescription As String, _
                                    ByRef ErrSourceDetail As String, _
                                    ByVal CNB As ADODB.Connection) As Recordset
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "select Matricule from vehicule Order by matricule"

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetMatricule_Vehicules = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function
Public Sub UpdateCompteurFT_Vehicule(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal Matricule As String, _
                                        ByVal compteur As String, _
                                        ByVal CNB As ADODB.Connection)
On Error GoTo ErrHandler
    
    CNB.Execute "Update Vehicule Set CompteurFT = " & SQLText(compteur) & " where Matricule=" & SQLText(Matricule)
    
Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub

'Suppression d'un véhicule
Public Sub Delete_Add_Vehicule(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vSupp As String, _
              ByVal CodeUser As String, _
              ByVal code As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_Vehicule"

    Dim SQL As String
    
    If vSupp <> "N" And vSupp <> "O" Then
        MsgBox "variable supp invalide", vbExclamation
        Exit Sub
    End If
    
On Error GoTo ErrHandler

    SQL = "Update vehicule set "
    SQL = SQL & " Supp = " & SQLText(vSupp)
    If vSupp = "O" Then
        SQL = SQL & ", UserDelete=" & SQLText(CodeUser)
    ElseIf vSupp = "N" Then
        SQL = SQL & ", UserAdd=" & SQLText(CodeUser)
    End If
    SQL = SQL & " where Code = " & SQLText(code)
    CNB.Execute SQL

Exit Sub
ErrHandler:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub


Public Function GetCompteur_Vehicule_ByMatricule(ByRef ErrNumber As Long, _
                                        ByRef ErrDescription As String, _
                                        ByRef ErrSourceDetail As String, _
                                        ByVal Matricule As String, _
                                        ByVal CNB As ADODB.Connection) As Recordset
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from vehicule where Matricule = " & SQLText(Matricule)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set GetCompteur_Vehicule_ByMatricule = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Sub Update_Disp_Vehicule(ByRef ErrNumber As Long, _
                                ByRef ErrDescription As String, _
                                ByRef ErrSourceDetail As String, _
                                ByVal vDesp As String, _
                                ByVal Matricule As String, _
                                ByVal CNB As ADODB.Connection)
On Error GoTo ErrHandler
    
    If vDesp <> "N" And vDesp <> "O" Then
        MsgBox "variable vDesp invalide", vbExclamation
        Exit Sub
    End If
        
    CNB.Execute "Update Vehicule Set disponible= " & SQLText(vDesp) & " where Matricule=" & SQLText(Matricule)
    
Exit Sub
ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Sub

'=========================== Vehicule_Vidange=========================
'=====================================================================

Public Sub Insert_VehVdg(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal PRs_Record As ADODB.Recordset)

    Const sPROC_NAME As String = OBJ_NAME & ".Insert_VehVdg"

    Dim SQL As String

    On Error GoTo ErrHandler
    
PRs_Record.MoveFirst
While Not PRs_Record.EOF
    SQL = "Insert into Vehicule_Vidange (Vehicule,Lubrifiant,UserInsert) values ( "
    
    
    If IsNull(PRs_Record!Vehicule) Then
        SQL = SQL & "Null"
    Else
        SQL = SQL & SQLText(PRs_Record!Vehicule)
    End If
    
    If IsNull(PRs_Record!Lubrifiant) Then
        SQL = SQL & "," & "0"
    Else
        SQL = SQL & "," & Val(PRs_Record!Lubrifiant)
    End If
    
    SQL = SQL & "," & SQLText(PRs_Record!USERINSERT) & ")"

    CNB.Execute SQL
PRs_Record.MoveNext
Wend
    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

'Suppression d'un véhicule
Public Sub Delete_VehVdg(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal vcode As String)

    Const sPROC_NAME As String = OBJ_NAME & ".Delete_VehVdg"

    Dim SQL As String

    On Error GoTo ErrHandler

    SQL = "Delete from Vehicule_Vidange where Vehicule = " & SQLText(vcode)
   
    CNB.Execute SQL

    Exit Sub

ErrHandler:

    ErrNumber = Err.Number
    ErrDescription = Err.Description
    ErrSourceDetail = sPROC_NAME & vbCrLf & Err.Source

End Sub

Public Function Get_VehVdg(ByRef ErrNumber As Long, _
                        ByRef ErrDescription As String, _
                        ByRef ErrSourceDetail As String, _
                        ByVal CNB As ADODB.Connection, _
                        ByVal vcode As String) As Recordset
    
    Const sPROC_NAME As String = OBJ_NAME & ".Get_VehVdg"
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from Vehicule_Vidange where Vehicule = " & SQLText(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set Get_VehVdg = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

Public Function Get_VdgVehByLub(ByRef ErrNumber As Long, _
                        ByRef ErrDescription As String, _
                        ByRef ErrSourceDetail As String, _
                        ByVal CNB As ADODB.Connection, _
                        ByVal vcode As String) As Recordset
    
    Const sPROC_NAME As String = OBJ_NAME & ".Get_VdgVehByLub"
    
    Dim SQL As String
    Dim LRs_Record As Recordset
    
On Error GoTo ErrHandler

    SQL = "Select * from Vehicule_Vidange where Lubrifiant = " & Val(vcode)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open SQL, CNB, adOpenStatic, adLockPessimistic
    Set Get_VdgVehByLub = LRs_Record
    Set LRs_Record = Nothing

Exit Function
ErrHandler:
    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function

'Retourne l'abreviation du Matricule du véhicule
Public Function Get_Abrev(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal txt As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_Abrev"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "Select ABREVIATION from Vehicule where Matricule = " & SQLText(txt)

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

'Retourne matricule du véhicule par son abreviation
Public Function Get_MatricAbrev(ByRef ErrNumber As Long, _
              ByRef ErrDescription As String, _
              ByRef ErrSourceDetail As String, _
              ByVal CNB As ADODB.Connection, _
              ByVal txt As String) As Recordset

    Const sPROC_NAME As String = OBJ_NAME & ".Get_MatricAbrev"

    Dim StrSql As String
    Dim LRs_Record As Recordset

    On Error GoTo ErrHandler

    StrSql = "Select Matricule from Vehicule where ABREVIATION  = " & SQLText(txt)

    Set LRs_Record = New Recordset
    LRs_Record.CursorLocation = adUseClient
    LRs_Record.Open StrSql, CNB, adOpenKeyset
    Set Get_MatricAbrev = LRs_Record
    Set LRs_Record = Nothing

    Exit Function

ErrHandler:

    Set LRs_Record = Nothing
    ErrNumber = Err.Number
    ErrDescription = Err.Description
End Function


