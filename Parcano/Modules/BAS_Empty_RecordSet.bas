Attribute VB_Name = "BAS_Empty_RecordSet"
Option Explicit
    Private Const OBJ_NAME As String = "BAS_Clt"
'----------------------------------------------------------------------------------------------------------------------------------
'Programme Chauffeurs***
'----------------------------------------------------------------------------------------------------------------------------------
Public Function CreateEmptyRS_Ass_ProgChauf() As Recordset
    Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_Ass_ProgChauf"
    Dim L_RS As ADODB.Recordset
    
    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Code", 202, 25, adFldIsNullable
    L_RS.Fields.Append "CodeConducteur", 202, 50, adFldIsNullable
    L_RS.Fields.Append "CodeVehicule", 202, 30, adFldIsNullable
    L_RS.Fields.Append "DateCreation", 202, 25, adFldIsNullable
    L_RS.Fields.Append "DateProgramme", 202, 25, adFldIsNullable
    L_RS.Fields.Append "UserInsert", 202, 30, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserDelete", 202, 30, adFldIsNullable
    L_RS.Fields.Append "UserAdd", 202, 30, adFldIsNullable
    L_RS.Open
    
    Set CreateEmptyRS_Ass_ProgChauf = L_RS
End Function
Public Function CreateEmptyRS_Det_ProgChauf() As Recordset
    Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_Det_ProgChauf"
    Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "CodeProgChauf", 202, 25, adFldIsNullable
    L_RS.Fields.Append "CodeFournisseur", 202, 25, adFldIsNullable
    L_RS.Fields.Append "TxtCommande", 202, 200, adFldIsNullable
    L_RS.Fields.Append "TxtPaiement", 202, 50, adFldIsNullable
    L_RS.Fields.Append "TxtObservation", 202, 200, adFldIsNullable
    L_RS.Fields.Append "ProgOrder", 5, 25, adFldIsNullable
    L_RS.Open
    
    Set CreateEmptyRS_Det_ProgChauf = L_RS
End Function



Public Function CreateEmptyRS_DetailBC() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_DetailBC"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", adVarChar, 25
    L_RS.Fields.Append "Vehicule", adVarChar, 25
    L_RS.Fields.Append "Energie", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "CompteurCarburant", adInteger, 10, adFldIsNullable
    L_RS.Fields.Append "Litre", adDouble, 4, adFldIsNullable
    L_RS.Fields.Append "prixLitre", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "PrixHT", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "TVA", adNumeric, 19, adFldIsNullable
    L_RS.Fields.Append "Observation", adVarChar, 255, adFldIsNullable
    L_RS.Fields.Append "AnomalieConsom", adDouble, 19, adFldIsNullable

    L_RS.Open
    Set CreateEmptyRS_DetailBC = L_RS

End Function

Public Function CreateEmptyRS_AssBC() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_AssBC"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", adVarChar, 25
    L_RS.Fields.Append "DateDoc", adDate, 16, adFldIsNullable
    L_RS.Fields.Append "Station", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "Conducteur", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "Litre", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "Valeur", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "NumFact", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "transf", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "HEURE", adDate, 16, adFldIsNullable
    L_RS.Fields.Append "NBC", adInteger, 10, adFldIsNullable
    L_RS.Fields.Append "dateop", adDate, 16, adFldIsNullable
    L_RS.Fields.Append "UserInsert", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserDelete", adVarChar, 50, adFldIsNullable
    
    L_RS.Open
    Set CreateEmptyRS_AssBC = L_RS

End Function

Public Function CreateEmptyRS_LubBV() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_LubBV"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "Libelle", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "THT", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "TVA", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "prixTTC", adDouble, 19, adFldIsNullable
    
    L_RS.Open
    Set CreateEmptyRS_LubBV = L_RS

End Function

Public Function CreateEmptyRS_BV() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_BV"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", adVarChar, 25
    L_RS.Fields.Append "DateDoc", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "Vehicule", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "Station", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "Conducteur", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "Valeur", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "NumFact", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "Transf", adVarChar, 25, adFldIsNullable
    L_RS.Fields.Append "Heure", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "NBC", adInteger, 10, adFldIsNullable
    L_RS.Fields.Append "dateOp", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "CompteurVidange", adInteger, 10, adFldIsNullable
    L_RS.Fields.Append "NBKLMvid", adInteger, 10, adFldIsNullable
    L_RS.Fields.Append "UserInsert", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserDelete", adVarChar, 50, adFldIsNullable
    L_RS.Open
    Set CreateEmptyRS_BV = L_RS

End Function

Public Function CreateEmptyRS_Vehicule() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_Vehicule"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Code", adVarChar, 25
    L_RS.Fields.Append "Matricule", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "ABr", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "TYPE", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "Marque", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "Puissance", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "Energie", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "NumCartGris", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "DateCartGris", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "LieuCartGris", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "NumSerie", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "DateCircul", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "NumAssur", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "FournisAssur", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "AgenceAssur", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "DateDebAssur", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "DAteFinAssur", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "DateDebVisite", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "DAteFinVisite", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "DateDebTax", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "DateFinTax", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "DateSortie", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "Obs", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "genre", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "carrosserie", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "PlaceAssis", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "PlaceDebout", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "Cylindre", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "NbrEssieux", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "PTAC", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "PTRA", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "PoidsVide", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "TYPECOMM", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "Charge", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "NbKlmVidange", adInteger, 19, adFldIsNullable
    L_RS.Fields.Append "CompteurCarburant", adInteger, 19, adFldIsNullable
    L_RS.Fields.Append "CompteurVidange", adInteger, 19, adFldIsNullable
    L_RS.Fields.Append "CompteurFT", adInteger, 19, adFldIsNullable
    L_RS.Fields.Append "Actif", adInteger, 19, adFldIsNullable
    L_RS.Fields.Append "disponible", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "PicBox", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserInsert", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserDelete", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserAdd", adVarChar, 50, adFldIsNullable
    L_RS.Open
    Set CreateEmptyRS_Vehicule = L_RS

End Function
Public Function CreateEmptyRS_VehVdg() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_VehVdg"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", adInteger, 19
    L_RS.Fields.Append "Vehicule", adVarChar, 20
    L_RS.Fields.Append "Lubrifiant", adInteger, 19, adFldIsNullable
    L_RS.Fields.Append "UserInsert", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserDelete", adVarChar, 50, adFldIsNullable
   
    L_RS.Open
    Set CreateEmptyRS_VehVdg = L_RS

End Function
Public Function CreateEmptyRS_AssBCRepar() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_AssBCRepar"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", adVarChar, 25
    L_RS.Fields.Append "DateCreation", adDate, 16
    L_RS.Fields.Append "Fournisseur", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "Conducteur", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserInsert", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserDelete", adVarChar, 50, adFldIsNullable
    L_RS.Open
    Set CreateEmptyRS_AssBCRepar = L_RS

End Function

Public Function CreateEmptyRS_DetBCRepar() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_DetBCRepar"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    'L_RS.Fields.Append "Code", adInteger, 19
    L_RS.Fields.Append "Numero", adVarChar, 25
    L_RS.Fields.Append "désignation", adVarChar, 50
    L_RS.Fields.Append "Qté", adInteger, 19
    L_RS.Fields.Append "Vehicule", adVarChar, 50
    L_RS.Fields.Append "Observation", adVarChar, 254, adFldIsNullable
    L_RS.Open
    Set CreateEmptyRS_DetBCRepar = L_RS

End Function

Public Function CreateEmptyRS_AssPRepar() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_AssPRepar"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", adVarChar, 25
    L_RS.Fields.Append "Type", adVarChar, 50
    L_RS.Fields.Append "DatePiece", adDate, 19
    L_RS.Fields.Append "RemisePiece", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "totTTC", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "DateOperation", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "refPiece", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "timbre", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "Fournisseur", adVarChar, 50, adFldIsNullable
'    L_RS.Fields.Append "NumFact", adVarChar, 20, adFldIsNullable
    L_RS.Fields.Append "PrixMOeuvre", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "TVA_MOeuvre", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "UserInsert", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", adVarChar, 50, adFldIsNullable
    L_RS.Open
    Set CreateEmptyRS_AssPRepar = L_RS

End Function

Public Function CreateEmptyRS_DetPRepar() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_DetPRepar"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", adVarChar, 20
    L_RS.Fields.Append "Designation", adVarChar, 50
    L_RS.Fields.Append "Qte", adInteger, 19, adFldIsNullable
    L_RS.Fields.Append "Vehicule", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "PUHT", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "Remise", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "TVA", adDouble, 19, adFldIsNullable
    L_RS.Open
    Set CreateEmptyRS_DetPRepar = L_RS

End Function

Public Function CreateEmptyRS_Prod_Lub() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_Prod_Lub"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", adInteger, 19
    L_RS.Fields.Append "Libelle", adVarChar, 50
    L_RS.Fields.Append "prixht", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "DatePrix", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "tva", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "Type_PL", adVarChar, 20, adFldIsNullable
    L_RS.Fields.Append "Actif", adVarChar, 1, adFldIsNullable
    L_RS.Fields.Append "OperateurSaisi", adVarChar, 20, adFldIsNullable
    
    L_RS.Open
    Set CreateEmptyRS_Prod_Lub = L_RS

End Function

Public Function CreateEmptyRS_Energie() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_Energie"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Code", adVarChar, 20
    L_RS.Fields.Append "Libelle", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "tht", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "tva", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "Prix", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "UserInsert", adVarChar, 20, adFldIsNullable
    
    L_RS.Open
    Set CreateEmptyRS_Energie = L_RS

End Function

Public Function CreateEmptyRS_Station() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_Station"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Code", adVarChar, 50
    L_RS.Fields.Append "Libelle", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "Type", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "Adresse", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "Ville", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "CPOSTAL", adInteger, 19, adFldIsNullable
    L_RS.Fields.Append "Activite", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "TELEPHONE", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "MOBILE", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "FAX", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "EMAIL", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "Actif", adInteger, 19, adFldIsNullable
    L_RS.Fields.Append "UserInsert", 202, 30, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserDelete", 202, 30, adFldIsNullable
    L_RS.Fields.Append "UserAdd", 202, 30, adFldIsNullable
    L_RS.Open
    Set CreateEmptyRS_Station = L_RS

End Function


Public Function CreateEmptyRS_Fact() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_Fact"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", adVarChar, 20
    L_RS.Fields.Append "dateDoc", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "Station", adVarChar, 20, adFldIsNullable
    L_RS.Fields.Append "PeriodeDu", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "Periodeau", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "TTC_BC", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "TTC_BV", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "TTC", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "NBC", adInteger, 19, adFldIsNullable
    L_RS.Fields.Append "TTC_PR", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "dateOp", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "timbre", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "TTC_AV", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "TTC_BR", adDouble, 19, adFldIsNullable
    L_RS.Fields.Append "UserInsert", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserDelete", 202, 50, adFldIsNullable
    
    L_RS.Open
    Set CreateEmptyRS_Fact = L_RS

End Function


'----------------------------------------------------------------------------------------------------------------------------------
'Programme Traffic***
'----------------------------------------------------------------------------------------------------------------------------------
Public Function CreateEmptyRS_DispoPerso() As Recordset
    Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_DispoPerso"
    Dim L_RS As ADODB.Recordset
    
    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", 202, 25, adFldIsNullable
    L_RS.Fields.Append "Conducteur", 202, 50, adFldIsNullable
    L_RS.Fields.Append "Etat", 202, 30, adFldIsNullable
    L_RS.Fields.Append "HDebut", 202, 25, adFldIsNullable
    L_RS.Open
    
    Set CreateEmptyRS_DispoPerso = L_RS
End Function
Public Function CreateEmptyRS_Traffic() As Recordset
    Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_Traffic"
    Dim L_RS As ADODB.Recordset
    
    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", 202, 25, adFldIsNullable
    L_RS.Fields.Append "Vehicule", 202, 50, adFldIsNullable
    L_RS.Fields.Append "CompteurSortie", 202, 30, adFldIsNullable
    L_RS.Fields.Append "CompteurEntre", 202, 30, adFldIsNullable
    L_RS.Fields.Append "Conducteur", 202, 25, adFldIsNullable
    L_RS.Fields.Append "Destination", 202, 25, adFldIsNullable
    L_RS.Fields.Append "HeureSortie", 202, 50, adFldIsNullable
    L_RS.Fields.Append "OperateurSortie", 202, 30, adFldIsNullable
    L_RS.Fields.Append "HeureEntre", 202, 50, adFldIsNullable
    L_RS.Fields.Append "OperateurEntre", 202, 30, adFldIsNullable
    L_RS.Fields.Append "Observation", 202, 200, adFldIsNullable
    L_RS.Fields.Append "UserInsert", 202, 30, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserDelete", 202, 30, adFldIsNullable
    L_RS.Open
    Set CreateEmptyRS_Traffic = L_RS
End Function
'----------------------------------------------------------------------------------------------------------------------------------
'Conducteur / Personnel***
'----------------------------------------------------------------------------------------------------------------------------------
Public Function CreateEmptyRS_Conducteur() As Recordset
    Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_Conducteur"
    Dim L_RS As ADODB.Recordset
    
    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", 202, 25, adFldIsNullable
    L_RS.Fields.Append "Libelle", 202, 50, adFldIsNullable
    L_RS.Fields.Append "ABr", 202, 50, adFldIsNullable
    L_RS.Fields.Append "CIN", 202, 25 ', adFldIsNullable
    L_RS.Fields.Append "Fonction", 202, 30, adFldIsNullable
    L_RS.Fields.Append "Telephone", 202, 25 ', adFldIsNullable
    L_RS.Fields.Append "Mobile", 202, 25, adFldIsNullable
    L_RS.Fields.Append "Permie", 202, 50
    L_RS.Fields.Append "DateLivr", adDate, 30
    L_RS.Fields.Append "LieuLivr", 202, 50
    L_RS.Fields.Append "Actif", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Disponible", 202, 50, adFldIsNullable
    L_RS.Fields.Append "PicBox", 202, 200, adFldIsNullable
    L_RS.Fields.Append "Supp", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserInsert", 202, 30, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserDelete", 202, 30, adFldIsNullable
    L_RS.Open
                
    Set CreateEmptyRS_Conducteur = L_RS
End Function
'----------------------------------------------------------------------------------------------------------------------------------
'Destination***
'----------------------------------------------------------------------------------------------------------------------------------
Public Function CreateEmptyRS_Destination() As Recordset
    Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_Destination"
    Dim L_RS As ADODB.Recordset
    
    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", 202, 25, adFldIsNullable
    L_RS.Fields.Append "Type", 202, 50, adFldIsNullable
    L_RS.Fields.Append "Libelle", 202, 30, adFldIsNullable
    L_RS.Fields.Append "Actif", 5, 30, adFldIsNullable
    L_RS.Fields.Append "MaxDuree", 202, 25, adFldIsNullable
    L_RS.Fields.Append "MaxCompteur", 5, 25, adFldIsNullable
    L_RS.Fields.Append "MinCompteur", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Ord", 5, 25, adFldIsNullable
    L_RS.Fields.Append "UserInsert", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", 202, 30, adFldIsNullable
    L_RS.Fields.Append "Temps", 202, 10, adFldIsNullable
    L_RS.Open
                
    Set CreateEmptyRS_Destination = L_RS
End Function
'----------------------------------------------------------------------------------------------------------------------------------
'Utilisateur***
'----------------------------------------------------------------------------------------------------------------------------------
Public Function CreateEmptyRS_USER() As Recordset
    Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_USER"
    Dim L_RS As ADODB.Recordset
    
    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Code", 202, 20, adFldIsNullable
    L_RS.Fields.Append "MP", 202, 50, adFldIsNullable
    L_RS.Fields.Append "NomPrn", 202, 50, adFldIsNullable
    L_RS.Fields.Append "Ins_BC", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Maj_BC", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Supp_BC", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Consult_BC", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Ins_BV", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Maj_BV", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Supp_BV", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Consult_BV", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Consult_Alerte", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Ins_BCR", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Maj_BCR", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Supp_BCR", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Consult_BCR", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Ins_PR", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Maj_PR", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Supp_PR", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Consult_PR", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Ins_FF", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Maj_FF", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Supp_FF", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Consult_FF", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Consult_SC", 5, 50, adFldIsNullable
'    L_RS.Fields.Append "Consult_SR", 5, 30, adFldIsNullable
'    L_RS.Fields.Append "Consult_ST", 5, 30, adFldIsNullable
'    L_RS.Fields.Append "Consult_EHS", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Ins_FT", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Maj_FT", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Supp_FT", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Consult_FT", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Consult_Sup", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Ins_Vehicule", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Maj_vehicule", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Supp_vehicule", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Consult_vehicule", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Ins_Fournisseur", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Maj_Fournisseur", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Supp_Fournisseur", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Conslt_Fournisseur", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Ins_TC", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Maj_TC", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Supp_TC", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Consult_TC", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Ins_TV", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Maj_TV", 5, 50, adFldIsNullable
    L_RS.Fields.Append "supp_TV", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Consult_TV", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Ins_Destination", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Maj_Destination", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Supp_Destination", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Consult_Destination", 5, 30, adFldIsNullable
'    L_RS.Fields.Append "Ins_Lub", 5, 30, adFldIsNullable
'    L_RS.Fields.Append "Maj_Lub", 5, 25, adFldIsNullable
'    L_RS.Fields.Append "Supp_Lub", 5, 50, adFldIsNullable
'    L_RS.Fields.Append "Consult_Lub", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Ins_Produit", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Maj_produit", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Supp_Produit", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Consult_Produit", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Ins_Personnel", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Maj_Personnel", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Supp_personnel", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Consult_personnel", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Ins_Utilisateur", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Maj_Utilisateur", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Supp_Utilisateur", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Consult_Utilisateur", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Actif", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Maj_Disp", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Maj_Compt", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Consult_Compteurs", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Ins_PCH", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Maj_PCH", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Supp_PCH", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Consult_PCH", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Ins_PLING", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Maj_PLING", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Supp_PLING", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Consult_PLING", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Ins_Conge", 5, 50, adFldIsNullable
    L_RS.Fields.Append "Maj_Conge", 5, 30, adFldIsNullable
    L_RS.Fields.Append "Supp_Conge", 5, 25, adFldIsNullable
    L_RS.Fields.Append "Consult_Conge", 5, 25, adFldIsNullable
    L_RS.Fields.Append "UserInsert", 202, 50, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", 202, 50, adFldIsNullable
    L_RS.Open
    Set CreateEmptyRS_USER = L_RS
    
End Function
'----------------------------------------------------------------------------------------------------------------------------------
'PLANNING***
'----------------------------------------------------------------------------------------------------------------------------------
Public Function CreateEmptyRS_PLANNING() As Recordset
    Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_PLANNING"
    Dim L_RS As ADODB.Recordset
    
    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", 202, 25, adFldIsNullable
    L_RS.Fields.Append "Datedu", 202, 50, adFldIsNullable
    L_RS.Fields.Append "Dateau", 202, 50, adFldIsNullable
    L_RS.Fields.Append "DateCreat", 202, 50, adFldIsNullable
    L_RS.Fields.Append "DateEdit", 202, 50, adFldIsNullable
    L_RS.Fields.Append "Tournee", 202, 50, adFldIsNullable
    L_RS.Fields.Append "Jour", 202, 30, adFldIsNullable
    L_RS.Fields.Append "HeureEntre", 202, 50, adFldIsNullable
    L_RS.Fields.Append "Conducteur", 202, 50, adFldIsNullable
    L_RS.Fields.Append "Vehicule", 202, 50, adFldIsNullable
    L_RS.Fields.Append "DateJour", 202, 50, adFldIsNullable
    L_RS.Fields.Append "userinsert", 202, 30, adFldIsNullable
    L_RS.Fields.Append "userupdate", 202, 50, adFldIsNullable
    L_RS.Open
                
    Set CreateEmptyRS_PLANNING = L_RS
End Function
Public Function CreateEmptyRS_TmpPLANNING() As Recordset
    Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_TmpPLANNING"
    Dim L_RS As ADODB.Recordset
    
    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Datedu", 202, 50, adFldIsNullable
    L_RS.Fields.Append "Dateau", 202, 50, adFldIsNullable
    L_RS.Fields.Append "Tournee", 202, 50, adFldIsNullable
    L_RS.Fields.Append "Detail", 202, 300, adFldIsNullable
    L_RS.Open
                
    Set CreateEmptyRS_TmpPLANNING = L_RS
End Function

Public Function CreateEmptyRS_Conge() As Recordset

Const sPROC_NAME As String = OBJ_NAME & ".CreateEmptyRS_Conge"

Dim L_RS As ADODB.Recordset

    Set L_RS = New ADODB.Recordset
    L_RS.Fields.Append "Numero", adInteger, 19
    L_RS.Fields.Append "Conducteur", adVarChar, 20
    L_RS.Fields.Append "DateDu", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "DateAu", adDate, 19, adFldIsNullable
    L_RS.Fields.Append "Type", adVarChar, 10, adFldIsNullable
    L_RS.Fields.Append "Observation", adVarChar, 255, adFldIsNullable
    L_RS.Fields.Append "UserInsert", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserUpdate", adVarChar, 50, adFldIsNullable
    L_RS.Fields.Append "UserDelete", adVarChar, 50, adFldIsNullable

    L_RS.Open
    Set CreateEmptyRS_Conge = L_RS

End Function


