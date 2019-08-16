VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form Frm_Rpt_Apercus 
   Caption         =   "Aperçu avant impression ..."
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Rpt_Apercus.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4455
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "Frm_Rpt_Apercus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public Numero As String
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'|||||||||||||||||||||||||||||||||||||||||||||||||
'Compteurs***
Public Sub PrintOutAndApercu_Compteurs(ByVal LBit_mode As Byte)
    Dim LObj_Imp As New Imprimer
    Dim Report As New Rpt_Compteurs
    Dim rs As Command
On Error GoTo HandlErreur
    Set rs = LObj_Imp.PrintCompteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Report.FormulaFields.GetItemByName("SOCIETE").Text = SQLText("Centra Nord Bizerte")
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1
    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    MouseOff
Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'|||||||||||||||||||||||||||||||||||||||||||||||||
'Programe Chauffeur***
Public Sub PrintOutAndApercu_ProgChauf(ByVal LBit_mode As Byte)
    Dim Report As New RPT_PH
    Dim LObj_Imp As New Imprimer
    Dim Lrs_Imp As Command
On Error GoTo HandlErreur
    Set LObj_Imp = New Imprimer
    Set Lrs_Imp = LObj_Imp.PrintProgChauf(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Numero)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Imp = Nothing
    
    Report.FormulaFields.GetItemByName("SOCIETE").Text = SQLText("Centra Nord Bizerte")
    Report.Database.AddADOCommand CNR, Lrs_Imp
    Report.AutoSetUnboundFieldSource 1
    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'|||||||||||||||||||||||||||||||||||||||||||||||||
'Trafic Anomalie***
Public Sub PrintOutAndApercu_AnomalieTrafic(ByVal LBit_mode As Byte, _
                                            ByVal DATEDEBUT As Date, _
                                            ByVal DateFin As Date, _
                                            ByVal Conducteur As String, _
                                            ByVal VEHICULE As String, _
                                            ByVal DESTINATION As String, _
                                            ByVal User As String, _
                                            ByVal AnKm As String, _
                                            ByVal AnDuree As String, _
                                            ByVal AnTotal As String, _
                                            ByVal anomali As Boolean)
                                            
    Dim Report As New Rpt_Anomalie
    Dim rs As Command
    Dim LObj_Find As Imprimer
    Dim Name_Table As String
    Dim YearTrafic As Integer
On Error GoTo HandlErreur
    For YearTrafic = Year(DATEDEBUT) To Year(DateFin)
        Name_Table = "FicheTraffic"
        If YearTrafic < Year(Date) Then Name_Table = "FicheTraffic_" & YearTrafic
        Set LObj_Find = New Imprimer
        Set rs = LObj_Find.PrintAnomalieTrafic(ErrNumber, ErrDescription, ErrSourceDetail, Name_Table, DATEDEBUT, DateFin, Conducteur, VEHICULE, DESTINATION, anomali, CNB)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
    Next
    Report.FormulaFields.GetItemByName("SOCIETE").Text = SQLText("Centra Nord Bizerte")
    Report.FormulaFields.GetItemByName("User").Text = SQLText(User)
    Report.FormulaFields.GetItemByName("DateDebut").Text = SQLText(DATEDEBUT)
    Report.FormulaFields.GetItemByName("DateFin").Text = SQLText(DateFin)
    Report.FormulaFields.GetItemByName("AnKm").Text = SQLText(AnKm)
    Report.FormulaFields.GetItemByName("AnDuree").Text = SQLText(AnDuree)
    Report.FormulaFields.GetItemByName("AnTotal").Text = SQLText(AnTotal)
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1
    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'|||||||||||||||||||||||||||||||||||||||||||||||||
'PLANNING***
Public Sub PrintOutAndApercu_PLANNING(ByVal LBit_mode As Byte, ByVal DateDu As Date, ByVal DateAu As Date, ByVal TypePLNG As String)
    Dim Report
    If TypePLNG = "Abreviation" Then
         Set Report = New RPT_PLNG
    Else
        Set Report = New RPT_PLNG_Normal
    End If
    Dim rs As Command, LObj_Find As Imprimer
On Error GoTo HandlErreur
    Set LObj_Find = New Imprimer
    Set rs = LObj_Find.PrintPLANNING(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DateDu)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    
    Report.FormulaFields.GetItemByName("SOCIETE").Text = SQLText("Centra Nord Bizerte")
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1
Set rs = Nothing
    Dim Report1 As New Rpt_REPOS
On Error GoTo HandlErreur
    Set LObj_Find = New Imprimer
    Set rs = LObj_Find.PrintREPOS(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DateDu)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing

    Report.Subreport1.OpenSubreport.Database.AddADOCommand CNR, rs
    Report.Subreport1.OpenSubreport.AutoSetUnboundFieldSource 1
    Set rs = Nothing
    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    MouseOff
Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub

















'============== Impression bon de carburant ==============
'========================================================
Public Sub PrintOutAndApercu_BC(ByVal LBit_mode As Byte)

On Error GoTo HandlErreur

'==============
'==============
Dim LObj_CReport As CReport
Dim Report As New Rpt_BC_Details1
Dim rs As Command

Set LObj_CReport = New CReport
Set rs = LObj_CReport.PrintOutBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Numero)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
Report.FormulaFields.GetItemByName("SOCIETE").Text = SQLText("Centra Nord Bizerte")
Report.Database.AddADOCommand CNR, rs
Report.AutoSetUnboundFieldSource 1
    
'==============
'==============

    Dim Report2 As New Rpt_BC_Details2
    Dim RSS As Command
    
    Set RSS = LObj_CReport.PrintOutBC2(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Numero)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    
    Report.Subreport1.OpenSubreport.Database.AddADOCommand CNR, RSS
    Report.Subreport1.OpenSubreport.AutoSetUnboundFieldSource 1

'==============
'==============

CRViewer1.ReportSource = Report
CRViewer1.Zoom (180)

If LBit_mode = 1 Then
    Report.PrintOut False
Else
    CRViewer1.ViewReport
End If
    
    Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub
'========================================================
'Impression bon de vidange
'========================================================
Public Sub PrintOutAndApercu_BV2(ByVal LBit_mode As Byte)
   
Dim LObj_CReport As CReport
Dim Report As New Rpt_BV2
Dim rs As Command

On Error GoTo HandlErreur
Set LObj_CReport = New CReport
Set rs = LObj_CReport.PrintOutBV(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Numero)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If

    '*
    Report.FormulaFields.GetItemByName("SOCIETE").Text = SQLText("Centra Nord Bizerte")
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1

    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    
Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub
'========================================================
'Impression Facture
'========================================================
Public Sub PrintOutAndApercu_RECAP_Facture(ByVal LBit_mode As Byte, _
                                            ByVal tht As Double, _
                                            ByVal tva As Double, _
                                            ByVal ttc As Double)

   
    Dim LObj_CReport As CReport
    Dim Report As New RPT_RECAP_FACTURE
On Error GoTo HandlErreur

    Dim rs As Command
    
    On Error GoTo HandlErreur
    Set LObj_CReport = New CReport
    Set rs = LObj_CReport.PrintOutFact(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Numero)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If

    '*
    Report.FormulaFields.GetItemByName("tht").Text = SQLText(tht)
    Report.FormulaFields.GetItemByName("tva").Text = SQLText(tva)
    Report.FormulaFields.GetItemByName("ttc").Text = SQLText(ttc)
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1
 
    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    
    Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub

'========================================================
'Impression bon de commande de réparation
'========================================================
Public Sub PrintOutAndApercu_BCReparation(ByVal LBit_mode As Byte)
   
Dim LObj_CReport As CReport
Dim Report As New Rpt_BCREparation
Dim rs As Command

On Error GoTo HandlErreur
Set LObj_CReport = New CReport
Set rs = LObj_CReport.PrintOutBRepar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Numero)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
    
    '*
    Report.FormulaFields.GetItemByName("SOCIETE").Text = SQLText("Centra Nord Bizerte")
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1

    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    
'    MouseOff
    
    Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub
'========================================================
'Impression bon de pièce de réparation
'========================================================
Public Sub PrintOutAndApercu_PieceRepa(ByVal LBit_mode As Byte, _
                            ByVal TotHTBrut As Double, _
                            ByVal TotRemLigne As Double, _
                            ByVal RemiseP As Double, _
                            ByVal TotHtNet As Double, _
                            ByVal TotTva As Double, _
                            ByVal TotTTC As Double, _
                            ByVal MainOeuvre As Double, _
                            ByVal TvaMOeuvre As Double)

Dim LObj_CReport As CReport
Dim Report As New Rpt_PieceReparation
Dim rs As Command
    ' Open the data connection
On Error GoTo HandlErreur
Set LObj_CReport = New CReport
Set rs = LObj_CReport.PrintOutPieceRepar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Numero, MainOeuvre)
If ErrNumber <> 0 Then
    MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
    ErrNumber = 0
    Exit Sub
End If
    '*
    Report.FormulaFields.GetItemByName("TotHTBrut").Text = SQLText(TotHTBrut)
    Report.FormulaFields.GetItemByName("TotRemLigne").Text = SQLText(TotRemLigne)
    Report.FormulaFields.GetItemByName("RemiseP").Text = SQLText(RemiseP)
    Report.FormulaFields.GetItemByName("TotHtNet").Text = SQLText(TotHtNet)
    Report.FormulaFields.GetItemByName("TotTva").Text = SQLText(TotTva)
    Report.FormulaFields.GetItemByName("TotTTC").Text = SQLText(TotTTC)
'    Report.FormulaFields.GetItemByName("MainOeuvre").Text = SQLText(MainOeuvre)
'    Report.FormulaFields.GetItemByName("TvaMOeuvre").Text = SQLText(TvaMOeuvre)
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1

    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If

    Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub




'=================================================
'Statistiques Services ***
'=================================================
Public Sub PrintOutAndApercu_StatService(ByVal LBit_mode As Byte, _
                                            ByVal DATEDEBUT As Date, _
                                            ByVal DateFin As Date, _
                                            ByVal Conducteur As String, _
                                            ByVal User As String)
                                            
    Dim Report As New Rpt_StatService
    Dim rs As Command
    Dim LObj_Find As Imprimer
    Dim Name_Table As String
    Dim YearTrafic As Integer
On Error GoTo HandlErreur

     For YearTrafic = Year(DATEDEBUT) To Year(DateFin)
        Name_Table = "FicheTraffic"
        If YearTrafic < Year(Date) Then Name_Table = "FicheTraffic_" & YearTrafic

    Set LObj_Find = New Imprimer
    Set rs = LObj_Find.Print_StatServices(ErrNumber, ErrDescription, ErrSourceDetail, Name_Table, DATEDEBUT, DateFin, Conducteur, CNB)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    
    Next
    
    Report.FormulaFields.GetItemByName("User").Text = SQLText(User)
    Report.FormulaFields.GetItemByName("DateDebut").Text = SQLText(DATEDEBUT)
    Report.FormulaFields.GetItemByName("DateFin").Text = SQLText(DateFin)
    
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1

    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    
Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub

'=================================================
'REPOS***
'=================================================
Public Sub PrintOutAndApercu_REPOS(ByVal LBit_mode As Byte, ByVal DateDu As Date, ByVal DateAu As Date)
    Dim Report As New Rpt_REPOS
    Dim rs As Command
    Dim LObj_Find As Imprimer

On Error GoTo HandlErreur

    Set LObj_Find = New Imprimer
    Set rs = LObj_Find.PrintREPOS(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DateDu)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    
    
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1

    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    
Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub
'=======================Conge***======================
'=====================================================
Public Sub PrintOutAndApercu_Conge(ByVal LBit_mode As Byte, _
                                    ByVal DATEDEBUT As Date, _
                                    ByVal DateFin As Date, _
                                    ByVal Conducteur As String, _
                                    ByVal User As String)
                                            
    Dim Report As New Rpt_Conge
    Dim rs As Command
    Dim LObj_Find As CReport
On Error GoTo HandlErreur

    Set LObj_Find = New CReport
    Set rs = LObj_Find.PrintOut_CongeConduc(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DATEDEBUT, DateFin)
    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If
    Set LObj_Find = Nothing
    
    Report.FormulaFields.GetItemByName("DateDebut").Text = SQLText(DATEDEBUT)
    Report.FormulaFields.GetItemByName("DateFin").Text = SQLText(DateFin)
    Report.FormulaFields.GetItemByName("User").Text = SQLText(User)
    
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1

    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    
Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub

'=================================================
'Statistiques carburant***
'=================================================
Public Sub PrintOutAndApercu_StatCarb(ByVal LBit_mode As Byte, _
                                    ByVal DATEDEBUT As Date, _
                                    ByVal DateFin As Date, _
                                    ByVal VCode As String, _
                                    ByVal User As String, _
                                    ByVal TotLitre As Double, _
                                    ByVal Total As Double)
                                            
    Dim Report As New Rpt_StatCarb
    Dim rs As Command
    Dim LObj_Find As CReport
On Error GoTo HandlErreur

    Set LObj_Find = New CReport
    If VCode = "Tous" Or VCode = "" Then
        Set rs = LObj_Find.Get_StatistBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DATEDEBUT, DateFin)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
    Else
        Set rs = LObj_Find.Get_StatistBCVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DATEDEBUT, DateFin, VCode)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
    End If
    
    Report.FormulaFields.GetItemByName("DateDebut").Text = SQLText(DATEDEBUT)
    Report.FormulaFields.GetItemByName("DateFin").Text = SQLText(DateFin)
    Report.FormulaFields.GetItemByName("User").Text = SQLText(User)
    Report.FormulaFields.GetItemByName("TotLitre").Text = SQLText(TotLitre)
    Report.FormulaFields.GetItemByName("Total").Text = SQLText(Total)
    
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1

'==============
'==============

 Dim Report2 As New Rpt_StatBCTot
    Dim RSS As Command
    
    Set LObj_Find = New CReport
    If VCode = "Tous" Or VCode = "" Then
        Set RSS = LObj_Find.Print_StatistBCTot(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DATEDEBUT, DateFin)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
    Else
        Set RSS = LObj_Find.Print_StatistBCTotVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DATEDEBUT, DateFin, VCode)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
    End If
    
    Report.Subreport1.OpenSubreport.Database.AddADOCommand CNR, RSS
    Report.Subreport1.OpenSubreport.AutoSetUnboundFieldSource 1
'==============
'==============
    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    
Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub

'=================================================
'Statistiques Réparation***
'=================================================
Public Sub PrintOutAndApercu_StatRep(ByVal LBit_mode As Byte, _
                                    ByVal DATEDEBUT As Date, _
                                    ByVal DateFin As Date, _
                                    ByVal VCode As String, _
                                    ByVal User As String, _
                                    ByVal nbrRep As Double, _
                                    ByVal Total As Double)
                                            
    Dim Report As New Rpt_StatRepar
    Dim rs As Command
    Dim LObj_Find As CReport
    
On Error GoTo HandlErreur

    Set LObj_Find = New CReport
    If VCode = "Tous" Or VCode = "" Then
        Set rs = LObj_Find.Print_DetRepStatist(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DATEDEBUT, DateFin)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
    Else
        Set rs = LObj_Find.Print_DetRepStatVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode, DATEDEBUT, DateFin)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing
    End If
    
    Report.FormulaFields.GetItemByName("DateDebut").Text = SQLText(DATEDEBUT)
    Report.FormulaFields.GetItemByName("DateFin").Text = SQLText(DateFin)
    Report.FormulaFields.GetItemByName("User").Text = SQLText(User)
    Report.FormulaFields.GetItemByName("NbrRep").Text = SQLText(nbrRep)
    Report.FormulaFields.GetItemByName("Total").Text = SQLText(Total)
    
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1
    
'==============
'==============
    Dim Report2 As New Rpt_Repar
    Dim RSS As Command
    
    Set LObj_Find = New CReport
    If VCode = "Tous" Or VCode = "" Then
        Set RSS = LObj_Find.Print_RepStat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, DATEDEBUT, DateFin)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
    Else
        Set RSS = LObj_Find.Print_RepStatVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode, DATEDEBUT, DateFin)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
    End If
    
    Report.Subreport1.OpenSubreport.Database.AddADOCommand CNR, RSS
    Report.Subreport1.OpenSubreport.AutoSetUnboundFieldSource 1

'==============
'==============
    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    
Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub

Public Sub PrintAlertes(ByVal LBit_mode As Byte)


    Dim Report As New Rpt_Alertes
    Dim rs As Command
    Dim LObj_Find As CReport

On Error GoTo HandlErreur

    Set LObj_Find = New CReport

    Set rs = LObj_Find.Print_Alertes(ErrNumber, ErrDescription, ErrSourceDetail, CNB)
        If ErrNumber <> 0 Then
            MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
            ErrNumber = 0
            Exit Sub
        End If
        Set LObj_Find = Nothing

    Report.FormulaFields.GetItemByName("SOCIETE").Text = SQLText("Centra Nord Bizerte")
    Report.Database.AddADOCommand CNR, rs
    Report.AutoSetUnboundFieldSource 1
    CRViewer1.ReportSource = Report
    CRViewer1.Zoom (180)
    If LBit_mode = 1 Then
        Report.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    MouseOff
Exit Sub
HandlErreur:
    MsgBox Err.Description, vbInformation
End Sub
