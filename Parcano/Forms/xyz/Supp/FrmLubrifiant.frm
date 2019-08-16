VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Begin VB.Form FrmLubrifiant 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Parcano"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLubrifiant.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   10260
   Begin SToolBox.SGrid Grid 
      Height          =   3975
      Left            =   1680
      TabIndex        =   10
      Top             =   3360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7011
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisableIcons    =   -1  'True
      MaxVisibleRows  =   0
   End
   Begin SToolBox.SCommand cmdFindMatricule 
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   1560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmLubrifiant.frx":0ECA
      ButtonType      =   1
   End
   Begin VB.TextBox txt_libelle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "M"
      Top             =   2280
      Width           =   7695
   End
   Begin VB.TextBox txt_Matricule 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   1680
      TabIndex        =   0
      Tag             =   "M"
      Top             =   1560
      Width           =   2295
   End
   Begin SToolBox.SCommand CmdSave 
      Height          =   495
      Left            =   9720
      TabIndex        =   3
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmLubrifiant.frx":121D
   End
   Begin SToolBox.SCommand CmdDelete 
      Height          =   495
      Left            =   8880
      TabIndex        =   4
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmLubrifiant.frx":139F
   End
   Begin SToolBox.SCommand CmdFind 
      Height          =   495
      Left            =   9240
      TabIndex        =   5
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmLubrifiant.frx":16F2
   End
   Begin SToolBox.SCommand CmdAjout 
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   2820
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   " Ajouter"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmLubrifiant.frx":1A45
      ButtonType      =   1
   End
   Begin SToolBox.SCommand CmdEnleve 
      Height          =   495
      Left            =   8280
      TabIndex        =   12
      Top             =   2820
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   "Enlever"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmLubrifiant.frx":1BC7
      ButtonType      =   1
   End
   Begin SToolBox.SCommand CmdAdd 
      Height          =   495
      Left            =   8520
      TabIndex        =   13
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmLubrifiant.frx":1D49
   End
   Begin VB.Line Line3 
      X1              =   7080
      X2              =   1680
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      X1              =   7080
      X2              =   7080
      Y1              =   2760
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   9360
      X2              =   7080
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Produits:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   885
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Libelle :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Matricule"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiche type vidange"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2685
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmLubrifiant.frx":1ECB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "FrmLubrifiant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAdd_Click()
On Error GoTo Err

Dim rQ As New ADODB.Recordset
Dim SQL As String
SQL = "Select * from utilisateur where Ins_TV = 1 and code= " & LInt_UserId
rQ.Open SQL, CNB, adOpenDynamic
If rQ.EOF Then
    rQ.Close
    MsgBox "Accès refusé.", vbExclamation
    Exit Sub
End If


If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Call ViderZone(FrmLubrifiant)
'txt_Matricule.Enabled = False
txt_Matricule.Text = "Auto"
txt_libelle.SetFocus
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdAjout_Click()
With FrmSaisieProduit
    .Show
End With
End Sub

Private Sub CmdDelete_Click()

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim vcode As String

On Error GoTo Err
   If txt_Matricule.Text = "Auto" Then
        If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        Else
            txt_Matricule.SetFocus
            Exit Sub
        End If
    End If
    If txt_Matricule.Text <> "Auto" Then
        Dim rQ As New ADODB.Recordset
        SQL = "Select * from utilisateur where supp_TV = 1 and code= " & LInt_UserId
        rQ.Open SQL, CNB, adOpenDynamic
        If rQ.EOF Then
            rQ.Close
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
    End If
    If MsgBox("Confirmez vous la suppression de ce type de vidange", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    vcode = txt_Matricule.Text
    SQL = "Delete from lubrifiant where code =" & SQLText(vcode)
    CNB.Execute SQL
    SQL = "Delete from detlubrifiant where codelubrifiant  =" & SQLText(vcode)
    CNB.Execute SQL
    txt_Matricule.SetFocus
    End If
    
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdEnleve_Click()

Dim ii As Integer

On Error GoTo Err
        
       ii = Grid.SelectedRow
       If ii > 0 Then
        Grid.RemoveRow ii
        
       End If
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub CmdFind_Click()
On Error Resume Next
If txt_Matricule.Text = "Auto" Then
    If MsgBox("Annuler la création en cour.?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

Unload FrmFind
With FrmFind
    .StrSource = "Lubrifiant"
    .Show
End With
End Sub

Private Sub cmdFindMatricule_Click()
Unload FrmFind
With FrmFind
    .StrSource = "Lubrifiant"
    .Show
End With
End Sub

Private Sub CmdSave_Click()

Dim SQL As String
Dim rs As New ADODB.Recordset
Dim vcodep
Dim ii
Dim Qte As Integer

On Error GoTo Err

    If Left(CheckMandatory(FrmLubrifiant), 1) = 1 Then
       Exit Sub
    End If
    If Grid.Rows = 0 Then
        MsgBox "Aucun produit saisi pour ce type de vidange" & vbNewLine & "Impossible de continuer", vbExclamation
        Exit Sub
    End If
    If txt_Matricule.Text <> "Auto" Then
        Dim rQ As New ADODB.Recordset
        SQL = "Select * from utilisateur where Maj_TV = 1 and code= " & LInt_UserId
        rQ.Open SQL, CNB, adOpenDynamic
        If rQ.EOF Then
            rQ.Close
            MsgBox "Accès refusé.", vbExclamation
            Exit Sub
        End If
    End If
    If MsgBox("Confirmez vous l'enregistrement", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
    'Delete l'enregistrement
    vcode = txt_Matricule.Text
    CNB.BeginTrans
    SQL = "Delete from Lubrifiant where code =" & SQLText(vcode)
    CNB.Execute SQL
    SQL = "Delete from DetLubrifiant where codeLubrifiant =" & SQLText(vcode)
    CNB.Execute SQL
    If vcode = "Auto" Then
    LInt_NumCompteur = Crement_Compteur(ErrNumber, ErrDescription, ErrSourceDetail, CNB, "NextValCounter", "F_Lubrifiant")
    If ErrNumber <> 0 Then
       MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
       ErrNumber = 0
       Exit Sub
    End If
    Set LObj_Compteur = Nothing
    'Insertion enregistrement assiette
    txt_Matricule.Text = Format(LInt_NumCompteur, "00000")
    End If

    'Insertion enregistrement
    SQL = "Insert into Lubrifiant  (Code,Libelle) values ("
    SQL = SQL & SQLText(txt_Matricule.Text)
    SQL = SQL & "," & SQLText(txt_libelle.Text)
    SQL = SQL & ")"
    CNB.Execute SQL
    'Insert details
    For ii = 1 To Grid.Rows
        vcodep = Grid.CellText(ii, 1)
        Qte = Grid.CellText(ii, 3)
        SQL = "Insert into detLubrifiant (CodeLubrifiant,Codeproduit,Qte) values("
        SQL = SQL & SQLText(vcode) & "," & SQLText(vcodep) & "," & Qte & ")"
        CNB.Execute SQL
    Next
    CNB.CommitTrans
    MsgBox "Enregistrement terminé avec succé  ", vbQuestion, App.ProductName
    txt_Matricule.SetFocus
    End If
    
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub

Private Sub Form_Load()
Me.Width = 10380
Me.Height = 7935
Call Initgrid
Me.Move 0, 0
Me.WindowState = 2
End Sub

Private Sub Form_Resize()
On Error Resume Next
Image1.Width = Me.Width
CmdSave.Left = Me.Width - 700
CmdFind.Left = Me.Width - 1100
CmdDelete.Left = Me.Width - 1500
CmdAdd.Left = Me.Width - 1900
End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error GoTo erreur
   Dim i As Integer
   Dim MSG ' Déclare la variable.
   ' Définit le texte du message.
   MSG = "Voulez-vous vraiment quitter?"
   ' Si l'utilisateur clique sur Non, met fin à l'événement QueryUnload.
   If MsgBox(MSG, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
      Cancel = True
   Else
   Unload Me
   End If
   
   Exit Sub
erreur:
   MsgBox Err.Description, 48
End Sub

Private Sub txt_libelle_GotFocus()
If Len(Trim(txt_Matricule.Text)) = 0 Then
    MsgBox "N° matricule obligatoire      ", vbInformation
    txt_Matricule.SetFocus
End If
End Sub

Private Sub Txt_Libelle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txt_Matricule_GotFocus()
Call ViderZone(FrmLubrifiant)
Grid.ClearRows
End Sub

Private Sub txt_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub Initgrid()
With Grid
    ' Allow the grid to be grouped, but
    ' don't show the grouping box
    .HideGroupingBox = True
    .AllowGrouping = True
    ' Group rows will be shown by
    ' a gradient underline
    .GroupRowBackColor = vbWindowBackground
    .GroupRowForeColor = vbWindowText
    
    .GridLineColor = vbWindowBackground
    .GridFillLineColor = vbWindowBackground
    .GridLines = True
    
    .SelectionAlphaBlend = True
    .SelectionOutline = True
    .DrawFocusRectangle = False
    
    .AddColumn "Code", "Code", , , 60, False, , , , , , CCLSortNumeric
    .AddColumn "Libelle", "Libelle", , , 240
    .AddColumn "Qte", "Qte", , , 60
    .AddColumn "Prix", "Prix.TTC", , , 80
  
    .AddColumn "Q", ""
    .StretchLastColumnToFit = True
End With



End Sub

Public Sub AfficheRow(ByVal vcode As String)

Dim SQL As String
Dim rs As New ADODB.Recordset
Call ViderZone(FrmLubrifiant)
Grid.ClearRows
SQL = "Select * from Lubrifiant where code = " & SQLText(vcode)
rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    'Charge
    txt_Matricule.Text = rs("Code")
    txt_libelle.Text = rs("Libelle")
Else
    MsgBox "Code introuvable", vbInformation
    txt_Matricule.SetFocus
    Exit Sub
End If
rs.Close
'Charge details
SQL = "SELECT DetLubrifiant.CodeProduit, Produit.Libelle, Produit.Prix,DetLubrifiant.Qte"
SQL = SQL & " From DetLubrifiant"
SQL = SQL & " INNER JOIN  Produit ON DetLubrifiant.CodeProduit = Produit.Code"
SQL = SQL & " WHERE DetLubrifiant.CodeLubrifiant = " & SQLText(vcode)

rs.Open SQL, CNB, adOpenKeyset
If Not rs.EOF Then
    While Not rs.EOF
        With Grid
            .AddRow
            .CellDetails .Rows, 1, rs("CodeProduit")
            .CellDetails .Rows, 2, rs("Libelle")
            .CellDetails .Rows, 3, rs("Qte")
            .CellDetails .Rows, 4, Format(rs("Prix"), "#,##0.000")
        End With
        rs.MoveNext
    Wend
End If
rs.Close

End Sub

Private Sub txt_Matricule_LostFocus()

On Error GoTo Err

If Len(Trim(txt_Matricule.Text)) > 0 Then Call AfficheRow(txt_Matricule.Text)
    
Exit Sub
Err:
MsgBox Err.Description, vbInformation

End Sub


