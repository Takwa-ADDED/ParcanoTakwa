VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PAR_DOWNLOAD 
   Caption         =   "Téléchargement"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PAR_DOWNLOAD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      Begin MSComctlLib.ProgressBar PROG 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txt_Nouvelle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   3165
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   930
         Width           =   6015
      End
      Begin VB.Label Lab 
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.CommandButton Fermer 
      Caption         =   "Fermer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "PAR_DOWNLOAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VER As Long

Public THEFILE As String
Public TOFILE As String
Dim FOIS As Byte
Dim TELECHARGEMENT As Boolean

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_NOSIZE = &H1
Const SWP_SHOWWINDOW = &H40
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1

Private Sub Form_Load()

    Dim EXE As String
    EXE = App.EXEName
    FOIS = 0
    TELECHARGEMENT = False
    Me.Caption = " Mise à jour du " & EXE
    
End Sub
Private Sub Form_Activate()

On Error GoTo ER
 
    Me.Refresh
    If FOIS = 0 Then
        Call AFFICHE_NOUVEAUTE_VERSION
        ShowOn True
        FOIS = FOIS + 1

        If CopyFileAny(THEFILE, TOFILE) = True Then
            If CopyDLL(THEFILE, TOFILE, "GestionParc.dll") = True Then
                SaveLog_maj THEFILE, "Version " & GetOnlyVersion(THEFILE) & " télécharger par " & ComputerName
                Lab(9).Caption = "Téléchargement terminé avec succès"
                TELECHARGEMENT = True
            Else
                On Error Resume Next
                Kill TOFILE
                On Error GoTo ER
                Lab(9).Caption = "Erreur pendant le Téléchargement de DLL"
                TELECHARGEMENT = False
                Unload Me
                frmconnexionx.Show
                Exit Sub
            End If
        Else
            Lab(9).Caption = "Erreur pendant le Téléchargement"
            TELECHARGEMENT = False
        End If
        
        Pause_Timer 2
        Fermer.Enabled = True
    End If
    
Exit Sub
ER:
    MsgBox Err.Description, vbInformation
    TELECHARGEMENT = False
End Sub
Private Sub AFFICHE_NOUVEAUTE_VERSION()

    Dim A%, buffer%, temp$, fRead&, fSize&, b%
    Dim FOIS As Long
    Dim currentFilename As String
    Dim N As String
    N = ""
    FOIS = 0
    currentFilename = GetSetting("CentraNord", "GestParc", "CHEMIN_MAJ", "") & "Description.txt"
    A = FreeFile
    buffer = 4048
    Open currentFilename For Binary Access Read As A

    fSize = FileLen(currentFilename)

    While fRead < fSize
        If buffer > (fSize - fRead) Then buffer = (fSize - fRead)
        temp = Space(buffer)
        Get A, , temp
        N = N & temp
        fRead = fRead + buffer
    Wend
    
    Close A
    txt_Nouvelle.Text = N
End Sub

Private Sub ShowOn(ByVal TopMost As Boolean)
On Error Resume Next
    If TopMost Then
        Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
        SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
    Else
        Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
    End If
End Sub
Private Sub Fermer_Click()
    On Error Resume Next
    Shell ("regsvr32 " & App.Path & "\GestionParc.dll /s")
    Unload Me
End Sub

Private Sub Pause_Timer(PauseTime As Single)
Dim start
    start = Timer   ' Définit l'heure de début.
    Do While Timer < start + PauseTime
        DoEvents    ' Donne le contrôle à d'autres processus.
    Loop
End Sub

Private Function CopyDLL(ByVal currentFilename As String, ByVal newFilename As String, DLLFILE As String) As Boolean
    
    currentFilename = Replace(LCase(currentFilename), LCase(App.EXEName & ".exe"), DLLFILE)
    newFilename = Replace(LCase(newFilename), LCase("000000_" & App.EXEName & ".exe"), DLLFILE)
    
    CopyDLL = CopyFileAny(currentFilename, newFilename)
    
End Function

Private Function CopyFileAny(currentFilename As String, newFilename As String) As Boolean
    Dim A%, buffer%, temp$, fRead&, fSize&, b%
    On Error GoTo ErrHan:
    A = FreeFile
    buffer = 4048
    Open currentFilename For Binary Access Read As A
    b = FreeFile
    Open newFilename For Binary Access Write As b
    fSize = FileLen(currentFilename)
    
        PROG.Max = fSize

        While fRead < fSize
            
            Lab(9).Caption = Format((Int(fRead) / fSize) / 100, "0") & " %"
            PROG.Value = Int(fRead)
    
            If buffer > (fSize - fRead) Then buffer = (fSize - fRead)
            temp = Space(buffer)
            Get A, , temp
            Put b, , temp
            fRead = fRead + buffer
        Wend
                
        Close b
        Close A
        CopyFileAny = True
        
        MouseOff
        Exit Function
ErrHan:
        Close b
        Close A
        CopyFileAny = False
        
        MouseOff
        Me.Refresh
End Function

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    If TELECHARGEMENT = True Then
        'Shell TOFILE, vbNormalFocus
        'fermer Parcano.exe et ouvrir 000000_Parcano.exe
        Shell ("regsvr32 " & App.Path & "\GestionParc.dll /s"), vbNormalFocus
        Shell (App.Path & "\000000_Parcano.exe"), vbNormalFocus
        MouseOff
        End
    End If
    
End Sub

Private Sub MouseOn()
    Screen.MousePointer = vbHourglass
End Sub

Private Sub MouseOff()
    Screen.MousePointer = vbDefault
End Sub
  
Private Function ComputerName() As String

    ' Retourne le nom de l'ordinateur
    Dim stTmp As String, lgTmp As Long
    stTmp = Space$(250)
    lgTmp = 251
    Call GetComputerName(stTmp, lgTmp)
    ComputerName = Split(stTmp, Chr$(0))(0)
    
End Function

