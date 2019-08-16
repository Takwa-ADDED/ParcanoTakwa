Attribute VB_Name = "check_ver"
Option Explicit

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
Private Declare Function lstrcpy Lib "Kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Type FILEINFO
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OriginalFileName As String
    ProductName As String
    ProductVersion As String
End Type

Private Enum VerisonReturnValue
    eOK = 1
    eNoVersion = 2
End Enum
Public Sub CHECK_NEW_VERSION()
    
    Dim LOCALI As String
    Dim SERVER As String
    Dim EXE As String
       
    EXE = App.EXEName
'    SaveSetting "CentraNord", "GestParc", "EXE", EXE
    SERVER = GetSetting("CentraNord", "GestParc", "CHEMIN_MAJ", "") & EXE & ".exe" 'Renvoie une valeur de clé à partir de l'entrée d'une application dans la base de registre de Windows
'    SaveSetting "CentraNord", "GestParc", "CHEMIN_MAJ", SERVER
'    SERVER = "\\srv-files\sce informatique\Parcano exe\" & EXE & ".exe"
    If EXE = "" Or SERVER = "" Then Exit Sub
    
    LOCALI = App.Path & "\" & EXE & ".exe"   'C:\Parcano app.path
   
    Screen.MousePointer = vbHourglass
    
    If ExisteFile(App.Path & "\000000_Parcano" & "\000000_" & EXE & ".exe") = True Then
        'Effacer le fichier temp
        On Error Resume Next
        Kill App.Path & "\000000_Parcano" & "\000000_" & EXE & ".exe"
        On Error GoTo 0
    End If
    
        
    If LCase(Mid(EXE, 1, 6)) = LCase("000000") Then
        'Copier le new vers le vrai
        CopyFileAny LOCALI, _
        App.Path & "\" _
        & Replace(EXE, "000000_", "", , , vbTextCompare) _
        & ".exe"
    
    Else
        If GetOnlyVersion(SERVER) = "Inconnue" Then
            MsgBox "Mise à jour automatique du [" & EXE & "]" & vbLf & vbLf & "Aucune version trouvée dans l'emplacement indiqué." & vbLf & vbLf & "Veillez contacter votre Administrateur", vbCritical, "Mise à jour automatique"
        End If
        
        'Verifier si il y a une nouvelle version
        If IsNew(GetOnlyVersion(LOCALI), GetOnlyVersion(SERVER)) = True Then
            If MsgBox("Votre version est " & GetOnlyVersion(LOCALI) & "." _
            & vbLf & "Nouvelle version " & GetOnlyVersion(SERVER) _
            & " du " & EXE & " est disponible, Télécharger cette version ?" & vbLf & vbLf _
            & "EXENAME:" & App.EXEName & vbLf & GetALLVersion(SERVER), vbQuestion + vbYesNo, EXE) = vbYes Then
                PAR_DOWNLOAD.THEFILE = SERVER
                PAR_DOWNLOAD.TOFILE = App.Path & "\000000_" & EXE & ".exe"
                PAR_DOWNLOAD.Show 1
            Else
                frmconnexionx.Show
            End If
        End If
    End If
'    frmconnexionx.Show
    
    Screen.MousePointer = vbDefault
    
End Sub
Public Function TEST_NEW_VERSION() As Boolean

    On Error GoTo ER
    
    Dim LOCALI As String
    Dim SERVER As String
    Dim EXE As String
    
    EXE = App.EXEName
    LOCALI = App.Path & "\" & EXE & ".exe"
    SERVER = GetSetting("CentraNord", "GestParc", "CHEMIN_MAJ", "") & EXE & ".exe"
    
    If EXE = "" Or SERVER = "" Then Exit Function
    
    'MsgBox "" & GetOnlyVersion(LOCALI)
    If IsNew(GetOnlyVersion(LOCALI), GetOnlyVersion(SERVER)) = True Then
        TEST_NEW_VERSION = True
    Else
        TEST_NEW_VERSION = False
    End If
    
Exit Function
ER:
    TEST_NEW_VERSION = False

End Function
Private Function IsNew(OLDFILE As String, NEWFILE As String) As Boolean

    Dim T As Integer
    Dim CHAINE1 As String
    Dim CHAINE2 As String
    Dim NOUVELLE As Boolean
    
    NOUVELLE = False
    OLDFILE = OLDFILE & "."
    NEWFILE = NEWFILE & "."
    
    For T = 1 To NB_MOT(OLDFILE, , ".")
        CHAINE1 = NB_MOT(OLDFILE, T, ".")
        CHAINE2 = NB_MOT(NEWFILE, T, ".")
        If Val(CHAINE2) > Val(CHAINE1) Then
            NOUVELLE = True
        End If
    Next T
    
    IsNew = NOUVELLE
    
End Function
Public Function GetOnlyVersion(strFile As String) As String
    Dim udtFileInfo As FILEINFO
    Dim temp As String
    On Error Resume Next
    
    If GetFileVersionInformation(strFile, udtFileInfo) = eNoVersion Then
        GetOnlyVersion = "Inconnue"
        Exit Function
    End If

    temp = udtFileInfo.FileVersion
    
    Dim VERSION1 As Integer
    Dim VERSION2 As Integer
    Dim VERSION3 As Integer
    
    VERSION1 = Format(NB_MOT(temp, 1, "."), "0")
    VERSION2 = Format(NB_MOT(temp, 2, "."), "0")
    VERSION3 = Format(NB_MOT(temp, 3, "."), "0")
    
    GetOnlyVersion = VERSION1 & "." & VERSION2 & "." & VERSION3
End Function
Private Function GetFileVersionInformation(ByRef pstrFieName As String, ByRef tFileInfo As FILEINFO) As VerisonReturnValue

    Dim lBufferLen As Long, lDummy As Long
    Dim sBuffer() As Byte
    Dim lVerPointer As Long
    Dim lRet As Long
    Dim Lang_Charset_String As String
    Dim HexNumber As Long
    Dim i As Integer
    Dim strTemp As String
    
    tFileInfo.CompanyName = ""
    tFileInfo.FileDescription = ""
    tFileInfo.FileVersion = ""
    tFileInfo.InternalName = ""
    tFileInfo.LegalCopyright = ""
    tFileInfo.OriginalFileName = ""
    tFileInfo.ProductName = ""
    tFileInfo.ProductVersion = ""
    lBufferLen = GetFileVersionInfoSize(pstrFieName, lDummy)

    If lBufferLen < 1 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If
    ReDim sBuffer(lBufferLen)
    lRet = GetFileVersionInfo(pstrFieName, 0&, lBufferLen, sBuffer(0))

    If lRet = 0 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If
    lRet = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)

    If lRet = 0 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If
    Dim bytebuffer(255) As Byte
    MoveMemory bytebuffer(0), lVerPointer, lBufferLen
    HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
    Lang_Charset_String = Hex(HexNumber)
    
    Do While Len(Lang_Charset_String) < 8
        Lang_Charset_String = "0" & Lang_Charset_String
    Loop
    
    Dim strVersionInfo(7) As String
    strVersionInfo(0) = "CompanyName"
    strVersionInfo(1) = "FileDescription"
    strVersionInfo(2) = "FileVersion"
    strVersionInfo(3) = "InternalName"
    strVersionInfo(4) = "LegalCopyright"
    strVersionInfo(5) = "OriginalFileName"
    strVersionInfo(6) = "ProductName"
    strVersionInfo(7) = "ProductVersion"
    Dim buffer As String

    For i = 0 To 7
    
        buffer = String(255, 0)
        strTemp = "\StringFileInfo\" & Lang_Charset_String _
        & "\" & strVersionInfo(i)
        
        lRet = VerQueryValue(sBuffer(0), strTemp, _
        lVerPointer, lBufferLen)

        If lRet <> 0 Then
            lstrcpy buffer, lVerPointer
            buffer = Mid$(buffer, 1, InStr(buffer, vbNullChar) - 1)
            Select Case i
                Case 0
                tFileInfo.CompanyName = buffer
                Case 1
                tFileInfo.FileDescription = buffer
                Case 2
                tFileInfo.FileVersion = buffer
                Case 3
                tFileInfo.InternalName = buffer
                Case 4
                tFileInfo.LegalCopyright = buffer
                Case 5
                tFileInfo.OriginalFileName = buffer
                Case 6
                tFileInfo.ProductName = buffer
                Case 7
                tFileInfo.ProductVersion = buffer
            End Select
        End If
Next i
GetFileVersionInformation = eOK
End Function
'-----------

Private Function GetALLVersion(strFile As String) As String

    Dim udtFileInfo As FILEINFO
    Dim temp As String
    
    On Error Resume Next
    
    If GetFileVersionInformation(strFile, udtFileInfo) = eNoVersion Then
        GetALLVersion = "No version available For this file"
        Exit Function
    End If
    
    temp = "Company Name: " & udtFileInfo.CompanyName & vbCrLf
    temp = temp & "File Description:" & udtFileInfo.FileDescription & vbCrLf
    temp = temp & "File Version:" & udtFileInfo.FileVersion & vbCrLf
    temp = temp & "Internal Name: " & udtFileInfo.InternalName & vbCrLf
    temp = temp & "Legal Copyright: " & udtFileInfo.LegalCopyright & vbCrLf
    temp = temp & "Original FileName:" & udtFileInfo.OriginalFileName & vbCrLf
    temp = temp & "Product Name:" & udtFileInfo.ProductName & vbCrLf
    temp = temp & "Product Version: " & udtFileInfo.ProductVersion & vbCrLf
    GetALLVersion = temp
    
End Function

Private Function NB_MOT(MOT As String, Optional GetMot As Integer = -1, Optional SEPARATEUR As String = " ") As String

    Dim OLD_MOT As String
    Dim NB As Integer
    Dim x As Integer
    Dim THEMOT As String
    Dim MyMOT As String
    
    
    MyMOT = Trim(MOT) & SEPARATEUR
    OLD_MOT = MyMOT
    
    NB = 0
    x = InStr(1, OLD_MOT, SEPARATEUR)
    While x > 1
        x = InStr(1, OLD_MOT, SEPARATEUR)
        If x <> 0 Then
                NB = NB + 1
                If NB = GetMot Then THEMOT = Trim(Mid(OLD_MOT, 1, InStr(1, OLD_MOT, SEPARATEUR)))
                OLD_MOT = Mid(OLD_MOT, x + 1, Len(OLD_MOT) - x)
        End If
    Wend
    
    If GetMot = -1 Then NB_MOT = Trim(Str(NB)) Else NB_MOT = Trim(Replace(THEMOT, SEPARATEUR, " "))

End Function

 Public Function CopyFileAny(currentFilename As String, newFilename As String) As Boolean
    
    Dim A%, buffer%, temp$, fRead&, b%
    Dim fSize As Double
    Dim FOIS As Long
    On Error GoTo ErrHan:
    FOIS = 0
    
    A = FreeFile
    buffer = 4048
    Open currentFilename For Binary Access Read As A
    b = FreeFile
    Open newFilename For Binary Access Write As b

    fSize = LOF(A)
   
        While fRead < fSize
            If buffer > (fSize - fRead) Then buffer = (fSize - fRead)
            temp = Space(buffer)
            Get A, , temp
            Put b, , temp
            fRead = fRead + buffer
        Wend
        
        Close b
        Close A
        CopyFileAny = True
        
        Exit Function
ErrHan:

        FOIS = FOIS + 1
        Pause_Timer 1
        If FOIS < 10 Then Resume
        Close b
        Close A
        CopyFileAny = False
    
End Function

Private Sub Pause_Timer(PauseTime As Single)
Dim start
    start = Timer   ' Définit l'heure de début.
    Do While Timer < start + PauseTime
        DoEvents    ' Donne le contrôle à d'autres processus.
    Loop
End Sub

Public Sub CHECK_NEW_IndexPCT()
    
    Dim LOCALI As String
    Dim SERVER As String
    Dim EXE As String
    Dim Okay As Boolean
        
    EXE = "Index-PCT2008"
    SERVER = GetSetting("CentraNord", "GestParc", "CHEMIN_MAJ", "")
    
    If EXE = "" Or SERVER = "" Then Exit Sub
    
    LOCALI = App.Path & "\IndexPCT2008\"
    
    If SERVER = "" Then Exit Sub

    SERVER = SERVER & "\IndexPCT2008\"
    Screen.MousePointer = vbHourglass
    
    If ExisteFile(SERVER & EXE & ".exe") = False Then
        'Copy
        If CopyFileAny(SERVER & "INDEXPCT2008.mdb", LOCALI & "INDEXPCT2008.mdb") = True Then
            Okay = True
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub
Sub SaveLog_maj(THEFILE As String, Text As String)
    
    Debug.Print Text
    On Error GoTo ER
    
    Open THEFILE & ".log" For Append As #1
    Print #1, Date & "-" & Time & ":" & Text
    Close #1
    Exit Sub
ER:

    On Error Resume Next
    Close #1

End Sub

Private Sub alerteON(ETAT As Boolean)

On Error GoTo ER
Dim buttonTool As ActiveBar2LibraryCtl.Tool

If ETAT = True Then
    Set buttonTool = Frm_Main.ACB_Main.Tools("lblAffichage")
    With buttonTool
        .Caption = ""
        .SetPicture ddITNormal, LoadResPicture("nouvelle", 0)
        .Style = ddSIconText
    End With
    Frm_Main.ACB_Main.ApplyAll buttonTool
    Frm_Main.ACB_Main.RecalcLayout
Else
    Set buttonTool = Frm_Main.ACB_Main.Tools("lblAffichage")
    With buttonTool
        .Caption = ""
        .SetPicture ddITNormal, Nothing
        .Style = ddSIconText
    End With
    Frm_Main.ACB_Main.ApplyAll buttonTool
    Frm_Main.ACB_Main.RecalcLayout
End If
Exit Sub
ER:
End Sub

    
