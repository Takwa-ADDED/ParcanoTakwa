Attribute VB_Name = "Main_version"
Function SameLength(ByVal Ver1 As String, ByVal Ver2 As String) As String
    Dim diff As Integer, num2 As Integer
    diff = UBound(Split(Ver2, ".")) - UBound(Split(Ver1, "."))
    SameLength = Ver1 & String((diff + Abs(diff)) / 2, ".")
End Function

Function CompareVersions(Ver1 As String, Ver2 As String) As Boolean
    Ver1 = SameLength(Ver1, vel2)
    Ver2 = SameLength(Ver2, Ver1)
    If Ver1 = Ver2 Then CompareVersions = True: Exit Function
    Dim xSplit1() As String, xSplit2() As String
    xSplit1 = Split(Ver1, ".")
    xSplit2 = Split(Ver2, ".")
    For x = 0 To UBound(xSplit1)
        If Val(xSplit1(x)) > Val(xSplit2(x)) Then
            CompareVersions = True: Exit Function
        ElseIf Val(xSplit1(x)) < Val(xSplit2(x)) Then
            CompareVersions = False: Exit Function
        End If
    Next x
End Function

Private Sub MiseAjour()

Dim AppPathServer, AppPathLocal, VersionLocal As String
Dim fso As Object
Dim fichier As String, NvVer As String
 
AppPathServer = "\\srv-files\sce informatique\Parcano exe\Parcano.exe"
AppPathLocal = App.Path  '& "\Parcano.exe"

'Ancienne version
VersionLocal = App.Major & "." & App.Minor & "." & App.Revision

NvVer = ""
Set fso = CreateObject("Scripting.FileSystemObject", "srv-files")
Do Until NvVer <> ""
   NvVer = fso.GetFileVersion(AppPathServer) 'version de l'application.exe sur le serveur
   DoEvents
Loop

'App.Path & "\" 'Chemin de l'application.exe

If Not (CompareVersions(NvVer, VersionLocal)) Then

    fichier = "Parcano_" & NvVer & ".exe"
    Name AppPathServer & "\Parcano.exe" As AppPathServer & fichier
    Do Until fso.GetFileVersion(AppPathServer & fichier) <> ""
        DoEvents
    Loop
    
    'Copier l'ancienne version dans un dossier "OldVersions"
    'If Not ExistFolder(AppPath & "OldVersions") Then MkDir AppPath & "OldVersions"
    'FileCopy AppPath & fichier, AppPath & "OldVersions" & "\" & fichier
    'Do Until fso.GetFileVersion(AppPath & "OldVersions" & "\" & fichier) <> ""
    '    DoEvents
    'Loop
    
    'Supprime l'ancienne version
    Kill AppPathLocal & fichier
    
    Name AppPathLocal & "Parcano_" & NvVer As AppPathLocal & "Parcano.exe"
    Do Until fso.GetFileVersion(AppPathLocal & "Parcano.exe") <> ""
        DoEvents
    Loop
    
    MsgBox "Mise à jour effectué, cliquer sur 'OK' pour relancer l'application ! "
    Shell AppPathLocal & "Parcano.exe"  'Lancer le programme exécutable
    End
End If

End Sub
