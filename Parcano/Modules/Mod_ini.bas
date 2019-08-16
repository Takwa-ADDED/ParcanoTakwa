Attribute VB_Name = "Mod_ini"
'Pour le nom de fichier ex: "c:\test.ini"

Option Explicit

Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName$) As Long

Public Function GetIni(Section As String, Variable As String) As String
Dim strRetour As String
Dim fichier As String
fichier = "parcano.ini"
strRetour = String(255, Chr(0))
Dim Longueur As Integer
Longueur = GetPrivateProfileString(Section, Variable, "", strRetour, Len(strRetour), fichier)
GetIni = Left$(strRetour, Longueur)
End Function

Function WriteIni(Section As String, Variable As String, valeur As String) As Integer
Dim fichier As String
fichier = "parcano.ini"
WritePrivateProfileString Section, Variable, valeur, fichier
End Function
