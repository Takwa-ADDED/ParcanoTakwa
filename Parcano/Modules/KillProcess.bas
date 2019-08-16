Attribute VB_Name = "KillProcess"
Option Explicit

Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "Kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function GetCurrentProcess Lib "Kernel32" () As Long
Private Declare Function ProcessFirst Lib "Kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "Kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "Kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function apiGetClassName Lib "user32" Alias _
                "GetClassNameA" (ByVal hWnd As Long, _
                ByVal lpClassname As String, _
                ByVal nMaxCount As Long) As Long
Private Declare Function apiGetDesktopWindow Lib "user32" Alias _
                "GetDesktopWindow" () As Long
Private Declare Function apiGetWindow Lib "user32" Alias _
                "GetWindow" (ByVal hWnd As Long, _
                ByVal wCmd As Long) As Long
Private Declare Function apiGetWindowLong Lib "user32" Alias _
                "GetWindowLongA" (ByVal hWnd As Long, ByVal _
                nIndex As Long) As Long
                
Private Declare Function apiGetWindowText Lib "user32" Alias _
                "GetWindowTextA" (ByVal hWnd As Long, ByVal _
                lpString As String, ByVal aint As Long) As Long
                
Private Const mcGWCHILD = 5
Private Const mcGWHWNDNEXT = 2
Private Const mcGWLSTYLE = (-16)
Private Const mcWSVISIBLE = &H10000000
Private Const mconMAXLEN = 255

Public Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SwitchToThisWindow Lib "user32" (ByVal hWnd As Long, ByVal hWindowState As Long) As Long

Const MAX_PATH As Integer = 260
Private Type LUID
    LowPart As Long
    HighPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type
Private Function ProcessTerminate(Optional lProcessID As Long, Optional lHwndWindow As Long) As Boolean
     
     Dim lhwndProcess As Long
     Dim lExitCode As Long
     Dim lRetVal As Long
     Dim lhThisProc As Long
     Dim lhTokenHandle As Long
     Dim tLuid As LUID
     Dim tTokenPriv As TOKEN_PRIVILEGES, tTokenPrivNew As TOKEN_PRIVILEGES
     Dim lBufferNeeded As Long
     Const PROCESS_ALL_ACCESS = &H1F0FFF, PROCESS_TERMINAT = &H1
     Const ANYSIZE_ARRAY = 1, TOKEN_ADJUST_PRIVILEGES = &H20
     Const TOKEN_QUERY = &H8, SE_DEBUG_NAME As String = "SeDebugPrivilege"
     Const SE_PRIVILEGE_ENABLED = &H2
     On Error Resume Next
     If lHwndWindow Then
        lRetVal = GetWindowThreadProcessId(lHwndWindow, lProcessID)
     End If
     If lProcessID Then
        lhThisProc = GetCurrentProcess
        OpenProcessToken lhThisProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lhTokenHandle
        LookupPrivilegeValue "", SE_DEBUG_NAME, tLuid
        tTokenPriv.PrivilegeCount = 1
        tTokenPriv.TheLuid = tLuid
        tTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
        AdjustTokenPrivileges lhTokenHandle, False, tTokenPriv, Len(tTokenPrivNew), tTokenPrivNew, lBufferNeeded
        lhwndProcess = OpenProcess(PROCESS_TERMINAT, 0, lProcessID)
        If lhwndProcess Then
         ProcessTerminate = CBool(TerminateProcess(lhwndProcess, lExitCode))
         Call CloseHandle(lhwndProcess)
        End If
     End If
     On Error GoTo 0
End Function

Public Function KillProcessus(ByVal sProcessNameExe As String) As String

     Dim i As Integer
     Dim hSnapshot As Long
     Dim uProcess As PROCESSENTRY32
     Dim r As Long
     Dim nom(1 To 100)
     Dim NUM(1 To 100)
     Dim nr As Integer
     
     Const TH32CS_SNAPPROCESS As Long = 2&
     nr = 0
     hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
     If hSnapshot = 0 Then Exit Function
     uProcess.dwSize = Len(uProcess)
     r = ProcessFirst(hSnapshot, uProcess)
     Do While r
        nr = nr + 1
        nom(nr) = uProcess.szexeFile
        NUM(nr) = uProcess.th32ProcessID
        r = ProcessNext(hSnapshot, uProcess)
     Loop
     For i = 1 To nr
         If InStr(UCase(nom(i)), UCase(sProcessNameExe)) <> 0 Then
            ProcessTerminate (NUM(i))
            Exit For
         End If
     Next i
     
End Function

Function fEnumWindows(ETAT As Long, ByVal N As String) As Boolean

Dim lngx As Long, lngLen As Long
Dim lngStyle As Long, strCaption As String
Dim Okay As Boolean
Dim K

lngx = GetActiveWindow
lngx = apiGetDesktopWindow()
'Return the first child to Desktop
lngx = apiGetWindow(lngx, mcGWCHILD)

fEnumWindows = False
    
    Do While Not lngx = 0
        strCaption = fGetCaption(lngx)
        If Len(strCaption) > 0 Then
            lngStyle = apiGetWindowLong(lngx, mcGWLSTYLE)
            If lngStyle And mcWSVISIBLE Then
                If fGetCaption(lngx) = N Then
                    fEnumWindows = True
                    
                    If ETAT = 0 Then
                        CloseWindow lngx
                        Exit Do
                    Else
                        Call SwitchToThisWindow(lngx, vbNormalFocus)
                        Exit Do
                    End If
                End If
            End If

        End If
        lngx = apiGetWindow(lngx, mcGWHWNDNEXT)
    Loop
    
End Function
Function fEnumWindowsClose(ETAT As Long, ByVal N As String, F As Form)

Dim lngx As Long, lngLen As Long
Dim lngStyle As Long, strCaption As String
Dim Okay As Boolean
Dim K

lngx = GetActiveWindow
lngx = apiGetDesktopWindow()
'Return the first child to Desktop
lngx = apiGetWindow(lngx, mcGWCHILD)
    
    Do While Not lngx = 0
        strCaption = fGetCaption(lngx)
        If Len(strCaption) > 0 Then
            lngStyle = apiGetWindowLong(lngx, mcGWLSTYLE)
            If lngStyle And mcWSVISIBLE Then
                If fGetCaption(lngx) = N Then
                    If ETAT = 0 Then
                        Unload F
                        Exit Do
                    Else
                        Call SwitchToThisWindow(lngx, vbNormalFocus)
                        Exit Do
                    End If
                End If
            End If

        End If
        lngx = apiGetWindow(lngx, mcGWHWNDNEXT)
    Loop
    
End Function

Function fEnumWindows_Row(ETAT As Long, ByVal N As String) As Boolean

Dim lngx As Long, lngLen As Long
Dim lngStyle As Long, strCaption As String
Dim Okay As Boolean
Dim K

lngx = GetActiveWindow
lngx = apiGetDesktopWindow()
'Return the first child to Desktop
lngx = apiGetWindow(lngx, mcGWCHILD)
N_XLSNG = 0
fEnumWindows_Row = False
    
    Do While Not lngx = 0
        strCaption = fGetCaption(lngx)
        If Len(strCaption) > 0 Then
            lngStyle = apiGetWindowLong(lngx, mcGWLSTYLE)
            If lngStyle And mcWSVISIBLE Then
                If fGetCaption(lngx) = N Then
                    fEnumWindows_Row = True
                    If ETAT = 0 Then
                        CloseWindow lngx
                        Exit Do
                    Else
                        N_XLSNG = lngx
                        Exit Do
                    End If
                End If
            End If

        End If
        lngx = apiGetWindow(lngx, mcGWHWNDNEXT)
    Loop
    
End Function
Private Function fGetClassName(hWnd As Long) As String
    Dim strBuffer As String
    Dim intCount As Integer
     
    strBuffer = String$(mconMAXLEN - 1, 0)
    intCount = apiGetClassName(hWnd, strBuffer, mconMAXLEN)
    If intCount > 0 Then
        fGetClassName = Left$(strBuffer, intCount)
    End If
End Function
Private Function fGetCaption(hWnd As Long) As String
    
    Dim strBuffer As String
    Dim intCount As Integer

    strBuffer = String$(mconMAXLEN - 1, 0)
    intCount = apiGetWindowText(hWnd, strBuffer, mconMAXLEN)

   If intCount > 0 Then
        fGetCaption = Left$(strBuffer, intCount)
    End If
    
End Function
