Attribute VB_Name = "modGlobals"
Option Explicit
Global Framed As Long
Global Split As Boolean
Global Processor As Integer, ScrollValue As Integer, ScrollMin As Integer, ScrollMax As Integer, ApplyEffect As Boolean, CopyImage As Boolean
Global ApplyFlag As Boolean
Global Hidder As Boolean
Global Initt As Boolean
Global AVIInput As Boolean

Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long


Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Const SW_SHOWNORMAL = 1          '  Activates and displays a window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when displaying the window for the first time.

'declare for moving the form










'Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
      
      Declare Function FindWindow _
       Lib "user32" Alias "FindWindowA" _
       (ByVal lpClassName As String, _
       ByVal lpWindowName As String) _
       As Long
'declarations...
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
(ByVal hwndOwner As Long, _
ByVal nFolder As Long, pidl As Long) As Long
' Ret: 0=success

Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
(pidl As Long, ByVal pszPath As String) As Long
Public Declare Function GlobalFree Lib "kernel32" _
(ByVal hMem As Long) As Long

Private Const MAX_PATH = 260
'Public Const CSIDL_DESKTOP = &H0

'implementation...
Public Function GetShellFolderPath(ByVal CSIDL As Long) As String
Dim pID As Long
Dim sTmp As String

If SHGetSpecialFolderLocation(0&, CSIDL, pID) = 0& Then
sTmp = String(MAX_PATH + 2, 0)
If SHGetPathFromIDList(ByVal pID, sTmp) <> 0& Then
    GetShellFolderPath = Left$(sTmp, InStr(1, sTmp, vbNullChar) - 1)
End If
End If
If pID <> 0& Then GlobalFree pID
End Function

      

Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
 If Topmost = True Then 'Make the window topmost
  SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
 Else
  SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
  SetTopMostWindow = False
 End If
End Function


Public Function ShellExecLaunchFile(ByVal strPathFile As String, ByVal strOpenInPath As String, ByVal strArguments As String) As Long

    Dim Scr_hDC As Long
    
    'Get the Desktop handle
    Scr_hDC = GetDesktopWindow()
    
    'Launch File
    ShellExecLaunchFile = ShellExecute(Scr_hDC, "Open", strPathFile, "", strOpenInPath, SW_SHOWNORMAL)

End Function

