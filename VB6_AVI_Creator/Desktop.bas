Attribute VB_Name = "Desktop"
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

'usage:
'Debug.Print GetShellFolderPath(CSIDL_DESKTOP)

Public Sub Delay(HowLong As Date)
TempTime = DateAdd("s", HowLong, Now)
While TempTime > Now
DoEvents 'Allows windows to handle other stuff
Wend
End Sub

