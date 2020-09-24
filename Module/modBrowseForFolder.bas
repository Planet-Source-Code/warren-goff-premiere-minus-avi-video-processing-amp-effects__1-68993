Attribute VB_Name = "modBrowseForFolder"
'Browse For Folder API Call

'Enum for the Flags of the BrowseForFolder API function
Enum BrowseForFolderFlags
    BIF_RETURNONLYFSDIRS = &H1
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_STATUSTEXT = &H4
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_EDITBOX = &H10
    BIF_RETURNFSANCESTORS = &H8
End Enum

'BrowseInfo is a type used with the SHBrowseForFolder API call
Private Type BROWSEINFO
     hwndOwner As Long
     pidlRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

'Shell APIs from Shell32.dll file:
'SHBrowseForFolder - Gets the Browse For Folder Dialog
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long


'lstrcat API function appends a string to another - that means that some API functions
'need their string in the numeric way like this does, so its kind of converts strings
'to numbers
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Function BrowseForFolder(hWnd As Long, Optional Title As String, Optional flags As BrowseForFolderFlags) As String

    'Variables for use:
     Dim iNull As Integer
     Dim IDList As Long
     Dim Result As Long
     Dim Path As String
     Dim bi As BROWSEINFO
     
     If flags = 0 Then flags = BIF_RETURNONLYFSDIRS
     
    'Type Settings
     With bi
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(Title, "")
        .ulFlags = flags
     End With

    'Execute the BrowseForFolder shell API and display the dialog
     IDList = SHBrowseForFolder(bi)
     
    'Get the info out of the dialog
     If IDList Then
        Path = String$(300, 0)
        Result = SHGetPathFromIDList(IDList, Path)
        iNull = InStr(Path, vbNullChar)
        If iNull Then Path = Left$(Path, iNull - 1)
     End If

    'If Cancel button was clicked, error occured or My Computer was selected then Path = ""
     BrowseForFolder = Path
End Function
