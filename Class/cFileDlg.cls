VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFileDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************************
'*  VB file:   CmnDlg.bas... VB32 wrapper for Win32 common dialog
'*                           functions.
'*  created:        1997 by Ray Mercer
'*  modified:       8/98 by Ray Mercer (added browse for folders)
'*  modified:       10/21/98 by Ray Mercer (added comments)
'*  modified:       11/19/98 by Ray Mercer (major enhancements)
'*  modified:       01/28/99 by Ray Mercer (added CenterScreen())
'*  modified:       12/31/99 by Ray Mercer (changed to Class Module)
'*  modified:       02/13/00 by Ray Mercer
'*                   changed name to cFileDlg.cls
'*                   simplified interface and removed extra code
'*                   this class now supports only FileOpen and FileSave dialogs
'*
'*
'*  original functions based on code found in Bruce McKinney's book
'*  "Hardcore Visual Basic"
'*  enhancements on 11/19/98 based on code by Brad Martinez (especially
'*  useful comments)
'*
'*  Copyright (C) 1997 - 2000 Ray Mercer.  All rights reserved.
'*  Latest version can be downloaded from http://www.shrinkwrapvb.com
'****************************************************************

Private Const MAX_PATH = 1024
Private Const MAX_FILE = 512
'
'Private Type SHITEMID
'    cb As Long     'Size of the ID (including cb itself)
'    abID As Byte   'The item ID (variable length)
'End Type
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type

'most of these are also in
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders
Public Enum SPECIAL_FOLDERS
    'Windows desktop virtual folder at the root of the name space
    vbCSIDL_DESKTOP = &H0&   'File system directory that contains the
   'user's program groups (which are also file   'system directories)
    vbCSIDL_PROGRAMS = &H2&
   'Control Panel - virtual folder containing
   'icons for the control panel applications
    vbCSIDL_CONTROLS = &H3&
   'Printers folder - virtual folder containing    'installed printers.
    vbCSIDL_PRINTERS = &H4&   'File system directory that serves as a
   'common repository for documents (My Documents folder)
    vbCSIDL_PERSONAL = &H5&
   'File system directory that contains the
   'user's favorite Internet Explorer URLs
    vbCSIDL_FAVORITES = &H6&
   'File system directory that corresponds to the
   'user's Startup program group
    vbCSIDL_STARTUP = &H7&
   'File system directory that contains the
   'user's most recently used documents (Recent folder)
    vbCSIDL_RECENT = &H8&   'File system directory that contains
   'Send To menu items    Public Const
    vbCSIDL_SENDTO = &H9&
   'Recycle bin file system directory containing file
   'objects in the user's recycle bin. The location of
   'this directory is not in the registry; it is marked
   'with the hidden and system attributes to prevent the
   'user from moving or deleting it.
    vbCSIDL_BITBUCKET = &HA&
   'File system directory containing Start menu items
    vbCSIDL_STARTMENU = &HB&
   'File system directory used to physically store
   'file objects on the desktop (not to be confused
   'with the desktop folder itself).
    vbCSIDL_DESKTOPDIRECTORY = &H10&
   'My Computer - virtual folder containing everything
   'on the local computer: storage devices, printers,
   'and Control Panel. The folder may also contain    'mapped network drives.
    vbCSIDL_DRIVES = &H11&
   'Network Neighborhood - virtual folder representing
   'the top level of the network hierarchy
    vbCSIDL_NETWORK = &H12&
   'File system directory containing objects that
   'appear in the network neighborhood
    vbCSIDL_NETHOOD = &H13&
   'Virtual folder containing fonts
    vbCSIDL_FONTS = &H14&
   'File system directory that serves as a
   'common repository for document templates    '(ShellNew folder.)
    vbCSIDL_TEMPLATES = &H15&
End Enum
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
                                                (ByVal hWndOwner As Long, _
                                                ByVal nFolder As SPECIAL_FOLDERS, _
                                                pidl As Long) As Long 'returns NOERROR on success

'Converts an item identifier list to a file system path.
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Const NOERROR As Long = &H0

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hWndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    Flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" _
    Alias "GetOpenFileNameA" (filestruct As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" _
    Alias "GetSaveFileNameA" (filestruct As OPENFILENAME) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" _
    Alias "GetFileTitleA" (ByVal szFile As String, _
    ByVal szTitle As String, ByVal cbBuf As Integer) As Integer
'VFW "customized" File Dialogs
Private Declare Function GetOpenFileNamePreview Lib "MSVFW32" _
    Alias "GetOpenFileNamePreviewA" (filestruct As OPENFILENAME) As Long
Private Declare Function GetSaveFileNamePreview Lib "MSVFW32" _
    Alias "GetSaveFileNamePreviewA" (filestruct As OPENFILENAME) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1&
    OFN_OVERWRITEPROMPT = &H2&
    OFN_HIDEREADONLY = &H4&
    OFN_NOCHANGEDIR = &H8&
    OFN_SHOWHELP = &H10&
    OFN_ENABLEHOOK = &H20&
    OFN_ENABLETEMPLATE = &H40&
    OFN_ENABLETEMPLATEHANDLE = &H80&
    OFN_NOVALIDATE = &H100&
    OFN_ALLOWMULTISELECT = &H200&
    OFN_EXTENSIONDIFFERENT = &H400&
    OFN_PATHMUSTEXIST = &H800&
    OFN_FILEMUSTEXIST = &H1000&
    OFN_CREATEPROMPT = &H2000&
    OFN_SHAREAWARE = &H4000&
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF&

    CDERR_GENERALCODES = &H0&
    CDERR_STRUCTSIZE = &H1&
    CDERR_INITIALIZATION = &H2&
    CDERR_NOTEMPLATE = &H3&
    CDERR_NOHINSTANCE = &H4&
    CDERR_LOADSTRFAILURE = &H5&
    CDERR_FINDRESFAILURE = &H6&
    CDERR_LOADRESFAILURE = &H7&
    CDERR_LOCKRESFAILURE = &H8&
    CDERR_MEMALLOCFAILURE = &H9&
    CDERR_MEMLOCKFAILURE = &HA&
    CDERR_NOHOOK = &HB&
    CDERR_REGISTERMSGFAIL = &HC&

    PDERR_PRINTERCODES = &H1000&
    PDERR_SETUPFAILURE = &H1001&
    PDERR_PARSEFAILURE = &H1002&
    PDERR_RETDEFFAILURE = &H1003&
    PDERR_LOADDRVFAILURE = &H1004&
    PDERR_GETDEVMODEFAIL = &H1005&
    PDERR_INITFAILURE = &H1006&
    PDERR_NODEVICES = &H1007&
    PDERR_NODEFAULTPRN = &H1008&
    PDERR_DNDMMISMATCH = &H1009&
    PDERR_CREATEICFAILURE = &H100A&
    PDERR_PRINTERNOTFOUND = &H100B&
    PDERR_DEFAULTDIFFERENT = &H100C&

    CFERR_CHOOSEFONTCODES = &H2000&
    CFERR_NOFONTS = &H2001&
    CFERR_MAXLESSTHANMIN = &H2002&

    FNERR_FILENAMECODES = &H3000&
    FNERR_SUBCLASSFAILURE = &H3001&
    FNERR_INVALIDFILENAME = &H3002&
    FNERR_BUFFERTOOSMALL = &H3003&

    CCERR_CHOOSECOLORCODES = &H5000&
End Enum

 
   
Private Const sEmpty As String = ""

'Class member variables
Private m_FileMustExist As Boolean
Private m_MultiSelect As Boolean
Private m_ReadOnly As Boolean 'read only
Private m_HideReadOnly As Boolean
Private m_Filter As String
Private m_FilterIndex As Long
Private m_InitDir As String
Private m_DlgTitle As String
Private m_DefaultExt As String
Private m_Flags As Long
Private m_OverwritePrompt As Boolean
Private m_hWnd As Long

Private Sub Class_Initialize()
'Initialize default values
m_hWnd = -1&
m_FileMustExist = True
m_MultiSelect = False
m_HideReadOnly = True
m_DlgTitle = App.title
m_OverwritePrompt = True
m_InitDir = GetSpecialFolderLocation(vbCSIDL_PERSONAL) 'default to My Documents folder

End Sub

Property Get OwnerHwnd() As Long
    OwnerHwnd = m_hWnd
End Property
Property Let OwnerHwnd(ByVal vHwnd As Long)
    m_hWnd = vHwnd
End Property

Property Get FileMustExist() As Boolean
    FileMustExist = m_FileMustExist
End Property
Property Let FileMustExist(ByVal vNewValue As Boolean)
    m_FileMustExist = vNewValue
End Property

Property Get MultiSelect() As Boolean
    MultiSelect = m_MultiSelect
End Property
Property Let MultiSelect(ByVal vNewValue As Boolean)
    m_MultiSelect = vNewValue
End Property

Property Get ReadOnly() As Boolean
    ReadOnly = m_ReadOnly
End Property

Property Get HideReadOnly() As Boolean
    HideReadOnly = m_HideReadOnly
End Property
Property Let HideReadOnly(ByVal vNewValue As Boolean)
    m_HideReadOnly = vNewValue
End Property

Property Get Filter() As String
    Filter = m_Filter
End Property
Property Let Filter(ByVal vFilterString As String)
    m_Filter = vFilterString
End Property

Property Get FilterIndex() As Long
    FilterIndex = m_FilterIndex
End Property
Property Let FilterIndex(ByVal vIndex As Long)
    m_FilterIndex = vIndex
End Property

Property Get InitDirectory() As String
    InitDirectory = m_InitDir
End Property
Property Let InitDirectory(ByVal DirPath As String)
    m_InitDir = DirPath
End Property

Property Let InitDirectorySpecial(ByVal SpecialDirectory As SPECIAL_FOLDERS)
    m_InitDir = GetSpecialFolderLocation(SpecialDirectory)
End Property

Property Get DlgTitle() As String
    DlgTitle = m_DlgTitle
End Property
Property Let DlgTitle(ByVal title As String)
    m_DlgTitle = title
End Property

Property Get DefaultExt() As String
    DefaultExt = m_DefaultExt
End Property
Property Let DefaultExt(ByVal fileExt As String)
    m_DefaultExt = fileExt
End Property

Property Get Flags() As EOpenFile
    Flags = m_Flags
End Property
Property Let Flags(ByVal vFlags As EOpenFile)
    m_Flags = vFlags
End Property

Property Get OverwritePrompt() As Boolean
    OverwritePrompt = m_OverwritePrompt
End Property
Property Let OverwritePrompt(ByVal vShowPrompt As Boolean)
    m_OverwritePrompt = vShowPrompt
End Property

Public Function VBGetOpenFileName(filename As String, _
                           Optional FileTitle As String) As Boolean

    Dim opfile As OPENFILENAME
    Dim s As String
    Dim afFlags As Long
    Dim ch As String
    Dim i As Integer
    
With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    .Flags = (-m_FileMustExist * OFN_FILEMUSTEXIST) Or _
             (-m_MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-m_ReadOnly * OFN_READONLY) Or _
             (-m_HideReadOnly * OFN_HIDEREADONLY) Or _
             (m_Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If m_hWnd <> -1 Then .hWndOwner = m_hWnd
    ' InitDir can take initial directory string
    .lpstrInitialDir = m_InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = m_DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = m_DlgTitle
    
    ' To make Windows-style filter, replace | and : with nulls
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = m_FilterIndex

    ' Pad file and file title buffers to maximum path
    s = filename & String$(MAX_PATH - Len(filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields set to zero
    
    If GetOpenFileName(opfile) Then
        VBGetOpenFileName = True
        filename = left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
        FileTitle = left$(.lpstrFileTitle, InStr(.lpstrFileTitle, vbNullChar) - 1)
        m_Flags = .Flags
        ' Return the filter index
        m_FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        m_Filter = FilterLookup(.lpstrFilter, m_FilterIndex)
        If (.Flags And OFN_READONLY) Then m_ReadOnly = True
        'save directory as init directory for user
        m_InitDir = .lpstrFile
    Else
        VBGetOpenFileName = False
        filename = vbNullChar
        FileTitle = vbNullChar
        Flags = 0
        FilterIndex = -1
        Filter = vbNullChar
    End If
End With
End Function


Public Function VBGetOpenFileNamePreview(filename As String, _
                           Optional FileTitle As String) As Boolean

    Dim opfile As OPENFILENAME
    Dim s As String
    Dim afFlags As Long
    Dim ch As String
    Dim i As Integer
    
With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    .Flags = (-m_FileMustExist * OFN_FILEMUSTEXIST) Or _
             (-m_MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-m_ReadOnly * OFN_READONLY) Or _
             (-m_HideReadOnly * OFN_HIDEREADONLY) Or _
             (m_Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If m_hWnd <> -1 Then .hWndOwner = m_hWnd
    ' InitDir can take initial directory string
    .lpstrInitialDir = m_InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = m_DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = m_DlgTitle
    
    ' To make Windows-style filter, replace | and : with nulls
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = m_FilterIndex

    ' Pad file and file title buffers to maximum path
    s = filename & String$(MAX_PATH - Len(filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields set to zero
    
    If GetOpenFileNamePreview(opfile) Then
        VBGetOpenFileNamePreview = True
        filename = left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
        FileTitle = left$(.lpstrFileTitle, InStr(.lpstrFileTitle, vbNullChar) - 1)
        m_Flags = .Flags
        ' Return the filter index
        m_FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        m_Filter = FilterLookup(.lpstrFilter, m_FilterIndex)
        If (.Flags And OFN_READONLY) Then m_ReadOnly = True
        'save directory as init directory for user
        m_InitDir = .lpstrFile
    Else
        VBGetOpenFileNamePreview = False
        filename = vbNullChar
        FileTitle = vbNullChar
        Flags = 0
        FilterIndex = -1
        Filter = vbNullChar
    End If
End With
End Function

Public Function VBGetSaveFileName(filename As String, _
                           Optional FileTitle As String) As Boolean
            
    Dim opfile As OPENFILENAME, s As String
With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    .Flags = (-m_OverwritePrompt * OFN_OVERWRITEPROMPT) Or _
             OFN_HIDEREADONLY Or _
             (m_Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If m_hWnd <> -1 Then .hWndOwner = m_hWnd
    ' InitDir can take initial directory string
    .lpstrInitialDir = m_InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = m_DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = m_DlgTitle
    
    ' Make new filter with bars (|) replacing nulls and double null at end
    Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = m_FilterIndex

    ' Pad file and file title buffers to maximum path
    s = filename & String$(MAX_PATH - Len(filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields zero
    
    If GetSaveFileName(opfile) Then
        VBGetSaveFileName = True
        filename = left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
        FileTitle = left$(.lpstrFileTitle, InStr(.lpstrFileTitle, vbNullChar) - 1)
        m_Flags = .Flags
        ' Return the filter index
        m_FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        m_Filter = FilterLookup(.lpstrFilter, FilterIndex)
        'save directory as init directory for user
        m_InitDir = .lpstrFile
    Else
        VBGetSaveFileName = False
        filename = vbNullChar
        FileTitle = vbNullChar
        m_Flags = 0
        m_FilterIndex = 0
        m_Filter = vbNullChar
    End If
End With
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long
    Dim iEnd As Long
    Dim s As String
    
    iStart = 1
    If sFilters = vbNullChar Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function

Private Function StrZToStr(s As String) As String
    Dim TempString As String
    
    TempString = left$(s, InStr(s, vbNullChar) - 1)
    If TempString = "" Then
        'if VB string is accidently passed in there will be no NULL
        'so just pass back the original string in that case
        StrZToStr = s
    Else
        StrZToStr = TempString
    End If
End Function

'This fuction is courtesy of Randy Birch and VBNet <www.mvps.org/vbnet>
'however I changed it a bit to fit my class
Private Function GetSpecialFolderLocation(CSIDL As SPECIAL_FOLDERS) As String

   Dim sPath As String
   Dim pidl As Long
   
  'fill the idl structure with the specified folder item
   If SHGetSpecialFolderLocation(m_hWnd, CSIDL, pidl) = NOERROR Then
     
     'if the pidl is returned, initialize
     'and get the path from the id list
      sPath = Space$(MAX_PATH)
      
      If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then

        'free the pidl and return the path
         Call CoTaskMemFree(ByVal VarPtr(pidl))
         GetSpecialFolderLocation = left(sPath, InStr(sPath, Chr$(0)) - 1)
         
      End If
    
    End If
   
End Function

' Here are a couple of utility functions often needed when file dialogs are used

' Test file existence with error trapping
Public Function ExistFile(ByVal sSpec As String) As Boolean
    On Error Resume Next
    Call fileLen(sSpec)
    ExistFile = (Err = 0)
End Function

'Get FileTitle (filename without path) from any full path
Public Function VBGetFileTitle(sFile As String) As String
    Dim sFileTitle As String, cFileTitle As Integer

    cFileTitle = MAX_PATH
    sFileTitle = String$(MAX_PATH, 0)
    cFileTitle = GetFileTitle(sFile, sFileTitle, MAX_PATH)
    If cFileTitle Then
        VBGetFileTitle = ""
    Else
        VBGetFileTitle = left$(sFileTitle, InStr(sFileTitle, vbNullChar) - 1)
    End If

End Function


