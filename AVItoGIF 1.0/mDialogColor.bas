Attribute VB_Name = "mDialogColor"
'================================================
' Module:        mDialogColor.bas
' Author:        -
' Dependencies:  None
' Last revision: 2003.03.28
'================================================

Option Explicit

'-- API:

Private Type tChooseColor
    lStructSize    As Long
    hwndOwner      As Long
    hInstance      As Long
    rgbResult      As Long
    lpCustColors   As Long
    Flags          As Long
    lCustData      As Long
    lpfnHook       As Long
    lpTemplateName As String
End Type

Private Const CC_RGBINIT   As Long = &H1
Private Const CC_FULLOPEN  As Long = &H2
Private Const CC_ANYCOLOR  As Long = &H100
Private Const CC_NORMAL    As Long = CC_ANYCOLOR Or CC_RGBINIT
Private Const CC_EXTENDED  As Long = CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN

Private Declare Function ChooseColor Lib "comdlg32" Alias "ChooseColorA" (Color As tChooseColor) As Long

'//

Private m_CustomColors(15) As Long
Private m_Initialized      As Boolean

Public Function SelectColor(ByVal hWndParent As Long, ByVal DefaultColor As Long, Optional ByVal ShowDlgEx As Boolean = 0) As Long
 
  Dim tCC  As tChooseColor
  Dim lRet As Long
  Dim lIdx As Long
     
    '-- Initialize custom colors
    If (m_Initialized = 0) Then
        m_Initialized = -1
        For lIdx = 0 To 15
            m_CustomColors(lIdx) = RGB(lIdx * 17, lIdx * 17, lIdx * 17)
        Next lIdx
    End If
    
    '-- Prepare struct.
    With tCC
        .lStructSize = Len(tCC)
        .hwndOwner = hWndParent
        .rgbResult = DefaultColor
        .lpCustColors = VarPtr(m_CustomColors(0))
        .Flags = IIf(ShowDlgEx, CC_EXTENDED, CC_NORMAL)
    End With
        
    '-- Show Color dialog
    lRet = ChooseColor(tCC)
     
    '-- Get color / Cancel
    If (lRet) Then
        SelectColor = tCC.rgbResult
      Else
        SelectColor = -1
    End If
End Function
