Attribute VB_Name = "mCommon"
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Const OFN_HIDEREADONLY = &H4

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Function CommonOpenFile(ByVal hWndParent As Long, ByVal Filter As String, ByVal InitDir As String, ByVal Title As String) As String
  Dim OFName As OPENFILENAME
  Dim Tam As Long
    
  OFName.lStructSize = Len(OFName)
  OFName.hwndOwner = hWndParent
  OFName.hInstance = App.hInstance
  OFName.lpstrFilter = Filter
  OFName.lpstrFile = Space$(254)
  OFName.nMaxFile = 255
  OFName.lpstrFileTitle = Space$(254)
  OFName.nMaxFileTitle = 255
  OFName.lpstrInitialDir = InitDir
  OFName.lpstrTitle = Title
  OFName.flags = OFN_HIDEREADONLY
  If GetOpenFileName(OFName) Then
    OFName.lpstrFile = Trim(OFName.lpstrFile)
    Tam = Len(OFName.lpstrFile)
    CommonOpenFile = Mid(OFName.lpstrFile, 1, Tam - 1) 'cut char 0
  Else
    CommonOpenFile = ""
  End If
End Function


