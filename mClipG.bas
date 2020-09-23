Attribute VB_Name = "mClipG"
'+--------------------------------------------------------+
'| Name            : mClipG - Graphics Clipboard Functions|
'| Author          : Manuel Augusto Nogueira dos Santos   |
'| Dates           : 03/06/2001                           |
'| Description     : Copy/Paste with images               |
'+--------------------------------------------------------+
Option Explicit
'-------------------------------------------Windows API
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Const SRCCOPY = &HCC0020

Public Enum ClipboardFormats
   [_First] = 1
   CF_TEXT = 1
   CF_BITMAP = 2
   CF_METAFILEPICT = 3
   CF_SYLK = 4
   CF_DIF = 5
   CF_TIFF = 6
   CF_OEMTEXT = 7
   CF_DIB = 8
   CF_PALETTE = 9
   CF_PENDATA = 10
   CF_RIFF = 11
   CF_WAVE = 12
   CF_UNICODETEXT = 13
   CF_ENHMETAFILE = 14
   CF_HDROP = 15
   CF_LOCALE = 16
   CF_MAX = 17
   [_Last] = 17
End Enum

Public Sub Image2Clipboard(ByVal mHDC As Long, ByVal SizeX As Long, ByVal SizeY As Long)
  Dim NewHDC  As Long
  Dim ClipBMP As Long
  Dim OldBMP  As Long
  
  'create
  NewHDC = CreateCompatibleDC(mHDC)
  If (NewHDC = 0) Then Exit Sub
  ClipBMP = CreateCompatibleBitmap(mHDC, SizeX, SizeY)
  If (ClipBMP = 0) Then Exit Sub
  OldBMP = SelectObject(NewHDC, ClipBMP)
  'copy to Bitmap
  Call BitBlt(NewHDC, 0, 0, SizeX, SizeY, mHDC, 0, 0, SRCCOPY)
  Call SelectObject(NewHDC, OldBMP)
  'set Clipboard data
  Call EmptyClipboard
  Call OpenClipboard(0&)
  Call SetClipboardData(CF_BITMAP, ClipBMP)
  Call CloseClipboard
  'close
  Call DeleteObject(NewHDC)
End Sub
