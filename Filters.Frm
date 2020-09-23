VERSION 5.00
Begin VB.Form frmFilters 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   Caption         =   "Filters"
   ClientHeight    =   2070
   ClientLeft      =   4425
   ClientTop       =   2250
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   138
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   0
      ScaleHeight     =   282
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   5055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   60
      Top             =   60
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   0
      ScaleHeight     =   282
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   5055
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenFile 
         Caption         =   "Load Graphic..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuRetouch 
      Caption         =   "&Retouch"
      Begin VB.Menu mnuDefinition 
         Caption         =   "Definition"
         Begin VB.Menu mnuBlur 
            Caption         =   "Smooth"
         End
         Begin VB.Menu mnuBlurMore 
            Caption         =   "Blur"
         End
         Begin VB.Menu mnuSharpen 
            Caption         =   "Sharpen"
         End
         Begin VB.Menu mnuSharpenMore 
            Caption         =   "Sharpen More"
         End
         Begin VB.Menu mnuDiffuse 
            Caption         =   "Diffuse"
         End
         Begin VB.Menu mnuDiffuseMore 
            Caption         =   "Diffuse More"
         End
         Begin VB.Menu mnuPixelize 
            Caption         =   "Pixelize"
         End
      End
      Begin VB.Menu mnuEdges 
         Caption         =   "Edges"
         Begin VB.Menu mnuEmboss 
            Caption         =   "Emboss"
         End
         Begin VB.Menu mnuEmbossMore 
            Caption         =   "Emboss More"
         End
         Begin VB.Menu mnuEngrave 
            Caption         =   "Engrave"
         End
         Begin VB.Menu mnuEngraveMore 
            Caption         =   "Engrave More"
         End
         Begin VB.Menu mnuRelief 
            Caption         =   "Relief"
         End
         Begin VB.Menu mnuEdge 
            Caption         =   "Edge Enhance"
         End
         Begin VB.Menu mnuContour 
            Caption         =   "Contour"
         End
         Begin VB.Menu mnuConnection 
            Caption         =   "Connected Contour"
         End
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Colors"
         Begin VB.Menu mnuGreyScale 
            Caption         =   "GreyScale"
         End
         Begin VB.Menu mnuBlackWhite 
            Caption         =   "Black && White"
            Begin VB.Menu mnuBW1 
               Caption         =   "Nearest Color"
            End
            Begin VB.Menu mnuBW2 
               Caption         =   "Enhanced Diffusion"
            End
            Begin VB.Menu mnuBW3 
               Caption         =   "Ordered Dither"
            End
            Begin VB.Menu mnuBW4 
               Caption         =   "Floyd-Steinberg"
            End
            Begin VB.Menu mnuBW5 
               Caption         =   "Burke"
            End
            Begin VB.Menu mnuBW6 
               Caption         =   "Stucki"
            End
         End
         Begin VB.Menu mnuNegative 
            Caption         =   "Negative"
         End
         Begin VB.Menu mnuSwapColors 
            Caption         =   "Swap Colors"
            Begin VB.Menu mnuSwapBank 
               Caption         =   "RGB -> BRG"
               Index           =   1
            End
            Begin VB.Menu mnuSwapBank 
               Caption         =   "RGB -> GBR"
               Index           =   2
            End
            Begin VB.Menu mnuSwapBank 
               Caption         =   "RGB -> RBG"
               Index           =   3
            End
            Begin VB.Menu mnuSwapBank 
               Caption         =   "RGB -> BGR"
               Index           =   4
            End
            Begin VB.Menu mnuSwapBank 
               Caption         =   "RGB -> GRB"
               Index           =   5
            End
         End
         Begin VB.Menu mnuAqua 
            Caption         =   "Aqua"
         End
         Begin VB.Menu mnuAddNoise 
            Caption         =   "Add Noise"
         End
         Begin VB.Menu mnuGamma 
            Caption         =   "Gamma Correction"
         End
      End
      Begin VB.Menu mnuIntensity 
         Caption         =   "Intensity"
         Begin VB.Menu mnuBrighter 
            Caption         =   "Brighter"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuDarker 
            Caption         =   "Darker"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuContrast1 
            Caption         =   "Increase Contrast"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuContrast2 
            Caption         =   "Decrease Contrast"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuDilate 
            Caption         =   "Dilate"
         End
         Begin VB.Menu mnuErode 
            Caption         =   "Erode"
         End
         Begin VB.Menu mnuStretch 
            Caption         =   "Contrast Stretch"
         End
         Begin VB.Menu mnuSaturationI 
            Caption         =   "Increase Saturation"
            Shortcut        =   {F11}
         End
         Begin VB.Menu mnuSaturationD 
            Caption         =   "Decrease Saturation"
            Shortcut        =   {F12}
         End
      End
   End
End
Attribute VB_Name = "frmFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mProgress As Long

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call mnuExit_Click
End Sub

Private Sub mnuCopy_Click()
  Call Image2Clipboard(Picture1.HDC, Picture1.ScaleWidth, Picture1.ScaleHeight)
End Sub

Private Sub mnuExit_Click()
  Unload frmFilters
  End
End Sub

Private Sub mnuOpenFile_Click()
  Dim FileName As String
  
  FileName = CommonOpenFile(Me.hwnd, "", App.Path, "Open")
  If FileName <> "" Then
    Picture1.Picture = LoadPicture(FileName)
    Picture2.Width = Picture1.Width
    Picture2.Height = Picture1.Height
    Call BitBlt(Picture2.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
    'resizing the form....
    frmFilters.Width = (Picture1.ScaleWidth + 8) * Screen.TwipsPerPixelX
    frmFilters.Height = (Picture1.ScaleHeight + 46) * Screen.TwipsPerPixelY
    Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
    Me.Refresh
    
  End If
End Sub

Private Sub mnuRedo_Click()
  Call BitBlt(Me.HDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.HDC, 0, 0, SRCCOPY)
  Call BitBlt(Picture2.HDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Call BitBlt(Picture1.HDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuUndo_Click()
  Call BitBlt(Me.HDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.HDC, 0, 0, SRCCOPY)
  Call BitBlt(Picture2.HDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Call BitBlt(Picture1.HDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub Timer1_Timer()
  Me.Caption = "Filters " & mProgress
  If mProgress = 100 Then
    Me.Caption = "Filters"
    Timer1.Enabled = False
  End If
End Sub

Private Sub mnuSharpenMore_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iSHARPEN, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuSharpen_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iSHARPEN, Picture1.Image, 2, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuEmboss_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iEMBOSS, Picture1.Image, RGB(0, 120, 120), mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuEmbossMore_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iEMBOSSMORE, Picture1.Image, RGB(0, 120, 120), mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuBlur_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iSMOOTH, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuBlurMore_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iBLUR, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuDiffuseMore_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iDIFFUSE, Picture1.Image, 12, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuDiffuse_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iDIFFUSE, Picture1.Image, 6, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuEdge_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iEDGE, Picture1.Image, 1, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuEngrave_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iENGRAVE, Picture1.Image, RGB(0, 120, 120), mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuEngraveMore_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iENGRAVEMORE, Picture1.Image, RGB(0, 120, 120), mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuContour_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iCONTOUR, Picture1.Image, RGB(255, 255, 255), mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuGreyScale_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iGREYSCALE, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuRelief_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iRELIEF, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuBrighter_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iBRIGHTNESS, Picture1.Image, 10, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuDarker_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iBRIGHTNESS, Picture1.Image, -10, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuPixelize_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iPIXELIZE, Picture1.Image, 3, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuNegative_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iNEGATIVE, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuSwapBank_Click(Index As Integer)
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iSWAPBANK, Picture1.Image, Index, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuContrast1_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iCONTRAST, Picture1.Image, 10, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuContrast2_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iCONTRAST, Picture1.Image, -10, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuBW1_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iCOLDEPTH1, Picture1.Image, RGB(180, 180, 180), mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuBW2_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iCOLDEPTH2, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuBW3_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iCOLDEPTH3, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuBW4_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iCOLDEPTH4, Picture1.Image, 15, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuBW5_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iCOLDEPTH5, Picture1.Image, 15, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuBW6_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iCOLDEPTH6, Picture1.Image, 15, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuAqua_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iAQUA, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuDilate_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iDILATE, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuErode_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iERODE, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuConnection_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iCONNECTION, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuStretch_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iSTRETCH, Picture1.Image, 0, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuAddNoise_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iADDNOISE, Picture1.Image, 50, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuSaturationD_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iSATURATION, Picture1.Image, -20, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub

Private Sub mnuSaturationI_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iSATURATION, Picture1.Image, 15, mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub


Private Sub mnuGamma_Click()
  Timer1.Enabled = True
  Call BitBlt(Picture2.HDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.HDC, 0, 0, SRCCOPY)
  Call FilterG(iGAMMA, Picture1.Image, InputBox("Gamma Factor * 100"), mProgress)
  Call BitBlt(Me.HDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.HDC, 0, 0, SRCCOPY)
  Me.Refresh
  
End Sub


