VERSION 5.00
Begin VB.Form frmDisableIcon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disabled Icon Options"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optIcon 
      Caption         =   "Use this >"
      Height          =   210
      Index           =   5
      Left            =   105
      TabIndex        =   21
      Top             =   3660
      Width           =   1020
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   5
      Left            =   1185
      Picture         =   "frmDisableIcon.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3465
      Width           =   540
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   4
      Left            =   1170
      Picture         =   "frmDisableIcon.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2880
      Width           =   540
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "Use this >"
      Height          =   210
      Index           =   4
      Left            =   90
      TabIndex        =   18
      Top             =   3075
      Width           =   1020
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   3
      Left            =   1185
      Picture         =   "frmDisableIcon.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2280
      Width           =   540
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "Use this >"
      Height          =   210
      Index           =   3
      Left            =   105
      TabIndex        =   16
      Top             =   2475
      Width           =   1020
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "Use this >"
      Height          =   210
      Index           =   2
      Left            =   105
      TabIndex        =   15
      Top             =   1875
      Width           =   1020
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "Use this >"
      Height          =   210
      Index           =   1
      Left            =   105
      TabIndex        =   14
      Top             =   1290
      Width           =   1020
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "Use this >"
      Height          =   210
      Index           =   0
      Left            =   105
      TabIndex        =   13
      Top             =   675
      Width           =   1020
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   2
      Left            =   1185
      Picture         =   "frmDisableIcon.frx":0CC6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
      Width           =   540
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   1
      Left            =   1185
      Picture         =   "frmDisableIcon.frx":1108
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1080
      Width           =   540
   End
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   840
      Index           =   4
      Left            =   1935
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1095
      Width           =   1080
   End
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   840
      Index           =   3
      Left            =   1935
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2835
      Width           =   1080
   End
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   840
      Index           =   2
      Left            =   1935
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1965
      Width           =   1080
   End
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   555
      Index           =   1
      Left            =   1935
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3705
      Width           =   1080
   End
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   840
      Index           =   0
      Left            =   1935
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   225
      Width           =   1080
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   0
      Left            =   1185
      Picture         =   "frmDisableIcon.frx":154A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Gray disabled icon using DrawState && DSS_Mono"
      Height          =   510
      Index           =   4
      Left            =   3180
      TabIndex        =   10
      Top             =   1275
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Disabled Icon (Custom colors #2)"
      Height          =   225
      Index           =   3
      Left            =   3165
      TabIndex        =   8
      Top             =   3135
      Width           =   2460
   End
   Begin VB.Label Label1 
      Caption         =   "Disabled Icon (Custom colors #1)"
      Height          =   225
      Index           =   2
      Left            =   3165
      TabIndex        =   6
      Top             =   2295
      Width           =   2460
   End
   Begin VB.Label Label1 
      Caption         =   "Disabled Icon with color. Actually combo of DrawIconEx && BitBlt APIs"
      Height          =   435
      Index           =   1
      Left            =   3150
      TabIndex        =   5
      Top             =   3795
      Width           =   2565
   End
   Begin VB.Label Label1 
      Caption         =   "Standard gray disabled icon using DrawState && DSS_Disabled"
      Height          =   540
      Index           =   0
      Left            =   3165
      TabIndex        =   4
      Top             =   420
      Width           =   2415
   End
End
Attribute VB_Name = "frmDisableIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' note that the lParam parameter was changed to Any vs Long so it can accept Icon handles and also string values
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, lParam As Any, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Const DSS_DISABLED As Long = &H20
Private Const DSS_MONO As Long = &H80
Private Const DSS_NORMAL As Long = &H0
Private Const DST_ICON As Long = &H3
Private Const DST_TEXT As Long = &H1

Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Const DI_NORMAL As Long = &H3

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Const WHITENESS As Long = &HFF0062

Private IconID As Integer

Private Sub MakeDisabledIcons()
picDest(0).Cls
picDest(1).Cls
picDest(2).Cls
picDest(3).Cls
picDest(4).Cls

' This is the generally accepted way of creating a disabled icon and text & is very quick
DrawState picDest(0).hdc, 0&, 0&, ByVal picIcon(IconID).Picture.handle, 0&, 10, 0, 0, 0, DST_ICON Or DSS_DISABLED
DrawState picDest(0).hdc, 0&, 0&, ByVal "Disabled", 0&, 10, 36, 0, 0, DST_TEXT Or DSS_DISABLED

' declare brushes used in the samples below
Dim hBrushEdge1 As Long, hBrushIcon1 As Long
Dim hBrushEdge2 As Long, hBrushIcon2 As Long
Dim hBrushIcon3 As Long
' the hBrushEdge will be the icon outline and text shadow colors
' the hBrushIcon will be the icon filler and text forecolor

Dim Cx As Long, Cy As Long
' note that I don't set these in the samples below because I want to draw the icon using the same
' dimensions as the source icon. If you wanted to stretch the destination icon to a different size,
' then you would supply values for the cx & cy parameters

' brushes for samples
hBrushEdge1 = CreateSolidBrush(vbWhite)
hBrushIcon1 = CreateSolidBrush(&HC0C0C0)

hBrushEdge2 = CreateSolidBrush(&HC0E0FF)

hBrushIcon2 = CreateSolidBrush(&H808080)
hBrushIcon3 = CreateSolidBrush(&H808080)

' this routine simply duplicates the standard gray disabled icons & text shown above...
' Note: wouldn't use this 'cause it takes 2 API calls vs 1 for each icon or text you want disabled with default gray
' This example only shows you how to go about it a different way to achieve the same result
DrawState picDest(4).hdc, hBrushEdge1, 0&, ByVal picIcon(IconID).Picture.handle, 0&, 11, 1, Cx, Cy, DST_ICON Or DSS_MONO
DrawState picDest(4).hdc, hBrushIcon3, 0&, ByVal picIcon(IconID).Picture.handle, 0&, 10, 0, Cx, Cy, DST_ICON Or DSS_MONO
DrawState picDest(4).hdc, hBrushEdge1, 0&, ByVal "Disabled", 0&, 11, 37, Cx, Cy, DST_TEXT Or DSS_MONO
DrawState picDest(4).hdc, hBrushIcon3, 0&, ByVal "Disabled", 0&, 10, 36, Cx, Cy, DST_TEXT Or DSS_MONO

' this routine is the custom color sample 1.
' Here the outline is shifted down to the right by 1 pixel before the "disabled" icon is displayed
DrawState picDest(2).hdc, hBrushEdge1, 0&, ByVal picIcon(IconID).Picture.handle, 0&, 11, 1, Cx, Cy, DST_ICON Or DSS_MONO
DrawState picDest(2).hdc, hBrushIcon1, 0&, ByVal picIcon(IconID).Picture.handle, 0&, 10, 0, Cx, Cy, DST_ICON Or DSS_MONO
DrawState picDest(2).hdc, hBrushEdge1, 0&, ByVal "Disabled", 0&, 11, 37, Cx, Cy, DST_TEXT Or DSS_MONO
DrawState picDest(2).hdc, hBrushIcon1, 0&, ByVal "Disabled", 0&, 10, 36, Cx, Cy, DST_TEXT Or DSS_MONO

' this routine is the custom color sample 2.
' Here the outline is shifted up to the left by 1 pixel before the "disabled" icon is displayed
DrawState picDest(3).hdc, hBrushEdge2, 0&, ByVal picIcon(IconID).Picture.handle, 0&, 10, 0, Cx, Cy, DST_ICON Or DSS_MONO
DrawState picDest(3).hdc, hBrushIcon2, 0&, ByVal picIcon(IconID).Picture.handle, 0&, 11, 1, Cx, Cy, DST_ICON Or DSS_MONO
DrawState picDest(3).hdc, hBrushEdge2, 0&, ByVal "Disabled", 0&, 11, 37, Cx, Cy, DST_TEXT Or DSS_MONO
DrawState picDest(3).hdc, hBrushIcon2, 0&, ByVal "Disabled", 0&, 10, 36, Cx, Cy, DST_TEXT Or DSS_MONO

' Let's trash the brushes
DeleteObject hBrushIcon1
DeleteObject hBrushEdge1
DeleteObject hBrushIcon2
DeleteObject hBrushEdge2
DeleteObject hBrushIcon3

' call routine to fade colored icon into a grayish brush
Call ColorDisableIcon

picDest(0).Refresh
picDest(1).Refresh
picDest(2).Refresh
picDest(3).Refresh
picDest(4).Refresh
End Sub


Private Sub ColorDisableIcon()

Const MAGICROP = &HB8074A

Dim hMemDC As Long
Dim hBitmap As Long, picWidth As Long, picHeight As Long
Dim hOldBackColor As Long, hbrShadow As Long
    
picWidth = picDest(1).Width / Screen.TwipsPerPixelX
picHeight = picDest(1).Height / Screen.TwipsPerPixelY

    ' Create a temporary DC and bitmap to hold the image
    hMemDC = CreateCompatibleDC(picDest(1).hdc)
    hBitmap = SelectObject(hMemDC, CreateCompatibleBitmap(picDest(1).hdc, picWidth, picHeight))
    
    ' copy the source DC background to the memory DC
    BitBlt hMemDC, 0, 0, picWidth, picHeight, picDest(1).hdc, 0, 0, vbSrcCopy
    
    ' this will help create the fade/blend effect
    PatBlt hMemDC, 0, 0, picWidth, picHeight, WHITENESS

    ' draw the icon onto the temp DC, using same dimensions as source
    DrawIconEx hMemDC, 10, 0, picIcon(IconID).Picture.handle, 0, 0, 0, 0, DI_NORMAL

    ' now set some colors that will be used to create the soft blend of icon icolors
    hOldBackColor = SetBkColor(picDest(1).hdc, vbWhite)
    hbrShadow = SelectObject(picDest(1).hdc, CreateSolidBrush(8421504))
    
    ' copy the finished product back to the source DC
    BitBlt picDest(1).hdc, 0, 0, picWidth, picHeight, hMemDC, 0, 0, MAGICROP
  
    ' replace, reset & delete memory objects
    SetBkColor picDest(1).hdc, hOldBackColor
    DeleteObject SelectObject(picDest(1).hdc, hbrShadow)
    DeleteObject SelectObject(hMemDC, hBitmap)
    DeleteDC hMemDC

End Sub

Private Sub Form_Load()
Randomize Timer
optIcon(Int(Rnd * optIcon.UBound)) = True
End Sub

Private Sub optIcon_Click(Index As Integer)
IconID = Index
Call MakeDisabledIcons
End Sub

