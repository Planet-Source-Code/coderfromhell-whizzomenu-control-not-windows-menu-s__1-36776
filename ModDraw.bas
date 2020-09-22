Attribute VB_Name = "ModDraw"
'********************************************************************************
'Various API calls used to draw directly to the hDC of the control or form
'********************************************************************************
    Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
    Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
    Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
    Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
    Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
    Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
    Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
    Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
    Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
    Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Integer) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
    Declare Function AlphaBlending Lib "Alphablending.dll" _
                         (ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, _
                          ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, _
                          ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal AlphaSource As Long) As Long
    Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
    Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
    Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
    Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
    Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function GetActiveWindow Lib "user32" () As Long
    Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'    Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
'    Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long
'
'    Private Type typSHFILEINFO
'      hIcon As Long
'      iIcon As Long
'      dwAttributes As Long
'      szDisplayName As String * 260
'      szTypeName As String * 80
'    End Type
'
'Private Const SHGFI_DISPLAYNAME = &H200
'Private Const SHGFI_EXETYPE = &H2000
'Private Const SHGFI_SYSICONINDEX = &H4000
'Private Const SHGFI_SHELLICONSIZE = &H4
'Private Const SHGFI_TYPENAME = &H400
'Private Const SHGFI_LARGEICON = &H0
'Private Const SHGFI_SMALLICON = &H1
'Private Const ILD_TRANSPARENT = &H1
'Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

    
    
    
Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_DESKTOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40


Type PointAPI
    x As Long
    y As Long
End Type

Enum XorY
    zX = 0
    zY = 1
End Enum

'Constants for DrawIconEX
Public Const DI_NORMAL = &H3
'Functions for GetIcon
'Contstants for GetIcon
Private Const WM_GETICON = &H7F
Private Const GCL_HICON = (-14)
Private Const GCL_HICONSM = (-34)
Private Const WM_QUERYDRAGICON = &H37

'********************************************************************************
'RECT type is used in most graphical related API calls
'********************************************************************************
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'********************************************************************************
'Constants for DRAWEDGE API CALL
'********************************************************************************
Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_ADJUST = &H2000
Public Const BF_BOTTOM = &H8
Public Const BF_DIAGONAL = &H10
Public Const BF_FLAT = &H4000
Public Const BF_LEFT = &H1
Public Const BF_MIDDLE = &H800
Public Const BF_MONO = &H8000
Public Const BF_RIGHT = &H4
Public Const BF_SOFT = &H1000
Public Const BF_TOP = &H2
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)

'********************************************************************************
'Constants for DRAWFRAMECONTROL API call
'********************************************************************************
Public Const DFC_BUTTON = 4
Public Const DFC_CAPTION = 1
Public Const DFC_MENU = 2
Public Const DFC_SCROLL = 3
Public Const DFCS_ADJUSTRECT = &H2000
Public Const DFCS_BUTTON3STATE = &H8
Public Const DFCS_BUTTONCHECK = &H0
Public Const DFCS_BUTTONPUSH = &H10
Public Const DFCS_BUTTONRADIO = &H4
Public Const DFCS_BUTTONRADIOIMAGE = &H1
Public Const DFCS_BUTTONRADIOMASK = &H2
Public Const DFCS_CAPTIONCLOSE = &H0
Public Const DFCS_CAPTIONHELP = &H4
Public Const DFCS_CAPTIONMAX = &H2
Public Const DFCS_CAPTIONMIN = &H1
Public Const DFCS_CAPTIONRESTORE = &H3
Public Const DFCS_CHECKED = &H400
Public Const DFCS_FLAT = &H4000
Public Const DFCS_INACTIVE = &H100
Public Const DFCS_MENUARROW = &H0
Public Const DFCS_MENUARROWRIGHT = &H4
Public Const DFCS_MENUBULLET = &H2
Public Const DFCS_MENUCHECK = &H1
Public Const DFCS_MONO = &H8000
Public Const DFCS_PUSHED = &H200
Public Const DFCS_SCROLLCOMBOBOX = &H5
Public Const DFCS_SCROLLDOWN = &H1
Public Const DFCS_SCROLLLEFT = &H2
Public Const DFCS_SCROLLRIGHT = &H3
Public Const DFCS_SCROLLSIZEGRIP = &H8
Public Const DFCS_SCROLLSIZEGRIPRIGHT = &H10
Public Const DFCS_SCROLLUP = &H0

'********************************************************************************
'DrawText constants for Format option of DRAWTEXT API
'********************************************************************************
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Public Const DT_DISPFILE = 6            '  Display-file
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_INTERNAL = &H1000
Public Const DT_LEFT = &H0
Public Const DT_METAFILE = 5            '  Metafile, VDM
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_PLOTTER = 0             '  Vector plotter
Public Const DT_RASCAMERA = 3           '  Raster camera
Public Const DT_RASDISPLAY = 1          '  Raster display
Public Const DT_RASPRINTER = 2          '  Raster printer
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10

Public GlobalControl As wMenu

'Global Const WS_MOUSEBUTTON_1 = 1
'Global Const WS_MOUSEBUTTON_2 = 3
'Global Const WS_MOUSEBUTTON_3 = 5
'Global Const WS_MOUSEBUTTON_4 = 10


Public Function GetIcon(hwnd As Long) As Long
    Call SendMessageTimeout(hwnd, WM_GETICON, 0, 0, 0, 1000, GetIcon)
    If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICONSM)
    If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_GETICON, 1, 0, 0, 1000, GetIcon)
    If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICON)
    If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_QUERYDRAGICON, 0, 0, 0, 1000, GetIcon)
End Function

Sub BltDesktop(sourceX As Integer, sourceY As Integer, targetBox As Object, Optional width As Integer = -1, Optional height As Integer = -1)
    If width = -1 Then width = targetBox.width
    If height = -1 Then height = targetBox.height
    deskHDC = GetDC(0)
    BitBlt targetBox.hDC, 0, 0, width, height, deskHDC, sourceX, sourceY, vbSrcCopy
    ReleaseDC 0, deskHDC
End Sub


Public Function ToTwip(Orientation As XorY, zSource As Single) As Single
    If Orientation = zX Then
        ToTwip = zSource * Screen.TwipsPerPixelX
    Else
        ToTwip = zSource * Screen.TwipsPerPixelY
    End If
End Function

Public Function ToPixel(Orientation As XorY, zSource As Single) As Single
    If Orientation = zX Then
        ToPixel = zSource / Screen.TwipsPerPixelX
    Else
        ToPixel = zSource / Screen.TwipsPerPixelY
    End If
End Function

Public Sub OnTop(hwnd As Long, ByVal Left As Long, ByVal Top As Long, ByVal width As Long, ByVal height As Long)
    SetWindowPos hwnd, HWND_TOPMOST, Left, Top, width, height, SWP_NOACTIVATE '+ SWP_NOZORDER
End Sub


'Private Function ExtractIcon(FileName As String, PictureBox As PictureBox, PixelsXY As Integer) As Long
'    Dim SmallIcon As Long
'    Dim NewHDC As Long
'    Dim FileInfo As typSHFILEINFO
'
'
'    If PixelsXY = 16 Then
'        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
'    Else
'        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
'    End If
'
'    If SmallIcon <> 0 Then
'      With PictureBox
'        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
'        .Refresh
'      End With
'
'
'
'      ExtractIcon = IconIndex
'    End If
'End Function
'
