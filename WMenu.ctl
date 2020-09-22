VERSION 5.00
Begin VB.UserControl WMenu 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "WMenu.ctx":0000
   Begin VB.PictureBox IconHolder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1440
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   480
      Picture         =   "WMenu.ctx":0312
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   600
      Width           =   540
   End
   Begin VB.Image arrow3 
      Height          =   240
      Left            =   2040
      Picture         =   "WMenu.ctx":061C
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image arrow2 
      Height          =   240
      Left            =   1680
      Picture         =   "WMenu.ctx":0766
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image arrow1 
      Height          =   240
      Left            =   1320
      Picture         =   "WMenu.ctx":08B0
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "WMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WhizzoMenus As Collection

Enum ButtonCaptureConstants
    WS_MouseButton1 = 0
    WS_MouseButton2 = 1
    WS_MouseButton3 = 2
    WS_MouseButton4 = 3
End Enum

Enum ButtonClickConstants
    WS_ButtonDown = 1
    WS_ButtonUp = 2
End Enum

Enum WSBorderStyleConstants
    ws_colouredline = 0
    WS_ColouredLineAndEdge = 1
    WS_Edge = 2
End Enum

Enum WSEdgeConstants
    WS_Raised = 0
    WS_RaisedOuter = 1
    WS_RaisedInner = 2
    WS_Sunken = 3
    WS_SunkenOuter = 4
    WS_SunkenInner = 5
    WS_Flat = 6
End Enum

Event ButtonOver(wMenu As wsMenu, wButton As wsMenuButton)
Event ButtonDown(wMenu As wsMenu, wButton As wsMenuButton, MouseButton As Integer)
Event ButtonUp(wMenu As wsMenu, wButton As wsMenuButton, MouseButton As Integer)
Event MenuClose(Menu As wsMenu)
Event CloseAll()


Private Sub UserControl_Paint()
    UserControl_Resize
    UserControl.Cls
    Dim MyRect As RECT
    SetRect MyRect, 0, 0, ScaleWidth, ScaleHeight
    DrawEdge UserControl.hDC, MyRect, BDR_RAISED, BF_RECT
End Sub

Private Sub UserControl_Resize()
    If UserControl.width <> (Picture1.width + 10) * Screen.TwipsPerPixelX Then UserControl.width = (Picture1.width + 10) * Screen.TwipsPerPixelX: Exit Sub
    If UserControl.height <> (Picture1.height + 10) * Screen.TwipsPerPixelY Then UserControl.height = (Picture1.height + 10) * Screen.TwipsPerPixelY: Exit Sub
    Picture1.Move 5, 5
End Sub

Friend Sub eButtonOver(wMenu As wsMenu, wButton As wsMenuButton)
    RaiseEvent ButtonOver(wMenu, wButton)
End Sub
Friend Sub eButtonDown(wMenu As wsMenu, wButton As wsMenuButton, MouseButton As Integer)
    RaiseEvent ButtonDown(wMenu, wButton, MouseButton)
End Sub
Friend Sub eButtonUp(wMenu As wsMenu, wButton As wsMenuButton, MouseButton As Integer)
    RaiseEvent ButtonUp(wMenu, wButton, MouseButton)
End Sub
Friend Sub eMenuClose(wMenu As wsMenu)
    RaiseEvent MenuClose(wMenu)
End Sub

Public Function AddMenu(ByVal Key As Variant, Optional ByVal BackColor As Long = vbButtonFace) As wsMenu
    Dim NewMenu As New wsMenu
    On Error GoTo ErrHandler
    Set GlobalControl = Me
    With NewMenu
        .Key = Key
        .BackColor = BackColor
        .hArrow = arrow3.Picture.Handle
    End With
    If Me.WhizzoMenus Is Nothing Then Set Me.WhizzoMenus = New Collection
    Me.WhizzoMenus.Add NewMenu, Key
    Set AddMenu = NewMenu
    Exit Function
ErrHandler:
    Set NewMenu = Nothing
    Set AddMenu = Nothing
End Function

Public Sub RemoveMenu(Index As Variant)
    Me.WhizzoMenus.Remove Index
End Sub

Friend Function TextWidth(ByVal Str As String) As Single
    TextWidth = UserControl.TextWidth(Str)
End Function

Friend Sub CloseAllMenus()
    Dim Mnu As wsMenu
    For Each Mnu In Me.WhizzoMenus
        If Not Mnu.Closed Then Mnu.CloseMenu
    Next
    RaiseEvent CloseAll
End Sub
