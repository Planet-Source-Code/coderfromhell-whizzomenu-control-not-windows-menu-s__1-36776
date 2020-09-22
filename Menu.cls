VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public MyWindow As Object               'This is the ascociated form
Public MenuButtons As Collection        'This is a collection of buttons in this menu
Public OverButton As wsMenuButton       'This is the button the cursor is over
Public Parent As Object                 'This is the menu that opened this menu
Public ActiveChild As wsMenu            'This is the menu a expanded button has opened
Private m_NewMenuButton As wsMenuButton

Public Left As Long                   'Menu Position on Screen
Public Top As Long                    '"    "        "  "
Public width As Long
Public height As Long

Dim m_Closed As Boolean

Public BackColor As Long
Public BorderColor As Long
Public TextColor As Long
Public HoverColor As Long
Public HoverTextColor As Long

Public BorderStyle As WSBorderStyleConstants
Public EdgeStyle As WSEdgeConstants

Public Key As Variant

Private Const ButtonHeight = 20
Private Const ButtonSepHeight = 10
Private Const Border = 2

Public MouseButtonAllowed As ButtonCaptureConstants
Public MouseClickWhen As ButtonClickConstants

Public hArrow As Long

Dim HasEmpty As Boolean


Public Function AddMenuButton(ByVal Caption As String, ByVal Key As Variant, Optional ByVal hIcon As Long = -1, Optional Child As wsMenu, Optional ByVal Before As Variant, Optional ByVal After As Variant) As wsMenuButton
    Dim NewButton As wsMenuButton
    On Error GoTo ErrHandler
    If Me.MenuButtons Is Nothing Then Set Me.MenuButtons = New Collection
    Set NewButton = New wsMenuButton
    With NewButton
        .Caption = Caption
        .Key = Key
        Set .Parent = Me
        Set .MyWindow = MyWindow
        If Me.height = 0 Then Me.height = Border
        .SetPosition Border, Me.height, Me.width - Border, Me.height + ButtonHeight
        If Not hIcon = -1 Then .hIcon = hIcon
        If Not Child Is Nothing Then Set .Child = Child
    End With
    Me.MenuButtons.Add NewButton, Key, Before, After
'    WorkOutScale Caption
    Set m_NewMenuButton = NewButton
    Set AddMenuButton = NewButton
    If HasEmpty And Key <> "\EmptyMenu\" Then RemoveMenuButton "\EmptyMenu\": HasEmpty = False
    Exit Function
ErrHandler:
    'Debug.Print Err.Description
End Function

Public Sub RemoveMenuButton(ByVal Index As Variant)
    Me.MenuButtons.Remove Index
    If MenuButtons.Count = 0 Then
        HasEmpty = True
        Me.AddMenuButton "Empty", "\EmptyMenu\"
        Me.NewMenuButton.Enabled = False
    End If
End Sub

Private Sub WorkOutScale()
    Dim NewWid As Long, NewHgt As Long, tmp As Long, MB As wsMenuButton
    Dim TmpHgt As Long, TopCount As Long
    
    NewHgt = Border
    For Each MB In Me.MenuButtons
        If MB.Seperator Then
            NewHgt = NewHgt + ButtonSepHeight
        Else
            NewHgt = NewHgt + ButtonHeight
        End If
        tmp = CLng(GlobalControl.TextWidth(MB.Caption)) + 50
        If tmp > NewWid Then NewWid = tmp
    Next
    TopCount = Border
    For Each MB In Me.MenuButtons
        If MB.Seperator Then
            TmpHgt = ButtonSepHeight
        Else
            TmpHgt = ButtonHeight
        End If
        MB.SetPosition Border, TopCount, NewWid - Border, TopCount + TmpHgt
        TopCount = TopCount + TmpHgt
    Next
    Me.width = NewWid
    Me.height = NewHgt
End Sub

Public Sub CloseMenu()
    If Not MyWindow Is Nothing Then Unload MyWindow
    Set MyWindow = Nothing
    Set OverButton = Nothing
    Set Parent = Nothing
    If Not ActiveChild Is Nothing Then ActiveChild.CloseMenu
    Set ActiveChild = Nothing
    m_Closed = True
    If Not Me.Parent Is Nothing Then
        Set Me.Parent.ActiveChild = Nothing
    End If
    GlobalControl.eMenuClose Me
End Sub

Public Sub ShowMenu(ByVal X As Single, ByVal Y As Single)
    Dim MB As wsMenuButton
    If Not m_Closed Then
        Do Until m_Closed
            DoEvents
        Loop
    End If
    Set MyWindow = New frmMenu
    WorkOutScale
    Load MyWindow
    MyWindow.BackColor = Me.BackColor
    OnTop MyWindow.hwnd, X, Y, Me.width, Me.height + Border
    'MyWindow.Width = Width * Screen.TwipsPerPixelX
    'MyWindow.Height = (Height + Border) * Screen.TwipsPerPixelY
    Me.Left = X
    Me.Top = Y
    'MyWindow.Left = x * Screen.TwipsPerPixelX
    'MyWindow.Top = y * Screen.TwipsPerPixelY
    Set MyWindow.MyMenu = Me
    ActivateButtons
    MyWindow.Show
    m_Closed = False
    Me.DrawMenu
'    SetActiveWindow MyWindow.hwnd
End Sub

Private Sub ActivateButtons()
    Dim MB As wsMenuButton
    If Not Me.MenuButtons Is Nothing Then
        For Each MB In Me.MenuButtons
            Set MB.MyWindow = MyWindow
            MB.Over = False
            
        Next
    End If
End Sub

Public Sub HideMenu()
    If Not MyWindow Is Nothing Then Unload MyWindow
    Set MyWindow = Nothing
End Sub

Private Sub Class_Initialize()
    'Set MenuButtons = New Collection
    m_Closed = True
    MouseButtonAllowed = WS_MouseButton1
    MouseClickWhen = WS_ButtonDown
    Me.BackColor = vbWhite
    Me.BorderColor = vbBlue
    Me.HoverColor = vbBlue
    Me.TextColor = vbBlue
    Me.HoverTextColor = vbWhite
    Me.BorderStyle = ws_colouredline
    Me.EdgeStyle = WS_RaisedOuter
    HasEmpty = True
    Me.AddMenuButton "Empty", "\EmptyMenu\"
    Me.NewMenuButton.Enabled = False
End Sub

Private Sub Class_Terminate()
    If Not m_Closed Then CloseMenu
End Sub

Public Sub MouseDown(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonClick Button, X, Y, WS_ButtonDown
End Sub

Public Sub MouseUp(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
    ButtonClick Button, X, Y, WS_ButtonUp
End Sub

Public Sub MouseMove(ByVal X As Single, ByVal Y As Single)
    IsOverButton X, Y
End Sub

Private Sub IsOverButton(ByVal X As Single, ByVal Y As Single)
    Dim MB As wsMenuButton
    If OverButton Is Nothing Then Set OverButton = New wsMenuButton
    For Each MB In Me.MenuButtons
        With MB
            If X >= .Left And X <= .Right Then
                If Y >= .Top And Y <= .Bottom Then
                    If OverButton Is MB Then Exit Sub
                    OverButton.Over = False
                    OverButton.Draw
                    If Not Me.ActiveChild Is Nothing Then
                        Me.ActiveChild.CloseMenu
                        Set Me.ActiveChild = Nothing
                    End If
                    Set OverButton = MB
                    OverButton.Over = True
                    OverButton.Draw
                    Call GlobalControl.eButtonOver(Me, OverButton)
                    If Not OverButton.Child Is Nothing Then
                        OverButton.Child.ShowMenu Me.Left + Me.width, Me.Top + OverButton.Top
                        Set Me.ActiveChild = OverButton.Child
                    End If
                    Set MB = Nothing
                    Exit Sub
                End If
            End If
        End With
    Next
    If Not OverButton Is Nothing Then
        If Me.ActiveChild Is Nothing Then
            OverButton.Over = False
            OverButton.Draw
            Set OverButton = Nothing
        End If
    End If
End Sub

Private Sub ButtonClick(Button As Integer, X As Single, Y As Single, UorD As ButtonClickConstants)
    If Button - 1 <> Me.MouseButtonAllowed Then Exit Sub
    If UorD <> Me.MouseClickWhen Then Exit Sub
    Dim MB As wsMenuButton
    For Each MB In Me.MenuButtons
        With MB
            If .Enabled And Not .Seperator Then
                If X >= .Left And X <= .Right Then
                    If Y >= .Top And Y <= .Bottom Then
                        If Not OverButton Is Nothing Then OverButton.Over = False
                        If Not OverButton.Child Is Nothing Then
                            OverButton.Child.ShowMenu Me.Left + Me.width, Me.Top + OverButton.Top
                            Set Me.ActiveChild = OverButton.Child
                        Else
                            GlobalControl.CloseAllMenus
                        End If
                        If UorD = WS_ButtonDown Then
                            GlobalControl.eButtonDown Me, MB, Button
                        Else
                            GlobalControl.eButtonUp Me, MB, Button
                        End If
                    End If
                End If
            End If
        End With
    Next MB
End Sub

Public Sub DrawMenu()
    On Error GoTo ErrHandler
    If Not MyWindow Is Nothing Then
        Dim rRect As RECT
        MyWindow.Cls
        SetRect rRect, 0, 0, Me.width, Me.height + Border
        Select Case Me.BorderStyle
            Case 0      'Colored Line
                DrawBorderLine rRect, False
            Case 1      'Colored Line + Edge
                DrawBorderLine rRect, True
                DrawBorderEdge rRect, True
            Case 2      'Edge
                DrawBorderEdge rRect, False
        End Select
        Dim MB As wsMenuButton
        For Each MB In Me.MenuButtons
            MB.Draw
        Next
    End If
ErrHandler:
End Sub

Private Sub DrawBorderLine(rRect As RECT, HasEdge As Boolean)
    MyWindow.BackColor = Me.BackColor
    SetColor BorderColor, 0, vbFSTransparent
    Rectangle MyWindow.hDC, rRect.Left, rRect.Top, rRect.Right, rRect.Bottom
    If Not HasEdge Then
        SetColor vbWhite, 0, vbFSTransparent
        Rectangle MyWindow.hDC, rRect.Left + 1, rRect.Top + 1, rRect.Right - 1, rRect.Bottom - 1
    End If
End Sub

Private Sub DrawBorderEdge(rRect As RECT, HasLine As Boolean)
    If HasLine Then ShrinkRect rRect, 1
    Dim Bdr As Long, Bf As Long
    Select Case Me.EdgeStyle
        Case 0: Bdr = BDR_RAISED:           Bf = BF_RECT + BF_SOFT
        Case 1: Bdr = BDR_RAISEDOUTER:      Bf = BF_RECT + BF_SOFT
        Case 2: Bdr = BDR_RAISEDINNER:      Bf = BF_RECT + BF_SOFT
        Case 3: Bdr = BDR_SUNKEN:           Bf = BF_RECT + BF_SOFT
        Case 4: Bdr = BDR_SUNKENOUTER:      Bf = BF_RECT + BF_SOFT
        Case 5: Bdr = BDR_SUNKENINNER:      Bf = BF_RECT + BF_SOFT
        Case 6: Bdr = BDR_RAISED:           Bf = BF_RECT + BF_SOFT + BF_FLAT
    End Select
    DrawEdge MyWindow.hDC, rRect, Bdr, Bf
End Sub

Friend Sub SetColor(ByVal ForeColor As Long, ByVal FillColor As Long, ByVal Fillstyle As FillStyleConstants)
    If Not MyWindow Is Nothing Then
        MyWindow.ForeColor = ForeColor
        MyWindow.FillColor = FillColor
        MyWindow.Fillstyle = Fillstyle
    End If
End Sub

Public Property Get Closed() As Boolean
    Closed = m_Closed
End Property

Private Sub ShrinkRect(rRect As RECT, aAmount As Long)
    With rRect
        .Left = .Left + aAmount
        .Top = .Top + aAmount
        .Right = .Right - aAmount
        .Bottom = .Bottom - aAmount
    End With
End Sub

Public Sub Win95Style()
    Me.BorderStyle = WS_Edge
    Me.EdgeStyle = WS_Raised
    Me.BackColor = vbButtonFace
    Me.TextColor = vbButtonText
    Me.HoverColor = vbHighlight
    Me.HoverTextColor = vbHighlightText
End Sub

Public Property Get NewMenuButton() As wsMenuButton
    Set NewMenuButton = m_NewMenuButton
End Property

Public Sub GetButtonScreenPos(ByVal wButton As Variant, Left As Long, Top As Long, Right As Long, Bottom As Long)
    On Error GoTo XS
    If Not MyWindow Is Nothing Then
        Dim wp As RECT
        GetWindowRect MyWindow.hwnd, wp
        If Not Me.MenuButtons(wButton) Is Nothing Then
            With Me.MenuButtons(wButton)
                Left = wp.Left + .Left
                Top = wp.Top + .Top
                Right = wp.Right + .Right
                Bottom = wp.Bottom + .Bottom
            End With
        End If
    End If
XS:
End Sub
