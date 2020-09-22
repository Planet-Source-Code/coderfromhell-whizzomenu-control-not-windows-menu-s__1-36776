VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsMenuButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public MyWindow As Object       'This is the asociated form.

Public Parent As wsMenu         'This is the menu that the button is on.
Public Child As wsMenu          'If Expandable then this is the menu to open.

'Public IsExpandable As Boolean  'Returns/Sets whether this button opens another menu.

Public Left As Long             'Position of button.
Public Top As Long              '   "
Public Right As Long            '   "
Public Bottom As Long           '   "

Public Caption As String        'Stores Buttons Text.
Public hIcon As Long            'Stores the buttons Icon handle.
Public Tag As String            'Stores other data.

Public Visible As Boolean       'Returns/Sets whether this button is visible.
Public Enabled As Boolean

Public Over As Boolean          'Returns/Sets whether the cursor is over this button.
Public Key As Variant

Public Seperator As Boolean

Public ID As Long


Private Sub Class_Initialize()
    Caption = ""
    hIcon = 0
    Tag = ""
    Visible = True
    Enabled = True
End Sub

Private Sub Class_Terminate()
    Set MyWindow = Nothing
    Set Parent = Nothing
    Set Child = Nothing
End Sub

Public Sub SetPosition(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Me.Left = Left
    Me.Top = Top
    Me.Right = Right
    Me.Bottom = Bottom
End Sub

Public Sub Draw()
    If Not MyWindow Is Nothing Then
        'Draw Back of button
        If Over And Enabled And Not Seperator Then
            Parent.SetColor Parent.HoverColor, Parent.HoverColor, vbFSSolid
        Else
            Parent.SetColor Parent.BackColor, Parent.BackColor, vbFSSolid
        End If
        Rectangle MyWindow.hDC, Left, Top, Right, Bottom
        If Not Seperator Then
            'Draw button text
            If Me.Caption <> "" Then
                Dim rRect As RECT
                SetRect rRect, Left + 24, Top, Right - 20, Bottom
                If Enabled Then
                    If Over Then
                        Parent.SetColor Parent.HoverTextColor, Parent.HoverTextColor, vbFSSolid
                    Else
                        Parent.SetColor Parent.TextColor, Parent.TextColor, vbFSSolid
                    End If
                Else
                    MyWindow.ForeColor = vbGrayText
                    'rRect.Left = rRect.Left + 1: rRect.Top = rRect.Top + 1
                    DrawText MyWindow.hDC, Me.Caption, Len(Me.Caption), rRect, DT_SINGLELINE + DT_LEFT + DT_VCENTER
                    MyWindow.ForeColor = vbWhite
                    'rRect.Left = rRect.Left - 1: rRect.Top = rRect.Top - 1
                    rRect.Left = rRect.Left + 1: rRect.Top = rRect.Top + 1
                End If
                DrawText MyWindow.hDC, Me.Caption, Len(Me.Caption), rRect, DT_SINGLELINE + DT_LEFT + DT_VCENTER
            End If
            'Draw Arrow
            If Not Me.Child Is Nothing Then
                Dim Lft As Long, Tp As Long
                Lft = Me.Right - 12
                Tp = Me.Top + 7
                DrawArrow Tp, Lft
                'DrawIconEx MyWindow.hDC, Lft, Tp, Parent.hArrow, 16, 16, 0, 0, DI_NORMAL
            End If
            'Draw the icon
            If Not hIcon = 0 Then
                DrawIconEx MyWindow.hDC, Me.Left + 2, Me.Top + 2, Me.hIcon, 16, 16, 0, 0, DI_NORMAL
                'DrawIcon MyWindow.hDC, Me.Left + 2, Me.Top + 2, Me.hIcon
            End If
        Else
            MyWindow.Line (Me.Left + 4, Me.Top + 4)-(Me.Right - 4, Me.Top + 4), Parent.TextColor
        End If
    End If
End Sub

Private Sub DrawArrow(ByVal Top As Long, ByVal Left As Long)
    MyWindow.Line (Left, Top + 3)-(Left + 3, Top + 3)
    MyWindow.Line (Left + 4, Top)-(Left + 4, Top + 7)
    MyWindow.Line (Left + 5, Top + 1)-(Left + 5, Top + 6)
    MyWindow.Line (Left + 6, Top + 2)-(Left + 6, Top + 5)
    MyWindow.Line (Left + 7, Top + 3)-(Left + 7, Top + 4)
End Sub

