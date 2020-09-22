VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00E6D8BB&
   BorderStyle     =   0  'None
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
   ControlBox      =   0   'False
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   168
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   960
      Top             =   1440
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MyMenu As wsMenu

Private Sub Form_Activate()
    Timer1.Interval = 50
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MyMenu.MouseDown Button, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MyMenu.MouseMove X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MyMenu.MouseUp Button, X, Y
End Sub

Private Sub Form_Paint()
    MyMenu.DrawMenu
End Sub

Private Sub Timer1_Timer()
    Dim hWin As Long, WfP As Long, CP As PointAPI
    hWin = GetActiveWindow()
    GetCursorPos CP
    WfP = WindowFromPoint(CP.X, CP.Y)
    If WfP <> Me.hwnd Then
        MyMenu.MouseMove -5, -5
    End If
    If hWin <> Me.hwnd Then
        If MyMenu.ActiveChild Is Nothing Then
            MyMenu.CloseMenu
        ElseIf MyMenu.ActiveChild.Closed Then
            Set MyMenu.ActiveChild = Nothing
            MyMenu.CloseMenu
        End If
    End If
End Sub

