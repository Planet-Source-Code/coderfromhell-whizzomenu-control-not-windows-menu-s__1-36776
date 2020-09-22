VERSION 5.00
Object = "{5DD6A649-644F-4488-A5F7-0EEAFCEF584A}#7.0#0"; "WhizzoMenu.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin WhizzoMenu.WMenu WMenu2 
      Left            =   3360
      Top             =   1440
      _ExtentX        =   1217
      _ExtentY        =   1217
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":045C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&File"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin WhizzoMenu.WMenu WMenu1 
      Left            =   840
      Top             =   720
      _ExtentX        =   1217
      _ExtentY        =   1217
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   2400
      Picture         =   "frmMain.frx":1168
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "frmMain.frx":15AA
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   1440
      Picture         =   "frmMain.frx":19EC
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   960
      Picture         =   "frmMain.frx":1E2E
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim MainM As wsMenu
Dim Side1 As wsMenu
Dim Side2 As wsMenu

Private Sub Command1_Click()
    Dim MM As wsMenu
    Set MM = Me.WMenu1.AddMenu("FileMenu", vbWhite)
    If Not MM Is Nothing Then
        With MM
            .BackColor = vbWhite
            .BorderColor = vbRed
            .TextColor = vbBlack
            .HoverColor = vbRed
            .HoverTextColor = vbWhite
            .MouseButtonAllowed = WS_MouseButton1
            .MouseClickWhen = WS_ButtonUp
            Dim i As Integer
            For i = 0 To 30
                .AddMenuButton "Hey! This is Menu Item " & i, "mb" & i
                If right(CStr(i), 1) = "0" Then .MenuButtons("mb" & i).Enabled = False
                If right(CStr(i), 1) = "4" Then .MenuButtons("mb" & i).Seperator = True
            Next i
            Dim rrect As RECT
            GetWindowRect Command1.hwnd, rrect
            .ShowMenu rrect.left, rrect.bottom
        End With
    End If
End Sub

Private Sub Command2_Click()
    Dim R As RECT
    GetWindowRect Command2.hwnd, R
    WMenu1.WhizzoMenus("Main").ShowMenu R.left, R.bottom
End Sub

Private Sub Form_Load()
    Set Side2 = Me.WMenu1.AddMenu("Side2", Me.BackColor)
    Side2.MouseButtonAllowed = WS_MouseButton1
    With Side2
        .AddMenuButton "Menu Button 1", "b1"
        .AddMenuButton "Menu Button 2", "b2"
        .MenuButtons("b2").Seperator = True
        .AddMenuButton "Menu Button 3", "b3"
        .MenuButtons("b3").Enabled = False
        .AddMenuButton "Menu Button 4", "b4"
        .BorderStyle = WS_ColouredLineAndEdge
        .EdgeStyle = WS_RaisedOuter
        .BorderColor = vbRed
        .BackColor = RGB(220, 220, 220)
        .TextColor = vbBlack
        .HoverColor = RGB(150, 150, 150)
        .HoverTextColor = vbWhite
    End With
    Set Side1 = Me.WMenu1.AddMenu("Side1", Me.BackColor)
    Side1.MouseButtonAllowed = WS_MouseButton1
    With Side1
        .AddMenuButton "Menu Button 1", "b1"
        .AddMenuButton "Menu Button 2", "b2"
        .AddMenuButton "Menu Button 3", "b3"
        .AddMenuButton "Menu Button 4", "b4", , Side2
    End With
    Set MainM = Me.WMenu1.AddMenu("Main", Me.BackColor)
    MainM.MouseButtonAllowed = WS_MouseButton1
    With MainM
        .AddMenuButton "Menu Button 1", "b1", ImageList1.ListImages(1).Picture.Handle
        .AddMenuButton "Menu Button 2", "b2", ImageList1.ListImages(2).Picture.Handle, Side1
        .AddMenuButton "Menu Button 3", "b3", ImageList1.ListImages(3).Picture.Handle
        .MenuButtons("b3").Enabled = False
        .AddMenuButton "Menu Button 4", "b4", ImageList1.ListImages(4).Picture.Handle
        .AddMenuButton "", "seperator1"
        .NewMenuButton.Seperator = True
        .AddMenuButton "E&xit", "Exit"
        .Win95Style
    End With
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Dim Xy As POINTAPI
        GetCursorPos Xy
        MainM.ShowMenu Xy.X, Xy.Y
    End If
End Sub

Private Sub WMenu1_ButtonDown(wMenu As WhizzoMenu.wsMenu, wButton As WhizzoMenu.wsMenuButton, MouseButton As Integer)
    If wMenu.Key = "Side2" And wButton.Key = "b3" Then
        MsgBox "It Works"
    ElseIf wMenu.Key = "Main" And wButton.Key = "Exit" Then
        Unload Me
    End If
End Sub

Private Sub WMenu1_ButtonUp(wMenu As WhizzoMenu.wsMenu, wButton As WhizzoMenu.wsMenuButton, MouseButton As Integer)
    If wMenu.Key = "FileMenu" And wButton.Key = "mb6" Then
        MsgBox "Hello! Thankyou for clicking number 6"
    End If
End Sub

Private Sub WMenu1_MenuClose(Menu As WhizzoMenu.wsMenu)
    If Menu.Key = "FileMenu" Then
        Me.WMenu1.RemoveMenu (Menu.Key)
    End If
End Sub
