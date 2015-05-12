VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parent Control"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "¬€ Àﬁ◊»“‹  ŒÃœ‹ﬁ“≈–"
      Height          =   495
      Left            =   3480
      MaskColor       =   &H000000FF&
      TabIndex        =   11
      Top             =   6480
      Width           =   2895
   End
   Begin VB.PictureBox picIcon 
      Height          =   615
      Left            =   6000
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox ExplButton 
      Height          =   50
      Left            =   6360
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   8400
      Top             =   4080
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   8040
      Top             =   4080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8280
      Top             =   4800
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   7920
      Top             =   4800
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "–¿«¡ÀŒ »–Œ¬¿“‹  ŒÃœ‹ﬁ“≈–"
      Height          =   495
      Left            =   3480
      MaskColor       =   &H0000C000&
      TabIndex        =   1
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "œÓÎÓ‚ËÌÍÓÏËÌÛÚ"
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   7920
      TabIndex        =   14
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   7920
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   7920
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "—ÓÓ·˘ÂÌËÂ ÔÓÎ¸ÁÓ‚‡ÚÂÎ˛"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   10455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intFH As Integer
Dim intST As Integer
Dim intPS As Integer
Dim intTC As Integer
Dim i As Integer
Dim timeuse As Integer
Dim filename As String
Dim password As String
Dim timecontrol As Integer
Dim load As Boolean

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Const EWX_LOGOFF As Long = &H0
Private Const EWX_SHUTDOWN As Long = &H1
Private Const EWX_REBOOT As Long = &H2
Private Const EWX_FORCE As Long = &H4
Private Const EWX_POWEROFF As Long = &H8
Private Const EWX_FORCEIFHUNG As Long = &H10

Private Type LUID
dwLowPart As Long
dwHighPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
udtLUID As LUID
dwAttributes As Long
End Type
Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
laa As LUID_AND_ATTRIBUTES
End Type



 Private Function EnableShutdownPrivledges() As Boolean
Dim hProcessHandle As Long
Dim hTokenHandle As Long
Dim lpv_la As LUID
Dim token As TOKEN_PRIVILEGES
hProcessHandle = GetCurrentProcess()
If hProcessHandle <> 0 Then
If OpenProcessToken(hProcessHandle, (&H20 Or &H8), hTokenHandle) <> 0 Then
If LookupPrivilegeValue(vbNullString, "SeShutdownPrivilege", lpv_la) <> 0 Then
With token
.PrivilegeCount = 1
.laa.udtLUID = lpv_la
.laa.dwAttributes = &H2
End With
If AdjustTokenPrivileges(hTokenHandle, False, token, ByVal 0&, ByVal 0&, ByVal 0&) <> 0 Then EnableShutdownPrivledges = True
End If
End If
End If
End Function
 




Private Sub Command1_Click()
If Text1.Text = password Then
Unload Me
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Call EnableShutdownPrivledges
Call ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub

Private Sub Form_Load()
load = True

i = 0
intST = FreeFile()
Open "settings.txt" For Input As intST
Input #intST, timeuse
Close intST

intPS = FreeFile()
Open "password.txt" For Input As intPS
Input #intPS, password
Close intPS

Label2.Caption = password
Label3.Caption = timeuse

    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net

    Dim Stretched As Boolean
    'picIcon.Visible = False
    'API uses pixels
    picIcon.ScaleMode = vbPixels
    'No border
    ExplButton.BorderStyle = 0
    'API uses pixels
    ExplButton.ScaleMode = vbPixels
    'Set graphic mode te 'persistent graphic'
    ExplButton.AutoRedraw = True
    'API uses pixels
    Me.ScaleMode = vbPixels
    'Set the button's caption
    Command3.Caption = "Set Mousecursor on X"

    ' If you set Stretched to true then stretch the icon to te Height and Width of the button
    ' If Stretched=False, the icon will be centered
  '  Stretched = False

  '  If Stretched = True Then
        ' Stretch the Icon
   '     ExplButton.PaintPicture picIcon.Picture, 1, 1, ExplButton.ScaleWidth - 2, ExplButton.ScaleHeight - 2
   ' ElseIf Stretched = False Then
        ' Center the picture of the icon
   '     ExplButton.PaintPicture picIcon.Picture, (ExplButton.ScaleWidth - picIcon.ScaleWidth) / 2, (ExplButton.ScaleHeight - picIcon.ScaleHeight) / 2
  '  End If
    ' Set icon as picture
   ' ExplButton.Picture = ExplButton.Image

intTC = FreeFile()
filename = Date & "tc.txt"
Open filename For Append As intTC
Close intTC

Open filename For Input As intTC
If FileLen(filename) <> 0 Then
Label8.Caption = FileLen(filename)
Input #intTC, timecontrol
End If
Close intTC
Label6.Caption = timecontrol

    
On Error GoTo First
intFH = FreeFile()
filename = Date & ".txt"
Open filename For Input As intST
Input #intST, alreadyuse
Close intST
Label1.Caption = alreadyuse
If alreadyuse = 1 Then

Timer2.Enabled = True
Label4.Caption = "¬–≈Ãﬂ –¿¡Œ“€ «¿  ŒÃœ‹ﬁ“≈–ŒÃ ”∆≈ «¿ ŒÕ◊»ÀŒ—‹!"
Form1.Visible = True
Form1.SetFocus
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End If
Exit Sub




First:
'Resume Next





End Sub


Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Timer1_Timer()

If load = True Then
load = False
i = timecontrol * 2
End If

Label7.Caption = i

If i < timeuse * 2 Then
i = i + 1
Label5.Caption = i

If i Mod 2 = 0 Then
intTC = FreeFile()
filename = Date & "tc.txt"
Open filename For Output As intTC
Print #intTC, i / 2
Close intTC
End If

Else
Label4.Caption = "¬–≈Ãﬂ –¿¡Œ“€ «¿  ŒÃœ‹ﬁ“≈–ŒÃ «¿ ŒÕ◊»ÀŒ—‹!"
intFH = FreeFile()
filename = Date & ".txt"
Open filename For Append As intFH
Print #intFH, "1"
Close intFH

Form1.Visible = True
Form1.SetFocus
Timer2.Enabled = True
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
 Dim Rec As RECT
    'Get Left, Right, Top and Bottom of Form1
    GetWindowRect Form1.hwnd, Rec
    'Set Cursor position on X
    SetCursorPos Rec.Right - 150, Rec.Top + 150
End If

End Sub

'This project needs
'a Form, called 'Form1'
'a Picture Box, called 'ExplButton' (50x50 pixels)
'a Picture Box with an icon in it, called 'picIcon'
'two timers (Timer1 and Timer2), both with interval 100
'Button, called 'Command1'
'In general section

Sub DrawButton(Pushed As Boolean)
    Dim Clr1 As Long, Clr2 As Long
    If Pushed = True Then
        'If Pushed=True then clr1=Dark Gray
        Clr1 = &H808080
        'If Pushed=True then clr1=White
        Clr2 = &HFFFFFF
    ElseIf Pushed = False Then
        'If Pushed=True then clr1=White
        Clr1 = &HFFFFFF
        'If Pushed=True then clr1=Dark Gray
        Clr2 = &H808080
    End If

    With Form1.ExplButton
        ' Draw the button
        Form1.ExplButton.Line (0, 0)-(.ScaleWidth, 0), Clr1
        Form1.ExplButton.Line (0, 0)-(0, .ScaleHeight), Clr1
        Form1.ExplButton.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(.ScaleWidth - 1, 0), Clr2
        Form1.ExplButton.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(0, .ScaleHeight - 1), Clr2
    End With
End Sub
Private Sub Command3_Click()
    Dim Rec As RECT
    'Get Left, Right, Top and Bottom of Form1
    GetWindowRect Form1.hwnd, Rec
    'Set Cursor position on X
    SetCursorPos Rec.Right - 15, Rec.Top + 15
End Sub
Private Sub ExplButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton True
End Sub
Private Sub ExplButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton False
End Sub
Private Sub ExplButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton False
End Sub

Private Sub Timer2_Timer()
Form1.SetFocus

    Dim Rec As RECT, Point As POINTAPI
        GetWindowRect Me.hwnd, Rec
            GetCursorPos Point

 '   Shell "Cmd /x/c taskkill /f /im taskmgr.exe", vbHide


    If Point.X >= Rec.Left And Point.X <= Rec.Right And Point.Y >= Rec.Top And Point.Y <= Rec.Bottom Then
        'Me.Caption = "MouseCursor is on form."
    Else
        ' The cursor is not located above the form
       ' Me.Caption = "MouseCursor is not on form."
        
    'Get Left, Right, Top and Bottom of Form1
        GetWindowRect Form1.hwnd, Rec
    'Set Cursor position on X
        SetCursorPos Rec.Right - 150, Rec.Top + 150
    End If
End Sub

Private Sub Timer4_Timer()
    Dim Rec As RECT, Point As POINTAPI
    ' Get Left, Right, Top and Bottom of Form1
    GetWindowRect Me.hwnd, Rec
    ' Get the position of the cursor
    GetCursorPos Point
    

    ' If the cursor is located above the form then
    If Point.X >= Rec.Left And Point.X <= Rec.Right And Point.Y >= Rec.Top And Point.Y <= Rec.Bottom Then
        'Me.Caption = "MouseCursor is on form."
    Else
        ' The cursor is not located above the form
        'Me.Caption = "MouseCursor is not on form."
    End If
End Sub
Private Sub Timer3_Timer()
    Dim Rec As RECT, Point As POINTAPI
    ' Get Left, Right, Top and Bottom of ExplButton
    GetWindowRect ExplButton.hwnd, Rec
    ' Get the position of the cursor
    GetCursorPos Point
    ' If the cursor isn't located above ExplButton then
    If Point.X < Rec.Left Or Point.X > Rec.Right Or Point.Y < Rec.Top Or Point.Y > Rec.Bottom Then ExplButton.Cls
End Sub


