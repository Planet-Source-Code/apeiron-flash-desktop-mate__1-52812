VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form DMate 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "DesktopMate"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   Icon            =   "DesktopMate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   127
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   122
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   1320
      Top             =   1440
   End
   Begin VB.Label lblDrag 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      _cx             =   3625
      _cy             =   2778
      FlashVars       =   ""
      Movie           =   "C:\Documents and Settings\Default\Desktop\DesktopFlashMate\snowman.swf"
      Src             =   "C:\Documents and Settings\Default\Desktop\DesktopFlashMate\snowman.swf"
      WMode           =   "Transparent"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   -1  'True
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
End
Attribute VB_Name = "DMate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Have a flash animation that wanders around your desktop.
' Shows some interesting things like communication between flash and VB,
'    and interacting with other top level windows.
' Easily customized, read the readme
' One note, I use some of the varibles in a non-standard way
'   that might not be obvious at first.  The integers like moveright
'   are used as a boolean (everything non-zero is treated as true by VB) but
'   it also holds the amount still needed to travel in that direction.
'  Have fun with it, and if you come up with a cool aniamtion or
'  expand it email it to me or post it.
'  If you find it useful or fun vote and leave a comment.

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Const MF_CHECKED = &H8&
Const MF_APPEND = &H100&
Const TPM_LEFTALIGN = &H0&
Const MF_DISABLED = &H2&
Const MF_GRAYED = &H1&
Const MF_SEPARATOR = &H800&
Const MF_STRING = &H0&
Const TPM_RETURNCMD = &H100&
Const TPM_RIGHTBUTTON = &H2&
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hwnd As Long, ByVal lptpm As Any) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Dim RMenu As Long

Private Const ULW_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_OPAQUE = &H4
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_WINDOWEDGE = &H100&
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Dim tRect As RECT
Dim OnTop As Boolean
Dim Trans As Boolean
Dim OldStyle As Long
Dim Falling As Boolean
Dim WasFalling As Boolean
Dim MoveLeft As Integer
Dim WasMoveLeft As Boolean
Dim MoveRight As Integer
Dim WasMoveRight As Boolean
Dim Waiting As Boolean
Dim WalkRect As RECT
Dim Climbing As Boolean
Dim ClimbLeft As Boolean
Dim ClimbRight As Boolean
Dim ClimbHWND As Long
Dim WasClimbing As Boolean

' How fast to fall
Const FALL_RATE = 150
' Percentage of the time to climb a window you bump into.
Const CLIMB_CHANCE = 25
' How fast to climb or walk
Const CLIMB_RATE = 80
Const WALK_RATE = 50
' Change these to your corresponding frames in your animation.
Const LEFT_FRAME = 1
Const RIGHT_FRAME = 1
Const COOLTHING1_FRAME = 64
Const COOLTHING2_FRAME = 50
Const FALLING_FRAME = 22
Const LANDED_FRAME = 12
Const CLIMBING_RIGHT_FRAME = 38
Const CLIMBING_LEFT_FRAME = 51
Const BEING_CARRIED = 1

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape
      Unload Me
      End
    Case vbKeyT
      OnTop = True
      FormOnTop Me.hwnd, True
    Case vbKeyU
      OnTop = False
      FormOnTop Me.hwnd, False
    Case vbKeyM
      Transparent (True)
    Case vbKeyN
      Transparent (False)
  End Select
End Sub

Private Sub Form_Load()
  Randomize (Now)
  Dim tWnd As Long
  tWnd = FindWindow("Shell_traywnd", vbNullString)
  GetWindowRect tWnd, tRect
  ' Get old window style so can be reset with the transparency off
  OldStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
  ShockwaveFlash1.Width = Me.ScaleWidth
  ShockwaveFlash1.Height = Me.ScaleHeight
  lblDrag.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
  ShockwaveFlash1.Movie = App.Path + "\snowman.swf"
  
  ' Win98, 95, me and NT, comment out the next two lines or it won't run
  ' the setWindowLayersAttribute call for transparency is 2000 XP only
  Trans = False
  Transparent (True)
  
  OnTop = True
  FormOnTop Me.hwnd, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub lblDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    WasMoveRight = False
    WasMoveLeft = False
    WasFalling = False
    MoveLeft = 0
    MoveRight = 0
    Falling = False
    Waiting = False
    WasClimbing = False
    Climbing = False
    Timer1.Enabled = True
    PlayFlashFrom (BEING_CARRIED)
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, ByVal 0&
  End If
End Sub

Private Sub lblDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Create the right click menu, thanks to allapi.net for the sample code
  If Button = 2 Then
    Dim Pt As POINTAPI
    Dim ret As Long
    RMenu = CreatePopupMenu()
    If OnTop Then
      AppendMenu RMenu, MF_STRING, 1, "Not On Top"
    Else
      AppendMenu RMenu, MF_STRING, 1, "Put On Top"
    End If
    AppendMenu RMenu, MF_SEPARATOR, 3, ByVal 0&
    If Trans Then
      AppendMenu RMenu, MF_STRING, 4, "Make Opaque"
    Else
      AppendMenu RMenu, MF_STRING, 4, "Make Transparent"
    End If
    AppendMenu RMenu, MF_STRING, 5, "Exit"
    GetCursorPos Pt
    Timer1.Enabled = False
    ret = TrackPopupMenuEx(RMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.X, Pt.Y, Me.hwnd, ByVal 0&)
    Select Case ret
      Case 1
        If OnTop Then
          FormOnTop Me.hwnd, False
        Else
          FormOnTop Me.hwnd, True
        End If
        OnTop = Not OnTop
      Case 2
      Case 4
        If Trans Then
          Transparent (False)
        Else
          Transparent (True)
        End If
      Case 5
        Unload Me
        End
    End Select
    DestroyMenu RMenu
    Timer1.Enabled = True
  End If
End Sub

Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)
  ' Resets animation when it was set to wait using the setWait function
  ' Communicates from flash to vb through actionscript (read the ReadMe.txt
  ' for more info)
  
  If command = "Done" And Waiting Then
    command = ""
    Waiting = False
    WasMoveRight = False
    WasMoveLeft = False
    WasFalling = False
    MoveLeft = 0
    MoveRight = 0
    Falling = False
    Waiting = False
    WasClimbing = False
    Climbing = False
    Timer1.Enabled = True
  End If
End Sub

Private Sub Timer1_Timer()
  ' "brains" of the animation, checks what the animation should be doing
  ' now.  If nothing else is already being done then it decides WhatToDo().
  ' The order is important here, example --climbing has a higher precedence than
  ' checking if the animation should be falling. Otherwise every time it starts to
  ' climb it will fall.
  
  Select Case True
    Case Climbing
      Climb
    Case CheckFalling2()
    Case MoveLeft
      Call MoveLeftNow
    Case MoveRight
      Call MoveRightNow
    Case Else ' Not doing anything else what should I do
      Call WhatToDo
  End Select
End Sub

Private Sub MoveLeftNow()
  If MoveLeft < 0 Then
    MoveLeft = 0
    WasMoveLeft = False
    Exit Sub
  End If
    ' Inititate correct frame only the first time
  If Not WasMoveLeft Then
    PlayFlashFrom LEFT_FRAME
  End If
  'WasFalling = False
  'Waiting = False
  If Me.Left < 0 Then
    Me.Left = 0
    MoveLeft = 0
    WasMoveLeft = False
    Exit Sub
  End If
  If Me.Left > Screen.Width - Me.Width Then
    Me.Left = Screen.Width - Me.Width
    WasMoveRight = False
    MoveRight = 0
  End If
  'ShockwaveFlash1.Play
  Dim Pt As POINTAPI, mWnd As Long, WR As RECT, nDC As Long
  Dim ParentWindow As Long
  Dim MeRect As RECT
  Dim winRECT As RECT
  GetWindowRect Me.hwnd, MeRect
    ' Set the point to check right of form now.
  Pt.X = MeRect.Left - 1
  Pt.Y = MeRect.Bottom - 10
  If MeRect.Left <= 0 Then
    MoveLeft = 0
    WasMoveLeft = False
    Exit Sub
  End If

    mWnd = WindowFromPoint(Pt.X, Pt.Y)
    'Get the window's position
    ParentWindow = GetParent(mWnd)
    If ParentWindow = 0 Then
      GetWindowRect mWnd, WR
      Dim MyStr As String
      MyStr = String(100, Chr$(0))
      GetWindowText mWnd, MyStr, 100
      MyStr = Left$(MyStr, InStr(MyStr, Chr$(0)) - 1)
      If MyStr = "DesktopMate" Then
        MsgBox "Hi buddy!"
        WasMoveLeft = False
        MoveLeft = 0
        Exit Sub
        ' Found another one like me what should I do stuff here
      Else
        ' Another non desktopmate top level window should I climb it?
        If MeRect.Left - 10 <= 0 Then
          MoveLeft = 0
          Me.Left = 0
          WasMoveLeft = False
          Exit Sub
        End If
        Dim ClimbIt As Integer
        ClimbIt = Rnd * 100
        ' Only climb the window if you walk into it fromn the left
        If ClimbIt < CLIMB_CHANCE And MyStr <> "FolderView" And MeRect.Right < Screen.Width * Screen.TwipsPerPixelX - 30 _
            And MeRect.Right > WR.Right Then
          ' If within the percentage of chance to climb then climb it
          ClimbHWND = mWnd
          'If Pt.X >= WR.Right - 30 And Pt.X < WR.Right + 30 Then
          'If MeRect.Right > WR.Right Then
            ClimbRight = True
          'End If
          'Else
          '  ClimbLeft = True
          'End If
          WasMoveLeft = False
          Climbing = True
          Exit Sub
        End If
      End If
      'Cls
      'Print MyStr
    End If
    
    If Pt.X > WalkRect.Left And Pt.X > 0 Then 'ParentWindow = 0 And a <= 5 And (Pt.y > WR.Left Or Pt.y < WR.Right) Then
      Me.Left = Me.Left - WALK_RATE
      MoveLeft = MoveLeft - WALK_RATE
      WasMoveLeft = True
    Else
      ' I'm at the edge of the window should I jump or fall
      Me.Left = Me.Left - WALK_RATE
      'MoveLeft = 0
      WasMoveLeft = False
    End If
End Sub

Private Sub MoveRightNow()
  If MoveRight < 0 Then
    MoveRight = 0
    WasMoveRight = False
    Exit Sub
  End If
  WasFalling = False
  Waiting = False
  If Me.Left < 0 Then
    Me.Left = 0
    MoveRight = 0
    WasMoveRight = False
    Exit Sub
  End If
  If Not WasMoveRight Then
    PlayFlashFrom RIGHT_FRAME
  End If
  If Me.Left > Screen.Width - Me.Width Then
    Me.Left = Screen.Width - Me.Width
    WasMoveRight = False
    MoveRight = 0
  End If
  Dim Pt As POINTAPI, mWnd As Long, WR As RECT, nDC As Long
  Dim ParentWindow As Long
  ShockwaveFlash1.Play
  Dim MeRect As RECT
  Dim winRECT As RECT
  GetWindowRect Me.hwnd, MeRect
  ' Set the point to check right of form now.
  Pt.X = MeRect.Right + 1
  Pt.Y = MeRect.Bottom - 10
  ' Inititate correct frame only the first time
  mWnd = WindowFromPoint(Pt.X, Pt.Y)
  'Get the window's position
  ParentWindow = GetParent(mWnd)
  If ParentWindow = 0 Then
      GetWindowRect mWnd, WR
      Dim MyStr As String
      MyStr = String(100, Chr$(0))
      GetWindowText mWnd, MyStr, 100
      MyStr = Left$(MyStr, InStr(MyStr, Chr$(0)) - 1)
      If MyStr = "DesktopMate" Then
        ' Found another one like me what should I do stuff here
        MsgBox "Hi buddy!"
        WasMoveRight = False
        MoveLeft = 0
        Exit Sub
      Else
        ' Another non desktopmate top level window should I climb it?
        If MeRect.Right + 10 >= Screen.Width \ Screen.TwipsPerPixelX Then
          MoveRight = 0
          Me.Left = (Screen.Width - Me.Width)
          WasMoveRight = False
          Exit Sub
        End If
        
        Dim ClimbIt As Integer
        ClimbIt = Rnd * 100
        ' Only climb if walking into the window from the right
        If ClimbIt < CLIMB_CHANCE And MyStr <> "FolderView" _
            And MeRect.Left < WR.Left Then
          ' If within the percentage of chance to climb then climb it
          ClimbHWND = mWnd
          'If MeRect.Left < WR.Left Then
          'If MeRect.Right >= WR.Left - 130 And MeRect.Right < WR.Left + 130 Then
            'ClimbRight = True
          'Else
            ClimbLeft = True
          'End If
          WasMoveRight = False
          Climbing = True
          Exit Sub
        End If
      End If
'      Cls
'      Print MyStr
    End If

    If Pt.X < WalkRect.Right And Pt.X < Screen.Width \ Screen.TwipsPerPixelX Then 'ParentWindow = 0 And a <= 5 And (Pt.y > WR.Left Or Pt.y < WR.Right) Then
      Me.Left = Me.Left + WALK_RATE
      MoveRight = MoveRight - WALK_RATE
      WasMoveRight = True
    Else
      ' I'm at the edge of the window should I jump or fall
      'MoveRight = 0
      WasMoveRight = False
      Me.Left = Me.Left + WALK_RATE 'Me.Width \ 2 + 1
    End If
End Sub

Private Function CheckFalling2() As Boolean
  Dim MeRect As RECT
  Dim ParentWindow As Long
  Dim WinBelow As Long
  Dim winRECT As RECT
  GetWindowRect Me.hwnd, MeRect
  WinBelow = WindowFromPoint(MeRect.Left + (Me.ScaleWidth \ 2), MeRect.Bottom + 1)
  ParentWindow = GetParent(WinBelow)

  ' Checks if below the bottom of the screen
  If MeRect.Bottom > tRect.Top Then
    If WasFalling Then
      Waiting = False
      WasMoveRight = False
      WasMoveLeft = False
      MoveLeft = 0
      MoveRight = 0
      Waiting = False
      WasClimbing = False
      Climbing = False
      Timer1.Enabled = True
      PlayFlashFrom (LANDED_FRAME)
      SetWait
      WasFalling = False
    End If
    Me.Top = (tRect.Top - Me.ScaleHeight) * Screen.TwipsPerPixelY
    CheckFalling2 = False
    Exit Function
  End If
  
  If ParentWindow = 0 Then
    GetWindowRect WinBelow, WalkRect
    If MeRect.Bottom > WalkRect.Top Then
      If MeRect.Bottom > WalkRect.Top + 30 Then
        Me.Top = Me.Top + FALL_RATE
        WasFalling = True
        CheckFalling2 = True
      Else
        Me.Top = (WalkRect.Top - Me.ScaleHeight) * Screen.TwipsPerPixelY
        If WasFalling Then
      Waiting = False
      WasMoveRight = False
      WasMoveLeft = False
      MoveLeft = 0
      MoveRight = 0
      Waiting = False
      WasClimbing = False
      Climbing = False
      Timer1.Enabled = True
      PlayFlashFrom (LANDED_FRAME)
      SetWait
      WasFalling = False
        End If
        CheckFalling2 = False
      End If
    End If
  Else
    If Not WasFalling Then
      PlayFlashFrom (FALLING_FRAME)
    End If
      WasMoveLeft = False
      WasMoveRight = False
      Me.Top = Me.Top + FALL_RATE
      WasFalling = True
      CheckFalling2 = True
  End If
End Function
Private Sub WhatToDo()
  Dim What As Integer
  What = Round(Rnd * 100)
  Select Case What
    Case Is < 45
      MoveLeft = Round(Rnd * Screen.Width)
    Case Is < 85
      MoveRight = Round(Rnd * Screen.Width)
    Case Is < 97
      PlayFlashFrom (64)
      SetWait
    Case Else
      Me.Top = 0
  End Select
End Sub

Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)
' Example: Call FormOnTop(me.hWnd, True)
  Dim wFlags As Long
  Dim placement As Long
  Const SWP_NOSIZE = &H1
  Const SWP_NOMOVE = &H2
  Const SWP_NOACTIVATE = &H10
  Const SWP_SHOWWINDOW = &H40
  Const HWND_TOPMOST = -1
  Const HWND_NOTOPMOST = -2
  wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
  Select Case bTopMost
  Case True
    placement = HWND_TOPMOST
  Case False
    placement = HWND_NOTOPMOST
  End Select
  SetWindowPos hWindow, placement, 0, 0, 0, 0, wFlags
End Sub

Private Function PlayFlashFrom(pintFrameNo As Integer)
  ShockwaveFlash1.GotoFrame (pintFrameNo)
  ShockwaveFlash1.StopPlay
  ShockwaveFlash1.Play
End Function

Private Sub Climb()
  Dim ClimbRect As RECT
  Dim MeRect As RECT
  GetWindowRect Me.hwnd, MeRect
  GetWindowRect ClimbHWND, ClimbRect
  If MeRect.Left <= 0 Or MeRect.Right > Screen.Width / Screen.TwipsPerPixelX Then
    ClimbLeft = False
    ClimbRight = False
    Climbing = False
    MoveLeft = 0
    MoveRight = 0
    WasMoveLeft = False
    WasMoveRight = False
    Exit Sub
  End If

  If ClimbRight Then
    ' Initiate the correct frame only once at start of climb
    If Not WasClimbing Then
      PlayFlashFrom CLIMBING_RIGHT_FRAME
    End If
    WasClimbing = True
    Me.Left = ClimbRect.Right * Screen.TwipsPerPixelX
  Else
    ' Initiate the correct frame only once at start of climb
    If Not WasClimbing Then
      PlayFlashFrom CLIMBING_LEFT_FRAME
    End If
    Me.Left = (ClimbRect.Left - Me.ScaleWidth) * Screen.TwipsPerPixelX
    WasClimbing = True
  End If
  If MeRect.Bottom <= ClimbRect.Top Then
    ' I'm at the top
    ''Me.Left = Me.Left - (Me.Width \ 2) - 25
    ''Me.Top = (ClimbRect.Top - Me.ScaleHeight) * Screen.TwipsPerPixelY
    ' Walk on the window
    If ClimbRight Then
      Me.Left = Me.Left - (Me.Width \ 2) - 25
      Me.Top = (ClimbRect.Top - Me.ScaleHeight) * Screen.TwipsPerPixelY
      MoveLeft = 1000
    Else
      Me.Left = Me.Left + (Me.Width \ 2) + 25
      Me.Top = (ClimbRect.Top - Me.ScaleHeight) * Screen.TwipsPerPixelY
      MoveRight = 1000
    End If
    Climbing = False
    WasClimbing = False
  Else
    ' If I'm at the top of the screen
    If MeRect.Top <= 1 Then
      WasClimbing = False
      Climbing = False
    Else
      Me.Top = Me.Top - CLIMB_RATE
    End If
  End If
End Sub

Private Sub Transparent(t As Boolean)
  Trans = Not Trans
  If t Then
      Me.BackColor = &HFFCCCC
      SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
      SetLayeredWindowAttributes Me.hwnd, &HFFCCCC, 0, ULW_COLORKEY
  Else
      Me.BackColor = &H8000000F
      SetWindowLong Me.hwnd, GWL_EXSTYLE, OldStyle
  End If
End Sub

Private Sub SetWait()
  Waiting = True
  Timer1.Enabled = False
End Sub
