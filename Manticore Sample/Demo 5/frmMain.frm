VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manticore Demo 5 - Hit ESC to Exit"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   474
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   0
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   0
      Top             =   5880
      Width           =   7095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Manticore Demo 5
'This demo shows many multi-frame animated sprites on the
'screen, using a modified Ball object from Demo 4.
'NOTE: This demo does not use true double-buffering.
'      Higher speeds of the sprite may result in
'      heavy flickering.
Dim BallArray(1 To 3) As Ball

'These Win32 API declarations are used in the sub RunDemo()
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const TIMERRATE = 5
Private Const VK_ESCAPE = &H1B
Private Const FRAMERATE = 24   'This is for the animated sprites

Public Sub CheckBalls()
    Dim i As Integer
    
    For i = 1 To 3
        'Check to see if we've gone off the right side
        If BallArray(i).TileSet(BallArray(i).CurrentFrame).Width + BallArray(i).X > Me.ScaleWidth Then
            BallArray(i).XSpeed = -BallArray(i).XSpeed
        End If
            
        'Check to see if we've gone off the bottom
        If BallArray(i).TileSet(BallArray(i).CurrentFrame).Height + BallArray(i).Y > Me.ScaleHeight Then
            BallArray(i).YSpeed = -BallArray(i).YSpeed
        End If
            
        'Check to see if we've gone off the left side
        If BallArray(i).X < 0 Then
            BallArray(i).XSpeed = -BallArray(i).XSpeed
        End If
            
        'Check to see if we've gone past the top
        If BallArray(i).Y < 0 Then
            BallArray(i).YSpeed = -BallArray(i).YSpeed
        End If
    Next i
    
End Sub


Public Sub DrawBalls()
    Dim i As Integer
    
    For i = 1 To 3
        BallArray(i).TileSet.Blitter.Blt Me.hDC, BallArray(i).X, BallArray(i).Y, BallArray(i).TileSet(BallArray(i).CurrentFrame).Width, _
                BallArray(i).TileSet(BallArray(i).CurrentFrame).Height, BallArray(i).TileSet(BallArray(i).CurrentFrame).hDC, 0, 0, BallArray(i).TileSet.Blitter.SRCCOPY
    Next i
    
End Sub


Public Sub DrawOver()
    'TheBall.TileSet.Blitter.Blt Me.hDC, TheBall.X, TheBall.Y, TheBall.TileSet(CurTile).Width, TheBall.TileSet(CurTile).Height, picCanvas.hDC, 0, 0, TheBall.TileSet.Blitter.SRCCOPY
    Dim i As Integer
    
    For i = 1 To 3
        BallArray(i).TileSet.Blitter.Blt Me.hDC, BallArray(i).X, BallArray(i).Y, BallArray(i).TileSet(BallArray(i).CurrentFrame).Width, BallArray(i).TileSet(BallArray(i).CurrentFrame).Height, _
            picCanvas.hDC, 0, 0, BallArray(i).TileSet.Blitter.SRCCOPY
    Next i
    
End Sub


Public Sub RunDemo()
    'This starts up and runs the demo.
    Dim CurTime As Long

    Dim FramesElapsed As Byte
    Dim PrevTime As Long
    Dim keyPressed As Long
    
    'Show the form
    Me.Show
    
    'Make sure our backbuffer is invisible.
    picCanvas.Visible = False

    'Counter for the number of frames elapsed. This
    'allows control over how "fast" the frames of the animated
    'sprite are displayed, so you can actually see the changes
    'occurring on-screen, rather than just a blur.
    FramesElapsed = 0
    Do
        'To exit this program, one must hit the
        'escape key.
        keyPressed = GetAsyncKeyState(VK_ESCAPE)
        
        'If keyPressed came back True, then
        'the user did hit escape, and we
        'must exit this loop, which, in
        'turn, exits the program.
        If keyPressed Then
            Exit Do
        End If
        
        'Get the current time count
        CurTime = GetTickCount
        
        'If the time difference in the loop
        'is greater than the synchronization rate (10ms),
        'then we can redraw everything.
        If CurTime - PrevTime > TIMERRATE Then
            'In here, we update the ball's position
            FramesElapsed = FramesElapsed + 1
            'First, draw over the old place the ball was
            DrawOver
            
            'Then update the X and Y position variables
            UpdateBalls
            
            If FramesElapsed = FRAMERATE Then
                UpdateFrames
                FramesElapsed = 0
            End If
            
            'Now we do a BltTile to move the ball on the window.
            DrawBalls
            
            CheckBalls
            
            'Now we force a refresh on the window,
            'and let Windows do its thing for the moment.
            Me.Refresh
            DoEvents
            
        Else
            'Otherwise, we let the system refresh itself.
            DoEvents
            Sleep 2
        End If
    Loop
    
    'If execution made it down here, then
    'the user pressed the ESC key.
    Unload Me


End Sub


Public Sub UpdateBalls()
    Dim i As Integer
    
    For i = 1 To 3
        BallArray(i).UpdatePosition
    Next i
End Sub

Public Sub UpdateFrames()
    Dim i As Integer
    
    For i = 1 To 3
        BallArray(i).NextTile
    Next i
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'This initializes all of the Ball objects, and loads
    'the frames into memory for usage.
    
    For i = 1 To 3
        Set BallArray(i) = New Ball
        Set BallArray(i).TileSet.Blitter = New clsBlitter
        'We set the individual indices here so
        'that the TileSet object knows what to
        'load in the ANIM.LST file.
        BallArray(i).TileSet.Index = i
        BallArray(i).TileSet.LoadTiles App.Path & "\", "anim.lst"
        BallArray(i).X = i * 64
        BallArray(i).Y = i * 64
        BallArray(i).YSpeed = 5
        BallArray(i).XSpeed = 10
    Next i
    RunDemo
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    For i = 1 To 3
        Set BallArray(i) = Nothing
    Next i
End Sub


