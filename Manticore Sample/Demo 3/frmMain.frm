VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Manticore Demo 3 - Hit ESC to Exit"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   493
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   0
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   497
      TabIndex        =   0
      Top             =   5640
      Width           =   7455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Manticore Demo 3
'Here, we expand on Demo 2 (Bouncing Ball with Sound),
'and display many balls at once on the screen.
'NOTE: This demo does not use true double-buffering.
'      Higher speeds of the sprite may result in
'      heavy flickering.
'For this demo, we will be using an array of Ball objects,
'created for this project.
Dim BallArray(1 To 3) As Ball
'We'll also just use 1 clsSound object
Dim BounceSounds As clsSound
'These Win32 API declarations are used in the sub RunDemo()
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const TIMERRATE = 10
Private Const VK_ESCAPE = &H1B

Public Sub CheckBalls()
    Dim i As Integer
    
    For i = 1 To 3
        'Check to see if we've gone off the right side
        If BallArray(i).Tile.Width + BallArray(i).X > Me.ScaleWidth Then
            BallArray(i).XSpeed = -BallArray(i).XSpeed
            BounceSounds.PlaySound "HitRight", SND_ASYNC
        End If
            
        'Check to see if we've gone off the bottom
        If BallArray(i).Tile.Height + BallArray(i).Y > Me.ScaleHeight Then
            BallArray(i).YSpeed = -BallArray(i).YSpeed
            BounceSounds.PlaySound "HitBottom", SND_ASYNC
        End If
            
        'Check to see if we've gone off the left side
        If BallArray(i).X < 0 Then
            BallArray(i).XSpeed = -BallArray(i).XSpeed
'            BounceSounds.PlaySound "HitLeft", SND_ASYNC
        End If
            
        'Check to see if we've gone past the top
        If BallArray(i).Y < 0 Then
            BallArray(i).YSpeed = -BallArray(i).YSpeed
            BounceSounds.PlaySound "HitTop", SND_ASYNC
        End If
        
    Next i



End Sub


Public Sub DrawBalls()
    Dim i As Integer
    
    For i = 1 To 3
        BallArray(i).Tile.BltTile Me.hDC, BallArray(i).X, BallArray(i).Y
    Next i

End Sub


Public Sub DrawOver()
    'Draws over the current positions of the Ball objects
    'on the screen.
    
    Dim i As Integer
    
    For i = 1 To 3
        BallArray(i).Tile.Blitter.Blt Me.hDC, BallArray(i).X, BallArray(i).Y, BallArray(i).Tile.Width, _
                BallArray(i).Tile.Height, picCanvas.hDC, 0, 0, BallArray(i).Tile.Blitter.SRCCOPY
    Next i
    
End Sub


Public Sub RunDemo()
    'This starts up and runs the demo.
    Dim CurTime As Long
    Dim PrevTime As Long
    Dim keyPressed As Long
    
    'Here we use the clsSound object we created to
    'load in all of our "bounce" sounds. I wanted a
    'different sound for each of the 4 sides of the window
    'when the ball hit them.
    BounceSounds.LoadSound App.Path & "\left_wall.wav", "HitLeft"
    BounceSounds.LoadSound App.Path & "\right_wall.wav", "HitRight"
    BounceSounds.LoadSound App.Path & "\top.wav", "HitTop"
    BounceSounds.LoadSound App.Path & "\bottom.wav", "HitBottom"
    
    'Show the form
    Me.Show
    'Make sure our backbuffer is invisible.
    picCanvas.Visible = False
    
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
            'In here, we update the balls' position

            DrawOver
            
            'Then update the X and Y position variables
            UpdateBalls
            
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
    'Updates the position of all of the Ball objects
    Dim i As Integer
    
    For i = 1 To 3
        BallArray(i).UpdatePosition
    Next i
    
End Sub

Private Sub Form_Load()

    Dim i As Integer

    'We initialize the Ball array here
    For i = 1 To 3
        Set BallArray(i) = New Ball
        BallArray(i).Tile.LoadTile App.Path & "\Block002.bmp", 32, 32
        BallArray(i).X = i * BallArray(i).Tile.Width
        BallArray(i).XSpeed = BallArray(i).XSpeed + i
        BallArray(i).YSpeed = BallArray(i).YSpeed + i
    Next i
    Set BounceSounds = New clsSound
    RunDemo
End Sub


