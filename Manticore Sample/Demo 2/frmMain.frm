VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manticore Demo 2 - Hit ESC to Exit"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   509
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   7695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Manticore Demo 2
'Here, we expand on Demo 1 (Bouncing Ball), utilizing a
'clsSound object to manage sounds of bouncing.
'NOTE: This demo does not use true double-buffering.
'      Higher speeds of the sprite may result in
'      heavy flickering.
'For this demo, we'll just use 1 Tile object
Dim ATile As Tile
'We'll also just use 1 clsSound object
Dim BounceSounds As clsSound
'These Win32 API declarations are used in the sub RunDemo()
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const TIMERRATE = 10
Private Const VK_ESCAPE = &H1B

Public Sub RunDemo()
    'This starts up and runs the demo.
    Dim CurTime As Long
    Dim PrevTime As Long
    Dim keyPressed As Long
    Dim rc As Integer
    Dim TileX As Long
    Dim TileY As Long
    Dim YSpeed As Integer
    Dim XSpeed As Integer
    
    'Set the initial position of the ball to the upper-left
    'of the window.
    TileX = 0
    TileY = 0
    
    'Here we use the clsSound object we created to
    'load in all of our "bounce" sounds. I wanted a
    'different sound for each of the 4 sides of the window
    'when the ball hit them.
    BounceSounds.LoadSound App.Path & "\left_wall.wav", "HitLeft"
    BounceSounds.LoadSound App.Path & "\right_wall.wav", "HitRight"
    BounceSounds.LoadSound App.Path & "\top.wav", "HitTop"
    BounceSounds.LoadSound App.Path & "\bottom.wav", "HitBottom"
    
    'Set the initial speeds here. I wanted it to move kinda
    'quickly, so I just set XSpeed greater than YSpeed.
    'Change these values around to see what happens.
    XSpeed = 10
    YSpeed = 6
    'Show the form
    Me.Show
    'Make sure our backbuffer is invisible.
    picCanvas.Visible = False
    'Load the tile in using the LoadTile method. Chenge the last two
    'values to make it smaller or bigger. 32x32 is the size
    'I want to use, but the bitmap is actually 64x64. It's a good
    'idea to make sure your graphics are proportional and that
    'they are sized in multiples of 2.
    rc = ATile.LoadTile(App.Path & "\Block002.bmp", 32, 32)
    'Check the return code here. If it isn't 1, then
    'something went wrong on the load. You don't have to do this,
    'but it is generally a good idea to do it when you
    'haven't fully tested it yet.
    If rc <> 1 Then
        MsgBox "ERROR DURING LOAD. CHECK BITMAP.", vbOKOnly + vbCritical, "ERROR"
        Unload Me
        Exit Sub
    End If
    
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

            'First, draw over the old place the ball was
            ATile.Blitter.Blt Me.hDC, TileX, TileY, ATile.Width, ATile.Height, picCanvas.hDC, 0, 0, ATile.Blitter.SRCCOPY
            'Then update the X and Y position variables
            TileX = TileX + XSpeed
            TileY = TileY + YSpeed
            
            'Now we do a BltTile to move the ball on the window.
            ATile.BltTile Me.hDC, TileX, TileY
            
            'Check to see if we've gone off the right side
            If ATile.Width + TileX > Me.ScaleWidth Then
                XSpeed = -XSpeed
                BounceSounds.PlaySound "HitRight", SND_ASYNC
            End If
            
            'Check to see if we've gone off the bottom
            If ATile.Height + TileY > Me.ScaleHeight Then
                YSpeed = -YSpeed
                BounceSounds.PlaySound "HitBottom", SND_ASYNC
            End If
            
            'Check to see if we've gone off the left side
            If TileX < 0 Then
                XSpeed = -XSpeed
                BounceSounds.PlaySound "HitLeft", SND_ASYNC
            End If
            
            'Check to see if we've gone past the top
            If TileY < 0 Then
                YSpeed = -YSpeed
                BounceSounds.PlaySound "HitTop", SND_ASYNC
            End If
            
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


Private Sub Form_Load()
    'Initialize everything here
    Set ATile = New Tile
    Set BounceSounds = New clsSound
    'Run the demo
    RunDemo
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set ATile = Nothing
    Set BounceSounds = Nothing
End Sub


