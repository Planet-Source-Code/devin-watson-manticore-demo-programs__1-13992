VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manticore Demo 4 - Hit ESC to Exit"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   407
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   120
      ScaleHeight     =   345
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   0
      Top             =   4920
      Width           =   7215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Manticore Demo 4
'This demo shows a multi-frame animated sprite on the
'screen, using a modified Ball object from Demo 3.
'NOTE: This demo does not use true double-buffering.
'      Higher speeds of the sprite may result in
'      heavy flickering.
Dim TheBall As Ball
'We'll also just use 1 clsSound object
Dim BounceSounds As clsSound
'These Win32 API declarations are used in the sub RunDemo()
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const TIMERRATE = 10
Private Const VK_ESCAPE = &H1B
Private Const FRAMERATE = 24   'This is for the animated sprite

Public Sub RunDemo()
    'This starts up and runs the demo.
    Dim CurTime As Long
    Dim CurTile As Long
    Dim FramesElapsed As Byte
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
    
    'Here, we create a new clsBlitter object for the
    'TileSet reference to use. This enables the
    'TileSet object to automatically set references to
    'the Tile objects it creates, and can also
    'be used one level up at the TileSet level.
    Set TheBall.TileSet.Blitter = New clsBlitter
    'We set the TileSet's Index property here, so
    'it knows to load the correct graphics in (see ANIM.LST;
    'the last number in each row denotes the index it should
    'load into. This was done in order to facilitate more
    'complex animation sequences, as well as multiple
    'TileSet objects sharing the same file.
    TheBall.TileSet.Index = 1
    'I've manually set the X and Y Speed variables here,
    'because I wanted it to move slowly enough that one
    'could see the animations with the sprite. You can
    'always speed the whole thing up, of course.
    TheBall.XSpeed = 8
    TheBall.YSpeed = 5
    
    'Calling the LoadTiles routine will
    'look through the data file to see what
    'graphics it needs to load up.
    TheBall.TileSet.LoadTiles App.Path & "\", "anim.lst"
    
    'Show the form
    Me.Show
    
    'Make sure our backbuffer is invisible.
    picCanvas.Visible = False
    CurTile = 0
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
            TheBall.TileSet.Blitter.Blt Me.hDC, TheBall.X, TheBall.Y, TheBall.TileSet(CurTile).Width, TheBall.TileSet(CurTile).Height, picCanvas.hDC, 0, 0, TheBall.TileSet.Blitter.SRCCOPY
            'Then update the X and Y position variables
            TheBall.UpdatePosition
            
            If FramesElapsed = FRAMERATE Then
                If CurTile < 3 Then
                    CurTile = CurTile + 1
                Else
                    CurTile = 1
                End If
                FramesElapsed = 1
            End If
            'Now we do a BltTile to move the ball on the window.
            TheBall.TileSet.Blitter.Blt Me.hDC, TheBall.X, TheBall.Y, TheBall.TileSet(CurTile).Width, TheBall.TileSet(CurTile).Height, TheBall.TileSet(CurTile).hDC, 0, 0, TheBall.TileSet.Blitter.SRCCOPY
            
            'Check to see if we've gone off the right side
            If TheBall.TileSet(CurTile).Width + TheBall.X > Me.ScaleWidth Then
                TheBall.XSpeed = -TheBall.XSpeed
                BounceSounds.PlaySound "HitRight", SND_ASYNC
            End If
            
            'Check to see if we've gone off the bottom
            If TheBall.TileSet(CurTile).Height + TheBall.Y > Me.ScaleHeight Then
                TheBall.YSpeed = -TheBall.YSpeed
                BounceSounds.PlaySound "HitBottom", SND_ASYNC
            End If
            
            'Check to see if we've gone off the left side
            If TheBall.X < 0 Then
                TheBall.XSpeed = -TheBall.XSpeed
                BounceSounds.PlaySound "HitLeft", SND_ASYNC
            End If
            
            'Check to see if we've gone past the top
            If TheBall.Y < 0 Then
                TheBall.YSpeed = -TheBall.YSpeed
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
    Set TheBall = New Ball
    Set BounceSounds = New clsSound
    RunDemo
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set TheBall = Nothing
    Set BounceSounds = Nothing
    End
End Sub


