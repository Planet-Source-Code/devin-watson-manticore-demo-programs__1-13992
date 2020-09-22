VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarTile As Tile

Private mlngX As Long
Private mlngY As Long
Private mintXSpeed As Integer
Private mintYSpeed As Integer

Public Property Get Tile() As Tile
    Set Tile = mvarTile
End Property

Public Sub UpdatePosition()
    mlngX = mlngX + mintXSpeed
    mlngY = mlngY + mintYSpeed
End Sub

Public Property Get X() As Long
    X = mlngX
End Property


Public Property Let X(varData As Long)
    mlngX = varData
End Property


Public Property Get XSpeed() As Integer
    XSpeed = mintXSpeed
End Property

Public Property Let XSpeed(varData As Integer)
    mintXSpeed = varData
End Property

Public Property Get Y() As Long
    Y = mlngY
End Property

Public Property Let Y(varData As Long)
    mlngY = varData
End Property

Public Property Get YSpeed() As Integer
    YSpeed = mintYSpeed
End Property

Public Property Let YSpeed(varData As Integer)
    mintYSpeed = varData
End Property

Private Sub Class_Initialize()
    Set mvarTile = New Tile
    mlngY = 0
    mlngX = 0
    mintXSpeed = 8
    mintYSpeed = 6
End Sub


Private Sub Class_Terminate()
    Set mvarTile = Nothing
End Sub


