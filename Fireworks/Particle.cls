VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Particle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'This class defines a particle and provides functions to manipulate the particle
Option Explicit
Public StartPosX As Integer, StartPosY As Integer, Color As Long 'starting position and color of particle
Public Size As Boolean 'size of particle either 0=single pixel or 1=2x2 pixels
Private dx(1) As Single, dy(1) As Single ' to create a 2 dimentional array of 2x2 pixels
Private SpeedX As Single, SpeedY As Single


Public Sub SetValues()
SpeedX = 8 * Rnd
SpeedY = 8 * Rnd
dx(0) = StartPosX + 1 - 2 * Rnd
dy(0) = StartPosY + 1 - 2 * Rnd


'The following adjustments of speeds is to make the explosion round
'without this adjustment the particles will expand in a square shape
If SpeedX > 2 And SpeedY > 2 Then
    
    Select Case Abs(SpeedX - SpeedY) 'difference between speeds

        Case Is < 1
        SpeedX = SpeedX - 2.1
        SpeedY = SpeedY - 2.1


        Case 1 To 2
        SpeedX = SpeedX - 1.55
        SpeedY = SpeedY - 1.55


        Case 2 To 3
        SpeedX = SpeedX - 1.1
        SpeedY = SpeedY - 1.1

        Case 3 To 4
        SpeedX = SpeedX - 0.85
        SpeedY = SpeedY - 0.85

        Case 4 To 5
        SpeedX = SpeedX - 0.4
        SpeedY = SpeedY - 0.4

        Case 5 To 6
        SpeedX = SpeedX - 0.3
        SpeedY = SpeedY - 0.3

        Case 6 To 7
        SpeedX = SpeedX - 0.15
        SpeedY = SpeedY - 0.15

        Case Is > 7
        SpeedX = SpeedX - 0.075
        SpeedY = SpeedY - 0.075


    End Select
    
End If

End Sub


Public Sub Move()

If Size = 0 Then

    SetPixel Form1.hdc, dx(0), dy(0), vbBlack 'set previous pixel to black before moving to next position
    If dx(0) > StartPosX Then dx(0) = dx(0) + SpeedX Else dx(0) = dx(0) - SpeedX
    If dy(0) > StartPosY Then dy(0) = dy(0) + SpeedY Else dy(0) = dy(0) - SpeedY
    SetPixel Form1.hdc, dx(0), dy(0), Color 'set pixel color at new position

Else

    'set previous pixels to black before moving to next position
    SetPixel Form1.hdc, dx(0), dy(0), vbBlack
    SetPixel Form1.hdc, dx(0), dy(1), vbBlack
    SetPixel Form1.hdc, dx(1), dy(0), vbBlack
    SetPixel Form1.hdc, dx(1), dy(1), vbBlack

    If dx(0) > StartPosX Then
        dx(0) = dx(0) + SpeedX
        dx(1) = dx(0) + 1
    Else
        dx(0) = dx(0) - SpeedX
        dx(1) = dx(0) + 1
    End If

    
    If dy(0) > StartPosY Then
        dy(0) = dy(0) + SpeedY
        dy(1) = dy(0) + 1
    Else
        dy(0) = dy(0) - SpeedY
        dy(1) = dy(0) + 1
    End If

    'set pixel color at new position
    SetPixel Form1.hdc, dx(0), dy(0), Color
    SetPixel Form1.hdc, dx(0), dy(1), Color
    SetPixel Form1.hdc, dx(1), dy(0), Color
    SetPixel Form1.hdc, dx(1), dy(1), Color

End If

End Sub




