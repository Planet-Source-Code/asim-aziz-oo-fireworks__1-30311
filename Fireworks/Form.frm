VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   4950
   ClientTop       =   5010
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Asim Aziz
'chirisoft@flashmail.com
'http://www.chirisoft.cjb.net
'Most of the code is self explainatory but i've left
'comments wherever necessary
'To vote for the code Goto http://www.planetsourcecode.com/vb/scripts/showcode.asp?txtCodeId=30311&lngWId=1

'This form is the main drawing area
Option Explicit
Dim StartPosX As Integer, StartPosY As Integer 'Starting position of explosion
Public MAXPARTICLES As Long 'number of particles
Dim ExpSize As Integer 'Size of explosion
Dim R As Integer, G As Integer, B As Integer, Color As Long 'Color values
Dim p As Particle
Dim DoExplosion As Boolean
Dim Particle_Collection As P_Collection


Private Sub Form_Load()
Dim i As Long, Size As Boolean
ShowCursor False 'hides the cursor
Randomize 'initialize the random number generator

DoExplosion = True
Set Particle_Collection = New P_Collection

Form2.ProgressBar1.Max = MAXPARTICLES


For i = 1 To MAXPARTICLES
    Form2.ProgressBar1.Value = i
    
    Select Case Rnd
        Case Is > 0.6
        Size = 1
        Case Else
        Size = 0
    End Select

    'Add a new particle to the particle collection with given values
    Particle_Collection.Add Size, Color, StartPosX, StartPosY
Next


ExpSize = 30 + 70 * Rnd
Unload Form2
Me.Show
Explode
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
ShowCursor True
DoExplosion = False
Unload Me
End
Else: Me.Refresh
End If
End Sub




Private Sub Form_Unload(Cancel As Integer)
Load Form2
Form2.Show
Form2.ProgressBar1.Visible = True
Form2.ProgressBar1.Max = MAXPARTICLES
Form2.Label1(0) = "Cleaing Up   Please Wait....."
Form2.Label1(1).Visible = False
Form2.Refresh

For Each p In Particle_Collection
    Form2.ProgressBar1.Value = Form2.ProgressBar1.Value + 1
    Set p = Nothing ' delete the instance of particle object
Next

Unload Form2
End
End Sub

Public Sub Explode()
Dim i As Integer, Rdec As Double, Gdec As Double, Bdec As Double
Randomize

Do While DoExplosion
    
    ExpSize = 10 + (MAXPARTICLES / 40) * Rnd
    R = 100 + 155 * Rnd
    G = 100 + 155 * Rnd
    B = 100 + 155 * Rnd
    Color = RGB(R, G, B)

    StartPosX = Form1.ScaleWidth * Rnd
    StartPosY = Form1.ScaleHeight * Rnd

    For Each p In Particle_Collection
        p.Color = Color
        p.StartPosX = StartPosX
        p.StartPosY = StartPosY
        p.SetValues
    Next



    Rdec = R / ExpSize
    Gdec = G / ExpSize
    Bdec = B / ExpSize


    For i = 1 To ExpSize
        If R >= Rdec Then R = R - Rdec
        If G >= Gdec Then G = G - Gdec
        If B >= Bdec Then B = B - Bdec
        Color = RGB(R, G, B)

        For Each p In Particle_Collection
            p.Color = Color
            p.Move
        Next

        DoEvents
    Next
Loop

End Sub

