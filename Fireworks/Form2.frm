VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1695
   ClientLeft      =   4455
   ClientTop       =   5505
   ClientWidth     =   4425
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      Max             =   1000
      Scrolling       =   1
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   1425
      MaxLength       =   7
      TabIndex        =   1
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press Enter To Accept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Index           =   1
      Left            =   1020
      TabIndex        =   2
      Top             =   1110
      Width           =   2385
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter The Number Of Particals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   270
      Width           =   4005
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Text1.SetFocus
Text1.SelText = 750
End Sub

Private Sub Form_Load()
Dim rgn As Long
rgn = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 50, 50)
SetWindowRgn hwnd, rgn, True
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE

End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteObject rgn
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    Form1.MAXPARTICLES = Val(Text1)
    ProgressBar1.Visible = True
    Label1(0) = "Initializing     Please Wait......"
    Label1(1).Visible = False
    Me.Refresh
    Load Form1
ElseIf KeyAscii = vbKeyEscape Then
    Unload Me
    End
End If

'To make sure user only enters numeric values
If Not KeyAscii = vbKeyBack Then If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub
