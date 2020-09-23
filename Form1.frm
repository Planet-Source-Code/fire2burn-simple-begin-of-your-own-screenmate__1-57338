VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1215
   ClientLeft      =   5925
   ClientTop       =   2835
   ClientWidth     =   1620
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   1620
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1170
      Top             =   615
   End
   Begin VB.Image Image4 
      Height          =   270
      Index           =   2
      Left            =   765
      Picture         =   "Form1.frx":0000
      Top             =   915
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image4 
      Height          =   270
      Index           =   1
      Left            =   405
      Picture         =   "Form1.frx":0552
      Top             =   915
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image4 
      Height          =   270
      Index           =   0
      Left            =   45
      Picture         =   "Form1.frx":0AA4
      Top             =   915
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   270
      Index           =   2
      Left            =   765
      Picture         =   "Form1.frx":0FF6
      Top             =   630
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   270
      Index           =   1
      Left            =   405
      Picture         =   "Form1.frx":1548
      Top             =   630
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   270
      Index           =   0
      Left            =   45
      Picture         =   "Form1.frx":1A9A
      Top             =   630
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   2
      Left            =   765
      Picture         =   "Form1.frx":1FEC
      Top             =   345
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   1
      Left            =   405
      Picture         =   "Form1.frx":253E
      Top             =   345
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   0
      Left            =   45
      Picture         =   "Form1.frx":2A90
      Top             =   345
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   270
      Index           =   2
      Left            =   765
      Picture         =   "Form1.frx":2FE2
      Top             =   60
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   270
      Index           =   1
      Left            =   405
      Picture         =   "Form1.frx":3534
      Top             =   60
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   270
      Index           =   0
      Left            =   45
      Picture         =   "Form1.frx":3A86
      Top             =   60
      Visible         =   0   'False
      Width           =   345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LWA_COLORKEY = 1
Private Const LWA_ALPHA = 2
Private Const LWA_BOTH = 3
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = -20
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal color As Long, ByVal x As Byte, ByVal alpha As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Dim nieuwplaatje As Integer
Dim animatienr As Integer

Sub SetTranslucent(ThehWnd As Long, color As Long, nTrans As Integer, flag As Byte)
On Error Resume Next
Dim attrib As Long
attrib = GetWindowLong(ThehWnd, GWL_EXSTYLE)
SetWindowLong ThehWnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
SetLayeredWindowAttributes ThehWnd, color, nTrans, flag
End Sub



Private Sub Form_Load()
Me.BackColor = RGB(255, 0, 255)
SetTranslucent Me.hwnd, RGB(255, 0, 255), 255, 1
nieuwplaatje = 0
animatienr = 1
End Sub

Private Sub Form_Click()
If animatienr = 4 Then
animatienr = 1
Else
animatienr = animatienr + 1
End If
End Sub

Private Sub Timer1_Timer()
'Me.Picture = Image1(nieuwplaatje)
If animatienr = 1 Then
WalkTOP
ElseIf animatienr = 2 Then
WalkRIGHT
ElseIf animatienr = 3 Then
WalkBOTTOM
ElseIf animatienr = 4 Then
WalkLEFT
End If
End Sub

Function WalkTOP()
If nieuwplaatje = Image1.Count - 1 Then
nieuwplaatje = -1
End If
nieuwplaatje = nieuwplaatje + 1
Me.Picture = Image1(nieuwplaatje)
End Function

Function WalkRIGHT()
If nieuwplaatje = Image2.Count - 1 Then
nieuwplaatje = -1
End If
nieuwplaatje = nieuwplaatje + 1
Me.Picture = Image2(nieuwplaatje)
End Function
Function WalkBOTTOM()
If nieuwplaatje = Image3.Count - 1 Then
nieuwplaatje = -1
End If
nieuwplaatje = nieuwplaatje + 1
Me.Picture = Image3(nieuwplaatje)
End Function
Function WalkLEFT()
If nieuwplaatje = Image4.Count - 1 Then
nieuwplaatje = -1
End If
nieuwplaatje = nieuwplaatje + 1
Me.Picture = Image4(nieuwplaatje)
End Function

