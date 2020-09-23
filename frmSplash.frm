VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer ex 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7320
      Top             =   2280
   End
   Begin VB.Timer lvl 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7680
      Top             =   2280
   End
   Begin VB.Timer play 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8040
      Top             =   2280
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0 by Alex Donavon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Label lblLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Level Editor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Label lblPlay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Play VB Frogger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim u As Boolean, d As Boolean, m As Boolean, c As Integer

Private Sub ex_Timer()
If c > 255 Then
u = False
End If
If u = True Then c = c + 24
Label1.ForeColor = RGB(0, c, 0)
End Sub

Private Sub Form_Load()
c = 0
u = True
lblPlay.ForeColor = RGB(255, 255, 255)
lblLE.ForeColor = RGB(255, 255, 255)
Label1.ForeColor = RGB(255, 255, 255)
PlayWav App.Path + "\Intro.wav"
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
play.Enabled = False
lvl.Enabled = False
ex.Enabled = False
lblPlay.ForeColor = RGB(255, 255, 255)
lblLE.ForeColor = RGB(255, 255, 255)
Label1.ForeColor = RGB(255, 255, 255)
c = 50
End Sub

Private Sub Label1_Click()
Unload frmMain
Unload frmLevelEdit
Unload frmLoadMap
Unload Me
Unload frmLevels
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
u = True
ex.Enabled = True
End Sub

Private Sub lblLE_Click()
Hide
frmLevelEdit.Show
End Sub

Private Sub lblLE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lvl.Enabled = True
u = True
End Sub

Private Sub lblPlay_Click()
Hide
frmMain.Show
End Sub

Private Sub lblPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
u = True
play.Enabled = True
End Sub

Sub Pause(Interval)
current = Timer
Do While Timer - current < Interval
DoEvents
Loop
End Sub

Private Sub lvl_Timer()
If c > 255 Then
u = False
End If
If u = True Then c = c + 24
lblLE.ForeColor = RGB(0, c, 0)
End Sub

Private Sub play_Timer()
If c > 255 Then
u = False
End If
If u = True Then c = c + 24
lblPlay.ForeColor = RGB(0, c, 0)
End Sub

