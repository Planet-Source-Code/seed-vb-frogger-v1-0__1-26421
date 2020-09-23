VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frogger!"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   7590
   ScaleWidth      =   8985
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FF00&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Width           =   2175
   End
   Begin VB.PictureBox MaskLeft 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   3600
      Picture         =   "frmMain.frx":00CE
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox RedLeft 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   5160
      Picture         =   "frmMain.frx":3824
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer KeyGet 
      Interval        =   10
      Left            =   4440
      Top             =   1920
   End
   Begin VB.Timer Draw 
      Interval        =   1
      Left            =   4800
      Top             =   1920
   End
   Begin VB.PictureBox Masks 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   480
      Picture         =   "frmMain.frx":6F7A
      ScaleHeight     =   795
      ScaleWidth      =   3075
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox Frogs 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   480
      Picture         =   "frmMain.frx":E4EC
      ScaleHeight     =   795
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5955
      Left            =   0
      Picture         =   "frmMain.frx":15A5E
      ScaleHeight     =   5955
      ScaleWidth      =   9015
      TabIndex        =   0
      Top             =   720
      Width           =   9015
      Begin VB.PictureBox YellowLeft 
         AutoRedraw      =   -1  'True
         Height          =   855
         Left            =   5160
         Picture         =   "frmMain.frx":1A64F
         ScaleHeight     =   795
         ScaleWidth      =   1515
         TabIndex        =   10
         Top             =   4200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox YellowRight 
         AutoRedraw      =   -1  'True
         Height          =   855
         Left            =   5160
         Picture         =   "frmMain.frx":1DDA5
         ScaleHeight     =   795
         ScaleWidth      =   1515
         TabIndex        =   9
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox BlueLeft 
         AutoRedraw      =   -1  'True
         Height          =   855
         Left            =   5160
         Picture         =   "frmMain.frx":214FB
         ScaleHeight     =   795
         ScaleWidth      =   1515
         TabIndex        =   8
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox BlueRight 
         AutoRedraw      =   -1  'True
         Height          =   855
         Left            =   5160
         Picture         =   "frmMain.frx":24C51
         ScaleHeight     =   795
         ScaleWidth      =   1515
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox MaskRight 
         AutoRedraw      =   -1  'True
         Height          =   855
         Left            =   3600
         Picture         =   "frmMain.frx":283A7
         ScaleHeight     =   795
         ScaleWidth      =   1515
         TabIndex        =   6
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox RedRight 
         AutoRedraw      =   -1  'True
         Height          =   855
         Left            =   5160
         Picture         =   "frmMain.frx":2BAFD
         ScaleHeight     =   795
         ScaleWidth      =   1515
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Label lblLives 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lives: 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VBFrogger v1.0 by Alex Donavon (Aedseed@aol.com)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1365
      TabIndex        =   13
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   6840
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrLevel As Integer, BeatLevel As Boolean, Lives As Integer

Private Sub cmdExit_Click()
Hide
frmSplash.Show
End Sub

Private Sub Draw_Timer() 'The timer that draws stuff! (O:
    If BeatLevel = True Then
        PlayWav App.Path + "\Pass.wav"
        CurrLevel = CurrLevel + 1
            If CurrLevel = 11 Then WinGame
        lblLevel.Caption = "Level " & CurrLevel
        LoadLevel Level(CurrLevel)
        u.Left = 250
        u.Top = 347
        BeatLevel = False
    End If
CheckPads
picBG.Picture = frmLevelEdit.picBG.Picture 'Redraw BG
MovePads 'Move the lilypads around
DrawYou 'Draw Frogger
End Sub

Private Sub Form_Load()
Randomize
u.Left = 250
u.Top = 347
BitBlt picBG.hDC, u.Left, u.Top, 50, 50, Masks.hDC, 0, 0, vbSrcAnd
BitBlt picBG.hDC, u.Left, u.Top, 50, 50, Frogs.hDC, 0, 0, vbSrcPaint
LoadLevel "Level1.map"
Me.Picture = Nothing 'i accidently added a pic to the bg via the properties and this is the only way to rid it.
AssignColors
LoadLevels
CurrLevel = 1
Lives = 3
LoadLevel Level(CurrLevel)
End Sub

Sub AssignColors()
Dim Rn
For i = 1 To 15
    Rn = Int(Rnd * 3)
    If Rn = 0 Then Lily(i).CarColor = 0
    If Rn = 1 Then Lily(i).CarColor = 1
    If Rn = 2 Then Lily(i).CarColor = 2
Next i
End Sub

Sub CheckPads() 'repositions the cars if they go offscreen
For i = 1 To CarCount
    If Lily(i).Left > 600 Then Lily(i).Left = -100
    If Lily(i).Left < -100 Then Lily(i).Left = 600
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmSplash.Show
End Sub

Private Sub KeyGet_Timer() 'handles keypresses
If KeyC = 39 Then 'Right Key
    LpX = 50: KeyC = 0
    If u.Left = 550 Then Exit Sub
    u.Left = u.Left + 50
End If
If KeyC = 37 Then 'Left Key
    LpX = 150: KeyC = 0
    If u.Left = 0 Then Exit Sub
    u.Left = u.Left - 50
End If
If KeyC = 40 Then 'Down Key
    LpX = 100: KeyC = 0
    If u.Top = 347 Then Exit Sub
    u.Top = u.Top + 50
End If
If KeyC = 38 Then 'Up Key
    LpX = 0: KeyC = 0 'there is no checker cause when you leave screen at top, u win!
    u.Top = u.Top - 50
End If
End Sub

Sub DrawYou()
    BitBlt picBG.hDC, u.Left, u.Top, 50, 50, Masks.hDC, LpX, 0, vbSrcAnd
    BitBlt picBG.hDC, u.Left, u.Top, 50, 50, Frogs.hDC, LpX, 0, vbSrcPaint
    'only after you've been drawn do we actually want to render a win or death
    If u.Top < -50 Then BeatLevel = True
For i = 1 To CarCount
    If Lily(i).Top = u.Top And Lily(i).Left - u.Left <= 50 And u.Left - Lily(i).Left <= 100 Then
        Lives = Lives - 1
        If Lives = -1 Then
            resp = MsgBox("Oops!  You lost all your lives!  Would you like to play again?", vbQuestion + vbYesNo, "Oops!")
            If resp = vbNo Then
                UnloadEm
            Else
                CurrLevel = 1
                lblLevel.Caption = "Level " & CurrLevel
                LoadLevel Level(CurrLevel)
                u.Left = 250
                u.Top = 347
                Lives = 3
                lblLives = "Lives: " & Lives
            End If
        End If
        lblLives.Caption = "Lives: " & Lives
        u.Left = 250
        u.Top = 347
        PlayWav App.Path + "\Ribbit.wav"
    End If
Next i
End Sub

Sub MovePads()
Dim L As Integer
For i = 1 To CarCount
L = Lily(i).Top
    If L = -3 Or L = 97 Or L = 197 Or L = 297 Or L = 397 Or L = 497 Then
    Lily(i).Left = Lily(i).Left - 2
        Select Case Lily(i).CarColor
            Case 0
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, MaskLeft.hDC, 0, 0, vbSrcAnd
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, RedLeft.hDC, 0, 0, vbSrcPaint
            Case 1
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, MaskLeft.hDC, 0, 0, vbSrcAnd
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, BlueLeft.hDC, 0, 0, vbSrcPaint
            Case 2
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, MaskLeft.hDC, 0, 0, vbSrcAnd
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, YellowLeft.hDC, 0, 0, vbSrcPaint
        End Select
    Else
    Lily(i).Left = Lily(i).Left + 2
        Select Case Lily(i).CarColor
            Case 0
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, MaskRight.hDC, 0, 0, vbSrcAnd
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, RedRight.hDC, 0, 0, vbSrcPaint
            Case 1
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, MaskRight.hDC, 0, 0, vbSrcAnd
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, BlueRight.hDC, 0, 0, vbSrcPaint
            Case 2
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, MaskRight.hDC, 0, 0, vbSrcAnd
                BitBlt picBG.hDC, Lily(i).Left, Lily(i).Top, 100, 47, YellowRight.hDC, 0, 0, vbSrcPaint
        End Select
    End If
Next i
End Sub

Private Sub Label2_Click()

End Sub

Private Sub picBG_DblClick()
frmLevelEdit.Show
End Sub

Private Sub picBG_KeyDown(KeyCode As Integer, Shift As Integer)
KeyC = KeyCode
End Sub

Sub WinGame()
resp = MsgBox("Congrats!  You beat VBFrogger v1.0!!!  Would you like to play again?", vbYesNo + vbInformation, "Cool!")
If resp = vbNo Then
UnloadEm
Else
CurrLevel = 1
lblLevel.Caption = CurrLevel
LoadLevel Level(CurrLevel)
u.Left = 250
u.Top = 347
Lives = 3
lblLives = "Lives: " & Lives
End If
End Sub
