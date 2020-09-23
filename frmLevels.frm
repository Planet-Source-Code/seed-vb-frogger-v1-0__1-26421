VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLevels 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Levels"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   4860
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command11 
      BackColor       =   &H0000FF00&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2543
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4680
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0000FF00&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0000FF00&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FF00&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FF00&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FF00&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   9
      Left            =   1440
      TabIndex        =   20
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   8
      Left            =   1440
      TabIndex        =   19
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   7
      Left            =   1440
      TabIndex        =   18
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   6
      Left            =   1440
      TabIndex        =   17
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   5
      Left            =   1440
      TabIndex        =   16
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   4
      Left            =   1440
      TabIndex        =   15
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   14
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   2
      Left            =   1440
      TabIndex        =   13
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   12
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   11
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000FF00&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   863
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 2:"
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
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 3:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 4:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 5:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 6:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 7:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 8:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 9:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 10:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Level 1:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RName(0 To 9) As String

Private Sub cmdOK_Click()
For f = 0 To 9
    If txt(f) = "" Then MsgBox "Please fill in all the levels!", vbCritical + vbOKOnly, "Error:": Exit Sub
Next f
'lotsa string manipulation here.  gets all the levels u chose and puts them in levels.dat
Dim Char$
For j = 0 To 9
    For P = 1 To Len(txt(j))
        Char = Mid$(txt(j), P, 1)
        If Char = "\" Then RName(j) = Mid(txt(j), P + 1, Len(txt(j)) - P)
    Next P
Next j
Open App.Path + "\Levels.dat" For Output As #1
    For i = 0 To 9
        Print #1, RName(i)
    Next i
Close #1
Hide
frmLevelEdit.Show
End Sub

Private Sub Command1_Click()
CD.ShowOpen
txt(0) = CD.filename
End Sub

Private Sub Command10_Click()
CD.ShowOpen
txt(1) = CD.filename
End Sub

Private Sub Command11_Click()
Unload Me
End Sub

Private Sub Command2_Click()
CD.ShowOpen
txt(9) = CD.filename
End Sub

Private Sub Command3_Click()
CD.ShowOpen
txt(8) = CD.filename
End Sub

Private Sub Command4_Click()
CD.ShowOpen
txt(7) = CD.filename
End Sub

Private Sub Command5_Click()
CD.ShowOpen
txt(6) = CD.filename
End Sub

Private Sub Command6_Click()
CD.ShowOpen
txt(5) = CD.filename
End Sub

Private Sub Command7_Click()
CD.ShowOpen
txt(4) = CD.filename
End Sub

Private Sub Command8_Click()
CD.ShowOpen
txt(3) = CD.filename
End Sub

Private Sub Command9_Click()
CD.ShowOpen
txt(2) = CD.filename
End Sub

Private Sub Form_Load()
BackColor = vbBlack
CD.Filter = "*.map (Frogger Maps)|*.map"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
frmLevelEdit.Show
End Sub
