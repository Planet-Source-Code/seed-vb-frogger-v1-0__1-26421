VERSION 5.00
Begin VB.Form frmLoadMap 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load Map"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox File1 
      BackColor       =   &H0000FF00&
      Height          =   2430
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H0000FF00&
      Caption         =   "&Load"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H0000FF00&
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmLoadMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLoad_Click()
On Error Resume Next
If Right(File1.filename, 4) <> ".map" Then Exit Sub
frmLevelEdit.LoadMap Dir1 + "\" + File1.filename
Me.Hide
End Sub

Private Sub Dir1_Change()
File1 = Dir1
End Sub

Private Sub Drive1_Change()
On Error GoTo errd
Dir1 = Drive1
Exit Sub
errd:
MsgBox "Hey, the drive is empty!", vbOKOnly + vbCritical, "Error:"
End Sub

Private Sub File1_DblClick()
cmdLoad_Click
End Sub
