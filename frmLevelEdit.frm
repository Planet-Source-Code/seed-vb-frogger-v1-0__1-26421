VERSION 5.00
Begin VB.Form frmLevelEdit 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Frogger Level Editor"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11160
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FF00&
      Caption         =   "&Exit"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000FF00&
      Caption         =   "&Clear"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H0000C000&
      Caption         =   "&Help"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1335
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
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FF00&
      Caption         =   "&Save"
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   855
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   0
      Picture         =   "frmLevelEdit.frx":0000
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   741
      TabIndex        =   0
      Top             =   0
      Width           =   11175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VBFrogger v1.0 Level Editor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6600
      MousePointer    =   2  'Cross
      TabIndex        =   5
      Top             =   6120
      Width           =   3735
   End
End
Attribute VB_Name = "frmLevelEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim cars As Integer, P As Integer
Dim px As Integer, py As Integer
Dim map(14, 8) As String

Private Sub cmdClear_Click()
picBG.Cls
cars = 0
NewMap
End Sub

Private Sub cmdExit_Click()
Hide
frmSplash.Show
End Sub

Private Sub cmdHelp_Click()
MsgBox _
"LEFT MOUSE: Add car." + vbCrLf + _
"RIGHT MOUSE: Delete car." + vbCrLf + _
"Load/Save/Clear Maps with the corresponding buttons" + vbCrLf + vbCrLf + _
"NOTES:" + vbCrLf + _
"All the cars will appear red, but will be randomly distributed a color at run time." + vbCrLf + _
"The cars all appear to face one way, but will travel in opposite directions at run time." + vbCrLf + _
"Use the white area to add cars that will not appear onscreen at first." _
, vbOKOnly + vbInformation + vbSystemModal, "Help:"
End Sub

Private Sub cmdLoad_Click()
frmLoadMap.Show
End Sub

Private Sub cmdSave_Click()
If cars < 15 Then MsgBox "Please use all the cars avaliable!": Exit Sub
Dim f As String
RET:
f = InputBox("What would you like to name your map?", "Save:")
If Len(f) = 0 Then Exit Sub
If Len(f) > 20 Or InStr(f, "?") > 0 Or InStr(f, ".") > 0 Or InStr(f, "\") > 0 Then
    MsgBox "Please enter a valid name!", vbOKOnly + vbCritical + vbSystemModal, "Error:"
    GoTo RET
End If
If Right(f, 4) <> ".map" Then f = f + ".map" 'if the filename the user specifies doesn't end with ".map" then add it to the end of the string
SaveMap App.Path + "\Levels\" + f ' save the map
resp = MsgBox("File saved successfully.  Would you like to add this level to the game?", vbYesNo + vbInformation, "Saved")
If resp = vbNo Then
Exit Sub
Else
'
frmLevels.Show
Hide
'
End If
End Sub

Private Sub Form_Load()
NewMap
cars = 0
End Sub

Sub NewMap()
For i = 1 To 14
    For j = 1 To 8
        map(i, j) = "O"
    Next j
Next i
DrawMap
End Sub
Sub LoadMap(filename As String)
    'On Error GoTo FixIt
    Dim tmp As String
    Open filename For Input As #1 'open the file and read the data for the level
    cars = 0
        For i = 1 To 8
            Line Input #1, tmp ' read a whole line from the file
            For j = 1 To 14
                map(j, i) = Mid(tmp, j, 1) ' chop the data out of each column of the string
                If map(j, i) = "C" Then cars = cars + 1
            Next j
        Next i
    Close #1
    DrawMap ' And refresh the map
    Exit Sub
FixIt:
    MsgBox "There was an error loading this level", vbExclamation, "ERROR!!"
End Sub
Sub SaveMap(filename As String)

    Dim tmp As String
    Open filename For Output As #2 'open the file and write the data for the level
        For i = 1 To 8
            For j = 1 To 14
                Print #2, map(j, i); 'Write this data to the file without adding a character return
            Next j
            Print #2, "" 'Generates a character return so you don't get one long line of stuff
        Next i
    Close #2
End Sub
Sub DrawMap()
picBG.Cls
    For i = 1 To 14
        For j = 1 To 8
            If map(i, j) = "C" Then
                BitBlt picBG.hDC, (i - 1) * 50 - 3, (j - 1) * 50 - 3, 100, 47, frmMain.MaskLeft.hDC, 0, 0, vbSrcAnd
                BitBlt picBG.hDC, (i - 1) * 50 - 3, (j - 1) * 50 - 3, 100, 47, frmMain.RedLeft.hDC, 0, 0, vbSrcPaint
            End If
        Next j
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmSplash.Show
End Sub

Private Sub Label1_Click()
cmdHelp_Click
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
End Sub

Private Sub picBG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
    px = Int(X / 50) + 1
    py = Int(Y / 50) + 1
    If px < 15 And py < 9 Then
        If Button = 1 Then ' the left mouse button was clicked so we add a car
            If cars < 15 Then
                If map(px, py) = "C" Then Exit Sub 'if car already there then leave
                If px - 1 <> 0 And px < 14 Then 'in other words, if its in the first column
                    If map(px - 1, py) = "C" Then
                        map(px - 1, py) = "O"
                        cars = cars - 1 'if there is a car just in front of the new one then get rid of it!
                    End If
                    If map(px + 1, py) = "C" Then
                        map(px + 1, py) = "O"
                        cars = cars - 1 ' likewise if there is one behind
                    End If
                End If
                map(px, py) = "C" ' put our new car into the array
                cars = cars + 1
            Else ' there are too many cars on the map so we'll tell the user to get rid of some
                MsgBox "There are already 15 cars on the map.", vbExclamation + vbOKOnly, "Error:"
            End If
        Else 'another mouse button was pressed so we'll delete a car if there is one there
            If map(px, py) = "C" Then map(px, py) = "O": cars = cars - 1
        End If
        DrawMap ' then refresh the map
    End If
End Sub


