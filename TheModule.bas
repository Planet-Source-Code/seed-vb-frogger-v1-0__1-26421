Attribute VB_Name = "TheModule"
Option Base 1
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Type LilyData
  Left As Integer
  Top As Integer
  CarColor As Integer
End Type
Public Type PlayerData
  Left As Integer
  Top As Integer
End Type
Public Level(10) As String
Public u As PlayerData
Public Lily(15) As LilyData
Public LpX As Integer, KeyC As Integer, CarCount As Integer
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Sub LoadLevel(filename As String)
On Error GoTo handle
CarCount = 1
    Dim tmp As String
    Open App.Path + "\Levels\" + filename For Input As #1 'open the file and read the data for the level
        For i = 1 To 8
            Line Input #1, tmp ' read a whole line from the file
            For j = 1 To 14
                If Mid(tmp, j, 1) = "C" And CarCount < 16 Then
                    Lily(CarCount).Left = (j - 1) * 50 - 3
                    Lily(CarCount).Top = (i - 1) * 50 - 3
                    CarCount = CarCount + 1
                    If CarCount = 16 Then
                        CarCount = 15
                        Close #1
                        Exit Sub
                    End If
                End If
            Next j
        Next i
    Close #1
    Exit Sub
handle: 'There was an error load the map so tell the user what the error was and use the default map
    MsgBox "Error loading level! Using default level", vbExclamation, "ERROR : " & error(Err)
End Sub

Public Sub LoadLevels()
Open App.Path + "\Levels.dat" For Input As #1
    For i = 1 To 10
        Line Input #1, Level(i)
    Next i
Close #1
End Sub

Function UnloadEm() 'unloads all forms
Unload frmMain
Unload frmLevelEdit
Unload frmLoadMap
Unload frmSplash
Unload frmLevels
End
End Function

Public Sub PlayWav(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub
