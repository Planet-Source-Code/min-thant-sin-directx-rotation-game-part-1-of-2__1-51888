Attribute VB_Name = "basCheckingGameStatus"
Option Explicit

Private Const msg1 As String = "Right Arrow Key = Next Level"
Private Const msg2 As String = "Left Arrow Key = Previous Level"
      
Public Function PuzzleSolved() As Boolean
      Dim I As Integer
      Dim Index As Integer
      
      PuzzleSolved = True
      
      '////////////////////////////////////////////////////////////////////////
      '/// Check if the current balls' colors in the board coordinates match
      '/// the colors in the game level file's board coordinates.
      '////////////////////////////////////////////////////////////////////////
      For I = 0 To UBound(LevelData())
            Index = GetBallFromCoord(LevelData(I).BallX, LevelData(I).BallY)
            If Index <> -1 Then
                  If Balls(Index).Color <> LevelData(I).BallColor Then
                        PuzzleSolved = False
                        Exit For
                  End If
            End If
      Next I
      
End Function

Public Sub CheckGameStatus()
      If Not PuzzleSolved Then Exit Sub
      
      Dim nWidth As Integer, nHeight As Integer
      Dim Message As String
      Dim TimeMarker As Double
      Dim KeyboardState(0 To 255) As Byte
      Dim PuzzleSolvedRect As RECT
      Dim EmptyRect As RECT
      
      boolPuzzleSolved = True
      
      'The PuzzleSolved.bmp graphic's width and height
      nWidth = 560:       nHeight = 41
      
      'Game Over graphics.
      With PuzzleSolvedRect
            .Left = 0
            .Top = 0
            .Right = .Left + nWidth
            .Bottom = .Top + nHeight
      End With
      
      DisplayLargeBalls
      DisplayPreviewBalls
      
      'Center the graphic and blit it.
      ddsBack.BltFast (SCREEN_WIDTH - nWidth) \ 2, (SCREEN_HEIGHT - nHeight) \ 2, _
      ddsPuzzleSolved, PuzzleSolvedRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                              
      Message = "Press Enter to continue or Esc to quit..."
      DisplayText Message, False
      
      ddsFront.Flip Nothing, 0
            
      'Ask the user to press either Enter or Esc key.
      Do
            diKeyBoard.GetDeviceState 256, KeyboardState(0)
            DisplayArrow
            DoEvents
      Loop Until ((KeyboardState(DIK_RETURN) And &H80) <> 0) Or _
                       ((KeyboardState(DIK_ESCAPE) And &H80) <> 0)
      
      'The player presses Esc key
      If (KeyboardState(DIK_ESCAPE) And &H80) <> 0 Then
            EndGame
      Else  'Enter key.
                  
            LoadGame True
      End If
End Sub

Public Sub LoadGame(boolIncrement As Boolean)
      Dim Message As String
      Dim TimeMarker As Double
      Dim IncVal As Integer
      Dim EmptyRect As RECT
      
      boolGameStarted = False
      
      If boolIncrement Then
            IncVal = 1
      Else
            IncVal = -1
      End If
      
      GameFileIndex = GameFileIndex + IncVal
                  
      If GameFileIndex > (NumGameLevelFiles - 1) Then GameFileIndex = 0
      If GameFileIndex < 0 Then GameFileIndex = (NumGameLevelFiles - 1)
      
      Call CreatePuzzle(GameFileIndex)
      
      Message = "Generating puzzle, please wait..."
      TimeMarker = dx.TickCount
      
      Do Until dx.TickCount > TimeMarker + 500
            ddsBack.BltColorFill EmptyRect, 0
            DisplayText Message, True
            ddsFront.Flip Nothing, DDFLIP_WAIT
            DoEvents
      Loop
            
      boolGameStarted = True
      boolPuzzleSolved = False
End Sub

Public Sub DisplayText(str As String, boolCenterY As Boolean)
      Dim StringWidth As Integer, StringHeight As Integer
      Dim Stringx As Integer, StringY As Integer
      
      StringWidth = frmMain.TextWidth(str)
      StringHeight = frmMain.TextHeight(str)
                  
      Stringx = (SCREEN_WIDTH - StringWidth) \ 2
      StringY = (SCREEN_HEIGHT - StringHeight)
      
      If boolCenterY Then
            StringY = (SCREEN_HEIGHT - StringHeight) \ 2
      End If
      
      ddsBack.SetForeColor RGB(125, 125, 125)
      ddsBack.DrawText Stringx, StringY, str, False
      
      ddsBack.SetForeColor vbCyan
      ddsBack.DrawText Stringx + 2, StringY - 2, str, False
End Sub
