Attribute VB_Name = "basMainGameLoop"
Option Explicit

'//////////////////////////////////////////////////////////////////////////////
'This is the main game loop
'//////////////////////////////////////////////////////////////////////////////
Public Sub StartGame()
      On Error GoTo ErrorHandler
      
      Dim I As Integer, J As Integer
      Dim Index As Integer
      Dim TargetTick As Long
      Dim EmptyRect As RECT
            
      boolEndGame = False
      boolCanRotateBalls = False
      boolPuzzleSolved = False
      boolProcessing = False
      
      Call CreatePuzzle(GameFileIndex)
      
      Do
            TargetTick = dx.TickCount
            
            GetInput
                  
            ddsBack.BltColorFill EmptyRect, 0
            DisplayBackground
            
            
            If boolCanRotateBalls Then
                  If boolRotateLeft Then
                        Call RotateBallsLeft(BallClicked)
                  Else
                        Call RotateBallsRight(BallClicked)
                  End If
            End If
            
            DisplayLargeBalls
            DisplayPreviewBalls
            
            If boolCanShowCircle Then
                  DisplayCircle
            End If
            
            DisplayArrow
            
            boolGameStarted = True
            
            'In case the animation is too fast, put the following code
            'Do Until dx.TickCount - TargetTick > 5
            'Loop
            
            ddsFront.Flip Nothing, DDFLIP_WAIT
            DoEvents
      Loop Until boolEndGame
      
      EndGame
      Exit Sub
ErrorHandler:
      Debug.Print "Error occurred in StartGame sub : " & Err.Description
      EndGame
End Sub
