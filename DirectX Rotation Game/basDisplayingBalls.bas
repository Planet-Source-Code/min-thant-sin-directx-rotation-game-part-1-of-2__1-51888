Attribute VB_Name = "basDisplayingBalls"
Option Explicit

'This is for displaying large balls
Public Sub DisplayLargeBalls()
      Dim I As Integer
      Dim BallRect As RECT
      
      For I = 0 To (TOTAL_BALLS - 1)
            If Balls(I).Exists Then
                  With BallRect
                        .Left = Balls(I).HomeLeft
                        .Top = Balls(I).HomeTop
                        .Right = .Left + BALL_WIDTH
                        .Bottom = .Top + BALL_HEIGHT
                  End With
                  
                  ddsBack.BltFast Balls(I).Left, Balls(I).Top, _
                  ddsBalls, BallRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
            
            DoEvents
      Next I
End Sub

'This is for displaying small preview balls
Public Sub DisplayPreviewBalls()
      Dim I As Integer
      Dim BallRect As RECT
      
      For I = 0 To (TOTAL_BALLS - 1)
            If PreviewBalls(I).Exists Then
                  With BallRect
                        .Left = PreviewBalls(I).HomeLeft
                        .Top = PreviewBalls(I).HomeTop
                        .Right = .Left + SMALLBALL_WIDTH
                        .Bottom = .Top + SMALLBALL_WIDTH
                  End With
                        
                  ddsBack.BltFast PreviewBalls(I).Left, PreviewBalls(I).Top, _
                  ddsPreviewBalls, BallRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
            
            DoEvents
      Next I
End Sub
