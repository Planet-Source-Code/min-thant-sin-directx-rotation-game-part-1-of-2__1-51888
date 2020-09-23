Attribute VB_Name = "basRotatingBalls"
Option Explicit

'///////////////////////////////////////////////////////////////////////////////
'/// Sorry for lack of comments.
'/// Think of the balls rotating clockwise.
'///////////////////////////////////////////////////////////////////////////////

Sub RotateBallsRight(ByVal BallIndex As Integer)
      Dim TopLeftBall As Integer, TopRightBall As Integer
      Dim BottomLeftBall As Integer, BottomRightBall As Integer
      Dim LeftBall As Integer, RightBall As Integer
      Dim I As Integer, RowIndex As Integer
      
      I = BallIndex     'BallIndex is too long, for example >> compare "Balls(BallIndex).XCoord" with "Balls(I).XCoord"
      RowIndex = Balls(I).YCoord
      
      If RowIndex < MIDDLE_ROW Then
            TopLeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord - 1)
            TopRightBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord - 1)
            BottomLeftBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord + 1)
            BottomRightBall = GetBallFromCoord(Balls(I).XCoord + 1, Balls(I).YCoord + 1)
            LeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord)
            RightBall = GetBallFromCoord(Balls(I).XCoord + 1, Balls(I).YCoord)
            
      Else
            If Balls(I).YCoord > MIDDLE_ROW Then
                  TopLeftBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord - 1)
                  TopRightBall = GetBallFromCoord(Balls(I).XCoord + 1, Balls(I).YCoord - 1)
                  BottomLeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord + 1)
                  BottomRightBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord + 1)
                  LeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord)
                  RightBall = GetBallFromCoord(Balls(I).XCoord + 1, Balls(I).YCoord)
            Else
                  TopLeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord - 1)
                  TopRightBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord - 1)
                  BottomLeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord + 1)
                  BottomRightBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord + 1)
                  LeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord)
                  RightBall = GetBallFromCoord(Balls(I).XCoord + 1, Balls(I).YCoord)
            End If
      End If
      
      If TopLeftBall = -1 Or TopRightBall = -1 Or _
           RightBall = -1 Or BottomRightBall = -1 Or _
           BottomLeftBall = -1 Or LeftBall = -1 Then
           boolCanRotateBalls = False
            boolProcessing = False
            Exit Sub
      End If
      
      'Top-left ball
      Balls(TopLeftBall).Left = Balls(TopLeftBall).Left + FAST_SPEED
      
      'Top-right ball
      Balls(TopRightBall).Left = Balls(TopRightBall).Left + SLOW_SPEED
      Balls(TopRightBall).Top = Balls(TopRightBall).Top + FAST_SPEED
      
      'Right ball
      Balls(RightBall).Left = Balls(RightBall).Left - SLOW_SPEED
      Balls(RightBall).Top = Balls(RightBall).Top + FAST_SPEED
      
      'Bottom-right ball
      Balls(BottomRightBall).Left = Balls(BottomRightBall).Left - FAST_SPEED
      
      'Bottom-left ball
      Balls(BottomLeftBall).Left = Balls(BottomLeftBall).Left - SLOW_SPEED
      Balls(BottomLeftBall).Top = Balls(BottomLeftBall).Top - FAST_SPEED
      
      'Left ball
      Balls(LeftBall).Left = Balls(LeftBall).Left + SLOW_SPEED
      Balls(LeftBall).Top = Balls(LeftBall).Top - FAST_SPEED
      
      RotateStep = RotateStep + 1
            
      'The balls have reached their destination, update the data.
      If RotateStep >= (BALL_WIDTH / FAST_SPEED) Then
            RotateStep = 0
            
            'Top-left ball
            With Balls(TopLeftBall)
                  .XCoord = .XCoord + 1
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = TopLeftBall
            End With
            
            'Top-right ball
            With Balls(TopRightBall)
                  If RowIndex < MIDDLE_ROW Then
                        .XCoord = .XCoord + 1
                        .YCoord = .YCoord + 1
                  Else
                        If RowIndex > MIDDLE_ROW Then
                              .YCoord = .YCoord + 1
                        Else
                              .XCoord = .XCoord + 1
                              .YCoord = .YCoord + 1
                        End If
                  End If
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = TopRightBall
            End With
            
            'Right ball
            With Balls(RightBall)
                  If RowIndex < MIDDLE_ROW Then
                        .YCoord = .YCoord + 1
                  Else
                        If RowIndex > MIDDLE_ROW Then
                              .XCoord = .XCoord - 1
                              .YCoord = .YCoord + 1
                        Else
                              .XCoord = .XCoord - 1
                              .YCoord = .YCoord + 1
                        End If
                  End If
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = RightBall
            End With
            
            'Bottom-right ball
            With Balls(BottomRightBall)
                  .XCoord = .XCoord - 1
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = BottomRightBall
            End With
            
            'Bottom-left ball
            With Balls(BottomLeftBall)
                  If RowIndex < MIDDLE_ROW Then
                        .XCoord = .XCoord - 1
                        .YCoord = .YCoord - 1
                  Else
                        If RowIndex > MIDDLE_ROW Then
                              .YCoord = .YCoord - 1
                        Else
                              .YCoord = .YCoord - 1
                        End If
                  End If
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = BottomLeftBall
            End With
            
            'Left ball
            With Balls(LeftBall)
                  If RowIndex < MIDDLE_ROW Then
                        .YCoord = .YCoord - 1
                  Else
                        If RowIndex > MIDDLE_ROW Then
                              .XCoord = .XCoord + 1
                              .YCoord = .YCoord - 1
                        Else
                              .YCoord = .YCoord - 1
                        End If
                  End If
                  
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = LeftBall
            End With
                        
            boolCanRotateBalls = False
            boolProcessing = False
            
            CheckGameStatus
      End If
End Sub

'///////////////////////////////////////////////////////////////////////////////
'/// Sorry for lack of comments.
'/// Think of the balls rotating COUNTER-clockwise.
'///////////////////////////////////////////////////////////////////////////////
Sub RotateBallsLeft(ByVal BallIndex As Integer)
      Dim TopLeftBall As Integer
      Dim TopRightBall As Integer
      Dim BottomLeftBall As Integer
      Dim BottomRightBall As Integer
      Dim LeftBall As Integer
      Dim RightBall As Integer
      Dim I As Integer
      Dim RowIndex As Integer
      
      I = BallIndex
      RowIndex = Balls(I).YCoord
      
      If RowIndex < MIDDLE_ROW Then
            TopLeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord - 1)
            TopRightBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord - 1)
            BottomLeftBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord + 1)
            BottomRightBall = GetBallFromCoord(Balls(I).XCoord + 1, Balls(I).YCoord + 1)
            LeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord)
            RightBall = GetBallFromCoord(Balls(I).XCoord + 1, Balls(I).YCoord)
            
      Else
            If Balls(I).YCoord > MIDDLE_ROW Then
                  TopLeftBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord - 1)
                  TopRightBall = GetBallFromCoord(Balls(I).XCoord + 1, Balls(I).YCoord - 1)
                  BottomLeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord + 1)
                  BottomRightBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord + 1)
                  LeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord)
                  RightBall = GetBallFromCoord(Balls(I).XCoord + 1, Balls(I).YCoord)
            Else
                  TopLeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord - 1)
                  TopRightBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord - 1)
                  BottomLeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord + 1)
                  BottomRightBall = GetBallFromCoord(Balls(I).XCoord, Balls(I).YCoord + 1)
                  LeftBall = GetBallFromCoord(Balls(I).XCoord - 1, Balls(I).YCoord)
                  RightBall = GetBallFromCoord(Balls(I).XCoord + 1, Balls(I).YCoord)
            End If
      End If
      
      If TopLeftBall = -1 Or TopRightBall = -1 Or _
           RightBall = -1 Or BottomRightBall = -1 Or _
           BottomLeftBall = -1 Or LeftBall = -1 Then
           boolCanRotateBalls = False
            boolProcessing = False
            Exit Sub
      End If
      
      'Top-left ball
      Balls(TopLeftBall).Left = Balls(TopLeftBall).Left - SLOW_SPEED
      Balls(TopLeftBall).Top = Balls(TopLeftBall).Top + FAST_SPEED
      
      'Top-right ball
      Balls(TopRightBall).Left = Balls(TopRightBall).Left - FAST_SPEED
      
      'Right ball
      Balls(RightBall).Left = Balls(RightBall).Left - SLOW_SPEED
      Balls(RightBall).Top = Balls(RightBall).Top - FAST_SPEED
      
      'Bottom-right ball
      Balls(BottomRightBall).Left = Balls(BottomRightBall).Left + SLOW_SPEED
      Balls(BottomRightBall).Top = Balls(BottomRightBall).Top - FAST_SPEED
      
      'Bottom-left ball
      Balls(BottomLeftBall).Left = Balls(BottomLeftBall).Left + FAST_SPEED
      
      'Left ball
      Balls(LeftBall).Left = Balls(LeftBall).Left + SLOW_SPEED
      Balls(LeftBall).Top = Balls(LeftBall).Top + FAST_SPEED
      
      RotateStep = RotateStep + 1
      
      'The balls have reached their destination, update the data.
      If RotateStep >= (BALL_WIDTH / FAST_SPEED) Then
            RotateStep = 0
            
            'Top-left ball
            With Balls(TopLeftBall)
                  If RowIndex < MIDDLE_ROW Then
                        .YCoord = .YCoord + 1
                  Else
                        If RowIndex > MIDDLE_ROW Then
                              .XCoord = .XCoord - 1
                              .YCoord = .YCoord + 1
                        Else
                              .YCoord = .YCoord + 1
                        End If
                  End If
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = TopLeftBall
            End With
            
            'Top-right ball
            With Balls(TopRightBall)
                  If RowIndex < MIDDLE_ROW Then
                        .XCoord = .XCoord - 1
                  Else
                        If RowIndex > MIDDLE_ROW Then
                              .XCoord = .XCoord - 1
                        Else
                              .XCoord = .XCoord - 1
                        End If
                  End If
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = TopRightBall
            End With
            
            'Right ball
            With Balls(RightBall)
                  If RowIndex < MIDDLE_ROW Then
                        .XCoord = .XCoord - 1
                        .YCoord = .YCoord - 1
                  Else
                        If RowIndex > MIDDLE_ROW Then
                              .YCoord = .YCoord - 1
                        Else
                              .XCoord = .XCoord - 1
                              .YCoord = .YCoord - 1
                        End If
                  End If
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = RightBall
            End With
            
            'Bottom-right ball
            With Balls(BottomRightBall)
                  If RowIndex < MIDDLE_ROW Then
                        .YCoord = .YCoord - 1
                  Else
                        If RowIndex > MIDDLE_ROW Then
                              .XCoord = .XCoord + 1
                              .YCoord = .YCoord - 1
                        Else
                              .XCoord = .XCoord + 1
                              .YCoord = .YCoord - 1
                        End If
                  End If
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = BottomRightBall
            End With
            
            'Bottom-left ball
            With Balls(BottomLeftBall)
                  .XCoord = .XCoord + 1
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = BottomLeftBall
            End With
            
            'Left ball
            With Balls(LeftBall)
                  If RowIndex < MIDDLE_ROW Then
                        .XCoord = .XCoord + 1
                        .YCoord = .YCoord + 1
                  Else
                        If RowIndex > MIDDLE_ROW Then
                              .YCoord = .YCoord + 1
                        Else
                              .YCoord = .YCoord + 1
                        End If
                  End If
                  
                  
                  .Left = .XCoord * (BALL_WIDTH + XGAP) + LargeRows(.YCoord).StartX
                  .Top = LargeRows(.YCoord).StartY
                  
                  .Right = .Left + BALL_WIDTH
                  .Bottom = .Top + BALL_HEIGHT
                  
                  Board(.XCoord, .YCoord) = LeftBall
            End With
      
            boolCanRotateBalls = False
            boolProcessing = False
            
            CheckGameStatus
      End If
End Sub
