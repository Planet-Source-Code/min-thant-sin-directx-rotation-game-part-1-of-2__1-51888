Attribute VB_Name = "basCreatingPuzzle"
Option Explicit

Public Sub InitializeRowsData()
      '//////////////////////////////////////////////////////////////////////
      '/// Rows initializations, specify each row's start x and start y position.
      '//////////////////////////////////////////////////////////////////////
      
      'Large balls' rows
      LargeRows(0).StartX = 291:         LargeRows(0).StartY = 175
      LargeRows(1).StartX = 263:         LargeRows(1).StartY = 223
      LargeRows(2).StartX = 234:         LargeRows(2).StartY = 270
      LargeRows(3).StartX = 205:         LargeRows(3).StartY = 319
      LargeRows(4).StartX = 175:         LargeRows(4).StartY = 366
      LargeRows(5).StartX = 205:         LargeRows(5).StartY = 415
      LargeRows(6).StartX = 234:         LargeRows(6).StartY = 463
      LargeRows(7).StartX = 263:         LargeRows(7).StartY = 511
      LargeRows(8).StartX = 291:         LargeRows(8).StartY = 559
      
      'Small balls' rows
      SmallRows(0).StartX = 786:        SmallRows(0).StartY = 140
      SmallRows(1).StartX = 770:        SmallRows(1).StartY = 164
      SmallRows(2).StartX = 757:        SmallRows(2).StartY = 187
      SmallRows(3).StartX = 741:        SmallRows(3).StartY = 211
      SmallRows(4).StartX = 728:        SmallRows(4).StartY = 236
      SmallRows(5).StartX = 741:        SmallRows(5).StartY = 259
      SmallRows(6).StartX = 757:        SmallRows(6).StartY = 283
      SmallRows(7).StartX = 770:        SmallRows(7).StartY = 308
      SmallRows(8).StartX = 786:        SmallRows(8).StartY = 332
End Sub

Public Sub CreatePuzzle(ByVal FileIndex As Integer)
      Dim I As Integer
      Dim Color As Integer                          'Ball's color
      Dim Index As Integer                         'Ball's index
      Dim Row As Integer, Col As Integer   'Ball coordinates
      Dim FilePath As String                        'Game level file path
            
      Index = 0
      
      For Row = 0 To (PUZZLE_HEIGHT - 1)
            For Col = 0 To (PUZZLE_WIDTH - 1)
                  Board(Col, Row) = Index
                  
                  With Balls(Index)
                        .Index = Index
                        .Exists = True
                        
                        .XCoord = Col
                        .YCoord = Row
                        
                        .Left = Col * (BALL_WIDTH + XGAP) + LargeRows(Row).StartX
                        .Top = LargeRows(Row).StartY
                        
                        .Right = .Left + BALL_WIDTH
                        .Bottom = .Top + BALL_HEIGHT
                  End With
                                    
                  With PreviewBalls(Index)
                        .Index = Index
                        .Exists = True
                        .XCoord = Col
                        .YCoord = Row
                        
                        .Left = Col * (SMALLBALL_WIDTH + 2) + SmallRows(Row).StartX
                        .Top = SmallRows(Row).StartY
                        
                        .Right = .Left + SMALLBALL_WIDTH
                        .Bottom = .Top + SMALLBALL_WIDTH
                  End With
                  
                  Index = Index + 1
            Next Col
      Next Row
      
      '//////////////////////////////////////////////////////////////////////////////////
      '/// Setting some of upper and lower balls' Exists flag to False.
      '/// The flag is set to False to "hide" the balls.
      '/// When blitting the balls, this Exists flag is checked to determine
      '/// whether to blit the ball or not. If the flag is True, the ball is blitted, and
      '/// if the flag is False, as you may have guessed by now, the ball is NOT blitted.
      '///////////////////////////////////////////////////////////////////////////////////
      
      'Upper balls
      For Row = 0 To 3
            For Col = (5 + Row) To 8
                  Index = Board(Col, Row)
                  Balls(Index).Exists = False
            Next Col
      Next Row
      
      'Lower balls                                      'The same loop.
      For Row = 8 To 5 Step -1                  'For Row = 5 To 8
            For Col = (5 + (8 - Row)) To 8            'For Col = (8 - (Row - 5)) To 8
                  Index = Board(Col, Row)
                  Balls(Index).Exists = False
            Next Col
      Next Row
      
      'For small preview balls
      For I = 0 To (TOTAL_BALLS - 1)
            PreviewBalls(I).Exists = Balls(I).Exists
      Next I
      
      '///////////////////////////////////////////////////////////////////////////////////
      '/// Update the Board data, to indicate that the balls having these
      '/// coordinates don't exist (their Exists flag are set to False).
      '///////////////////////////////////////////////////////////////////////////////////
      
      'Upper board coordinates
      For Row = 0 To 3
            For Col = (5 + Row) To 8
                  Board(Col, Row) = -1
            Next Col
      Next Row
      
      'Lower board coordinates
      For Row = 8 To 5 Step -1
            For Col = (5 + (8 - Row)) To 8
                  Board(Col, Row) = -1
            Next Col
      Next Row
      
      
      '///////////////////////////////////////////////////////////////////////////////
      '/// The following code opens a game level file, load the file into LevelData() structure,
      '/// and then update some of the balls' data with the data in LevelData()
      '///////////////////////////////////////////////////////////////////////////////
      
      'Get game level file path
      FilePath = AddSlash(frmMain.File1.Path) & frmMain.File1.List(FileIndex)
      
      'Read a level file and load it into LevelData()
      Open FilePath For Input As #1
      
            I = 0
            Do Until EOF(1)
                  Input #1, Col, Row, Color
                  
                  With LevelData(I)
                        .BallX = Col
                        .BallY = Row
                        .BallColor = Color
                  End With
                  
                  I = I + 1
                  DoEvents
            Loop
                  
      Close #1
      
      'Assuming data loaded into LevelData() correctly and successfully...
      For I = 0 To UBound(LevelData())
            Index = GetBallFromCoord(LevelData(I).BallX, LevelData(I).BallY)
            
            'If Index is -1, that means the ball at the coordinates BallX & BallY
            'does NOT exist (ie. Exists flag set to False)
            If Index <> -1 Then
            
                  'Update some of the balls' data...
                  With Balls(Index)
                        .Color = LevelData(I).BallColor
                        .HomeLeft = .Color * BALL_WIDTH
                        .HomeTop = 0
                  End With
                  
                  With PreviewBalls(Index)
                        .Color = LevelData(I).BallColor
                        .HomeLeft = .Color * SMALLBALL_WIDTH
                        .HomeTop = 0
                  End With
            End If
      Next I
      
      'Transpose the colors of the balls
      ShuffleBalls
End Sub

'////////////////////////////////////////////////////////////////////////////
'/// Add a "\" after the path if one does not exist.
'////////////////////////////////////////////////////////////////////////////
Public Function AddSlash(ByRef Path As String) As String
      AddSlash = Path
      If Right(Path, 1) <> "\" Then
            AddSlash = Path & "\"
      End If
End Function
