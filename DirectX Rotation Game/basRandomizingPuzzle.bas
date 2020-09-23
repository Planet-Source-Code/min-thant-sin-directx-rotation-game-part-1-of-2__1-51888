Attribute VB_Name = "basRandomizingPuzzle"
Option Explicit

Public Sub ShuffleBalls()
      Dim I As Integer
      Dim RndNum As Integer
      Dim tmpColor As Integer
      
      Randomize Timer
            
      'Transpose only the color.
      For I = 0 To (TOTAL_BALLS - 1)
            If Balls(I).Exists Then
            
                  RndNum = Int(Rnd * TOTAL_BALLS)
                  
                  If Balls(RndNum).Exists Then
                        tmpColor = Balls(RndNum).Color
                        Balls(RndNum).Color = Balls(I).Color
                        Balls(I).Color = tmpColor
                                          
                        'Update data
                        With Balls(RndNum)
                              .HomeLeft = .Color * BALL_WIDTH
                        End With
                        
                        With Balls(I)
                              .HomeLeft = .Color * BALL_WIDTH
                        End With
                        
                  End If
            End If
      Next I
End Sub
