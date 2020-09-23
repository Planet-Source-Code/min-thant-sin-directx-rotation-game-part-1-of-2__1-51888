Attribute VB_Name = "basDeclarations"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////
'DirectX
Public dx As New DirectX7

'DirectDraw
Public dd As DirectDraw7
Public ddsFront As DirectDrawSurface7               'Front surface
Public ddsBack As DirectDrawSurface7                 'Back buffer
Public ddsBalls As DirectDrawSurface7                 'Balls surface
Public ddsPreviewBalls As DirectDrawSurface7     'Small preview balls surface
Public ddsCircle As DirectDrawSurface7                  'White circle surface
Public ddsArrow As DirectDrawSurface7               'Custom cursor surface
Public ddsPuzzleSolved As DirectDrawSurface7    'When the puzzle is solved, PuzzleSolved.bmp is blitted.
Public ddsBackground As DirectDrawSurface7       'Background surface

'Direct Input
Public DI As DirectInput
Public diKeyBoard As DirectInputDevice

'////////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////////////
'My Constants
Public Const SCREEN_WIDTH As Integer = 1024
Public Const SCREEN_HEIGHT As Integer = 768
Public Const DISPLAY_DEPTH As Integer = 16

Public Const PUZZLE_WIDTH As Integer = 9        'Nine columns and...
Public Const PUZZLE_HEIGHT As Integer = 9       'nine rows

Public Const XGAP As Integer = 6    'Gap between two large balls

Public Const BALL_WIDTH As Integer = 52     'The width of large ball
Public Const BALL_HEIGHT As Integer = 52    'The height of large ball

Public Const SMALLBALL_WIDTH As Integer = 27    'Small ball's width
Public Const SMALLBALL_HEIGHT As Integer = 27   'Small ball's height

Public Const MIDDLE_ROW As Integer = 4

'The speed at which the balls move when rotating.
'SLOW_SPEED is used when the ball must move a longer distance.
'FAST_SPEED is used when the ball must move a shorter distance.
Public Const SLOW_SPEED As Single = 2     '2.2
Public Const FAST_SPEED As Single = 4       '4.2


'How many balls?
Public Const TOTAL_BALLS As Integer = (PUZZLE_WIDTH * PUZZLE_HEIGHT)


'////////////////////////////////////////////////////////////////////////////////////

'Types
Public Type POINTAPI
      X As Long
      Y As Long
End Type

Public Type LEVEL_INFO
      BallX As Integer
      BallY As Integer
      BallColor As Integer
End Type

Public Type ROW_INFO
      StartX As Integer
      StartY As Integer
End Type

Public Type BALL_INFO
      Index As Integer              'Ball's index
      Color As Integer                'Ball's color (green, red, yellow, and blue) in order.
      HomeLeft As Integer         'HomeLeft is determined in this way >> HomeLeft = Color * BALL_WIDTH
      HomeTop As Integer          'HomeTop will always be 0 in this game because there's only one row in the graphic LargeBalls.bmp
      XCoord As Integer             'X coordinate on the game board
      YCoord As Integer             'Y coordinate on the game board
      Left As Single                    'Left position on the screen
      Top As Single                     'Top position on the screen
      Right As Single
      Bottom As Single
      Exists As Boolean               'A flag to indicate whether the ball should be displayed or not
      'Rotating As Boolean            'Not used in this version.
End Type
'////////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////////////
'My variables
Public BallClicked As Integer       'The ball that is clicked by the player
Public BallHovered As Integer      'The ball on which the mouse is over.

Public ArrowX As Integer      'Custom cursor X position
Public ArrowY As Integer      'Custom cursor Y position
Public CircleX As Integer         'X position ofwhite circle that is displayed around the ball
Public CircleY As Integer         'Y position of white circle that is displayed around the ball

'RotateStep is calculated like this >>  BALL_WIDTH / FAST_SPEED
Public RotateStep As Integer

Public GameFileIndex As Integer           'Index in the File1 file list box
Public NumGameLevelFiles As Integer    'File1.ListCount

'When the player presses Esc key, this flag is set to False.
'In the main game loop StartGame(), this flag is tested to see if it contains
'False value, if it does, the loop ends and EndGame() is called.
Public boolEndGame As Boolean

'A flag to indicate if the balls can be rotated or not.
Public boolCanRotateBalls As Boolean

'A flag to indicate if the current puzzle has been solved or not.
Public boolPuzzleSolved As Boolean

'A flag to indicate if the Form_MouseDown() is
'still doing its job or the balls are still rotating.
Public boolProcessing As Boolean

'A flag to indicate whether to rotate the balls clockwise or counter-clockwise.
Public boolRotateLeft As Boolean

'A flag to indicate if the white circle should be shown
'when the mouse is moved over one of the balls.
Public boolCanShowCircle As Boolean
Public boolGameStarted As Boolean

Public CircleRect As RECT
Public ArrowRect As RECT
Public BackgroundRect As RECT

Public LargeRows(0 To 8) As ROW_INFO
Public SmallRows(0 To 8) As ROW_INFO

'Large balls info
Public Balls(0 To TOTAL_BALLS - 1) As BALL_INFO

'Small preview balls in the small board
Public PreviewBalls(0 To TOTAL_BALLS - 1) As BALL_INFO

'Board coordinates to hold ball's index
Public Board(0 To PUZZLE_WIDTH - 1, 0 To PUZZLE_HEIGHT - 1) As Integer

'To load a game level file
Public LevelData(0 To TOTAL_BALLS - 1) As LEVEL_INFO

'////////////////////////////////////////////////////////////////////////////////////

'APIs
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
