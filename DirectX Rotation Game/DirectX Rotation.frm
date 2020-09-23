VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   27.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "DirectX Rotation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   225
      Pattern         =   "*.rgl"
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTileMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   525
      Picture         =   "DirectX Rotation.frx":000C
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   0
      Top             =   300
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   $"DirectX Rotation.frx":1F62
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Index           =   1
      Left            =   1575
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   6915
   End
   Begin VB.Label Label1 
      Caption         =   $"DirectX Rotation.frx":20FF
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Index           =   0
      Left            =   1500
      TabIndex        =   2
      Top             =   225
      Visible         =   0   'False
      Width           =   6915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////////////////////////////////////////////////////
'/// This game is originally created by Diapraxas team.
'/// You can visit their website and register the game at :
'/// < http://www.diapraxas.net/rotation >
'/// Registration fee is $10.00 (ten dollars) currently.
'/// I have included their ReadMe.txt for your reference.
'////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////////
'/// Rotation Game by Min Thant Sin on Friday, February 20, 2004
'/// Comments and suggestions are welcome.
'/// Any bugs? Feel free to e-mail me at < minsin999@hotmail.com >
'////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////////
'/// NOTE : Error checkings and coding are done cursorily.
'///              Make sure you have the right graphics files, game level files
'///              in the right directory.
'////////////////////////////////////////////////////////////////////////////////

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      On Error GoTo ErrorHandler
      
      If Not boolGameStarted Then Exit Sub
      If boolPuzzleSolved Then Exit Sub
      If boolProcessing Then Exit Sub
      
      boolProcessing = True
      
      boolRotateLeft = False
      If Button = vbLeftButton Then boolRotateLeft = True
      
      Dim I As Integer
      Dim MaskX As Integer, MaskY As Integer
      Dim MouseX As Integer, MouseY As Integer
      Dim pt As POINTAPI
      
      boolCanRotateBalls = False
      
      'Get cursor position
      GetCursorPos pt
      MouseX = pt.X
      MouseY = pt.Y
      
      BallClicked = -1
      
      'Determine which ball the player has clicked
      For I = 0 To TOTAL_BALLS - 1
            If MouseX >= Balls(I).Left And MouseX <= Balls(I).Right Then
                  If MouseY >= Balls(I).Top And MouseY <= Balls(I).Bottom Then
                        If Balls(I).Exists Then
                              MaskX = (MouseX - Balls(I).Left)
                              MaskY = (MouseY - Balls(I).Top)
                              
                              'Determine if the player has clicked in the circle region of the ball.
                              If GetPixel(picTileMask.hdc, MaskX, MaskY) = vbBlack Then
                                    BallClicked = I
                                    Exit For
                              End If
                        End If
                  End If
            End If
      Next I
      
      If BallClicked = -1 Then
            boolProcessing = False
            Exit Sub
      End If
      
      boolCanRotateBalls = True
      Exit Sub
ErrorHandler:
      Debug.Print "Error occurred in Form_MouseDown : " & Err.Description
      EndGame
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyEscape Then
            boolEndGame = True
      End If
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'/// This sub shows the white circle around a ball.
'/// This sub also updates the custom cursor's position.
'////////////////////////////////////////////////////////////////////////////////////////
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      On Error GoTo ErrorHandler
      
      Dim I As Integer
      Dim MaskX As Integer, MaskY As Integer
      Dim MouseX As Integer, MouseY As Integer
      Dim pt As POINTAPI
      
      GetCursorPos pt
      ArrowX = pt.X
      ArrowY = pt.Y
      
      If Not boolGameStarted Then Exit Sub
      If boolPuzzleSolved Then Exit Sub
      
      MouseX = pt.X
      MouseY = pt.Y
      
      BallHovered = -1
      
      For I = 0 To (TOTAL_BALLS - 1)
            If MouseX >= Balls(I).Left And MouseX <= Balls(I).Right Then
                  If MouseY >= Balls(I).Top And MouseY <= Balls(I).Bottom Then
                        If Balls(I).Exists Then
                              MaskX = (MouseX - Balls(I).Left)
                              MaskY = (MouseY - Balls(I).Top)
                              
                              'Determine if the mouse x & y is in the circle region of the ball.
                              If GetPixel(picTileMask.hdc, MaskX, MaskY) = vbBlack Then
                                    BallHovered = I
                                    Exit For
                              End If
                        End If
                  End If
            End If
      Next I
                  
      boolCanShowCircle = False
      
      If BallHovered = -1 Then Exit Sub
            
      CircleX = Balls(BallHovered).Left
      CircleY = Balls(BallHovered).Top
      
      'Indicate that the white circle can be shown
      boolCanShowCircle = True
            
      Exit Sub
ErrorHandler:
      Debug.Print "Error occurred in Form_MouseDown : " & Err.Description
      EndGame
End Sub

'Just to make sure the cursor is shown again if the program crashes.
Private Sub Form_Terminate()
      ShowCursor True
End Sub

Private Sub Form_Unload(Cancel As Integer)
      ShowCursor True
End Sub
