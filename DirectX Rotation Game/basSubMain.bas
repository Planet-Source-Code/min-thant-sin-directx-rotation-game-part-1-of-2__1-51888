Attribute VB_Name = "basSubMain"
Option Explicit

Public Sub Main()
      'Circle graphic
      With CircleRect
            .Left = 0
            .Top = 0
            .Right = .Left + 52
            .Bottom = .Top + 52
      End With
            
      'Custom cursor graphic
      With ArrowRect
            .Left = 0
            .Top = 0
            .Right = .Left + 32
            .Bottom = .Top + 32
      End With
      
      'For blitting background.
      With BackgroundRect
            .Left = 0
            .Top = 0
            .Right = .Left + SCREEN_WIDTH
            .Bottom = .Top + SCREEN_HEIGHT
      End With
      
      Load frmMain
      
      'Set (File1) file list box's path
      frmMain.File1.Path = App.Path & "\Levels"
      NumGameLevelFiles = frmMain.File1.ListCount
      
      If NumGameLevelFiles = 0 Then
            MsgBox "No Rotation Game Level files were found!", vbInformation, "Sorry, couldn't proceed without them..."
            Unload frmMain
            Exit Sub
      End If
      
      GameFileIndex = 0
      Call InitializeRowsData
      
      ShowCursor False
      
      InitializeDirectInput
      InitializeDirectDraw
      StartGame
End Sub
