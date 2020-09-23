Attribute VB_Name = "basDirectInput"
Option Explicit

Public Sub InitializeDirectInput()
      Set DI = dx.DirectInputCreate()
      
      Set diKeyBoard = DI.CreateDevice("GUID_SysKeyboard")
      diKeyBoard.SetCommonDataFormat DIFORMAT_KEYBOARD
      diKeyBoard.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
      diKeyBoard.Acquire
End Sub

Public Sub GetInput()
      'Byte array to hold the state of the keyboard
      Dim KeyboardState(0 To 255) As Byte
    
      diKeyBoard.Acquire
      diKeyBoard.GetDeviceState 256, KeyboardState(0)
            
      If (KeyboardState(DIK_ESCAPE) And &H80) <> 0 Then
            EndGame
      End If
      
      If (KeyboardState(DIK_RIGHT) And &H80) <> 0 Then
            LoadGame True
      End If
      
      If (KeyboardState(DIK_LEFT) And &H80) <> 0 Then
            LoadGame False
      End If
      
End Sub
