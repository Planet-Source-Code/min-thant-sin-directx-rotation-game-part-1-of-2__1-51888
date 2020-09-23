Attribute VB_Name = "basDirectDraw"
Option Explicit

Public Sub InitializeDirectDraw()
      On Error GoTo ErrorHandler
    
      Dim ddsdFront As DDSURFACEDESC2
      Dim ddsCaps As DDSCAPS2
      Dim ddsdBalls As DDSURFACEDESC2
      Dim ddsdPreviewBalls As DDSURFACEDESC2
      Dim ddsdBackground As DDSURFACEDESC2
      Dim ddsdPuzzleSolved As DDSURFACEDESC2
      Dim ddsdArrow As DDSURFACEDESC2
      Dim ddsdCircle As DDSURFACEDESC2
      Dim ddKey As DDCOLORKEY
      
      Set dd = dx.DirectDrawCreate("")
      dd.SetCooperativeLevel frmMain.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
      dd.SetDisplayMode SCREEN_WIDTH, SCREEN_HEIGHT, DISPLAY_DEPTH, 0, DDSDM_DEFAULT
      
      ddsdFront.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
      ddsdFront.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE
      ddsdFront.lBackBufferCount = 1
      Set ddsFront = dd.CreateSurface(ddsdFront)
      
      'Back buffer
      ddsCaps.lCaps = DDSCAPS_BACKBUFFER
      Set ddsBack = ddsFront.GetAttachedSurface(ddsCaps)
      ddsBack.SetFont frmMain.Font
            
      'Balls, preview balls, custom cursor, circle, background, etc surfaces
      Set ddsBalls = dd.CreateSurfaceFromFile(App.Path & "\Graphics\LargeBalls.bmp", ddsdBalls)
      Set ddsPreviewBalls = dd.CreateSurfaceFromFile(App.Path & "\Graphics\PreviewBalls.bmp", ddsdPreviewBalls)
      Set ddsArrow = dd.CreateSurfaceFromFile(App.Path & "\Graphics\Arrow.bmp", ddsdArrow)
      Set ddsCircle = dd.CreateSurfaceFromFile(App.Path & "\Graphics\Circle.bmp", ddsdCircle)
      Set ddsPuzzleSolved = dd.CreateSurfaceFromFile(App.Path & "\Graphics\PuzzleSolved.bmp", ddsdPuzzleSolved)
      Set ddsBackground = CreateDDSFromBitmap(dd, App.Path & "\Graphics\GameBoard.jpg")
    
    
      Dim ddpf As DDPIXELFORMAT
      Dim KeyColor As Long
      
      ddsFront.GetPixelFormat ddpf
      KeyColor = ddpf.lGBitMask
    
      ddKey.low = KeyColor
      ddKey.high = KeyColor
      
      ddsBalls.SetColorKey DDCKEY_SRCBLT, ddKey
      ddsPreviewBalls.SetColorKey DDCKEY_SRCBLT, ddKey
      ddsArrow.SetColorKey DDCKEY_SRCBLT, ddKey
      ddsCircle.SetColorKey DDCKEY_SRCBLT, ddKey
      ddsPuzzleSolved.SetColorKey DDCKEY_SRCBLT, ddKey
      
      Exit Sub
ErrorHandler:
      Debug.Print Err.Description
      EndGame
End Sub
