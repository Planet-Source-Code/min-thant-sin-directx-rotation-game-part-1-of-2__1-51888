Attribute VB_Name = "basMiscellaneous"
Option Explicit

'Custom cursor
Public Sub DisplayArrow()
      ddsBack.BltFast ArrowX, ArrowY, _
      ddsArrow, ArrowRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Sub

'White circle
Public Sub DisplayCircle()
      ddsBack.BltFast CircleX, CircleY, _
      ddsCircle, CircleRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Sub

'Background
Public Sub DisplayBackground()
      ddsBack.BltFast 0, 0, ddsBackground, BackgroundRect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
End Sub

'Retrieving ball information
Public Function GetBallFromCoord(ByVal XCoord As Integer, ByVal YCoord As Integer) As Integer
      GetBallFromCoord = -1
      
      If XCoord >= 0 And XCoord <= PUZZLE_WIDTH - 1 Then
            If YCoord >= 0 And YCoord <= PUZZLE_HEIGHT - 1 Then
                  GetBallFromCoord = Board(XCoord, YCoord)
            End If
      End If
End Function

Public Sub EndGame()
      Set ddsFront = Nothing
      Set ddsBack = Nothing
      Set ddsBalls = Nothing
      Set ddsPreviewBalls = Nothing
      Set ddsBackground = Nothing
      Set ddsPuzzleSolved = Nothing
      
      dd.RestoreDisplayMode
      dd.SetCooperativeLevel 0, DDSCL_NORMAL
    
      Set dd = Nothing
      
      ShowCursor True
      
      End
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////
'/// This function is grabbed from DXSDK.
'/// Using this function, I can create DirectDraw surfaces from *jpg files.
'/// Background.bmp file is too large, so I use Background.jpg
'///////////////////////////////////////////////////////////////////////////////////////////
Public Function CreateDDSFromBitmap(dd As DirectDraw7, ByVal strFile As String, Optional VideoMem As Boolean) As DirectDrawSurface7
      'This function creates a direct draw surface from any valid file format that loadpicture uses, and returns
      'the newly created surface
      
      If Dir(strFile) = "" Then                                               'If the file doesn't exist
            Debug.Print "File not found! ** CreateDDSFromBitmap **"
            Exit Function
      End If
      
      Dim ddsd As DDSURFACEDESC2                                              'Surface description
      Dim dds As DirectDrawSurface7                                           'Created surface
      Dim hdcPicture As Long                                                  'Device context for picture
      Dim hdcSurface As Long                                                  'Device context for surface
      Dim Picture As StdPicture                                               'stdole2 StdPicture object
      
      Set Picture = LoadPicture(strFile)                                      'Load the bitmap
      
      With ddsd                                                               'Fill the surface description
            .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH                    'Tell Direct Draw that the caps element is valid, the height element is valid, and the width element is valid
            If VideoMem Then                                                    'If the videomem flag is set, then
                  .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY  'Create the surface in video memory
            Else
                  .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY 'Otherwise, creat the surface in system memory
            End If
            
            .lWidth = Screen.ActiveForm.ScaleX(Picture.Width, vbHimetric, vbPixels)
            'The width of the surface is set by scaling from the stdpicture objects vbhimetric scale mode to pixels
            .lHeight = Screen.ActiveForm.ScaleY(Picture.Height, vbHimetric, vbPixels)
            'The height of the surface is set by scaling from the stdpicture objects vbhimetric scale mode to pixels
      End With
      
      Set dds = dd.CreateSurface(ddsd)                                        'Create the surface
      
      hdcPicture = CreateCompatibleDC(ByVal 0&)                               'Create a memory device context
      
      SelectObject hdcPicture, Picture.Handle                                 'Select the bitmap into this memory device
      
      dds.restore                                                             'Restore the surface
      
      hdcSurface = dds.GetDC                                                  'Get the surface's DC
      
      StretchBlt hdcSurface, 0, 0, ddsd.lWidth, ddsd.lHeight, hdcPicture, 0, 0, Screen.ActiveForm.ScaleX(Picture.Width, vbHimetric, vbPixels), Screen.ActiveForm.ScaleY(Picture.Height, vbHimetric, vbPixels), vbSrcCopy
      'Copy from the memory device to the DirectDrawSurface
      
      dds.ReleaseDC hdcSurface                                                'Release the surface's DC
      
      DeleteDC hdcPicture                                                     'Release the memory device context
      
      Set Picture = Nothing                                                   'Release the picture object
      Set CreateDDSFromBitmap = dds                                           'Sets the function to the newly created direct draw surface
      
End Function
