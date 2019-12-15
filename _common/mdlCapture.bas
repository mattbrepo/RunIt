Attribute VB_Name = "mdlCapture"
Option Explicit
Option Base 0

''
'
Public Function CaptureASliceScreen(ByVal lX As Long, ByVal lY As Long, _
                                    ByVal lW As Long, ByVal lH As Long, _
                                    ByVal lWDst As Long, ByVal lHDst As Long) As Picture
Dim hWndScreen As Long

   hWndScreen = GetDesktopWindow()

   Set CaptureASliceScreen = CaptureWindow2(hWndScreen, True, _
                                    lX \ Screen.TwipsPerPixelX, lY \ Screen.TwipsPerPixelY, _
                                    lW \ Screen.TwipsPerPixelX, lH \ Screen.TwipsPerPixelY, _
                                    lWDst \ Screen.TwipsPerPixelX, lHDst \ Screen.TwipsPerPixelY)
End Function

''
'
Public Function CaptureScreen() As Picture
Dim hWndScreen As Long

   ' Get a handle to the desktop window
   hWndScreen = GetDesktopWindow()

   ' Call CaptureWindow to capture the entire desktop give the handle
   ' and return the resulting Picture object

   Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, _
      Screen.Width \ Screen.TwipsPerPixelX, _
      Screen.Height \ Screen.TwipsPerPixelY)
End Function

'###################################################################################

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CreateBitmapPicture
'    - Creates a bitmap type Picture object from a bitmap and palette
'
' hBmp
'    - Handle to a bitmap
'
' hPal
'    - Handle to a Palette
'    - Can be null if the bitmap doesn't use a palette
'
' Returns
'    - Returns a Picture object containing the bitmap
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
Dim r As Long
Dim Pic As PICTDESC
' IPicture requires a reference to "Standard OLE Types"
Dim ipic As IPicture
Dim IID_IDispatch As GUID

   ' Fill in with IDispatch Interface ID
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With

   ' Fill Pic with necessary parts
   With Pic
      .cbSize = Len(Pic)          ' Length of structure
      .pictType = vbPicTypeBitmap   ' Type of Picture (bitmap)
      .hIcon = hBmp              ' Handle to bitmap
      .hPal = hPal              ' Handle to palette (may be null)
   End With

   ' Create Picture object
   r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, ipic)

   ' Return the new Picture object
   Set CreateBitmapPicture = ipic
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureForm
'    - Captures an entire form including title bar and border
'
' frmSrc
'    - The Form object to capture
'
' Returns
'    - Returns a Picture object containing a bitmap of the entire
'      form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Function CaptureForm(frmSrc As Form) As Picture
   ' Call CaptureWindow to capture the entire form given it's window
   ' handle and then return the resulting Picture object
   Set CaptureForm = CaptureWindow(frmSrc.hwnd, False, 0, 0, _
      frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), _
      frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureClient
'    - Captures the client area of a form
'
' frmSrc
'    - The Form object to capture
'
' Returns
'    - Returns a Picture object containing a bitmap of the form's
' client area
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Function CaptureClient(frmSrc As Form) As Picture
   ' Call CaptureWindow to capture the client area of the form given
   ' it's window handle and return the resulting Picture object
   Set CaptureClient = CaptureWindow(frmSrc.hwnd, True, 0, 0, _
      frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), _
      frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureActiveWindow
'    - Captures the currently active window on the screen
'
' Returns
'    - Returns a Picture object containing a bitmap of the active
'      window
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Function CaptureActiveWindow() As Picture
Dim hWndActive As Long
Dim r As Long
Dim RectActive As RECT

    ' Get a handle to the active/foreground window
    hWndActive = GetForegroundWindow()

    ' Get the dimensions of the window
    r = GetWindowRect(hWndActive, RectActive)

    ' Call CaptureWindow to capture the active window given it's
    ' handle and return the Resulting Picture object
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, _
        RectActive.Right - RectActive.Left, _
        RectActive.Bottom - RectActive.Top)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PrintPictureToFitPage
'    - Prints a Picture object as big as possible
'
' Prn
'    - Destination Printer object
'
' Pic
'    - Source Picture object
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
   Const vbHiMetric As Integer = 8
   Dim PicRatio As Double
   Dim PrnWidth As Double
   Dim PrnHeight As Double
   Dim PrnRatio As Double
   Dim PrnPicWidth As Double
   Dim PrnPicHeight As Double

   ' Determine if picture should be printed in landscape or portrait
   ' and set the orientation
   If Pic.Height >= Pic.Width Then
      Prn.Orientation = vbPRORPortrait   ' Taller than wide
   Else
      Prn.Orientation = vbPRORLandscape  ' Wider than tall
   End If

   ' Calculate device independent Width to Height ratio for picture
   PicRatio = Pic.Width / Pic.Height

   ' Calculate the dimentions of the printable area in HiMetric
   PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
   PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
   ' Calculate device independent Width to Height ratio for printer
   PrnRatio = PrnWidth / PrnHeight

   ' Scale the output to the printable area
   If PicRatio >= PrnRatio Then
      ' Scale picture to fit full width of printable area
      PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
      PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, _
         Prn.ScaleMode)
   Else
      ' Scale picture to fit full height of printable area
      PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
      PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, _
         Prn.ScaleMode)
   End If

   ' Print the picture using the PaintPicture method
   Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
End Sub

Private Function CaptureWindow2(ByVal hWndSrc As Long, ByVal Client As Boolean, _
                               ByVal LeftSrc As Long, ByVal TopSrc As Long, _
                               ByVal WidthSrc As Long, ByVal HeightSrc As Long, _
                               ByVal lWDst As Long, ByVal lHDst As Long) As Picture
Dim hDCMemory As Long
Dim hBmp As Long
Dim hBmpPrev As Long
Dim r As Long
Dim hDCSrc As Long
Dim hPal As Long
Dim hPalPrev As Long
Dim RasterCapsScrn As Long
Dim HasPaletteScrn As Long
Dim PaletteSizeScrn As Long

   Dim LogPal As LOGPALETTE

   ' Depending on the value of Client get the proper device context
   If Client Then
      hDCSrc = GetDC(hWndSrc) ' Get device context for client area
   Else
      hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire window
   End If

   ' Create a memory device context for the copy process
   hDCMemory = CreateCompatibleDC(hDCSrc)
   
   ' Create a bitmap and place it in the memory DC
   hBmp = CreateCompatibleBitmap(hDCSrc, lWDst, lHDst) 'WidthSrc, HeightSrc)
   'hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   'hBmp = CreateCompatibleBitmap(hDCMemory, WidthSrc, HeightSrc)  'WidthSrc, HeightSrc)
   
   hBmpPrev = SelectObject(hDCMemory, hBmp)

   ' Get screen properties
   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster capabilities
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette support
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of palette

   ' If the screen has a palette make a copy and realize it
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      ' Create a copy of the system palette
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, _
          LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      ' Select the new palette into the memory DC and realize it
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
   End If

    r = StretchBlt(hDCMemory, 0, 0, lWDst, lHDst, hDCSrc, LeftSrc, TopSrc, WidthSrc, HeightSrc, vbSrcCopy)

    ' Remove the new copy of the  on-screen image
    hBmp = SelectObject(hDCMemory, hBmpPrev)

   ' If the screen has a palette get back the palette that was
   ' selected in previously
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If

   ' Release the device context resources back to the system
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)

   ' Call CreateBitmapPicture to create a picture object from the
   ' bitmap and palette handles.  Then return the resulting picture
   ' object.
   Set CaptureWindow2 = CreateBitmapPicture(hBmp, hPal)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureWindow
'    - Captures any portion of a window
'
' hWndSrc
'    - Handle to the window to be captured
'
' Client
'    - If True CaptureWindow captures from the client area of the
'      window
'    - If False CaptureWindow captures from the entire window
'
' LeftSrc, TopSrc, WidthSrc, HeightSrc
'    - Specify the portion of the window to capture
'    - Dimensions need to be specified in pixels
'
' Returns
'    - Returns a Picture object containing a bitmap of the specified
'      portion of the window that was captured
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''
Private Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, _
                            ByVal LeftSrc As Long, ByVal TopSrc As Long, _
                            ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
Dim hDCMemory As Long
Dim hBmp As Long
Dim hBmpPrev As Long
Dim r As Long
Dim hDCSrc As Long
Dim hPal As Long
Dim hPalPrev As Long
Dim RasterCapsScrn As Long
Dim HasPaletteScrn As Long
Dim PaletteSizeScrn As Long

   Dim LogPal As LOGPALETTE

   ' Depending on the value of Client get the proper device context
   If Client Then
      hDCSrc = GetDC(hWndSrc) ' Get device context for client area
   Else
      hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                    ' window
   End If

   ' Create a memory device context for the copy process
   hDCMemory = CreateCompatibleDC(hDCSrc)
   ' Create a bitmap and place it in the memory DC
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)

   ' Get screen properties
   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                      'capabilities
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                        'support
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                        ' palette

   ' If the screen has a palette make a copy and realize it
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      ' Create a copy of the system palette
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, _
          LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      ' Select the new palette into the memory DC and realize it
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
   End If

   ' Copy the on-screen image into the memory DC
   r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
      LeftSrc, TopSrc, vbSrcCopy)

' Remove the new copy of the  on-screen image
   hBmp = SelectObject(hDCMemory, hBmpPrev)

   ' If the screen has a palette get back the palette that was
   ' selected in previously
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If

   ' Release the device context resources back to the system
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)

   ' Call CreateBitmapPicture to create a picture object from the
   ' bitmap and palette handles.  Then return the resulting picture
   ' object.
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
