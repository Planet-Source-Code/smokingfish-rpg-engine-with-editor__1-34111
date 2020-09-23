Attribute VB_Name = "mdlGLOBAL"
'++++++++++++++++
'+RPG Engine... +
'+2002 by       +
'+SmokingFish   ++++++
'+mail@smokingfish.de+
'+++++++++++++++++++++
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As RasterOpConstants) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function TransparentBlt Lib "msimg32" _
                (ByVal hdcDest As Long, ByVal nXOriginDest As Long, _
                  ByVal nYOriginDest As Long, ByVal nWidthDest As Long, _
                  ByVal nHeightDest As Long, ByVal hdcSrc As Long, _
                  ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, _
                 ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
                 ByVal crTransparent As Long) As Long
'---------------------------------
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const MERGEPAINT = &HBB0226
Public Const SRCCOPY = &HCC0020
'---------------------------------
Public RPG As New ctlGame
Public SUBS As New ctlSubs
Public Frame As Integer
Public GG As String
Public Hoch, Runter, Links, Rechts, Mini As Boolean
Global ScreenWidth As Integer, ScreenHeight As Integer
Public FLAG(0 To 99) As String
Public Fscreen As Boolean
