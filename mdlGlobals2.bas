Attribute VB_Name = "mdlGlobals"
'++++++++++++++++
'+RPG Engine... +
'+2002 by       +
'+SmokingFish   ++++++
'+mail@smokingfish.de+
'+++++++++++++++++++++
Declare Function TransparentBlt Lib "msimg32" _
                (ByVal hdcDest As Long, ByVal nXOriginDest As Long, _
                  ByVal nYOriginDest As Long, ByVal nWidthDest As Long, _
                  ByVal nHeightDest As Long, ByVal hdcSrc As Long, _
                  ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, _
                 ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
                 ByVal crTransparent As Long) As Long
Global MAP As New clsMAP
Public TEXT2 As String
Public FARBE2 As String
Public TEXT As String
Public FARBE As String
Public Function Snap(Cordinate As Variant, Dimension As Integer) As Integer
Snap = (Cordinate \ Dimension) * Dimension
End Function
Public Function Snap2(Cordinate As String) As String
Snap2 = Cordinate
End Function
