Attribute VB_Name = "mdlChange"
Option Explicit

Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32

Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1
Private Const DISP_CHANGE_FAILED = -1
Private Const DISP_CHANGE_BADMODE = -2
Private Const DISP_CHANGE_NOTUPDATED = -3
Private Const DISP_CHANGE_BADFLAGS = -4
Private Const DISP_CHANGE_BADPARAM = -5

Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H2

Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000

Private Const BITSPIXEL = 12
Private Const PLANES = 14



Private Type DEVMODE
  dmDeviceName As String * CCDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type

Private Declare Function CreateIC Lib "gdi32" _
Alias "CreateICA" (ByVal lpDriverName As String, _
ByVal lpDeviceName As Any, ByVal lpOutput As Any, ByVal lpInitData As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, _
ByVal nIndex As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
    
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long

Function ChangeScreenSettings(lWidth As Integer, lHeight As Integer, lColors As Integer)
Dim tDevMode As DEVMODE, lTemp As Long, lIndex As Long
lIndex = 0
  lTemp = EnumDisplaySettings(0&, lIndex, tDevMode)
  With tDevMode
  .dmPelsWidth = lWidth
  .dmPelsHeight = lHeight
  .dmBitsPerPel = lColors
    
      lTemp = ChangeDisplaySettings(tDevMode, CDS_UPDATEREGISTRY)
   
  End With
Select Case lTemp
  Case DISP_CHANGE_SUCCESSFUL
    
  Case DISP_CHANGE_RESTART
    MsgBox "The computer must be restarted in order for the graphics mode to work", vbQuestion
  Case DISP_CHANGE_FAILED
    MsgBox "The display driver failed the specified graphics mode", vbCritical
  Case DISP_CHANGE_BADMODE
    MsgBox "The graphics mode is not supported", vbCritical
  Case DISP_CHANGE_NOTUPDATED
    MsgBox "Unable to write settings to the registry", vbCritical
  Case DISP_CHANGE_BADFLAGS
    MsgBox "An invalid set of flags was passed in", vbCritical
End Select
End Function

Function ChangeScreenSettingsback(lWidth As Integer, lHeight As Integer, lColors As Integer)
Dim tDevMode As DEVMODE, lTemp As Long, lIndex
lTemp = EnumDisplaySettings(0&, lIndex, tDevMode)
With tDevMode
  .dmPelsWidth = lWidth
  .dmPelsHeight = lHeight
  .dmBitsPerPel = lColors
      lTemp = ChangeDisplaySettings(tDevMode, CDS_UPDATEREGISTRY)
 End With
End Function

Public Function GetScreenResolution(xres, yres As String) As Boolean
xres = GetSystemMetrics(SM_CXSCREEN)
yres = GetSystemMetrics(SM_CYSCREEN)
End Function
Public Function GetAvailableColours(rColor) As String
Dim lHdc As Long, lPlanes As Long, lBitsPerPixel As Integer
lHdc = CreateIC("DISPLAY", 0&, 0&, 0&)
If lHdc = 0 Then
  GetAvailableColours = "Error"
  Exit Function
End If
lPlanes = GetDeviceCaps(lHdc, PLANES)
lBitsPerPixel = GetDeviceCaps(lHdc, BITSPIXEL)
rColor = lBitsPerPixel
lHdc = DeleteDC(lHdc)
Select Case lPlanes
  Case 1
    Select Case lBitsPerPixel
      Case 4: GetAvailableColours = "4 Bit, 16 Colours"
      Case 8: GetAvailableColours = "8 Bit, 256 Colours"
      Case 16: GetAvailableColours = "16 Bit, 65536 Colours"
      Case 24: GetAvailableColours = "24 Bit True Colour, 16.7 Million Colours"
      Case 32: GetAvailableColours = "32 Bit True Colour, 16.7 Million Colours"
    End Select
  Case 4
    GetAvailableColours = "16 Bit, 65536 Colours"
  Case Else
    GetAvailableColours = "Undetermined"
End Select

End Function


