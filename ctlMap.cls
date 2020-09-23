VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ctlGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'++++++++++++++++
'+RPG Engine... +
'+2002 by       +
'+SmokingFish   ++++++
'+mail@smokingfish.de+
'+++++++++++++++++++++
Private mProgress As Long
Private FLAG(0 To 999) As String
Public APPEND As Boolean
Public FULLSCREEN As Boolean
Public SETFILTER As Integer
Public MAPHEIGHT, MAPWIDTH As Integer
Private SCRIPTTITLE As String
Public Speed As Integer
Public t_PlayerX As Integer, t_PlayerY As Integer
Public OnlyOnLoad As Boolean
Public FACE As Integer
Private TILELINE As New Collection
Private TILELINE2 As New Collection
Public COLLX As New Collection
Public COLLY As New Collection
Private COLL As String
Public SCRIPTS As New Collection
Public SCRIPTNAMES As New Collection
Private TILESET As String
Private PLAYERSPRITE As String
Private PLAYERSMASK As String
Private TILEINFO() As String
Private TILEINFO2() As String
Private MARK As String
Private MARK2 As String
Private SCRIPTTEXT As String
Public NPCIMAGE1 As New Collection
Public NPCIMAGE2 As New Collection
Public NPCX As New Collection
Public NPCY As New Collection

Public Sub SetFlag(INDEX As Integer, TEXT As String)
FLAG(INDEX) = TEXT
End Sub

Public Function GetFlag(INDEX As Integer) As String
GetFlag = FLAG(INDEX)
End Function

Private Sub DrawTile(TileX As Long, TileY As Long, SourceX As Long, SourceY As Long)
BitBlt frmMain.picGame.HDC, TileX * 16, TileY * 16, 16, 16, frmMain.picTILESET.HDC, SourceX * 16, SourceY * 16, vbSrcCopy
'TransparentBlt frmMain.picGame.HDC, TileX * 16, TileY * 16, 16, 16, frmMain.picTILESET.HDC, SourceX * 16, SourceY * 16, 16, 16, &HFF00FF
End Sub
Private Sub DrawTile2(TileX As Long, TileY As Long, SourceX As Long, SourceY As Long)
'BitBlt frmMain.picGame.HDC, TileX * 16, TileY * 16, 16, 16, frmMain.picTILESET.HDC, SourceX * 16, SourceY * 16, vbSrcCopy
TransparentBlt frmMain.picGame.HDC, TileX * 16, TileY * 16, 16, 16, frmMain.picTILESET.HDC, SourceX * 16, SourceY * 16, 16, 16, &HFF00FF
End Sub
Private Sub DrawMap()
For I = 1 To TILELINE.Count - 1
TILEINFO = Split(TILELINE.Item(I), ",")
DrawTile SUBS.StringToLong(TILEINFO(0)), SUBS.StringToLong(TILEINFO(1)), SUBS.StringToLong(TILEINFO(2)), SUBS.StringToLong(TILEINFO(3))
Next I
End Sub
Private Sub DrawMapLayer2()
For I = 1 To TILELINE2.Count - 1
TILEINFO2 = Split(TILELINE2.Item(I), ",")
DrawTile2 SUBS.StringToLong(TILEINFO2(0)), SUBS.StringToLong(TILEINFO2(1)), SUBS.StringToLong(TILEINFO2(2)), SUBS.StringToLong(TILEINFO2(3))
Next I
End Sub
Public Sub MSG(TEXT As String)
Hoch = False
Runter = False
Rechts = False
Links = False
frmMain.picGame.Refresh
frmMain.Picture3.Visible = True
frmMain.lblMSG.Caption = TEXT
SUBS.Wait 3
frmMain.Picture3.Visible = False
End Sub
Public Sub StartMap(FileName As String)
frmMain.picGame.Cls
frmMain.picGame2.Cls
Hoch = False
Runter = False
Links = False
Rechts = False
For o = 1 To RPG.COLLX.Count - 1
RPG.COLLX.Remove (RPG.COLLX.Count)
Next o
For o = 1 To RPG.COLLY.Count - 1
RPG.COLLY.Remove (RPG.COLLY.Count)
Next o
For o = 1 To TILELINE.Count - 1
TILELINE.Remove (TILELINE.Count)
Next o
For o = 1 To TILELINE2.Count - 1
TILELINE2.Remove (TILELINE2.Count)
Next o
For o = 1 To RPG.SCRIPTS.Count - 1
RPG.SCRIPTS.Remove (o)
Next o
For o = 1 To RPG.SCRIPTNAMES.Count - 1
RPG.SCRIPTNAMES.Remove (o)
Next o
If RPG.SCRIPTNAMES.Count > 0 Then
RPG.SCRIPTNAMES.Remove 1
End If
If RPG.SCRIPTS.Count > 0 Then
RPG.SCRIPTS.Remove 1
End If
If RPG.COLLX.Count > 0 Then
RPG.COLLX.Remove 1
End If
If RPG.COLLY.Count > 0 Then
RPG.COLLY.Remove 1
End If
RPG.LoadMap FileName
RPG.FACE = 2
RPG.Speed = 4
RPG.PlayerX = 0
RPG.PlayerY = 0
DrawMap
DrawMapLayer2
For I = 1 To RPG.SCRIPTS.Count
CHECK = Left(RPG.SCRIPTS.Item(I), 8)
If CHECK = "'OnLoad'" Then
frmMain.SC.AddCode RPG.SCRIPTS.Item(I)
frmMain.SC.Run RPG.SCRIPTNAMES.Item(I)
End If
Next I
End Sub

Public Sub CheckScripts()
Dim CHECK As String
For I = 1 To RPG.SCRIPTS.Count
CHECK = Left(RPG.SCRIPTS.Item(I), 8)
If CHECK = "'OnLoad'" Then GoTo EX
CHECK = Left(RPG.SCRIPTS.Item(I), 7)
If CHECK = "'OnEnd'" Then GoTo EX
CHECK = Left(RPG.SCRIPTS.Item(I), 6)
If CHECK = "'Ever'" Then
frmMain.SC.AddCode RPG.SCRIPTS.Item(I)
frmMain.SC.Run RPG.SCRIPTNAMES(I)
End If
EX:
Next I
End Sub

Public Sub LoadMap(FileName As String)
Hoch = False
Runter = False
Links = False
Rechts = False
FileName = App.Path & "\DATA\MAPS\" & FileName
Open FileName For Input As #1
Input #1, TILESET
Input #1, PLAYERSPRITE
Input #1, PLAYERmask
Input #1, MAPHEIGHT, MAPWIDTH
MAPHEIGHT = MAPHEIGHT - 1
MAPWIDTH = MAPWIDTH - 1
Do
DoEvents
Input #1, X2, Y2, SX2, SY2, COLL
If COLL = "1" Then
COLLY.Add Y2 * 16
COLLX.Add X2 * 16
End If
TILELINE.Add X2 & "," & Y2 & "," & SX2 & "," & SY2 & "," & COLL
Loop Until X2 = "END_TILES"
Do
DoEvents
Input #1, X2, Y2, SX2, SY2
TILELINE2.Add X2 & "," & Y2 & "," & SX2 & "," & SY2
Loop Until X2 = "END_TILES2"
Do
DoEvents
Input #1, MARK
If MARK = "END_FILE" Then Exit Do
If MARK = "START_SCRIPT" Then
Input #1, SCRIPTTITLE
Do
DoEvents
Line Input #1, MARK2
If MARK2 = "END_SCRIPT" Then
SCRIPTS.Add SCRIPTTEXT
SCRIPTNAMES.Add SCRIPTTITLE
SCRIPTTEXT = ""
Exit Do
Else
SCRIPTTEXT = SCRIPTTEXT + MARK2 + vbCrLf
End If
Loop
End If
Loop
'---------
frmMain.picTILESET.Picture = LoadPicture(App.Path & "\DATA\IMAGES\TILES\" & TILESET)
frmMain.picPLAYER.Picture = LoadPicture(App.Path & "\DATA\IMAGES\PLAYER\" & PLAYERSPRITE)
frmMain.picPLAYER2.Picture = LoadPicture(App.Path & "\DATA\IMAGES\PLAYER\" & PLAYERmask)
SizePic
Close #1
End Sub

Public Sub SizePic()
frmMain.picGame.Move 0, 0, RPG.MAPWIDTH * 16, RPG.MAPHEIGHT * 16
frmMain.picGame.Move frmMain.ScaleWidth / 2 - frmMain.picGame.ScaleWidth / 2, frmMain.ScaleHeight / 2 - frmMain.picGame.ScaleHeight / 2, frmMain.picGame.Width, frmMain.picGame.Height
frmMain.picGame2.Move 0, 0, frmMain.Width, frmMain.Height
End Sub

Private Sub DrawPlayer(X As String, Y As String, FACE As String)
If Frame = 0 Then
If FACE = 1 Then
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER2.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER.HDC, 0, 0, SRCAND
ElseIf FACE = 2 Then
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER2.HDC, 0, 32, MERGEPAINT
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER.HDC, 0, 32, SRCAND
ElseIf FACE = 3 Then
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER2.HDC, 0, 64, MERGEPAINT
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER.HDC, 0, 64, SRCAND
ElseIf FACE = 4 Then
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER2.HDC, 0, 96, MERGEPAINT
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER.HDC, 0, 96, SRCAND
End If
ElseIf Frame = 1 Then
If FACE = 1 Then
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER2.HDC, 32, 0, MERGEPAINT
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER.HDC, 32, 0, SRCAND
ElseIf FACE = 2 Then
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER2.HDC, 32, 32, MERGEPAINT
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER.HDC, 32, 32, SRCAND
ElseIf FACE = 3 Then
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER2.HDC, 32, 64, MERGEPAINT
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER.HDC, 32, 64, SRCAND
ElseIf FACE = 4 Then
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER2.HDC, 32, 96, MERGEPAINT
BitBlt frmMain.picGame.HDC, X, Y, 32, 32, frmMain.picPLAYER.HDC, 32, 96, SRCAND
End If
End If
End Sub

Public Sub DrawChar()
DrawPlayer RPG.PlayerX, RPG.PlayerY, RPG.FACE
End Sub

Public Sub Render()
Do
DoEvents
DrawMap
DrawMapLayer2
DrawChar
If NPCIMAGE1.Count > 0 Then
For u = 1 To NPCIMAGE1.Count
frmMain.picNPC.Picture = LoadPicture(App.Path & "\DATA\IMAGES\NPC\" & NPCIMAGE1.Item(u))
frmMain.picNPC2.Picture = LoadPicture(App.Path & "\DATA\IMAGES\NPC\" & NPCIMAGE2.Item(u))
BitBlt frmMain.picGame.HDC, NPCX.Item(u), NPCY.Item(u), frmMain.picNPC.ScaleWidth, frmMain.picNPC.ScaleHeight, frmMain.picNPC2.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.picGame.HDC, NPCX.Item(u), NPCY.Item(u), frmMain.picNPC.ScaleWidth, frmMain.picNPC.ScaleHeight, frmMain.picNPC.HDC, 0, 0, SRCAND
Next u
End If
CheckScripts
frmMain.picGame.Refresh
Loop Until RPG.APPEND = True
End Sub

Public Sub AddNPC(Picture1 As String, Picture2 As String, X As Integer, Y As Integer)
NPCIMAGE1.Add Picture1
NPCIMAGE2.Add Picture2
NPCX.Add X
NPCY.Add Y
End Sub

Public Sub RemoveNPC(INDEX As Integer)
NPCIMAGE1.Remove INDEX
NPCIMAGE2.Remove INDEX
NPCX.Remove INDEX
NPCY.Remove INDEX
End Sub
Public Function DoesCollide(X As Integer, Y As Integer) As Boolean
    
    
For m = 1 To COLLX.Count
    If COLLX(m) = X And COLLY(m) = Y Then
        DoesCollide = True
        Exit Function
    End If
    
    
Next m
    
        DoesCollide = False
End Function

Public Property Get PlayerX() As Integer
PlayerX = t_PlayerX
End Property

Public Property Let PlayerX(ByVal vNewValue As Integer)
t_PlayerX = vNewValue
End Property


Public Property Get PlayerY() As Integer
PlayerY = t_PlayerY
End Property

Public Property Let PlayerY(ByVal vNewValue As Integer)
t_PlayerY = vNewValue
End Property

