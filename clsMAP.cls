VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMAP"
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
Private MapData(0 To 31, 0 To 31) As String
Private MapData2(0 To 31, 0 To 31) As String
Public TILESET As String
Public PLAYERSPRITE As String
Public PlayerMask As String
Public DumScript As String
Public DumScriptName As String
Private PosX(0 To 31) As String
Private PosY(0 To 31) As String
Private TileXX(0 To 31) As String
Private TileYY(0 To 31) As String
Private COLL(0 To 31, 0 To 31) As String
Public SCRIPTS As New Collection
Public ScriptNames As New Collection
Private MARK As String
Private MARK2 As String
Private SCRIPTTEXT As String
Public XPos, YPos, XTile, YTile, Col As String
Public Sub FillMap()
Dim lol, lol2
Dim a
If frmEditor.chkWalkable.Value = 1 Then a = 1 Else a = 0
For lol2 = 0 To 31
For lol = 0 To 31
SetTile Int(lol), Int(lol2), frmEditor.Shape1.Left / 16, frmEditor.Shape1.Top / 16, Int(a)
Next lol
Next lol2
End Sub

Public Function GetTile(XX As String, YY As String) As String
GetTile = MapData(XX, YY)
End Function

Public Sub SetTile(XX As String, YY As String, TileX As String, TileY As String, WALK As String)
On Error Resume Next
If frmEditor.Option1.Value = True Then
MapData(XX, YY) = TileX & "," & TileY & "," & WALK
MapData2(XX, YY) = ""
XX = XX * 16
YY = YY * 16
TileX = TileX * 16
TileY = TileY * 16
TransparentBlt frmEditor.imgTile.hDC, XX, YY, 16, 16, frmEditor.Picture1.hDC, TileX, TileY, 16, 16, &HFF00FF
TransparentBlt frmEditor.imgTile.hDC, XX, YY, 16, 16, frmEditor.picTileset.hDC, TileX, TileY, 16, 16, &HFF00FF
Else
MapData2(XX, YY) = TileX & "," & TileY
XX = XX * 16
YY = YY * 16
TileX = TileX * 16
TileY = TileY * 16
TransparentBlt frmEditor.imgTile.hDC, XX, YY, 16, 16, frmEditor.picTileset.hDC, TileX, TileY, 16, 16, &HFF00FF
End If
frmEditor.imgTile.Refresh
End Sub

Public Sub LoadMap()
'On Error GoTo errr
frmEditor.cd1.Filter = "*.MAP"
frmEditor.cd1.ShowOpen
If SCRIPTS.Count > 0 Then
For i = 1 To SCRIPTS.Count
ScriptNames.Remove SCRIPTS.Count
SCRIPTS.Remove SCRIPTS.Count
Next i
End If
Open App.Path & "\DATA\MAPS\" & frmEditor.cd1.FileTitle For Input As #1
Input #1, TILESET
Input #1, PLAYERSPRITE
Input #1, PlayerMask
Input #1, MapHeight, MapWidth
frmEditor.picTileset.Picture = LoadPicture(App.Path & "\DATA\IMAGES\TILES\" & TILESET)
Do
DoEvents
Input #1, XPos, YPos, XTile, YTile, Col
If XPos = "END_TILES" Then GoTo WEITER
frmEditor.Option1.Value = True
SetTile Str(XPos), Str(YPos), Str(XTile), Str(YTile), Str(Col)
WEITER:
Loop Until XPos = "END_TILES"
Do
DoEvents
Input #1, XPos, YPos, XTile, YTile
If XPos = "END_TILES2" Then GoTo WEITER2
frmEditor.Option2.Value = True
SetTile Str(XPos), Str(YPos), Str(XTile), Str(YTile), "0"
WEITER2:
Loop Until XPos = "END_TILES2"
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
ScriptNames.Add SCRIPTTITLE
SCRIPTTEXT = ""
Exit Do
Else
SCRIPTTEXT = SCRIPTTEXT + MARK2 + vbCrLf
End If
Loop
End If
Loop
Close #1
frmEditor.rtfScript.TEXT = MAP.SCRIPTS.Item(1)
frmEditor.rtfScript2.TEXT = MAP.SCRIPTS.Item(2)
frmEditor.rtfScript3.TEXT = MAP.SCRIPTS.Item(3)
Exit Sub
errr:
MsgBox "file not found"
Exit Sub
End Sub

Public Sub SaveMap()
On Error GoTo errr
Dim lol, lol2
frmEditor.cd1.Filter = "*.MAP"
frmEditor.cd1.ShowOpen
Open App.Path & "\DATA\MAPS\" & frmEditor.cd1.FileTitle For Output As #1
Print #1, TILESET
Print #1, PLAYERSPRITE
Print #1, PlayerMask
Print #1, "32", "32"
For lol2 = 0 To 31
For lol = 0 To 31
Print #1, lol & "," & lol2 & "," & MapData(lol, lol2)
Next lol
Next lol2
Print #1, "END_TILES" & "," & "," & "," & ","
For lol2 = 0 To 31
For lol = 0 To 31
If MapData2(lol, lol2) = "" Then
Else
Print #1, lol & "," & lol2 & "," & MapData2(lol, lol2)
End If
Next lol
Next lol2
Print #1, "END_TILES2" & "," & "," & "," & ","
Print #1, "START_SCRIPT"
Print #1, "OnLoad"
Print #1, frmEditor.rtfScript.TEXT
Print #1, "END_SCRIPT"
Print #1, "START_SCRIPT"
Print #1, "OnEnd"
Print #1, frmEditor.rtfScript2.TEXT
Print #1, "END_SCRIPT"
Print #1, "START_SCRIPT"
Print #1, "Ever"
Print #1, frmEditor.rtfScript3.TEXT
Print #1, "END_SCRIPT"
Print #1, "END_FILE"
Close #1
Exit Sub
errr:
MsgBox "error while saving"
Exit Sub
End Sub
