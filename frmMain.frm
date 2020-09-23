VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   568
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1080
      Top             =   480
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frmMain.frx":27A2
      ScaleHeight     =   975
      ScaleWidth      =   3375
      TabIndex        =   8
      Top             =   7080
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblMSG 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.PictureBox picGame2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Animation 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   720
      Top             =   480
   End
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.PictureBox picPLAYER 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picNPC2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picPLAYER2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picTILESET2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picNPC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picTILESET 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picGame 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   6720
      Picture         =   "frmMain.frx":342E
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7320
      Picture         =   "frmMain.frx":394F
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6120
      Picture         =   "frmMain.frx":3E93
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'++++++++++++++++
'+RPG Engine... +
'+2002 by       +
'+SmokingFish   ++++++
'+mail@smokingfish.de+
'+++++++++++++++++++++
Dim xres, yres As String
Dim rColor As String
Dim IniSectionName, IniFileName As String
Dim IniStat As Long
Dim TX, TY As Integer
Dim MoveNow As Boolean

Private Sub Animation_Timer()
If Frame = 0 Then Frame = 1 Else Frame = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload frmMain
End If
If KeyCode = vbKeyUp Then Hoch = True
If KeyCode = vbKeyDown Then Runter = True
If KeyCode = vbKeyLeft Then Links = True
If KeyCode = vbKeyRight Then Rechts = True
Animation.Enabled = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then Hoch = False
If KeyCode = vbKeyDown Then Runter = False
If KeyCode = vbKeyLeft Then Links = False
If KeyCode = vbKeyRight Then Rechts = False
Animation.Enabled = False
End Sub

Private Sub Form_Load()
Me.Show
Me.Refresh
Dim CHECK As String
SC.AddObject "RPG", RPG
SC.AddObject "SUBS", SUBS
SC.AddObject "FORM", frmMain
SC.AddObject "PicGame", picGame
Frame = 0
Open App.Path & "\CONFIG.CFG" For Input As #1
Input #1, GG
Close #1
RPG.StartMap GG
End Sub
  
Private Sub Form_LostFocus()
Hoch = False
Runter = False
Links = False
Rechts = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveNow = True
    TX = X
    TY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveNow Then
        Me.Top = Me.Top + Y * 15 - TY * 15
        Me.Left = Me.Left + X * 15 - TX * 15
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveNow = False
End Sub

Private Sub Form_Resize()
RPG.SizePic
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim CHECK As String
For I = 1 To RPG.SCRIPTS.Count
CHECK = Left(RPG.SCRIPTS.Item(I), 7)
If CHECK = "'OnEnd'" Then
SC.AddCode RPG.SCRIPTS.Item(I)
SC.Run RPG.SCRIPTNAMES(I)
End If
Next I
End
End Sub


Private Sub Scripts_Timer()
Dim CHECK As String
For I = 1 To RPG.SCRIPTS.Count
Debug.Print RPG.SCRIPTS.Item(I)
CHECK = Left(RPG.SCRIPTS.Item(I), 8)
Debug.Print CHECK
If CHECK = "'OnLoad'" Then GoTo EX
CHECK = Left(RPG.SCRIPTS.Item(I), 7)
Debug.Print CHECK
If CHECK = "'OnEnd'" Then GoTo EX
CHECK = Left(RPG.SCRIPTS.Item(I), 6)
If CHECK = "'Ever'" Then
SC.AddCode RPG.SCRIPTS.Item(I)
SC.Run RPG.SCRIPTNAMES(I)
End If
EX:
Next I
End Sub

Private Sub Timer1_Timer()
If Picture2.Left < Picture1.Width - Picture1.Width - Picture2.Width Then
    Picture2.Left = Picture1.Width - 1
    Picture2.Left = Picture2.Left - 5
Else
    Picture2.Left = Picture2.Left - 10
End If
End Sub


Private Sub Image1_Click()
frmMain.WindowState = 1
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Image3_Click()
MsgBox "By SmokingFish ,2002! mail@smokingfish.de / www.smokingfish.de", , App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Timer2_Timer()
Dim savex As Long
Dim savey As Long

savex = RPG.PlayerX
savey = RPG.PlayerY

If Hoch = True Then
RPG.PlayerY = RPG.PlayerY - RPG.Speed
RPG.FACE = 1
End If
If Runter = True Then
RPG.PlayerY = RPG.PlayerY + RPG.Speed
RPG.FACE = 3
End If
If Links = True Then
RPG.PlayerX = RPG.PlayerX - RPG.Speed
RPG.FACE = 4
End If
If Rechts = True Then
RPG.PlayerX = RPG.PlayerX + RPG.Speed
RPG.FACE = 2
End If




Dim semi_px As Long
Dim semi_py As Long
'################################

For semi_px = 1 To 32

If RPG.DoesCollide(RPG.PlayerX + semi_px - 6, RPG.PlayerY + 16 - 6) Then
RPG.PlayerX = savex
RPG.PlayerY = savey

Exit Sub
End If
Next

For semi_px = 1 To 32

If RPG.DoesCollide(RPG.PlayerX + semi_px - 6, RPG.PlayerY + 32 - 6) Then
RPG.PlayerX = savex
RPG.PlayerY = savey

Exit Sub
End If
Next

For semi_py = 16 To 32

If RPG.DoesCollide(RPG.PlayerX - 6, RPG.PlayerY + semi_py - 6) Then
RPG.PlayerX = savex
RPG.PlayerY = savey

Exit Sub
End If
Next

For semi_py = 16 To 32

If RPG.DoesCollide(RPG.PlayerX + 32 - 6, RPG.PlayerY + semi_py - 6) Then
RPG.PlayerX = savex
RPG.PlayerY = savey

Exit Sub
End If
Next
End Sub
