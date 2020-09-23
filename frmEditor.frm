VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MapEditor"
   ClientHeight    =   7965
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8295
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   8520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   30
      Top             =   5280
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   8520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   29
      Top             =   5040
      Width           =   255
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8880
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmever 
      Caption         =   "Script (Ever)"
      Height          =   7575
      Left            =   0
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   8295
      Begin RichTextLib.RichTextBox rtfScript3 
         Height          =   7215
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   12726
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmEditor.frx":27A2
      End
   End
   Begin VB.Frame frmOnend 
      Caption         =   "Script (OnEnd)"
      Height          =   7575
      Left            =   0
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   8295
      Begin RichTextLib.RichTextBox rtfScript2 
         Height          =   7215
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   12726
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmEditor.frx":2824
      End
   End
   Begin VB.Frame frmOnload 
      Caption         =   "Script (OnLoad)"
      Height          =   7575
      Left            =   0
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   8295
      Begin RichTextLib.RichTextBox rtfScript 
         Height          =   7215
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   12726
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmEditor.frx":28A6
      End
   End
   Begin MSComctlLib.TabStrip tabb 
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   661
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Map"
            Key             =   "tabMAP"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Script (OnLoad)"
            Key             =   "tabonload"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Script (OnEnd)"
            Key             =   "tabonend"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Script (Ever)"
            Key             =   "tabever"
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Frame frmEditor 
      Caption         =   "Editor"
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8295
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   5415
         Begin VB.PictureBox imgTile 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   7695
            Left            =   0
            Picture         =   "frmEditor.frx":2928
            ScaleHeight     =   513
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   513
            TabIndex        =   23
            Top             =   0
            Width           =   7695
            Begin VB.Shape Shape2 
               BorderColor     =   &H000000C0&
               BorderWidth     =   3
               Height          =   255
               Left            =   0
               Top             =   0
               Width           =   255
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Extra"
         Height          =   975
         Left            =   3960
         TabIndex        =   10
         Top             =   6480
         Width           =   4215
         Begin VB.TextBox txtSprite2 
            Height          =   285
            Left            =   1680
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtSprite 
            Height          =   285
            Left            =   1080
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "PlayerSprite"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "PlayerSprite (MASK)"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame info 
         Caption         =   "TileInfo"
         Height          =   975
         Left            =   3000
         TabIndex        =   5
         Top             =   6480
         Width           =   855
         Begin VB.Label lblY 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   9
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Y:"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "X:"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   135
         End
         Begin VB.Label lblX 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Options"
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   6480
         Width           =   2655
         Begin VB.OptionButton Option2 
            Caption         =   "Layer2"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Layer1"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.CheckBox chkWalkable 
            Caption         =   "UnWalkable"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image Image4 
            Height          =   255
            Left            =   2280
            Picture         =   "frmEditor.frx":BF6A
            Stretch         =   -1  'True
            Top             =   600
            Width           =   255
         End
         Begin VB.Image Image3 
            Height          =   255
            Left            =   2040
            Picture         =   "frmEditor.frx":C3AC
            Stretch         =   -1  'True
            Top             =   600
            Width           =   255
         End
         Begin VB.Image Image2 
            Height          =   255
            Left            =   1800
            Picture         =   "frmEditor.frx":C7EE
            Stretch         =   -1  'True
            Top             =   600
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   2040
            Picture         =   "frmEditor.frx":CC30
            Stretch         =   -1  'True
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "TileSet"
         Height          =   6135
         Left            =   5640
         TabIndex        =   1
         Top             =   240
         Width           =   2535
         Begin VB.Frame Frame5 
            Height          =   855
            Left            =   240
            TabIndex        =   28
            Top             =   5160
            Width           =   975
            Begin VB.Image Image5 
               Height          =   240
               Left            =   360
               Picture         =   "frmEditor.frx":D072
               Stretch         =   -1  'True
               Top             =   240
               Width           =   240
            End
            Begin VB.Image Image6 
               Height          =   255
               Left            =   120
               Picture         =   "frmEditor.frx":D4B4
               Stretch         =   -1  'True
               Top             =   480
               Width           =   255
            End
            Begin VB.Image Image7 
               Height          =   255
               Left            =   360
               Picture         =   "frmEditor.frx":D8F6
               Stretch         =   -1  'True
               Top             =   480
               Width           =   255
            End
            Begin VB.Image Image8 
               Height          =   255
               Left            =   600
               Picture         =   "frmEditor.frx":DD38
               Stretch         =   -1  'True
               Top             =   480
               Width           =   255
            End
         End
         Begin VB.PictureBox imgPreview 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1200
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   25
            Top             =   5760
            Width           =   255
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Load Set"
            Height          =   255
            Left            =   1560
            TabIndex        =   24
            Top             =   5760
            Width           =   855
         End
         Begin VB.PictureBox picTileset 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   2400
            Left            =   120
            Picture         =   "frmEditor.frx":E17A
            ScaleHeight     =   160
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   80
            TabIndex        =   2
            Top             =   240
            Width           =   1200
            Begin VB.Shape Shape1 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Height          =   255
               Left            =   0
               Top             =   0
               Width           =   255
            End
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnudel 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "&End"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuFill 
         Caption         =   "&Fill"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "frmEditor"
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

Private Sub Command1_Click()
On Error GoTo errr
cd1.Filter = "*.bmp,*.gif,*.jpeg"
cd1.ShowOpen
picTileset.Picture = LoadPicture(App.Path & "\DATA\IMAGES\TILES\" & cd1.FileTitle)
MAP.TILESET = cd1.FileTitle
Exit Sub
errr:
MsgBox "error while loading"
End Sub

Private Sub Form_Load()
rtfScript.TEXT = "'OnLoad'" & vbCrLf & "Sub OnLoad" & vbCrLf & vbCrLf & "End Sub"
rtfScript2.TEXT = "'OnEnd'" & vbCrLf & "Sub OnEnd" & vbCrLf & vbCrLf & "End Sub"
rtfScript3.TEXT = "'Ever'" & vbCrLf & "Sub Ever" & vbCrLf & vbCrLf & "End Sub"
MAP.TILESET = "TILESET1.BMP"
MAP.PLAYERSPRITE = "SPRITE2.BMP"
MAP.PlayerMask = "SPRITE2MASK.BMP"
txtSprite.TEXT = MAP.PLAYERSPRITE
txtSprite2.TEXT = MAP.PlayerMask
picTileset.Picture = LoadPicture(App.Path & "\DATA\IMAGES\TILES\" & MAP.TILESET)
picTileset_MouseDown 1, 0, 0, 0
DrawPreview
mnuFill_Click
imgTile.Refresh
End Sub


Private Sub Image1_Click()
imgTile.Top = imgTile.Top + 16 * 15
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgTile.Left = imgTile.Left + 16 * 15
End Sub

Private Sub Image3_Click()
imgTile.Top = imgTile.Top - 16 * 15
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgTile.Left = imgTile.Left - 16 * 15
End Sub

Private Sub Image5_Click()
picTileset.Top = picTileset.Top + 16 * 15
End Sub

Private Sub Image6_Click()
picTileset.Left = picTileset.Left + 16 * 15
End Sub

Private Sub Image7_Click()
picTileset.Top = picTileset.Top - 16 * 15
End Sub

Private Sub Image8_Click()
picTileset.Left = picTileset.Left - 16 * 15
End Sub

Private Sub imgTile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As String
If chkWalkable.Value = 1 Then a = 1 Else a = 0
Shape2.Left = Snap(X, 16)
Shape2.Top = Snap(Y, 16)
MAP.SetTile Shape2.Left / 16, Shape2.Top / 16, Shape1.Left / 16, Shape1.Top / 16, a
End Sub

Private Sub imgTile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As String
Shape2.Left = Snap(X, 16)
Shape2.Top = Snap(Y, 16)
lblX.Caption = Shape2.Left / 16
lblY.Caption = Shape2.Top / 16
If Button = 1 Then
If chkWalkable.Value = 1 Then a = 1 Else a = 0
MAP.SetTile Shape2.Left / 16, Shape2.Top / 16, Shape1.Left / 16, Shape1.Top / 16, a
End If
End Sub


Private Sub mnuEnd_Click()
a = MsgBox("Do you really want to Close this Program?", vbOKCancel)
If a = vbOK Then End
End Sub

Private Sub mnuFill_Click()
MAP.FillMap
End Sub

Private Sub mnuLoad_Click()
MAP.LoadMap
End Sub

Private Sub mnuSave_Click()
MAP.SaveMap
End Sub

Private Sub Option1_Click()
chkWalkable.Enabled = True
End Sub

Private Sub Option2_Click()
chkWalkable.Enabled = False
End Sub

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Left = Snap(X, 16)
Shape1.Top = Snap(Y, 16)
DrawPreview
End Sub

Private Sub DrawPreview()
imgPreview.PaintPicture picTileset.Picture, 0, 0, 16, 16, Shape1.Left, Shape1.Top, 16, 16
imgPreview.Refresh
End Sub
Private Sub tabb_Click()
Select Case tabb.SelectedItem.Index
Case "1"
frmEditor.Visible = True
frmOnload.Visible = False
frmOnend.Visible = False
frmever.Visible = False
Case "2"
frmEditor.Visible = False
frmOnload.Visible = True
frmOnend.Visible = False
frmever.Visible = False
Case "3"
frmEditor.Visible = False
frmOnload.Visible = False
frmOnend.Visible = True
frmever.Visible = False
Case "4"
frmEditor.Visible = False
frmOnload.Visible = False
frmOnend.Visible = False
frmever.Visible = True
End Select
End Sub

Private Sub txtSprite_Change()
MAP.PLAYERSPRITE = txtSprite.TEXT
End Sub

Private Sub txtSprite2_Change()
MAP.PlayerMask = txtSprite2.TEXT
End Sub
