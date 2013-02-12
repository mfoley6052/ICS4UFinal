VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "RTS"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSelAnim 
      Index           =   0
      Interval        =   150
      Left            =   11520
      Top             =   120
   End
   Begin VB.Label lblTest 
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   10920
      TabIndex        =   4
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblTest 
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   10920
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblTest 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   10920
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblTest 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   10920
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblTest 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   10920
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
   Begin VB.Image imgSel 
      Enabled         =   0   'False
      Height          =   1500
      Left            =   1440
      Picture         =   "frmMain.frx":0000
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   36
      Left            =   0
      Picture         =   "frmMain.frx":016A
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   35
      Left            =   0
      Picture         =   "frmMain.frx":0368
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   34
      Left            =   0
      Picture         =   "frmMain.frx":0566
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   33
      Left            =   0
      Picture         =   "frmMain.frx":0764
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   32
      Left            =   0
      Picture         =   "frmMain.frx":0962
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   31
      Left            =   0
      Picture         =   "frmMain.frx":0B60
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   30
      Left            =   0
      Picture         =   "frmMain.frx":0D5E
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   29
      Left            =   0
      Picture         =   "frmMain.frx":0F5C
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   28
      Left            =   0
      Picture         =   "frmMain.frx":115A
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   27
      Left            =   0
      Picture         =   "frmMain.frx":1358
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   26
      Left            =   0
      Picture         =   "frmMain.frx":1556
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   25
      Left            =   0
      Picture         =   "frmMain.frx":1754
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   24
      Left            =   0
      Picture         =   "frmMain.frx":1952
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   23
      Left            =   0
      Picture         =   "frmMain.frx":1B50
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   22
      Left            =   0
      Picture         =   "frmMain.frx":1D4E
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   21
      Left            =   0
      Picture         =   "frmMain.frx":1F4C
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   20
      Left            =   0
      Picture         =   "frmMain.frx":214A
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   19
      Left            =   0
      Picture         =   "frmMain.frx":2348
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   18
      Left            =   0
      Picture         =   "frmMain.frx":2546
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   17
      Left            =   0
      Picture         =   "frmMain.frx":2744
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   16
      Left            =   0
      Picture         =   "frmMain.frx":2942
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   15
      Left            =   0
      Picture         =   "frmMain.frx":2B40
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   14
      Left            =   0
      Picture         =   "frmMain.frx":2D3E
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   13
      Left            =   0
      Picture         =   "frmMain.frx":2F3C
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   12
      Left            =   0
      Picture         =   "frmMain.frx":313A
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   11
      Left            =   0
      Picture         =   "frmMain.frx":3338
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   10
      Left            =   0
      Picture         =   "frmMain.frx":3536
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   9
      Left            =   0
      Picture         =   "frmMain.frx":3734
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   8
      Left            =   0
      Picture         =   "frmMain.frx":3932
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   7
      Left            =   0
      Picture         =   "frmMain.frx":3B30
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   6
      Left            =   0
      Picture         =   "frmMain.frx":3D2E
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   5
      Left            =   0
      Picture         =   "frmMain.frx":3F2C
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   4
      Left            =   0
      Picture         =   "frmMain.frx":412A
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   3
      Left            =   0
      Picture         =   "frmMain.frx":4328
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   2
      Left            =   0
      Picture         =   "frmMain.frx":4526
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image imgTile 
      Enabled         =   0   'False
      Height          =   1500
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":4724
      Top             =   0
      Width           =   1500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbTile(0 To 2) As String
Dim mapData(0, 1 To 24, 1 To 24) As Integer
Dim intFieldSize As Integer
Dim u(0 To 1) As Integer
Dim intSelTile As Integer
Dim strSelType As String
Dim frmWidthRatio As Single
Dim frmHeightRatio As Single

Private Sub Form_Load()
frmWidthRatio = 15.3
frmHeightRatio = 15.95
u(0) = 50
dbTile(0) = "Plain"
strSelType = "R"
'0 - Terrain
'1 - Height
'2 - Effect
'Call defineMaps
u(1) = 100
dbTile(0) = "Plain"
strSelType = "R"
'0 - Terrain
'1 - Height
'2 - Effect
'Call defineMaps
Call getField
End Sub

Private Sub defineMaps()
'mapData(0, 1, 1)
'mapData(0, 1, 1)
End Sub

Private Sub getField()
Dim intMap As Integer
intMap = 0
If intMap = 0 Then
    intFieldSize = 6
End If
For Y = 1 To intFieldSize
    For X = 1 To intFieldSize
        imgTile(X + (intFieldSize * (Y - 1))).Top = (((Y - 1) * u(0)) / 2)
        imgTile(X + (intFieldSize * (Y - 1))).Left = ((X - 1) * u(1)) + ((u(1) * (Y Mod 2)) / 2)
    Next X
Next Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
intSelTile = -1
lblTest(0).Caption = X
Dim intSelTileTemp As Integer
intSelTileTemp = Int((X + 50) / u(1)) + Int(Y / ((u(0) / 2))) * intFieldSize
lblTest(1).Caption = Int((intSelTileTemp - 1) / intFieldSize) Mod 2
lblTest(2).Caption = (((X + 50) + (u(1) / 2)) - (u(1) * Int(((X + 50) + (u(1) / 2)) / u(1)))) / 2
lblTest(3).Caption = ((X - 50) + (u(1) / 2)) - (u(1) * Int(((X - 50) + (u(1) / 2)) / u(1)))
lblTest(4).Caption = (Y + (50 * (Int(Y / ((u(0) / 2))) / 2))) - (u(0) * Int(Y / (u(0) / 2)))
'lblTest(3).Caption = ((X + (u(1) / 2)) - (u(1) * Int((X + (u(1) / 2)) / u(1))) - 50) / 2
'lblTest(4).Caption = Y - (u(0) * Int(Y / (u(0) / 2)))
'lblTest(4).Caption = (Y + 100) - (u(0) * Int(Y / (u(0) / 2)))
If Int((intSelTileTemp - 1) / intFieldSize) Mod 2 = 0 Then
    If (X + (u(1) / 2)) - (u(1) * Int((X + (u(1) / 2)) / u(1))) < 50 Then
        If ((X + (u(1) / 2)) - (u(1) * Int((X + (u(1) / 2)) / u(1))) - 50) / -2 >= (Y + (50 * (Int(Y / ((u(0) / 2))) / 2))) - (u(0) * Int(Y / (u(0) / 2))) Then
            intSelTile = intSelTileTemp - intFieldSize
        End If
    Else
        If ((X + (u(1) / 2)) - (u(1) * Int((X + (u(1) / 2)) / u(1))) - 50) / 2 >= (Y + (50 * (Int(Y / ((u(0) / 2))) / 2))) - (u(0) * Int(Y / (u(0) / 2))) Then
            intSelTile = (intSelTileTemp + 1) - intFieldSize
        End If
    End If
Else
    If ((X - 50) + (u(1) / 2)) - (u(1) * Int(((X - 50) + (u(1) / 2)) / u(1))) < 50 Then
        intSelTileTemp = intSelTileTemp + 1
        If (((X + 50) + (u(1) / 2)) - (u(1) * Int(((X + 50) + (u(1) / 2)) / u(1)))) / 2 < (Y + (50 * (Int(Y / ((u(0) / 2))) / 2))) - (u(0) * Int(Y / (u(0) / 2))) Then
            intSelTile = (intSelTileTemp - 1) - intFieldSize
        End If
    Else
        If ((X + (u(1) / 2)) - (u(1) * Int((X + (u(1) / 2)) / u(1)))) / 2 >= (Y + (50 * (Int(Y / ((u(0) / 2))) / 2))) - (u(0) * Int(Y / (u(0) / 2))) Then
            intSelTile = intSelTileTemp - intFieldSize
        End If
    End If
End If
If intSelTile >= 0 Then
    strSelType = "B"
    imgSel.Top = imgTile(intSelTile).Top
    imgSel.Left = imgTile(intSelTile).Left
Else
    strSelType = "R"
    imgSel.Top = imgTile(intSelTileTemp).Top
    imgSel.Left = imgTile(intSelTileTemp).Left
End If
End Sub

Private Sub imgTile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim intSelTileTemp As Integer
intSelTileTemp = Int(X / u(1)) + Int((Y / u(0)) * intFieldSize)
'Y - (intSelTileTemp * u(0))


End Sub

Private Sub tmrSelAnim_Timer(Index As Integer)
Static count As Integer
Static blnRev As Boolean
If blnRev = False Then
    count = count + 1
End If
imgSel.Picture = LoadPicture(App.Path & "\Tile\Sel\Sel" & strSelType & count & ".gif")
If blnRev = True Then
    count = count - 1
End If
If count >= 5 Or count <= 0 Then
    If blnRev = False Then
        blnRev = True
    Else
        blnRev = False
    End If
End If
End Sub
