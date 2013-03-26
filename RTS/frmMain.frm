VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   4695
   ClientTop       =   1350
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Begin VB.Timer tmrJump 
      Index           =   3
      Interval        =   250
      Left            =   8520
      Top             =   3480
   End
   Begin VB.Timer tmrJump 
      Index           =   2
      Interval        =   250
      Left            =   8040
      Top             =   3480
   End
   Begin VB.Timer tmrJump 
      Index           =   1
      Interval        =   250
      Left            =   7560
      Top             =   3480
   End
   Begin VB.Timer tmrJump 
      Index           =   0
      Interval        =   250
      Left            =   7080
      Top             =   3480
   End
   Begin VB.PictureBox picSelMask 
      Height          =   810
      Index           =   4
      Left            =   8040
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSelMask 
      Height          =   810
      Index           =   3
      Left            =   6480
      Picture         =   "frmMain.frx":013D
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSelMask 
      Height          =   810
      Index           =   2
      Left            =   4920
      Picture         =   "frmMain.frx":027C
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSelMask 
      Height          =   810
      Index           =   1
      Left            =   3360
      Picture         =   "frmMain.frx":03BD
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSelMask 
      Height          =   810
      Index           =   0
      Left            =   1800
      Picture         =   "frmMain.frx":0500
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSelB 
      AutoRedraw      =   -1  'True
      Height          =   810
      Index           =   4
      Left            =   8040
      Picture         =   "frmMain.frx":0642
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSelB 
      AutoRedraw      =   -1  'True
      Height          =   810
      Index           =   3
      Left            =   6480
      Picture         =   "frmMain.frx":07B6
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSelB 
      AutoRedraw      =   -1  'True
      Height          =   810
      Index           =   2
      Left            =   4920
      Picture         =   "frmMain.frx":0944
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSelB 
      AutoRedraw      =   -1  'True
      Height          =   810
      Index           =   1
      Left            =   3360
      Picture         =   "frmMain.frx":0AAF
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picSelB 
      AutoRedraw      =   -1  'True
      Height          =   810
      Index           =   0
      Left            =   1800
      Picture         =   "frmMain.frx":0C0E
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picScene 
      AutoRedraw      =   -1  'True
      Height          =   1500
      Index           =   2
      Left            =   240
      Picture         =   "frmMain.frx":0D70
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picScene 
      AutoRedraw      =   -1  'True
      Height          =   1500
      Index           =   1
      Left            =   240
      Picture         =   "frmMain.frx":158D
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Timer tmrSel 
      Interval        =   125
      Left            =   5400
      Top             =   4320
   End
   Begin VB.PictureBox picScene 
      AutoRedraw      =   -1  'True
      Height          =   1500
      Index           =   0
      Left            =   240
      Picture         =   "frmMain.frx":2803
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      Height          =   9015
      Left            =   0
      ScaleHeight     =   597
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   797
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   12015
      Begin VB.PictureBox picBuffer 
         AutoRedraw      =   -1  'True
         Height          =   1500
         Left            =   3375
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   96
         TabIndex        =   15
         Top             =   240
         Width           =   1500
      End
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         Height          =   1500
         Left            =   1800
         Picture         =   "frmMain.frx":2F76
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   96
         TabIndex        =   14
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Image imgSel 
      Enabled         =   0   'False
      Height          =   1500
      Left            =   3480
      Picture         =   "frmMain.frx":308D
      Top             =   1920
      Width           =   1500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selType As String
Dim picCount As Integer
Dim dirJump(0 To 3) As String
Dim curX As Integer
Dim curY As Integer

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 123 Then 'F12
    If frmDbg.Visible = False Then
        frmDbg.Visible = True
        frmDbg.Show
        Me.SetFocus
    Else
        frmDbg.Visible = False
        frmDbg.Hide
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then 'Left
    If evalMove("L") = True Then
        dirJump(0) = "L"
        tmrJump(0).Enabled = True
    End If
ElseIf KeyCode = 38 Then 'Up
    If evalMove("U") = True Then
        dirJump(0) = "U"
    tmrJump(0).Enabled = True
    End If
ElseIf KeyCode = 39 Then 'Right
    If evalMove("R") = True Then
        dirJump(0) = "R"
    tmrJump(0).Enabled = True
    End If
ElseIf KeyCode = 40 Then 'Down
    If evalMove("D") = True Then
        dirJump(0) = "D"
    tmrJump(0).Enabled = True
    End If
End If
End Sub

Private Function evalMove(ByVal strDir As String) As Boolean
If strDir = "L" Then
    If curY + 1 Mod 2 = 0 Then
        If curX = 0 Then
            evalMove = False
        ElseIf curY < mapWidth Then
            evalMove = True
        End If
    ElseIf curY < mapWidth And curX > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDir = "U" Then
    If curY + 1 Mod 2 = 0 Then
        If curX > 0 Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curX > 0 And curY > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDir = "R" Then
    If curY + 1 Mod 2 = 0 Then
        If curX = mapWidth Then
            evalMove = False
        ElseIf curY > 0 Then
            evalMove = True
        End If
    Else
        evalMove = False
    End If
ElseIf strDir = "D" Then
    If curY + 1 Mod 2 = 0 Then
        If curX > 0 And curY < mapHeight Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curX < mapWidth And curY < mapHeight Then
        evalMove = True
    Else
        evalMove = False
    End If
End If
End Function

Private Sub tmrJump_Timer(Index As Integer)
If Index = 0 Then
    If dirJump(Index) = "L" Then
        If curY + 1 Mod 2 = 0 Then
            curX = curX - 1
        End If
        curY = curY - 1
    ElseIf dirJump(Index) = "U" Then
        If curY + 1 Mod 2 = 0 Then
            curX = curX - 1
        End If
        curY = curY + 1
    ElseIf dirJump(Index) = "R" Then
        If curY + 1 Mod 2 = 1 Then
            curX = curX + 1
        End If
        curY = curY + 1
    ElseIf dirJump(Index) = "D" Then
        If curY + 1 Mod 2 = 1 Then
            curX = curX + 1
        End If
        curY = curY - 1
    End If
End If
End Sub

Private Sub Form_Load()
Call DrawMap(1)
selType = "B"
frmBg.Show
curX = 2
curY = 2
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
selType = "O"
'flTextBox.Visible = True
'cmdCancelTextBox.Visible = True
'flTextBox.Movie = App.Path + "\Images\GUI\TextBoxB.swf"
End Sub

Private Sub Form_Resize()
Call DrawMap(1)
End Sub

Private Sub tmrSel_Timer()
Static blnRev As Boolean
Call PaintSelector(picCount)
If blnRev = False Then
    picCount = picCount + 1
End If
If blnRev = True Then
    picCount = picCount - 1
End If
If picCount >= 4 Or picCount <= 0 Then
    If blnRev = False Then
        blnRev = True
    Else
        blnRev = False
    End If
End If
End Sub

Private Function PaintSelector(ByVal imgIndex As Integer) As Integer
    'frmMain.PaintPicture picBackground.Image, Tile(intTileX, 0).oldX, Tile(0, intTileY).oldY, 100, 100, 0, 0, 100, 100, vbSrcCopy
    picBackground.PaintPicture frmMain.picScene(0).Image, Tile(curX, curY).X, Tile(curX, curY).Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
    picBuffer.PaintPicture frmMain.picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
    
    frmMain.PaintPicture frmMain.picMask.Image, Tile(curX, curY).X, Tile(curX, curY).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    frmMain.PaintPicture picBuffer.Image, Tile(curX, curY).X, Tile(curX, curY).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    'picBuffer.PaintPicture picBackground.Image, Tile(curX,curY).X, Tile(curX,curY).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    frmMain.PaintPicture picSelMask(imgIndex).Image, Tile(curX, curY).X, Tile(curX, curY).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    frmMain.PaintPicture picSelB(imgIndex).Image, Tile(curX, curY).X, Tile(curX, curY).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint

    'Tile(intTileX, intTileY).oldX = Tile(intTileX, intTileY).X
    'Tile(intTileX, intTileY).oldY = Tile(intTileX, intTileY).Y
End Function

Private Function limitVal(ByVal useX As Boolean, ByVal intCoord As Integer) As Integer
If useX = True Then
    If intCoord < 0 Then
        limitVal = 0
    Else
        limitVal = intCoord
    End If
Else
    If intCoord < 0 Then
        limitVal = 0
    Else
        limitVal = intCoord
    End If
End If
End Function
