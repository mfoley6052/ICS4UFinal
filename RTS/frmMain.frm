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
   Begin VB.Timer tmrScoreCheck 
      Interval        =   250
      Left            =   5040
      Top             =   360
   End
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   250
      Left            =   7920
      Top             =   360
   End
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   250
      Left            =   7440
      Top             =   360
   End
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   6960
      Top             =   360
   End
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   250
      Left            =   6480
      Top             =   360
   End
   Begin VB.Timer tmrSel 
      Enabled         =   0   'False
      Interval        =   125
      Left            =   5760
      Top             =   360
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      Height          =   9015
      Left            =   0
      ScaleHeight     =   597
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   749
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11295
      Begin VB.PictureBox picSelB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   0
         Left            =   1800
         Picture         =   "frmMain.frx":0000
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   15
         Top             =   1800
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   3360
         Picture         =   "frmMain.frx":0162
         ScaleHeight     =   98
         ScaleMode       =   0  'User
         ScaleWidth      =   102.083
         TabIndex        =   14
         Top             =   3360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   2
         Left            =   4920
         Picture         =   "frmMain.frx":02C1
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   3
         Left            =   6480
         Picture         =   "frmMain.frx":042C
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   4
         Left            =   8040
         Picture         =   "frmMain.frx":05BA
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   0
         Left            =   1800
         Picture         =   "frmMain.frx":072E
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   10
         Top             =   3360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   3360
         Picture         =   "frmMain.frx":0870
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   2
         Left            =   4920
         Picture         =   "frmMain.frx":09B3
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   3
         Left            =   6480
         Picture         =   "frmMain.frx":0AF4
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   7
         Top             =   3360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   4
         Left            =   8040
         Picture         =   "frmMain.frx":0C33
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   6
         Top             =   3360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picScene 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   240
         Picture         =   "frmMain.frx":0D70
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   5
         Top             =   3360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picScene 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   2
         Left            =   240
         Picture         =   "frmMain.frx":1FE6
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   4
         Top             =   1800
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picScene 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":2803
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picBuffer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   3375
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   2
         Top             =   240
         Width           =   1500
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   1800
         Picture         =   "frmMain.frx":2F76
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   1
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Label lblScore 
      Caption         =   "Score: 0000000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   16
      Top             =   120
      Width           =   2055
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
Dim intScore As Integer
Dim curX As Integer
Dim curY As Integer
Dim prevX As Integer
Dim prevY As Integer

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
    If (curY + 1) Mod 2 = 0 Then
        If curX = 0 Then
            evalMove = False
        ElseIf curY < mapWidth Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curY > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDir = "U" Then
    If (curY + 1) Mod 2 = 0 Then
        If curX < mapWidth Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curX < (mapWidth) And curY > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDir = "R" Then
    If (curY + 1) Mod 2 = 0 Then
        If curX = mapWidth Then
            evalMove = False
        ElseIf curY > 0 Then
            evalMove = True
        End If
    ElseIf curY < (mapWidth - 1) And curX < mapWidth Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDir = "D" Then
    If (curY + 1) Mod 2 = 0 Then
        If curX > 0 And curY < mapHeight Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curY < (mapHeight - 1) Then
        evalMove = True
    Else
        evalMove = False
    End If
End If
End Function

Private Sub tmrJump_Timer(Index As Integer)
If Index = 0 Then
    prevX = curX
    prevY = curY
    If dirJump(Index) = "L" Then
        If (curY + 1) Mod 2 = 0 Then
            curX = curX - 1
        End If
        curY = curY - 1
    ElseIf dirJump(Index) = "U" Then
        If (curY + 1) Mod 2 = 1 Then
            curX = curX + 1
        End If
        curY = curY - 1
    ElseIf dirJump(Index) = "R" Then
        If (curY + 1) Mod 2 = 1 Then
            curX = curX + 1
        End If
        curY = curY + 1
    ElseIf dirJump(Index) = "D" Then
        If (curY + 1) Mod 2 = 0 Then
            curX = curX - 1
        End If
        curY = curY + 1
    End If
    tmrJump(0).Enabled = False
    Call addScore(10)
End If
End Sub

Private Sub Form_Load()
Call DrawMap(1)
curX = 2
curY = 2
tmrSel.Enabled = True
End Sub

Private Sub addScore(ByVal intAdd As Integer)
intScore = intScore + intAdd
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

Private Sub tmrScoreCheck_Timer()
lblScore = "Score " & Format(intScore, "0000000")
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
    'paint over last sel
    picBackground.PaintPicture frmMain.picScene(0).Image, Tile(prevX, prevY).X, Tile(prevX, prevY).Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
    picBuffer.PaintPicture frmMain.picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
    frmMain.PaintPicture frmMain.picMask.Image, Tile(prevX, prevY).X, Tile(prevX, prevY).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    frmMain.PaintPicture picBuffer.Image, Tile(prevX, prevY).X, Tile(prevX, prevY).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    'paint over frame
    picBackground.PaintPicture frmMain.picScene(0).Image, Tile(curX, curY).X, Tile(curX, curY).Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
    picBuffer.PaintPicture frmMain.picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
    frmMain.PaintPicture frmMain.picMask.Image, Tile(curX, curY).X, Tile(curX, curY).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    frmMain.PaintPicture picBuffer.Image, Tile(curX, curY).X, Tile(curX, curY).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    'paint sel
    frmMain.PaintPicture picSelMask(imgIndex).Image, Tile(curX, curY).X, Tile(curX, curY).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    frmMain.PaintPicture picSelB(imgIndex).Image, Tile(curX, curY).X, Tile(curX, curY).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    frmDbg.txtTest(0).Text = "(" & curX & ", " & curY & ")"
End Function
