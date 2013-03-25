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
   Begin VB.Timer tmrTileDraw 
      Interval        =   1
      Left            =   6360
      Top             =   4560
   End
   Begin VB.PictureBox picSelMask 
      Height          =   810
      Index           =   4
      Left            =   8040
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Timer tmrSel 
      Interval        =   125
      Left            =   5400
      Top             =   4320
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      Height          =   1500
      Left            =   3480
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      Height          =   1500
      Left            =   1800
      Picture         =   "frmMain.frx":2803
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picScene 
      AutoRedraw      =   -1  'True
      Height          =   1500
      Index           =   0
      Left            =   240
      Picture         =   "frmMain.frx":291A
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
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   12015
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
Dim intTileXinit As Integer
Dim intTileYinit As Integer
Dim intTileX As Integer
Dim intTileY As Integer

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

Private Sub Form_Load()
Call DrawMap(1)
selType = "B"
frmBg.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
intTileXinit = X
intTileYinit = Y
End Sub

Private Sub cmdCancelTextBox_Click()
'flTextBox.Visible = False
'cmdCancelTextBox.Visible = False
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

Private Sub Form_Click()
MsgBox (intTileX & " " & intTileY)
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
    
    'frmMain.PaintPicture picBuffer.Image, Tile(intTileXinit / 100, 0).X, Tile(0, intTileYinit / 50).Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
    picBuffer.PaintPicture picBackground.Image, Tile(intTileXinit / 100, 0).X, Tile(0, intTileYinit / 50).Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
    frmMain.PaintPicture picSelMask(imgIndex).Image, Tile(intTileXinit / 100, 0).X, Tile(0, intTileYinit / 50).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    frmMain.PaintPicture picSelB(imgIndex).Image, Tile(intTileXinit / 100, 0).X, Tile(0, intTileYinit / 50).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint

    'Tile(intTileX, intTileY).oldX = Tile(intTileX, intTileY).X
    'Tile(intTileX, intTileY).oldY = Tile(intTileX, intTileY).Y
End Function

Private Sub tmrTileDraw_Timer()
Dim curTile As String
Static oldTile As String
Dim xStart As Integer
'Debug Form
frmDbg.txtSelType = "SelType: " & selType
frmDbg.txtX.Text = "X: " & intTileXinit
frmDbg.txtxMod.Text = intTileXinit Mod 100
frmDbg.txtYMod.Text = intTileYinit Mod 100
frmDbg.txtXd100 = intTileXinit / 100
frmDbg.txtYd50 = intTileYinit / 100
frmDbg.txtY.Text = "Y: " & intTileYinit
frmDbg.txtTest(0) = Int((intTileXinit Mod 50) / 2)
frmDbg.txtTest(1) = intTileYinit Mod 50
curTile = "(" & intTileX & "," & intTileY & ")"
'If Not (curTile = oldTile) Then
    If Tile(intTileXinit / 100, intTileYinit / 50).selectable = True Then
    
        'Tile Selection
        'imgSel.Top = Tile(0, intTileY).Y
        'imgSel.Left = Tile(intTileX, 0).X
        'Top half of odd tile
        If intTileYinit Mod 50 < 25 Then
            'Left half of odd tile
            If intTileXinit Mod 100 < 50 Then
                If 25 - Int((intTileXinit Mod 50) / 2) >= intTileYinit Mod 50 Then
                    'imgSel.Top = Tile(0, intTileY).Y - 25
                    'imgSel.Left = Tile(intTileX, 0).X - 50
                    'Call PaintSelector(picCount, Tile(intTileX, 0).X - 50, Tile(0, intTileY).Y - 25)
                    intTileX = Tile(intTileXinit / 100, 0).X - 50
                    intTileY = Tile(0, intTileYinit / 50).Y - 25
                    xStart = 0
                    frmDbg.txtSelectable.Text = "Tile(" & limitVal(True, (Tile(intTileXinit / 100, 0).X / 50) + 1) & "," & limitVal(False, (Tile(0, intTileYinit / 50).Y / 50) - 1) & "): " & Tile(limitVal(True, (Tile(intTileXinit / 100, 0).X / 50) + 1), limitVal(False, (Tile(0, intTileYinit / 50).Y / 50) - 1)).selectable
                Else
                    xStart = 50
                    frmDbg.txtSelectable.Text = "Tile(" & limitVal(True, (Tile(intTileXinit / 100, 0).X / 50)) & "," & limitVal(False, (Tile(0, intTileYinit / 50).Y / 50)) & "): " & Tile(limitVal(True, (Tile(intTileXinit / 100, 0).X / 50)), limitVal(False, (Tile(0, intTileYinit / 50).Y / 50))).selectable
                End If
            'Right half of odd tile
            ElseIf intTileXinit Mod 100 >= 50 Then
                If Int((intTileXinit Mod 50) / 2) >= intTileYinit Mod 50 Then
                    'imgSel.Top = Tile(0, intTileY).Y - 25
                    'imgSel.Left = Tile(intTileX, 0).X + 50
                    'Call PaintSelector(picCount, Tile(intTileX, 0).X + 50, Tile(0, intTileY).Y - 25)
                    intTileX = Tile(intTileXinit / 100, 0).X + 50
                    intTileY = Tile(0, intTileYinit / 50).Y - 25
                    xStart = 0
                    frmDbg.txtSelectable.Text = "Tile(" & limitVal(True, (Tile(intTileXinit / 100, 0).X / 50)) & "," & limitVal(False, (Tile(0, intTileYinit / 50).Y / 50) - 1) & "): " & Tile(limitVal(True, (Tile(intTileXinit / 100, 0).X / 50)), limitVal(False, (Tile(0, intTileYinit / 50).Y / 50) - 1)).selectable
                Else
                    xStart = 50
                    frmDbg.txtSelectable.Text = "Tile(" & limitVal(True, (Tile(intTileXinit / 100, 0).X / 50)) & "," & limitVal(False, (Tile(0, intTileYinit / 50).Y / 50) - 1) & "): " & Tile(limitVal(True, (Tile(intTileXinit / 100, 0).X / 50)), limitVal(False, (Tile(0, intTileYinit / 50).Y / 50) - 1)).selectable
                End If
            End If
        'Bottom half of even tile
        ElseIf intTileYinit Mod 50 >= 25 Then
            'Left half of odd tile
            If intTileXinit Mod 100 < 50 Then
                If 25 + Int((intTileXinit Mod 50) / 2) < intTileYinit Mod 50 Then
                    'imgSel.Top = Tile(0, intTileY).Y + 25
                    'imgSel.Left = Tile(intTileX, 0).X - 50
                    'Call PaintSelector(picCount, Tile(intTileX, 0).X - 50, Tile(0, intTileY).Y + 25)
                    intTileX = Tile(intTileXinit / 100, 0).X - 50
                    intTileY = Tile(0, intTileYinit / 50).Y + 25
                    xStart = 50
                    frmDbg.txtSelectable.Text = "Tile(" & limitVal(True, (Tile(intTileXinit / 100, 0).X / 50) - 1) & "," & limitVal(False, (Tile(0, intTileYinit / 50).Y / 50)) & "): " & Tile(limitVal(True, (Tile(intTileXinit / 100, 0).X / 50) - 1), limitVal(False, (Tile(0, intTileYinit / 50).Y / 50))).selectable
                Else
                    xStart = 0
                    frmDbg.txtSelectable.Text = "Tile(" & limitVal(True, (Tile(intTileXinit / 100, 0).X / 50)) & "," & limitVal(False, (Tile(0, intTileYinit / 50).Y / 50)) & "): " & Tile(limitVal(True, (Tile(intTileXinit / 100, 0).X / 50)), limitVal(False, (Tile(0, intTileYinit / 50).Y / 50))).selectable
                End If
            'Right half of odd tile
            ElseIf intTileXinit Mod 100 >= 50 Then
                If 50 - Int((intTileXinit Mod 50) / 2) < intTileYinit Mod 50 Then
                    'imgSel.Top = Tile(0, intTileY).Y + 25
                    'imgSel.Left = Tile(intTileX, 0).X + 50
                    'Call PaintSelector(picCount, Tile(intTileX, 0).X + 50, Tile(0, intTileY).Y + 25)
                    intTileX = Tile(intTileXinit / 100, 0).X + 50
                    intTileY = Tile(0, intTileYinit / 50).Y + 25
                    xStart = 50
                    frmDbg.txtSelectable.Text = "Tile(" & limitVal(True, (Tile(intTileXinit / 100, 0).X / 50) + 1) & "," & limitVal(False, (Tile(0, intTileYinit / 50).Y / 50)) & "): " & Tile(limitVal(True, (Tile(intTileXinit / 100, 0).X / 50) + 1), limitVal(False, (Tile(0, intTileYinit / 50).Y / 50))).selectable
                Else
                    xStart = 0
                    frmDbg.txtSelectable.Text = "Tile(" & limitVal(True, (Tile(intTileXinit / 100, 0).X / 50)) & "," & limitVal(False, (Tile(0, intTileYinit / 50).Y / 50)) & "): " & Tile(limitVal(True, (Tile(intTileXinit / 100, 0).X / 50)), limitVal(False, (Tile(0, intTileYinit / 50).Y / 50))).selectable
                End If
            End If
        End If
    End If
'End If
oldTile = curTile

End Sub

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
