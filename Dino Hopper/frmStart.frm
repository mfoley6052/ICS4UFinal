VERSION 5.00
Begin VB.Form frmStart 
   AutoRedraw      =   -1  'True
   Caption         =   "Dino Hopper"
   ClientHeight    =   9000
   ClientLeft      =   2535
   ClientTop       =   225
   ClientWidth     =   12000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTitleMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   1470
      Picture         =   "frmStart.frx":0000
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   604
      TabIndex        =   0
      Top             =   1500
      Visible         =   0   'False
      Width           =   9060
   End
   Begin VB.PictureBox picTitleMainMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   1470
      Picture         =   "frmStart.frx":6478
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   604
      TabIndex        =   1
      Top             =   1500
      Visible         =   0   'False
      Width           =   9060
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   840
      Index           =   2
      Left            =   8250
      TabIndex        =   4
      Top             =   4290
      Width           =   4500
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   1
      Left            =   3750
      TabIndex        =   3
      Top             =   4290
      Width           =   4500
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   840
      Index           =   0
      Left            =   -750
      TabIndex        =   2
      Top             =   4320
      Width           =   4500
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ReDim DefaultKey(4) As Integer
ReDim key(4) As Integer
key(0) = 37
key(1) = 38
key(2) = 39
key(3) = 40
key(4) = 32
DefaultKey(0) = 37
DefaultKey(1) = 38
DefaultKey(2) = 39
DefaultKey(3) = 40
DefaultKey(4) = 32
lblMenu(0).Caption = "Options"
lblMenu(0).Tag = 1
lblMenu(1).Caption = "Single Player"
lblMenu(1).Tag = 0
lblMenu(2).Caption = "Exit"
lblMenu(2).Tag = 2
frmStart.PaintPicture picTitleMainMask.Image, picTitleMainMask.Left, picTitleMainMask.Top, picTitleMainMask.Width, picTitleMainMask.Height, 0, 0, picTitleMainMask.Width, picTitleMainMask.Height, vbSrcAnd
frmStart.PaintPicture picTitleMain.Image, picTitleMain.Left, picTitleMain.Top, picTitleMain.Width, picTitleMain.Height, 0, 0, picTitleMain.Width, picTitleMain.Height, vbSrcPaint
End Sub

Private Sub lblMenu_Click(index As Integer)
If index = 0 Then 'left
    For x = 0 To 2
        If Val(lblMenu(x).Tag) < 2 Then
            lblMenu(x).Tag = Val(lblMenu(x).Tag) + 1
        Else
            lblMenu(x).Tag = 0
        End If
        If lblMenu(x).Tag = 0 Then
            lblMenu(x).Caption = "Single Player"
        ElseIf lblMenu(x).Tag = 1 Then
            lblMenu(x).Caption = "Options"
        Else
            lblMenu(x).Caption = "Exit"
        End If
    Next x
ElseIf index = 2 Then 'right
    For x = 0 To 2
        If Val(lblMenu(x).Tag) > 0 Then
            lblMenu(x).Tag = Val(lblMenu(x).Tag) - 1
        Else
            lblMenu(x).Tag = 2
        End If
        If lblMenu(x).Tag = 0 Then
            lblMenu(x).Caption = "Single Player"
        ElseIf lblMenu(x).Tag = 1 Then
            lblMenu(x).Caption = "Options"
        Else
            lblMenu(x).Caption = "Exit"
        End If
    Next x
Else ' Select
    If lblMenu(1).Tag = 2 Then
        End
    ElseIf lblMenu(1).Tag = 1 Then
        frmSettings.Show
    Else
        frmStart.Hide
        frmMain.Show
    End If
End If
End Sub
