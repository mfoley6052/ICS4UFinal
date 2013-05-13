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
   Begin VB.TextBox txtMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "@MS Gothic"
         Size            =   27.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4560
      Width           =   4215
   End
   Begin VB.TextBox txtMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "@MS Gothic"
         Size            =   27.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4560
      Width           =   4215
   End
   Begin VB.TextBox txtMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "@MS Gothic"
         Size            =   27.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4560
      Width           =   4215
   End
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
txtMenu(0).Text = "Options"
txtMenu(0).Tag = 1
txtMenu(1).Text = "Single Player"
txtMenu(1).Tag = 0
txtMenu(2).Text = "Exit"
txtMenu(2).Tag = 2
frmStart.PaintPicture picTitleMainMask.Image, picTitleMainMask.Left, picTitleMainMask.Top, picTitleMainMask.Width, picTitleMainMask.Height, 0, 0, picTitleMainMask.Width, picTitleMainMask.Height, vbSrcAnd
frmStart.PaintPicture picTitleMain.Image, picTitleMain.Left, picTitleMain.Top, picTitleMain.Width, picTitleMain.Height, 0, 0, picTitleMain.Width, picTitleMain.Height, vbSrcPaint
End Sub

Private Sub txtMenu_Click(Index As Integer)

If Index = 0 Then 'left
    For x = 0 To 2
        If Val(txtMenu(x).Tag) < 2 Then
            txtMenu(x).Tag = Val(txtMenu(x).Tag) + 1
        Else
            txtMenu(x).Tag = 0
        End If
        If txtMenu(x).Tag = 0 Then
            txtMenu(x).Text = "Single Player"
        ElseIf txtMenu(x).Tag = 1 Then
            txtMenu(x).Text = "Options"
        Else
            txtMenu(x).Text = "Exit"
        End If
    Next x
ElseIf Index = 2 Then
    For x = 0 To 2
        If Val(txtMenu(x).Tag) > 0 Then
            txtMenu(x).Tag = Val(txtMenu(x).Tag) - 1
        Else
            txtMenu(x).Tag = 2
        End If
        If txtMenu(x).Tag = 0 Then
            txtMenu(x).Text = "Single Player"
        ElseIf txtMenu(x).Tag = 1 Then
            txtMenu(x).Text = "Options"
        Else
            txtMenu(x).Text = "Exit"
        End If
    Next x
Else
    If txtMenu(1).Tag = 2 Then
        End
    ElseIf txtMenu(1).Tag = 1 Then
        frmSettings.Show
    Else
        frmStart.Hide
        frmMain.Show
    End If
End If
End Sub
