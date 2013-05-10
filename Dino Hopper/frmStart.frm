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
   Begin VB.Label lblSP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Single Player"
      BeginProperty Font 
         Name            =   "@Small Fonts"
         Size            =   24
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3750
      TabIndex        =   2
      Top             =   4500
      Width           =   4500
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmStart.PaintPicture picTitleMainMask.Image, picTitleMainMask.Left, picTitleMainMask.Top, picTitleMainMask.Width, picTitleMainMask.Height, 0, 0, picTitleMainMask.Width, picTitleMainMask.Height, vbSrcAnd
frmStart.PaintPicture picTitleMain.Image, picTitleMain.Left, picTitleMain.Top, picTitleMain.Width, picTitleMain.Height, 0, 0, picTitleMain.Width, picTitleMain.Height, vbSrcPaint
End Sub

Private Sub lblSP_Click()
frmMain.Show
frmStart.Visible = False
End Sub
