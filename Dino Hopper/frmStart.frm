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
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   5
      Left            =   3720
      TabIndex        =   7
      Top             =   7920
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   4
      Left            =   3720
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   3
      Left            =   3720
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
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
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   1
      Left            =   3720
      TabIndex        =   3
      Top             =   4290
      Width           =   4500
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
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
ReDim DefaultKey(14) As Integer
ReDim key(14) As Integer
key(0) = vbKeyLeft
key(1) = vbKeyUp
key(2) = vbKeyRight
key(3) = vbKeyDown
key(4) = vbKeyReturn
key(5) = vbKeyA
key(6) = vbKeyW
key(7) = vbKeyD
key(8) = vbKeyS
key(9) = vbKeySpace
key(10) = vbKeyNumpad4
key(11) = vbKeyNumpad8
key(12) = vbKeyNumpad6
key(13) = vbKeyNumpad5
key(14) = vbKeySeparator
DefaultKey(0) = vbKeyLeft
DefaultKey(1) = vbKeyUp
DefaultKey(2) = vbKeyRight
DefaultKey(3) = vbKeyDown
DefaultKey(4) = vbKeyReturn
DefaultKey(5) = vbKeyA
DefaultKey(6) = vbKeyW
DefaultKey(7) = vbKeyD
DefaultKey(8) = vbKeyS
DefaultKey(9) = vbKeySpace
DefaultKey(10) = vbKeyNumpad4
DefaultKey(11) = vbKeyNumpad8
DefaultKey(12) = vbKeyNumpad6
DefaultKey(13) = vbKeyNumpad5
DefaultKey(14) = vbKeySeparator
lblMenu(0).Caption = "MultiPlayer"
lblMenu(1).Caption = "Single Player"
lblMenu(2).Caption = "Options"
frmStart.PaintPicture picTitleMainMask.Image, picTitleMainMask.Left, picTitleMainMask.Top, picTitleMainMask.Width, picTitleMainMask.Height, 0, 0, picTitleMainMask.Width, picTitleMainMask.Height, vbSrcAnd
frmStart.PaintPicture picTitleMain.Image, picTitleMain.Left, picTitleMain.Top, picTitleMain.Width, picTitleMain.Height, 0, 0, picTitleMain.Width, picTitleMain.Height, vbSrcPaint
End Sub

Private Sub lblMenu_Click(index As Integer)
If index = 0 Then 'left
    lblMenu(3).Visible = False
    lblMenu(4).Visible = False
    lblMenu(5).Visible = False
    For X = 0 To 2
        If lblMenu(X).Caption = "Options" Then
            lblMenu(X).Caption = "Single Player"
            
        ElseIf lblMenu(X).Caption = "Single Player" Then
            lblMenu(X).Caption = "MultiPlayer"
            
        ElseIf lblMenu(X).Caption = "MultiPlayer" Then
            lblMenu(X).Caption = "Exit"
            
        Else
            lblMenu(X).Caption = "Options"
        End If
    Next X
ElseIf index = 2 Then 'right
    lblMenu(3).Visible = False
    lblMenu(4).Visible = False
    lblMenu(5).Visible = False
    For X = 0 To 2
        If lblMenu(X).Caption = "MultiPlayer" Then
            lblMenu(X).Caption = "Single Player"
            
        ElseIf lblMenu(X).Caption = "Exit" Then
            lblMenu(X).Caption = "MultiPlayer"
            
        ElseIf lblMenu(X).Caption = "Options" Then
            lblMenu(X).Caption = "Exit"
            
        Else
            lblMenu(X).Caption = "Options"
        End If
    Next X
ElseIf index = 1 Then ' Select
    If lblMenu(1).Caption = "Exit" Then
        End
    ElseIf lblMenu(1).Caption = "Exit" Then
        frmSettings.Show
    ElseIf lblMenu(1).Caption = "Single Player" Then
        lblMenu(3).Visible = True
        lblMenu(3).Caption = "Arcade Mode"
        lblMenu(4).Caption = "Turn-Based"
        lblMenu(5).Caption = "Puzzle Mode"
        lblMenu(5).Visible = True
        lblMenu(4).Visible = True
    ElseIf lblMenu(1).Caption = "MultiPlayer" Then
        lblMenu(3).Visible = True
        lblMenu(3).Caption = "Arcade Mode"
        lblMenu(4).Caption = "Turn-Based"
        lblMenu(4).Visible = True
        
    End If
ElseIf index = 3 Then
        gameMode = 0
        frmStart.Hide
        frmMain.Show
ElseIf index = 4 Then
        gameMode = 1
        frmStart.Hide
        frmMain.Show
ElseIf index = 5 Then
        gameMode = 2
        frmStart.Hide
        frmMain.Show
End If
End Sub
