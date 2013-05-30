VERSION 5.00
Begin VB.Form frmPause 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Paused"
   ClientHeight    =   7635
   ClientLeft      =   6360
   ClientTop       =   2205
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11565
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   5280
      Width           =   3375
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Main Menu"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   4560
      Width           =   3375
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Options"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   3840
      Width           =   3375
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Continue"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   3120
      Width           =   3375
   End
End
Attribute VB_Name = "frmPause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A bunch of complicated functions that I don't quite understand, it seems to use windows built in functions
'I got this off the internet
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Sub cmdGo_Click()
frmMain.Show
Unload Me
'Activate the timers and resume playing
End Sub

'Go back to the menu
Private Sub cmdMenu_Click()
frmStart.Show
Unload Me
End Sub

Private Sub cmdOpt_Click()
frmSettings.Show vbModal
End Sub

'Exit the game
Private Sub cmdQuit_Click()
End
End Sub
'Sets any thing that is vbCyan to transparent and sets the background to cyan
Private Sub Form_Load()
    Me.BackColor = vbCyan
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY
    frmPause.Left = frmMain.Left
    frmPause.Height = frmMain.Height
    frmPause.Width = frmMain.Width
    frmPause.Top = frmMain.Top
End Sub

