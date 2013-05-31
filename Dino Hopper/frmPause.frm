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
Call ChangeTimers("Go")
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
    Call ChangeTimers("Pause")
End Sub
Private Function ChangeTimers(ByVal io As String)
Dim setVal As Boolean
If io = "Pause" Then
    setVal = False
Else
    setVal = True
End If
With frmMain
    If .tmrAlternate.Enabled = True Or .tmrAlternate.Tag = "True" Then
        .tmrAlternate.Tag = .tmrAlternate.Enabled
        .tmrAlternate.Enabled = setVal
    End If
    If .tmrObj.Enabled = True Or .tmrObj.Tag = "True" Then
        .tmrObj.Tag = .tmrObj.Enabled
        .tmrObj.Enabled = setVal
    End If
    If .tmrObjEvent.Enabled = True Or .tmrObjEvent.Tag = "True" Then
        .tmrObjEvent.Tag = .tmrObjEvent.Enabled
        .tmrObjEvent.Enabled = setVal
    End If
    If .tmrTileAnim.Enabled = True Or .tmrTileAnim.Tag = "True" Then
        .tmrTileAnim.Tag = .tmrTileAnim.Enabled
        .tmrTileAnim.Enabled = setVal
    End If
    If .tmrTileAnimDelay.Enabled = True Or .tmrTileAnimDelay.Tag = "True" Then
        .tmrTileAnimDelay.Tag = .tmrTileAnimDelay.Enabled
        .tmrTileAnimDelay.Enabled = setVal
    End If
    For x = 0 To 3
        If .tmrChar(x).Enabled = True Or .tmrChar(x).Tag = "True" Then
            .tmrChar(x).Tag = .tmrChar(x).Enabled
            .tmrChar(x).Enabled = setVal
        End If
        If .tmrPow(x).Enabled = True Or .tmrPow(x).Tag = "True" Then
            .tmrPow(x).Tag = .tmrPow(x).Enabled
            .tmrPow(x).Enabled = setVal
        End If
    Next x
    For x = 1 To 3
        If .tmrCPUMove(x).Enabled = True Or .tmrCPUMove(x).Tag = "True" Then
            .tmrCPUMove(x).Tag = .tmrCPUMove(x).Enabled
            .tmrCPUMove(x).Enabled = setVal
        End If
    Next x
    
End With
End Function
