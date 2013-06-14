VERSION 5.00
Begin VB.Form frmPause 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Game Paused"
   ClientHeight    =   9555
   ClientLeft      =   4515
   ClientTop       =   1230
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblQuit 
      Alignment       =   2  'Center
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label lblOpt 
      Alignment       =   2  'Center
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Label lblGo 
      Alignment       =   2  'Center
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   3360
      Width           =   3375
   End
End
Attribute VB_Name = "frmPause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'transparency declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

'Sets any thing that is vbCyan to transparent and sets the background to cyan
Private Sub Form_Load()
    Me.BackColor = vbCyan
    lblGo.BackColor = vbCyan
    lblOpt.BackColor = vbCyan
    lblQuit.BackColor = vbCyan
    lblMenu.BackColor = vbCyan
    SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, vbCyan, 0&, LWA_COLORKEY
    frmPause.Left = frmMain.Left + 50
    frmPause.Top = frmMain.Top + 370
    Call ChangeTimers("Pause")
End Sub

'pause game
Private Function ChangeTimers(ByVal io As String)
Dim setVal As Boolean
Static tempTag(15) As Boolean

If io = "Pause" Then
    setVal = False
Else
    setVal = True
End If
With frmMain
    If .tmrAlternate.Enabled = True Or tempTag(0) Then
        tempTag(0) = .tmrAlternate.Enabled
        .tmrAlternate.Enabled = setVal
    End If
    If .tmrObj.Enabled = True Or tempTag(1) Then
        tempTag(1) = .tmrObj.Enabled
        .tmrObj.Enabled = setVal
    End If
    If .tmrObjEvent.Enabled = True Or tempTag(2) Then
       tempTag(2) = .tmrObjEvent.Enabled
        .tmrObjEvent.Enabled = setVal
    End If
    If .tmrTileAnim.Enabled = True Or tempTag(3) Then
        tempTag(3) = .tmrTileAnim.Enabled
        .tmrTileAnim.Enabled = setVal
    End If
    If .tmrTileAnimDelay.Enabled = True Or tempTag(4) = "True" Then
        tempTag(4) = .tmrTileAnimDelay.Enabled
        .tmrTileAnimDelay.Enabled = setVal
    End If
    For x = 0 To 3
        If .tmrChar(x).Enabled = True Or tempTag(5 + x) = "True" Then
            tempTag(5 + x) = .tmrChar(x).Enabled
            .tmrChar(x).Enabled = setVal
        End If
        If .tmrPow(x).Enabled = True Or tempTag(9 + x) = "True" Then
            tempTag(9 + x) = .tmrPow(x).Enabled
            .tmrPow(x).Enabled = setVal
        End If
    Next x
    For x = 1 To 3
        If .tmrCPUMove(x).Enabled = True Or tempTag(12 + x) = "True" Then
            tempTag(12 + x) = .tmrCPUMove(x).Enabled
            .tmrCPUMove(x).Enabled = setVal
        End If
    Next x
    
End With
End Function

'check if game is to start from pause
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If y > lblGo.Top And y < (lblGo.Top + lblGo.Height) And x > lblGo.Left And x < lblGo.Left + lblGo.Width Then
    Call lblGo_Click
End If
End Sub

'start from pause
Public Sub lblGo_Click()
frmMain.tmrRefresh.Enabled = False
Call ChangeTimers("Go")
Me.Hide
frmGUI.SetFocus
Unload Me
End Sub

'go to menu and end game
Private Sub lblMenu_Click()
frmStart.Show
Call getGameEnd
frmStart.Show
frmMain.Hide
Unload frmMain
Set frmMain = Nothing
frmGUI.Hide
Unload frmGUI
Set frmGUI = Nothing
Unload Me
End Sub

'get options
Private Sub lblOpt_Click()
frmSettings.Show vbModal
End Sub

'quit game
Private Sub lblQuit_Click()
End
End Sub
