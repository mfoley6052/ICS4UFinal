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
      Left            =   4440
      TabIndex        =   3
      Top             =   5280
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
      Left            =   4440
      TabIndex        =   2
      Top             =   4560
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
      Left            =   4440
      TabIndex        =   1
      Top             =   3840
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

'Sets any thing that is vbCyan to transparent and sets the background to cyan
Private Sub Form_Load()
    Me.BackColor = vbCyan
    lblGo.BackColor = vbCyan
    lblOpt.BackColor = vbCyan
    lblQuit.BackColor = vbCyan
    lblMenu.BackColor = vbCyan
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
    For X = 0 To 3
        If .tmrChar(X).Enabled = True Or tempTag(5 + X) = "True" Then
            tempTag(5 + X) = .tmrChar(X).Enabled
            .tmrChar(X).Enabled = setVal
        End If
        If .tmrPow(X).Enabled = True Or tempTag(9 + X) = "True" Then
            tempTag(9 + X) = .tmrPow(X).Enabled
            .tmrPow(X).Enabled = setVal
        End If
    Next X
    For X = 1 To 3
        If .tmrCPUMove(X).Enabled = True Or tempTag(12 + X) = "True" Then
            tempTag(12 + X) = .tmrCPUMove(X).Enabled
            .tmrCPUMove(X).Enabled = setVal
        End If
    Next X
    
End With
End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y > lblGo.Top And Y < (lblGo.Top + lblGo.Height) And X > lblGo.Left And X < lblGo.Left + lblGo.Width Then
    Call lblGo_Click
End If
End Sub

Private Sub lblGo_Click()
Call ChangeTimers("Go")
Me.Hide
frmMain.Show
Unload Me
End Sub

Private Sub lblMenu_Click()
frmStart.Show
Unload Me
End Sub

Private Sub lblOpt_Click()
frmSettings.Show vbModal
End Sub

Private Sub lblQuit_Click()
End
End Sub
