VERSION 5.00
Begin VB.Form frmHiscore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Top Scores"
   ClientHeight    =   5850
   ClientLeft      =   9015
   ClientTop       =   4110
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   3600
      TabIndex        =   24
      Top             =   240
      Width           =   735
   End
   Begin VB.ComboBox cmbGameMode 
      Height          =   315
      ItemData        =   "frmHiscore.frx":0000
      Left            =   1560
      List            =   "frmHiscore.frx":000D
      TabIndex        =   23
      Text            =   "Arcade"
      Top             =   240
      Width           =   1935
   End
   Begin VB.ComboBox cmbPlayers 
      Height          =   315
      ItemData        =   "frmHiscore.frx":002D
      Left            =   120
      List            =   "frmHiscore.frx":003A
      TabIndex        =   22
      Text            =   "1"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Clear"
      Height          =   255
      Left            =   3720
      TabIndex        =   20
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox txtScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   5040
      Width           =   3255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4560
      Width           =   3255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txtNam 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   9
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtNam 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   8
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtNam 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   7
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtNam 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   6
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtNam 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtNam 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtNam 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtNam 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtNam 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtNam 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmHiscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gameMode As String
Dim playerChoice As String
Dim playMode As String

Private Sub cmdBack_Click()
frmStart.Show
Unload Me
End Sub

Private Sub cmdLoad_Click()
playerChoice = cmbPlayers.Text
gameMode = cmbGameMode.Text
Call LoadScore
End Sub

Private Sub Form_Load()
gameMode = "Arcade"
playerChoice = "1"
playMode = "SP"
Call LoadScore
End Sub

Private Sub LoadScore()
Dim temp As String
If Val(playerChoice) > 1 Then
    playMode = "MP"
Else
    playMode = "SP"
End If
If numCPU = 0 Then
    playMode = "SOLO"
End If
Open App.Path & "\Scores\" & playMode & "\" & gameMode & "\" & playerChoice & ".sav" For Input As #1
For x = 0 To 9
    If Not EOF(1) Then
        Line Input #1, temp
        txtNam(x).Text = temp
        Line Input #1, temp
        txtScore(x).Text = temp
    End If
Next x
Close #1
End Sub

Public Sub WriteScore()
Open App.Path & "\Scores\SP\" & gameMode & "\" & playerChoice & ".sav" For Output As #1

Close #1
End Sub
