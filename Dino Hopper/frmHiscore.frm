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
   Begin VB.ComboBox cmbPlayMode 
      Height          =   315
      ItemData        =   "frmHiscore.frx":0000
      Left            =   2160
      List            =   "frmHiscore.frx":000D
      TabIndex        =   25
      Text            =   "SP"
      Top             =   240
      Width           =   1335
   End
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
      ItemData        =   "frmHiscore.frx":001F
      Left            =   720
      List            =   "frmHiscore.frx":002C
      TabIndex        =   23
      Text            =   "Arcade"
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox cmbPlayers 
      Height          =   315
      ItemData        =   "frmHiscore.frx":004C
      Left            =   120
      List            =   "frmHiscore.frx":0059
      TabIndex        =   22
      Text            =   "1"
      Top             =   240
      Width           =   495
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
Dim score(9) As Record
Private Type Record
    score As Long
    nam As String * 3
End Type

Private Sub cmdBack_Click()
frmStart.Show
Unload Me
End Sub

Private Sub cmdLoad_Click()
playerChoice = cmbPlayers.Text
gameMode = cmbGameMode.Text
playMode = cmbPlayMode.Text
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
If playMode <> "SOLO" Then
    Open App.Path & "\Scores\" & playMode & "\" & gameMode & "\" & playerChoice & ".sav" For Input As #1
Else
    Open App.Path & "\Scores\" & playMode & "\" & playerChoice & ".sav" For Input As #1
End If
For x = 0 To 9
    txtNam(x).Text = ""
    txtScore(x).Text = ""
    If Not EOF(1) Then
        Line Input #1, temp
        txtNam(x).Text = temp
        score(x).nam = Val(temp)
        Line Input #1, temp
        score(x).score = Val(temp)
        txtScore(x).Text = temp
    End If
Next x
Close #1
End Sub

Public Sub WriteScore(ByVal playerName As String, ByVal Pscore As Long)
Dim dump As String
Dim count As Integer
Dim sorted As Boolean
Dim upper As Long
Dim temp As Record
Open App.Path & "\Scores\" & playMode & "\" & gameMode & "\" & playerChoice & ".sav" For Input As #1
    Do Until EOF(1)
        Line Input #1, dump
        count = count + 1
    Loop
Close #1
    If count <= 16 Then
        Open App.Path & "\Scores\" & playMode & "\" & gameMode & "\" & playerChoice & ".sav" For Append As #1
            Print #1, playerName
            Print #1, Pscore
        Close #1
    Else
        'Find the lowest score, compare it to the score to be added, and then choose which one to keep
        Open App.Path & "\Scores\" & playMode & "\" & gameMode & "\" & playerChoice & ".sav" For Input As #1
        For x = 0 To 9
            Line Input #1, dump
            score(x).nam = dump
            Line Input #1, dump
            score(x).score = dump
        Next x
        Close #1
        Do Until sorted
            sorted = True
            For y = LBound(score) To upper 'Step - 1
                If score(y + 1).score < score(y).score Then
                    temp.score = score(y).score
                    temp.nam = score(y).nam
                    score(y).score = score(y + 1).score
                    score(y).nam = score(y + 1).nam
                    score(y + 1).score = temp.score
                    score(y + 1).nam = temp.nam
                    sorted = False
                End If
            Next y
            upper = upper - 1
        Loop
        If score(0).score < Pscore Then
            'Delete the low score AND the name of the player that got that score, then add the new score
            
        Else
            MsgBox ("Score didnt get into top 10")
        End If
    End If
End Sub
