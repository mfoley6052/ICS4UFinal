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
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   3000
      TabIndex        =   26
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
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

Dim ran As Boolean

Private Sub cmdAdd_Click()
'for testing
Call WriteScore(InputBox("playerName: "), InputBox("score"))
Call LoadScore
End Sub

Private Sub cmdBack_Click()
'back to main menu
Me.Hide
frmStart.Show
End Sub

Private Sub cmdLoad_Click()
'reset variables
For x = 0 To 9
    score(x).score = 0
    score(x).nam = ""
    txtNam(x).Text = ""
    txtScore(x).Text = ""
Next x
'get dropdown choices
playerChoice = cmbPlayers.Text
gMode = cmbGameMode.Text
playMode = cmbPlayMode.Text
'load the score
Call LoadScore
End Sub

Private Sub cmdReset_Click()
Dim temp As Integer
'make sure they want to erase the hiscores
temp = MsgBox("Are you sure you want to reset the hiscore?", vbYesNo)
If temp = vbYes Then
'Solo has a different path which is annoying, but its because it doesnt need as many save files
    If playMode <> "SOLO" Then
        Open App.Path & "\Scores\" & playMode & "\" & gMode & "\" & playerChoice & ".sav" For Output As #1
    Else
        Open App.Path & "\Scores\" & playMode & "\" & playerChoice & ".sav" For Output As #1
    End If
    'write a default score set
        For x = 0 To 9
            Print #1, "RST"
            Print #1, x * 1000
        Next x
    Close #1
End If
'load the scores
Call LoadScore
End Sub

Private Sub Form_Load()
'load the scores
Call LoadScore
End Sub

Private Sub LoadScore()
Dim temp As String
'the first time you load, use default values
If Not ran Then
    gMode = "Arcade"
    playerChoice = "1"
    playMode = "SP"
    ran = True
End If
'open file
If playMode <> "SOLO" Then
    frmHiscore.cmbGameMode.Enabled = True
    Open App.Path & "\Scores\" & playMode & "\" & gMode & "\" & playerChoice & ".sav" For Input As #1
Else
    frmHiscore.cmbGameMode.Enabled = False
    Open App.Path & "\Scores\" & playMode & "\" & playerChoice & ".sav" For Input As #1
End If
'clear text boxes, and read in the name and score of each record
For x = 0 To 9
    frmHiscore.txtNam(x).Text = ""
    frmHiscore.txtScore(x).Text = ""
    If Not EOF(1) Then
        Line Input #1, temp
        frmHiscore.txtNam(x).Text = temp
        score(x).nam = temp
        Line Input #1, temp
        score(x).score = Val(temp)
        frmHiscore.txtScore(x).Text = temp
    End If
Next x
Close #1
Dim sorted As Boolean
Dim upper As Long
Dim tempo As Record
upper = UBound(score) - 1
'better bubble sort on scores
Do Until sorted
    sorted = True
    For y = LBound(score) To upper
        If score(y + 1).score < score(y).score Then
            tempo.score = score(y).score
            tempo.nam = score(y).nam
            score(y).score = score(y + 1).score
            score(y).nam = score(y + 1).nam
            score(y + 1).score = tempo.score
            score(y + 1).nam = tempo.nam
            sorted = False
        End If
    Next y
    upper = upper - 1
Loop
Dim count As Integer
'output the scores to the textboxes
For x = 9 To 0 Step -1
frmHiscore.txtNam(count).Text = score(x).nam
frmHiscore.txtScore(count).Text = score(x).score
count = count + 1
Next x
End Sub
