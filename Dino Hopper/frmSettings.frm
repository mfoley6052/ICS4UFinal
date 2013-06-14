VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   5550
   ClientLeft      =   6450
   ClientTop       =   3285
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fControl 
      Caption         =   "Controls (P3)"
      Height          =   5415
      Index           =   2
      Left            =   4800
      TabIndex        =   24
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdScores 
         Caption         =   "Show High Scores"
         Height          =   375
         Left            =   840
         TabIndex        =   35
         Top             =   4920
         Width           =   1575
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   840
         TabIndex        =   29
         Text            =   "Numpad 4"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   840
         TabIndex        =   28
         Text            =   "Numpad 8"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   840
         TabIndex        =   27
         Text            =   "Numpad 6"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   840
         TabIndex        =   26
         Text            =   "Numpad 5"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   840
         TabIndex        =   25
         Text            =   "Numpad Return"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Left"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Right"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Up"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Down"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Action"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   615
      End
   End
   Begin VB.Frame fControl 
      Caption         =   "Controls (P2)"
      Height          =   5415
      Index           =   1
      Left            =   2400
      TabIndex        =   13
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   480
         TabIndex        =   36
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   840
         TabIndex        =   18
         Text            =   "A"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   840
         TabIndex        =   17
         Text            =   "W"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   840
         TabIndex        =   16
         Text            =   "D"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   840
         TabIndex        =   15
         Text            =   "S"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   840
         TabIndex        =   14
         Text            =   "Space Bar"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Left"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Right"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Up"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Down"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Action"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   615
      End
   End
   Begin VB.Frame fControl 
      Caption         =   "Controls (P1)"
      Height          =   5415
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   12
         Text            =   "Return"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdDef 
         Caption         =   "Default"
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdMap 
         Caption         =   "Rebind Keys"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Tag             =   "lock"
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   8
         Text            =   "Down Arrow"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   7
         Text            =   "Right Arrow"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   6
         Text            =   "Up Arrow"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Text            =   "Left Arrow"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Action"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Down"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Up"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Right"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Left"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDef_Click()
For x = LBound(key) To UBound(key)
    key(x) = DefaultKey(x)
    txtCTR(x).Text = txtCTR(x).Tag
Next x
End Sub

Private Sub cmdHelp_Click()
frmHelp.Show vbModal
End Sub

Private Sub cmdMap_Click()
If cmdMap.Tag = "lock" Then
    cmdMap.Tag = "edit"
    For x = txtCTR.LBound To txtCTR.UBound
        txtCTR(x).Enabled = True
    Next x
    cmdMap.Caption = "Save Binding"
Else
    cmdMap.Tag = "lock"
    For x = txtCTR.LBound To txtCTR.UBound
        txtCTR(x).Enabled = False
    Next x
    cmdMap.Caption = "Rebind Keys"
End If
End Sub

Private Sub cmdScores_Click()
frmHiscore.Show vbModal
Me.Hide
End Sub

Private Sub Form_Load()
For x = txtCTR.LBound To txtCTR.UBound
    If key(x) <> DefaultKey(x) Then
        txtCTR(x).Text = Chr(key(x))
    Else
        txtCTR(x).Tag = txtCTR(x).Text
    End If
Next x
End Sub

Private Sub txtCTR_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
key(Index) = KeyCode
txtCTR(Index).Enabled = False
txtCTR(Index).Text = Chr(KeyCode)
'MsgBox (KeyCode & ": " & Chr(KeyCode))
End Sub
