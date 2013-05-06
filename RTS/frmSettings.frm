VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   5550
   ClientLeft      =   6450
   ClientTop       =   3285
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fControl 
      Caption         =   "Controls"
      Height          =   5415
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdDef 
         Caption         =   "Default"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdMap 
         Caption         =   "Rebind Keys"
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Tag             =   "lock"
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   9
         Text            =   "Down Arrow"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   8
         Text            =   "Right Arrow"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   7
         Text            =   "Up Arrow"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtCTR 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   6
         Text            =   "Left Arrow"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Down"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Up"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Right"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Left"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame fGen 
      Caption         =   "General"
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDef_Click()
key(0) = 37
key(1) = 38
key(2) = 39
key(3) = 40
txtCTR(0).Text = "Left Arrow"
txtCTR(1).Text = "Up Arrow"
txtCTR(2).Text = "Right Arrow"
txtCTR(3).Text = "Down Arrow"
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

Private Sub Form_Load()
For x = txtCTR.LBound To txtCTR.UBound
    If key(x) <> DefaultKey(x) Then
        txtCTR(x).Text = Chr(key(x))
    End If
Next x
End Sub

Private Sub txtCTR_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
key(Index) = KeyCode
txtCTR(Index).Enabled = False
txtCTR(Index).Text = Chr(KeyCode)
'MsgBox (KeyCode & ": " & Chr(KeyCode))

End Sub
