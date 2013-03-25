VERSION 5.00
Begin VB.Form frmDbg 
   Caption         =   "Form1"
   ClientHeight    =   10590
   ClientLeft      =   7695
   ClientTop       =   1095
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   706
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   813
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtYMod2 
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   9720
      Width           =   1095
   End
   Begin VB.TextBox txtxMod2 
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   9720
      Width           =   1095
   End
   Begin VB.TextBox txtSelectable 
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   9000
      Width           =   1095
   End
   Begin VB.TextBox txtSelType 
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   9000
      Width           =   1095
   End
   Begin VB.TextBox txtXd100 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   10440
      Width           =   1095
   End
   Begin VB.TextBox txtYd50 
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   10440
      Width           =   1095
   End
   Begin VB.TextBox txtTest 
      Height          =   495
      Index           =   1
      Left            =   11040
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   9000
      Width           =   1095
   End
   Begin VB.TextBox txtTest 
      Height          =   495
      Index           =   0
      Left            =   9840
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   9000
      Width           =   1095
   End
   Begin VB.TextBox txtYMod 
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   9720
      Width           =   1095
   End
   Begin VB.TextBox txtY 
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   9000
      Width           =   1095
   End
   Begin VB.TextBox txtxMod 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   9720
      Width           =   1095
   End
   Begin VB.TextBox txtX 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   9000
      Width           =   1095
   End
   Begin VB.ListBox lstMap 
      Columns         =   8
      Height          =   8835
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12000
   End
   Begin VB.Label Label3 
      Caption         =   "mod 50"
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   9480
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "/100"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   10200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "mod 100"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   9480
      Width           =   2295
   End
End
Attribute VB_Name = "frmDbg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.WindowState = 2
End Sub
