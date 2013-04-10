VERSION 5.00
Begin VB.Form frmDbg 
   Caption         =   "Form1"
   ClientHeight    =   10590
   ClientLeft      =   4065
   ClientTop       =   750
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   706
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   813
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAiMode 
      Height          =   495
      Index           =   3
      Left            =   3720
      TabIndex        =   7
      Top             =   9720
      Width           =   975
   End
   Begin VB.TextBox txtAiMode 
      Height          =   495
      Index           =   2
      Left            =   3720
      TabIndex        =   6
      Top             =   9120
      Width           =   975
   End
   Begin VB.TextBox txtAiMode 
      Height          =   495
      Index           =   1
      Left            =   2640
      TabIndex        =   5
      Top             =   9720
      Width           =   975
   End
   Begin VB.TextBox txtAiMode 
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   4
      Top             =   9120
      Width           =   975
   End
   Begin VB.ListBox lstCoin 
      Height          =   8835
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
   Begin VB.TextBox txtTest 
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   9120
      Width           =   1095
   End
   Begin VB.TextBox txtTest 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   9120
      Width           =   1095
   End
   Begin VB.ListBox lstMap 
      Columns         =   8
      Height          =   8835
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4920
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
