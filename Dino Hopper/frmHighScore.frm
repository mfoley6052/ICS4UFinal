VERSION 5.00
Begin VB.Form frmHighScore 
   Caption         =   "Top Scores"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Menu"
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Hiscores"
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox txtHiScoreName 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtHiScoreName 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtHiScoreName 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtHiScoreName 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtHiScoreName 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtHiScoreName 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtHiScoreName 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtHiScoreName 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtHiScoreName 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtHiScoreName 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox txtHiScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   495
      Index           =   9
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   6120
      Width           =   3255
   End
   Begin VB.TextBox txtHiScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   495
      Index           =   8
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5520
      Width           =   3255
   End
   Begin VB.TextBox txtHiScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   495
      Index           =   7
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4920
      Width           =   3255
   End
   Begin VB.TextBox txtHiScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   495
      Index           =   6
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4320
      Width           =   3255
   End
   Begin VB.TextBox txtHiScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox txtHiScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox txtHiScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox txtHiScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtHiScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtHiScore 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "frmHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtHiScoreName_Change(Index As Integer)
txtHiScoreName(Index).Text = UCase(txtHiScoreName(Index).Text)
End Sub
