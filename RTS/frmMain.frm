VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   4695
   ClientTop       =   1350
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Begin VB.Timer tmrFrame 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   50
      Left            =   6240
      Top             =   840
   End
   Begin VB.Timer tmrFrame 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   50
      Left            =   6720
      Top             =   840
   End
   Begin VB.Timer tmrFrame 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   50
      Left            =   7200
      Top             =   840
   End
   Begin VB.Timer tmrFrame 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   1000
      Left            =   7680
      Top             =   840
   End
   Begin VB.Timer tmrHurt 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   9960
      Top             =   120
   End
   Begin VB.PictureBox picCharMaskIR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   7440
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   103
      Top             =   6000
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Timer tmrHurt 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   500
      Left            =   9960
      Top             =   1560
   End
   Begin VB.Timer tmrHurt 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   500
      Left            =   9960
      Top             =   1080
   End
   Begin VB.Timer tmrHurt 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   500
      Left            =   9960
      Top             =   600
   End
   Begin VB.Timer tmrChar 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   50
      Left            =   6120
      Top             =   1320
   End
   Begin VB.Timer tmrChar 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   50
      Left            =   5640
      Top             =   1320
   End
   Begin VB.Timer tmrChar 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   50
      Left            =   5640
      Top             =   840
   End
   Begin VB.Timer tmrCPUMove 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   1000
      Left            =   10440
      Top             =   1560
   End
   Begin VB.Timer tmrCPUMove 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   1000
      Left            =   10440
      Top             =   1080
   End
   Begin VB.Timer tmrCPUMove 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   1000
      Left            =   10440
      Top             =   600
   End
   Begin VB.Timer tmrCoinEvent 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   9480
      Top             =   360
   End
   Begin VB.PictureBox picSparkleMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   4
      Left            =   6240
      Picture         =   "frmMain.frx":00C5
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   42
      Top             =   5520
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer tmrCoin 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5040
      Top             =   840
   End
   Begin VB.Timer tmrTileAnimDelay 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   8880
      Top             =   360
   End
   Begin VB.Timer tmrTileAnim 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   8280
      Top             =   360
   End
   Begin VB.Timer tmrScoreCheck 
      Interval        =   250
      Left            =   5040
      Top             =   360
   End
   Begin VB.Timer tmrChar 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   50
      Left            =   5640
      Top             =   360
   End
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   0
      ScaleHeight     =   601
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   753
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11295
      Begin VB.PictureBox picCharMaskID 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   7440
         Picture         =   "frmMain.frx":0428
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   116
         Top             =   7680
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1ID 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   7440
         Picture         =   "frmMain.frx":04EA
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   115
         Top             =   6840
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskJD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   10320
         Picture         =   "frmMain.frx":06FA
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   114
         Top             =   7680
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1JD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   10320
         Picture         =   "frmMain.frx":07BB
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   113
         Top             =   6840
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   8880
         Picture         =   "frmMain.frx":09BA
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   112
         Top             =   6840
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   9600
         Picture         =   "frmMain.frx":0BC0
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   111
         Top             =   6840
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   8160
         Picture         =   "frmMain.frx":0DB3
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   110
         Top             =   6840
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   8880
         Picture         =   "frmMain.frx":0FBD
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   109
         Top             =   7680
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   9600
         Picture         =   "frmMain.frx":107C
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   108
         Top             =   7680
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   8160
         Picture         =   "frmMain.frx":1139
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   107
         Top             =   7680
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1IR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   7440
         Picture         =   "frmMain.frx":11FC
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   106
         Top             =   5160
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskJR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   10320
         Picture         =   "frmMain.frx":140A
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   105
         Top             =   6000
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1JR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   10320
         Picture         =   "frmMain.frx":14CA
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   104
         Top             =   5160
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   8160
         Picture         =   "frmMain.frx":16C2
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   102
         Top             =   6000
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   9600
         Picture         =   "frmMain.frx":1781
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   101
         Top             =   6000
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   8880
         Picture         =   "frmMain.frx":183C
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   100
         Top             =   6000
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   8160
         Picture         =   "frmMain.frx":18F9
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   99
         Top             =   5160
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   9600
         Picture         =   "frmMain.frx":1B00
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   98
         Top             =   5160
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   8880
         Picture         =   "frmMain.frx":1CEC
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   97
         Top             =   5160
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   9
         Left            =   7920
         Picture         =   "frmMain.frx":1EEA
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   96
         Top             =   3600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   8
         Left            =   7800
         Picture         =   "frmMain.frx":2027
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   95
         Top             =   3600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   7
         Left            =   7680
         Picture         =   "frmMain.frx":2166
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   94
         Top             =   3600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   6
         Left            =   7560
         Picture         =   "frmMain.frx":22A7
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   93
         Top             =   3600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   5
         Left            =   7440
         Picture         =   "frmMain.frx":23EA
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   92
         Top             =   3600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   9
         Left            =   7920
         Picture         =   "frmMain.frx":252C
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   91
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   5
         Left            =   7800
         Picture         =   "frmMain.frx":26A0
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   90
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   8
         Left            =   7680
         Picture         =   "frmMain.frx":2802
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   89
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   7
         Left            =   7560
         Picture         =   "frmMain.frx":2990
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   88
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   6
         Left            =   7440
         Picture         =   "frmMain.frx":2AFB
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   87
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":2C5A
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   86
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":2D17
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   85
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   2
         Left            =   960
         Picture         =   "frmMain.frx":2DDB
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   84
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   3
         Left            =   1320
         Picture         =   "frmMain.frx":2EAF
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   83
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   4
         Left            =   1680
         Picture         =   "frmMain.frx":2F66
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   82
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   5
         Left            =   2040
         Picture         =   "frmMain.frx":3015
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   81
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   6
         Left            =   2400
         Picture         =   "frmMain.frx":30B5
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   80
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   7
         Left            =   2760
         Picture         =   "frmMain.frx":3148
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   79
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   8
         Left            =   3120
         Picture         =   "frmMain.frx":31E0
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   78
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   9
         Left            =   3480
         Picture         =   "frmMain.frx":3273
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   77
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   10
         Left            =   3840
         Picture         =   "frmMain.frx":3310
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   76
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   11
         Left            =   4200
         Picture         =   "frmMain.frx":33BA
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   75
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   12
         Left            =   4560
         Picture         =   "frmMain.frx":3470
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   74
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   13
         Left            =   4920
         Picture         =   "frmMain.frx":3544
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   73
         Top             =   6720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":3626
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   72
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":39A7
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   71
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   2
         Left            =   960
         Picture         =   "frmMain.frx":3D2F
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   70
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   3
         Left            =   1320
         Picture         =   "frmMain.frx":40B9
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   69
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   4
         Left            =   1680
         Picture         =   "frmMain.frx":4440
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   68
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   5
         Left            =   2040
         Picture         =   "frmMain.frx":47CA
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   67
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   6
         Left            =   2400
         Picture         =   "frmMain.frx":4B50
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   66
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   7
         Left            =   2760
         Picture         =   "frmMain.frx":4ED7
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   65
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   8
         Left            =   3120
         Picture         =   "frmMain.frx":525D
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   64
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   9
         Left            =   3480
         Picture         =   "frmMain.frx":55E2
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   63
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   10
         Left            =   3840
         Picture         =   "frmMain.frx":5967
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   62
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   11
         Left            =   4200
         Picture         =   "frmMain.frx":5CF0
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   61
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   12
         Left            =   4560
         Picture         =   "frmMain.frx":607A
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   60
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   13
         Left            =   4920
         Picture         =   "frmMain.frx":6406
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   59
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":678D
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   58
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":684A
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   57
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   2
         Left            =   960
         Picture         =   "frmMain.frx":690E
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   56
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   3
         Left            =   1320
         Picture         =   "frmMain.frx":69E2
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   55
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   4
         Left            =   1680
         Picture         =   "frmMain.frx":6A99
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   54
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   5
         Left            =   2040
         Picture         =   "frmMain.frx":6B48
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   53
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   6
         Left            =   2400
         Picture         =   "frmMain.frx":6BE8
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   52
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   7
         Left            =   2760
         Picture         =   "frmMain.frx":6C7B
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   51
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   8
         Left            =   3120
         Picture         =   "frmMain.frx":6D13
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   50
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   9
         Left            =   3480
         Picture         =   "frmMain.frx":6DA6
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   49
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   10
         Left            =   3840
         Picture         =   "frmMain.frx":6E43
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   48
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   11
         Left            =   4200
         Picture         =   "frmMain.frx":6EED
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   47
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   12
         Left            =   4560
         Picture         =   "frmMain.frx":6FA3
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   46
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   13
         Left            =   4920
         Picture         =   "frmMain.frx":7077
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   45
         Top             =   6120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkleMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   6
         Left            =   6720
         Picture         =   "frmMain.frx":7159
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   44
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkleMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   5
         Left            =   6480
         Picture         =   "frmMain.frx":74AC
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   43
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   4
         Left            =   6240
         Picture         =   "frmMain.frx":7805
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   41
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   6
         Left            =   6720
         Picture         =   "frmMain.frx":785B
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   40
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   5
         Left            =   6480
         Picture         =   "frmMain.frx":78A0
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   39
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   3
         Left            =   6000
         Picture         =   "frmMain.frx":78F0
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   38
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   2
         Left            =   5760
         Picture         =   "frmMain.frx":7945
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   37
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkleMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   5280
         Picture         =   "frmMain.frx":799B
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   36
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkleMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   1
         Left            =   5520
         Picture         =   "frmMain.frx":7CEE
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   35
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkleMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   3
         Left            =   6000
         Picture         =   "frmMain.frx":8048
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   34
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkleMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   2
         Left            =   5760
         Picture         =   "frmMain.frx":83A8
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   33
         Top             =   5520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   1
         Left            =   5520
         Picture         =   "frmMain.frx":8709
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   32
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSparkle 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   5280
         Picture         =   "frmMain.frx":8759
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   31
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   13
         Left            =   4920
         Picture         =   "frmMain.frx":879E
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   30
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   12
         Left            =   4560
         Picture         =   "frmMain.frx":8880
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   29
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   11
         Left            =   4200
         Picture         =   "frmMain.frx":8954
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   28
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   10
         Left            =   3840
         Picture         =   "frmMain.frx":8A0A
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   27
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   9
         Left            =   3480
         Picture         =   "frmMain.frx":8AB4
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   26
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   8
         Left            =   3120
         Picture         =   "frmMain.frx":8B51
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   25
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   7
         Left            =   2760
         Picture         =   "frmMain.frx":8BE4
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   24
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   6
         Left            =   2400
         Picture         =   "frmMain.frx":8C7C
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   23
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   5
         Left            =   2040
         Picture         =   "frmMain.frx":8D0F
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   22
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   4
         Left            =   1680
         Picture         =   "frmMain.frx":8DAF
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   21
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   3
         Left            =   1320
         Picture         =   "frmMain.frx":8E5E
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   20
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   2
         Left            =   960
         Picture         =   "frmMain.frx":8F15
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   19
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":8FE9
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   18
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picCoinY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":90AD
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   17
         Top             =   4920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   0
         Left            =   9840
         Picture         =   "frmMain.frx":916A
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   9720
         Picture         =   "frmMain.frx":92CC
         ScaleHeight     =   100
         ScaleMode       =   0  'User
         ScaleWidth      =   104.166
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   2
         Left            =   9600
         Picture         =   "frmMain.frx":942B
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   3
         Left            =   9480
         Picture         =   "frmMain.frx":9596
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   4
         Left            =   9360
         Picture         =   "frmMain.frx":9724
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   0
         Left            =   9360
         Picture         =   "frmMain.frx":9898
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   10
         Top             =   3600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   9480
         Picture         =   "frmMain.frx":99DA
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   9
         Top             =   3600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   2
         Left            =   9600
         Picture         =   "frmMain.frx":9B1D
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   8
         Top             =   3600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   3
         Left            =   9720
         Picture         =   "frmMain.frx":9C5E
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   7
         Top             =   3600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   4
         Left            =   9840
         Picture         =   "frmMain.frx":9D9D
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   6
         Top             =   3600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picScene 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   240
         Picture         =   "frmMain.frx":9EDA
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   5
         Top             =   3360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picScene 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   2
         Left            =   240
         Picture         =   "frmMain.frx":B150
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   4
         Top             =   1800
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picScene 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":B96D
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picBuffer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   3375
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   2
         Top             =   240
         Width           =   1500
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   1800
         Picture         =   "frmMain.frx":C0E0
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   1
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "Score: 0000000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   9840
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selType(0 To 3) As String
Dim picCount(0 To 3) As Integer
Dim strDir(0 To 3) As String
Dim blnPlayerMoveable As Boolean
Dim counterLimit(1 To 3) As Integer
Dim frameCounter(0 To 3) As Integer
Dim frameLimit(0 To 3) As Integer
Dim strState(0 To 3) As String
Dim intScore As Integer
Dim spriteX(0 To 3) As Integer
Dim spriteY(0 To 3) As Integer
Dim curX(0 To 3) As Integer
Dim curY(0 To 3) As Integer
Dim prevX(0 To 3) As Integer
Dim prevY(0 To 3) As Integer
Dim coinTileCount As Integer
Dim blnClearPrevTile(0 To 3) As Boolean
Dim tileSwitch(0 To 100) As Boolean
Dim limswitch As Long
Dim smallestX As Integer
Dim smallestY As Integer

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 123 Then 'F12
    If frmDbg.Visible = False Then
        frmDbg.Visible = True
        frmDbg.Show
        Me.SetFocus
    Else
        frmDbg.Visible = False
        frmDbg.Hide
    End If
End If
End Sub

Private Function getJump(ByVal index As Integer, ByVal strDirJ As String)
If tmrFrame(index).Enabled = False Then
    strDir(index) = strDirJ
    tmrFrame(index).Enabled = True
End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If blnPlayerMoveable = True Then
    If KeyCode = 37 Then 'Left
        If evalMove(0, "L") = True Then
            Call getJump(0, "L")
        End If
    ElseIf KeyCode = 38 Then 'Up
        If evalMove(0, "U") = True Then
            Call getJump(0, "U")
        End If
    ElseIf KeyCode = 39 Then 'Right
        If evalMove(0, "R") = True Then
            Call getJump(0, "R")
        End If
    ElseIf KeyCode = 40 Then 'Down
        If evalMove(0, "D") = True Then
            Call getJump(0, "D")
        End If
    End If
End If
End Sub

Private Function evalMove(index As Integer, ByVal strDirMove As String) As Boolean
If strDirMove = "L" Then
    If (curY(index) + 1) Mod 2 = 0 Then
        If curX(index) = 0 Then
            evalMove = False
        ElseIf curY(index) < mapWidth Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curY(index) > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "U" Then
    If (curY(index) + 1) Mod 2 = 0 Then
        If curX(index) < mapWidth Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curX(index) < (mapWidth) And curY(index) > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "R" Then
    If (curY(index) + 1) Mod 2 = 0 Then
        If curX(index) = mapWidth Then
            evalMove = False
        ElseIf curY(index) > 0 Then
            evalMove = True
        End If
    ElseIf curY(index) < (mapWidth - 1) And curX(index) < mapWidth Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "D" Then
    If (curY(index) + 1) Mod 2 = 0 Then
        If curX(index) > 0 And curY(index) < mapHeight Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curY(index) < (mapHeight - 1) Then
        evalMove = True
    Else
        evalMove = False
    End If
End If
End Function

Private Sub tmrHurt_Timer(index As Integer)
Call getHurt(index)
curX(index) = prevX(index)
curY(index) = prevY(index)
blnPlayerMoveable = True
tmrHurt(index).Enabled = False
End Sub

Private Sub getHurt(ByVal enemyIndex As Integer)
MsgBox ("You got hurt.")
End Sub

Private Sub tmrCPUMove_Timer(index As Integer)
Static intCounter As Integer
'if counter limit is not reached by counter
If intCounter < counterLimit(index) Then
    'increase counter
    intCounter = intCounter + 1
'if counter limit is reached
Else
    'initiate cpu movement
    Call cpuAI(index)
    intCounter = 0
End If
End Sub
Private Function nearestCoin(ByVal index As Integer) As Long
smallestX = 0
smallestY = 0
For X = 0 To mapWidth
    For Y = 0 To mapHeight
    'if the tile has a coin then check if it is closer than the closest one, if it is make it the closest one
        If Tile(X, Y).coinEnabled = True Then
            If Abs(curX(index) - Tile(X, Y).X) < smallestX And Abs(curY(index) - Tile(X, Y).Y) < smallestY Then
                smallestX = Abs(curX(index) - Tile(X, Y).X)
                smallestY = Abs(curY(index) - Tile(X, Y).Y)
            End If
        End If
    Next Y
Next X
End Function
Private Function aiDecideMove(ByVal index As Integer) As Integer
'is it better to get rid of coins or chase player
Call nearestCoin(index)
If Abs(curX(index) - curX(0)) >= Abs(curX(index) - smallestX) And Abs(curY(index) - curY(0)) >= Abs(curY(index) - smallestY) Then
    aiDecideMove = 1
Else
    aiDecideMove = 0
End If
End Function

Private Function cpuAI(ByVal index As Integer)
'if the player score is over a certain amount(set in Form_load) then activate smart ai
'could add code to enable more modes based on different limits
If intScore >= limswitch Then
    mode = aiDecideMove(index)
Else
    mode = 1
End If
frmDbg.txtAiMode(index).Text = mode
'chase player
If mode = 1 Then
    'player x further than cpu x
    If curX(index) < curX(0) Then
        'player y lower than cpu y
        If curY(index) <= curY(0) Then
            If evalMove(index, "R") = True Then
                Call getJump(index, "R")
            End If
        'cpu y lower than player y
        ElseIf curY(index) > curY(0) Then
            If evalMove(index, "U") = True Then
                Call getJump(index, "U")
            End If
        End If
    'cpu x matches player x
    ElseIf curX(index) = curX(0) Then
        'ai y lower than player y
        If curY(index) < curY(0) Then
            'if y row is even
            If (curY(index) + 1) Mod 2 = 1 Then
                If evalMove(index, "D") = True Then
                    Call getJump(index, "D")
                End If
            'if y row is odd
            Else
                If evalMove(index, "R") = True Then
                    Call getJump(index, "R")
                End If
            End If
        'cpu y lower than player y
        ElseIf curY(index) > curY(0) Then
            'if y row is even
            If (curY(index) + 1) Mod 2 = 1 Then
                If evalMove(index, "L") = True Then
                    Call getJump(index, "L")
                End If
            'if y row is odd
            Else
                If evalMove(index, "U") = True Then
                    Call getJump(index, "U")
                End If
            End If
        End If
    'cpu x further than player x
    ElseIf curX(index) > curX(0) Then
        'player y lower than cpu y
        If curY(index) <= curY(0) Then
            If evalMove(index, "D") = True Then
                Call getJump(index, "D")
            End If
        'cpu y lower than player y
        ElseIf curY(index) > curY(0) Then
            If evalMove(index, "L") = True Then
                Call getJump(index, "L")
            End If
        End If
    End If
'stomp nearest coin
ElseIf mode = 0 Then
    If curX(index) < smallestX Then
        If curY(index) <= smallestY Then
            If evalMove(index, "R") Then
                Call getJump(index, "R")
            End If
        ElseIf curY(index) > smallestY Then
            If evalMove(index, "U") Then
                Call getJump(index, "U")
            End If
        End If
    ElseIf curX(index) = smallestX Then
        If curY(index) < smallestY Then
            'if y row is even
            If (curY(index) + 1) Mod 2 = 1 Then
                If evalMove(index, "D") = True Then
                    Call getJump(index, "D")
                End If
            'if y row is odd
            Else
                If evalMove(index, "R") = True Then
                    Call getJump(index, "R")
                End If
            End If
        End If
    ElseIf curX(index) > smallestX Then
        If curY(index) <= smallestY Then
            If evalMove(index, "D") = True Then
                Call getJump(index, "D")
            End If
        'cpu y lower than coin
        ElseIf curY(index) > smallestY Then
            If evalMove(index, "L") = True Then
                Call getJump(index, "L")
            End If
        End If
    End If
End If
End Function

Private Sub getJumpComplete(ByVal index As Integer)
Dim pScore As Integer
prevX(index) = curX(index)
prevY(index) = curY(index)
If (index = 0 And blnPlayerMoveable = True) Or index > 0 Then
    If strDir(index) = "L" Then
        'if y row is odd
        If (curY(index) + 1) Mod 2 = 0 Then
            curX(index) = curX(index) - 1
        End If
        curY(index) = curY(index) - 1
    ElseIf strDir(index) = "U" Then
        'if y row is even
        If (curY(index) + 1) Mod 2 = 1 Then
            curX(index) = curX(index) + 1
        End If
        curY(index) = curY(index) - 1
    ElseIf strDir(index) = "R" Then
        'if y row is even
        If (curY(index) + 1) Mod 2 = 1 Then
            curX(index) = curX(index) + 1
        End If
        curY(index) = curY(index) + 1
    ElseIf strDir(index) = "D" Then
        'if y row is odd
        If (curY(index) + 1) Mod 2 = 0 Then
            curX(index) = curX(index) - 1
        End If
        curY(index) = curY(index) + 1
    End If
End If
If index > 0 Then
    If curX(index) = curX(0) And curY(index) = curY(0) Then
        blnPlayerMoveable = False
        tmrHurt(index).Enabled = True
    End If
Else
End If
If index = 0 Then
    If Tile(curX(0), curY(0)).coinEnabled = True Then
        'play coin sound
        If Tile(curX(0), curY(0)).coinType = "Y" Then
            pScore = 100
        ElseIf Tile(curX(0), curY(0)).coinType = "R" Then
            pScore = 250
        ElseIf Tile(curX(0), curY(0)).coinType = "B" Then
            pScore = 500
        End If
    Else
            pScore = 10
    End If
    addScore (pScore)
End If
If Tile(curX(index), curY(index)).coinEnabled = True Then
    Tile(curX(index), curY(index)).coinEnabled = False
    Tile(curX(index), curY(index)).coinTimer = 0
    coinTileCount = coinTileCount - 1
End If
blnClearPrevTile(index) = True
strState(index) = "I"
End Sub

Private Sub Form_Load()
Call DrawMap(1)
blnClearPrevTile(0) = False
blnClearPrevTile(1) = False
blnPlayerMoveable = False
For t = 0 To 100
    tileSwitch(t) = False
Next t
limswitch = 1000
End Sub

Private Sub addScore(ByVal intAdd As Integer)
intScore = intScore + intAdd
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'selType = "O"
'flTextBox.Visible = True
'cmdCancelTextBox.Visible = True
'flTextBox.Movie = App.Path + "\Images\GUI\TextBoxB.swf"
End Sub

Private Sub Form_Resize()
Call DrawMap(1)
End Sub

Private Sub tmrScoreCheck_Timer()
lblScore = "Score " & Format(intScore, "0000000")
End Sub

Private Sub tmrCoin_Timer()
Static frameCount As Integer
frmDbg.lstCoin.Clear
For c = 0 To tileCount - 1
    frmDbg.lstCoin.AddItem (c & ": " & Tile(getTileFromInt(True, c), getTileFromInt(False, c)).coinEnabled & ", " & Tile(getTileFromInt(True, c), getTileFromInt(False, c)).coinTimer)
    If Tile(getTileFromInt(True, c), getTileFromInt(False, c)).coinEnabled = True Then
        Call PaintCoin(Tile(getTileFromInt(True, c), getTileFromInt(False, c)).coinType, frameCount, getTileFromInt(True, c), getTileFromInt(False, c))
        'If coinTimer (frame advancements on coin) is under 120, add 1 to it
        If Tile(getTileFromInt(True, c), getTileFromInt(False, c)).coinTimer < 80 Then
            Tile(getTileFromInt(True, c), getTileFromInt(False, c)).coinTimer = Tile(getTileFromInt(True, c), getTileFromInt(False, c)).coinTimer + 1
        'if coinTimer is 120 (or greater), disable coin and clear tile
        Else
            Tile(getTileFromInt(True, c), getTileFromInt(False, c)).coinTimer = 0
            Tile(getTileFromInt(True, c), getTileFromInt(False, c)).coinEnabled = False
            Call clearTile(getTileFromInt(True, c), getTileFromInt(False, c), False)
            coinTileCount = coinTileCount - 1
        End If
    End If
Next c
frameCount = frameCount + 1
If frameCount = 28 Then
    frameCount = 0
End If
End Sub

Private Function clearTile(ByVal IntX As Integer, ByVal intY As Integer, ByVal bypassForCoin As Boolean)
'if bypassForCoin is false or coin not enabled on tile (doesn't paint if bypassForCoin is true and coin is enabled on tile)
If bypassForCoin = False Or Tile(IntX, intY).coinEnabled = False Then
    'paint over tile
    picBackground.PaintPicture frmMain.picScene(0).Image, Tile(IntX, intY).X, Tile(IntX, intY).Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
    picBuffer.PaintPicture frmMain.picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
    frmMain.PaintPicture frmMain.picMask.Image, Tile(IntX, intY).X, Tile(IntX, intY).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    frmMain.PaintPicture picBuffer.Image, Tile(IntX, intY).X, Tile(IntX, intY).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
End If
End Function


Private Function PaintCoin(ByVal strType As String, ByVal intFrame As Integer, ByVal intCoinX As Integer, ByVal intCoinY As Integer)
Dim intXOffset As Integer
Dim intYOffset As Integer
Dim intFrameOffset As Integer
intXOffset = 41
intYOffset = -1
If intFrame > 13 Then
    intFrameOffset = -14
End If
Call clearTile(intCoinX, intCoinY, False)
'paint coin
frmMain.PaintPicture picCoinMask(intFrame + intFrameOffset).Image, Tile(intCoinX, intCoinY).X + intXOffset, Tile(intCoinX, intCoinY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
If strType = "Y" Then
    frmMain.PaintPicture picCoinY(intFrame + intFrameOffset).Image, Tile(intCoinX, intCoinY).X + intXOffset, Tile(intCoinX, intCoinY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
ElseIf strType = "R" Then
    frmMain.PaintPicture picCoinR(intFrame + intFrameOffset).Image, Tile(intCoinX, intCoinY).X + intXOffset, Tile(intCoinX, intCoinY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
ElseIf strType = "B" Then
    frmMain.PaintPicture picCoinB(intFrame + intFrameOffset).Image, Tile(intCoinX, intCoinY).X + intXOffset, Tile(intCoinX, intCoinY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
End If
'paint sparkle
If intFrame > 12 And intFrame < 20 Then
    frmMain.PaintPicture picSparkleMask(intFrame - 13).Image, Tile(intCoinX, intCoinY).X + intXOffset, Tile(intCoinX, intCoinY).Y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcAnd
    frmMain.PaintPicture picSparkle(intFrame - 13).Image, Tile(intCoinX, intCoinY).X + intXOffset, Tile(intCoinX, intCoinY).Y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcPaint
End If
End Function

Private Function PaintSelector(ByVal index As Integer, ByVal imgIndex As Integer) As Integer
'paint over last sel
If blnClearPrevTile(index) = True Then
    'prevX, prevX
    Call clearTile(prevX(index), prevY(index), True)
    If prevY(index) > 0 Then
        Dim prevXclear As Integer
        Dim prevYclear As Integer
        'prevX, prevY - 1
        prevYclear = prevY(index) - 1
        'even y
        If prevY(index) + 1 Mod 2 = 1 Then
            'prevX - 1, prevY - 1
            prevXclear = prevX(index) - 1
            Call clearTile(prevXclear, prevYclear, True)
        'odd y
        ElseIf prevY(index) + 1 Mod 2 = 0 Then
            'prevX + 1, prevY - 1
            prevXclear = prevX(index) + 1
            Call clearTile(prevXclear, prevYclear, True)
        End If
    End If
    blnClearPrevTile(index) = False
End If
'paint over frame
Call clearTile(curX(index), curY(index), True)
If curY(index) > 0 Then
    If curX(index) > 0 Then
        'curX - 1, curY - 1
        Call clearTile(curX(index) - 1, curY(index) - 1, True)
    End If
    'odd y
    If curY(index) + 1 Mod 2 = 1 Then
        If curX(index) < mapWidth Then
            'curX + 1, curY - 1
            Call clearTile(curX(index) + 1, curY(index) - 1, True)
        End If
    'even y
    Else
        If curX(index) < mapWidth Then
            'curX, curY - 1
            Call clearTile(curX(index), curY(index) - 1, True)
        End If
        If curX(index) < mapWidth - 1 Then
            'curX + 1, curY - 1
            Call clearTile(curX(index) + 1, curY(index) - 1, True)
        End If
    End If
End If
If curY(index) < mapHeight - 1 Then
    If curX(index) > 0 Then
        'curX - 1, curY + 1
        Call clearTile(curX(index) - 1, curY(index) + 1, True)
    End If
    'odd y
    If curY(index) + 1 Mod 2 = 1 Then
        If curX(index) < mapWidth Then
            'curX + 1, curY + 1
            Call clearTile(curX(index) + 1, curY(index) + 1, True)
        End If
    'even y
    Else
        If curX(index) < mapWidth Then
            'curX, curY + 1
            Call clearTile(curX(index), curY(index) + 1, True)
        End If
        If curX(index) < mapWidth - 1 Then
            'curX + 1, curY + 1
            Call clearTile(curX(index) + 1, curY(index) + 1, True)
        End If
    End If
End If
'paint sel
frmMain.PaintPicture picSelMask(imgIndex + 5 * index).Image, Tile(curX(index), curY(index)).X, Tile(curX(index), curY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
frmMain.PaintPicture picSel(imgIndex + 5 * index).Image, Tile(curX(index), curY(index)).X, Tile(curX(index), curY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
Call PaintCharSprite(index, spriteX(index), spriteY(index))
End Function

Private Sub PaintCharSprite(ByVal index As Integer, ByVal charX As Integer, ByVal charY As Integer)
Static counterC As Integer
If strState(index) = "I" Then
    If strDir(index) = "L" Then
        frmMain.PaintPicture picCharMaskIR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "U" Then
        frmMain.PaintPicture picCharMaskID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "R" Then
        frmMain.PaintPicture picCharMaskIR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "D" Then
        frmMain.PaintPicture picCharMaskID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    End If
ElseIf strState(index) = "C" Then
    If counterC < 3 Then
        counterC = counterC + 1
    End If
    Dim frameC As Integer
    frameC = Int((counterC + 1) / 2)
    If strDir(index) = "L" Then
        frmMain.PaintPicture picCharMaskCR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "U" Then
        frmMain.PaintPicture picCharMaskCD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "R" Then
        frmMain.PaintPicture picCharMaskCR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "D" Then
        frmMain.PaintPicture picCharMaskCD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    End If
ElseIf strState(index) = "J" Then
    counterC = 0
    If strDir(index) = "L" Then
        frmMain.PaintPicture picCharMaskJR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "U" Then
        frmMain.PaintPicture picCharMaskJD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "R" Then
        frmMain.PaintPicture picCharMaskJR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "D" Then
        frmMain.PaintPicture picCharMaskJD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If index = 0 Then
            frmMain.PaintPicture picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index > 0 Then
            frmMain.PaintPicture picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    End If
End If
End Sub

Private Sub tmrChar_Timer(index As Integer)
'reverse boolean for select
Static blnRev(0 To 3) As Boolean
'call selection paint
Call PaintSelector(index, picCount(index))
'if jump timer is started
If tmrFrame(index).Enabled = True Then
    'show each jump frame for p1
    'If Index = 0 Then
        'MsgBox (intFrame(0))
    'End If
    If strDir(index) = "L" Then
        Call getCharJumpAnim(index, Tile(curX(index), curY(index)).X - 25, Tile(curX(index), curY(index)).Y - 115, Tile(curX(index), curY(index)).X + 25, Tile(curX(index), curY(index)).Y - 15)
    ElseIf strDir(index) = "R" Then
        Call getCharJumpAnim(index, Tile(curX(index), curY(index)).X + 75, Tile(curX(index), curY(index)).Y + 85, Tile(curX(index), curY(index)).X + 25, Tile(curX(index), curY(index)).Y - 15)
    ElseIf strDir(index) = "U" Then
        Call getCharJumpAnim(index, Tile(curX(index), curY(index)).X + 75, Tile(curX(index), curY(index)).Y - 115, Tile(curX(index), curY(index)).X + 25, Tile(curX(index), curY(index)).Y - 15)
    ElseIf strDir(index) = "D" Then
        Call getCharJumpAnim(index, Tile(curX(index), curY(index)).X - 25, Tile(curX(index), curY(index)).Y + 85, Tile(curX(index), curY(index)).X + 25, Tile(curX(index), curY(index)).Y - 15)
    End If
    If frameCounter(index) = 1 Then
        strState(index) = "C"
    ElseIf frameCounter(index) = 5 Then
        strState(index) = "J"
    ElseIf frameCounter(index) = 10 Then
        strState(index) = "I"
    End If
Else
    spriteX(index) = Tile(curX(index), curY(index)).X + 25
    spriteY(index) = Tile(curX(index), curY(index)).Y - 15
End If
If blnRev(index) = False Then
    picCount(index) = picCount(index) + 1
End If
If blnRev(index) = True Then
    picCount(index) = picCount(index) - 1
End If
If picCount(index) >= 4 Or picCount(index) <= 0 Then
    If blnRev(index) = False Then
        blnRev(index) = True
    Else
        blnRev(index) = False
    End If
End If
End Sub

Private Function getCharJumpAnim(ByVal index As Integer, ByVal IntNewX As Integer, ByVal intNewY As Integer, ByVal intOldX As Integer, ByVal intOldY As Integer)
'if frame 5 to 10
If frameCounter(index) >= 5 And frameCounter(index) <= 10 Then
    spriteX(index) = intOldX + ((frameCounter(index) - 5) * Int((IntNewX - intOldX) / 5))
    '5 to 7 is jump up
    If frameCounter(index) < 8 Then
        spriteY(index) = (intOldY + ((frameCounter(index) - 5) * Int((intNewY - intOldY) / 5))) - (4 * (frameCounter(index) - 5))
    '8 to 10 is fall to ground
    ElseIf frameCounter(index) <= 10 Then
        spriteY(index) = (intOldY + ((frameCounter(index) - 5) * Int((intNewY - intOldY) / 5))) - (4 * (10 - frameCounter(index)))
    End If
End If
End Function

Private Sub tmrFrame_Timer(index As Integer)
frameCounter(index) = frameCounter(index) + 1
If frameCounter(index) = frameLimit(index) Then
    frameCounter(index) = 0
    getJumpComplete (index)
    tmrFrame(index).Enabled = False
End If
End Sub

Private Function getTileAnim(ByVal intFrame As Integer, ByVal IntX As Integer, ByVal intY As Integer)
picBuffer.PaintPicture frmMain.picBackground.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
frmMain.PaintPicture picBuffer.Image, Tile(IntX, intY).X, (Tile(IntX, intY).Y - 400) + (intFrame - 1) * 50, 100, 100, 0, 0, 100, 100, vbSrcCopy
'paint tile mask with new y
PaintPicture frmMain.picMask.Picture, Tile(IntX, intY).X, (Tile(IntX, intY).Y - 400) + intFrame * 50, 100, 100, 0, 0, 100, 100, vbSrcAnd
'paint tile with new y
PaintPicture frmMain.picScene(0).Picture, Tile(IntX, intY).X, (Tile(IntX, intY).Y - 400) + intFrame * 50, 100, 100, 0, 0, 100, 100, vbSrcPaint
If intFrame >= 8 And IntX = 0 And intY = 0 Then
    Call gameStart
    tmrTileAnim.Enabled = False
End If
End Function

Private Sub tmrTileAnim_Timer()
Static intCounter As Integer
For X = 0 To tileCount - 1
    If tileSwitch(X) = True And intCounter - (1 * ((tileCount - 1) - X)) <= 8 Then
        Call getTileAnim(intCounter - (1 * ((tileCount - 1) - X)), getTileFromInt(True, X), getTileFromInt(False, X))
    End If
Next X
intCounter = intCounter + 1
End Sub

Private Sub tmrTileAnimDelay_Timer()
Static intCounter As Integer
Static IntX As Integer
Static intY As Integer
frmMain.tmrTileAnim.Enabled = True
For z = 0 To intCounter - getAbs(intCounter - ((mapWidth * mapHeight) + (Int(mapHeight / 2))) - 1)
    tileSwitch((tileCount - 1) - z) = True
Next z
If IntX = mapWidth - ((intY + 1) Mod 2) Then
    intY = intY + 1
    IntX = 0
ElseIf IntX < mapWidth + ((mapHeight + 1) Mod 2) Then
    IntX = IntX + 1
End If
If intCounter >= ((mapWidth * mapHeight) + (Int(mapHeight / 2))) - 1 Then
    tmrTileAnimDelay.Enabled = False
End If
intCounter = intCounter + 1
End Sub

Private Function getTileFromInt(ByVal blnX As Boolean, ByVal intInput As Integer) As Integer
If blnX = True Then
    If intInput >= mapWidth Then
        If intInput >= (mapWidth * 2) + 1 Then
            Dim a As Integer
            a = 2
            Do Until intInput < (mapWidth * a) + Int(a / 2)
                a = a + 1
            Loop
            getTileFromInt = intInput - ((mapWidth * (a - 1)) + Int((a - 1) / 2))
        Else
            getTileFromInt = intInput - mapWidth
        End If
    Else
        getTileFromInt = intInput
    End If
ElseIf blnX = False Then
    If intInput >= mapWidth Then
        If intInput >= (mapWidth * 2) + 1 Then
            Dim b As Integer
            a = 2
            Do Until intInput < (mapWidth * b) + Int(b / 2)
                b = b + 1
            Loop
            getTileFromInt = b - 1
        Else
            getTileFromInt = 1
        End If
    Else
        getTileFromInt = 0
    End If
End If
End Function

Private Function getAbs(ByVal valInput As Single) As Integer
If valInput < 0 Then
    valInput = 0
End If
getAbs = valInput
End Function

Private Function randInt(ByVal min As Integer, ByVal max As Integer) As Integer
randInt = Int(Rnd() * max) + min
End Function

Private Sub tmrCoinEvent_Timer()
Dim intRand As Integer
If coinTileCount < tileCount - 1 Then
    intRand = Int(Rnd() * tileCount)
    Do While (Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).X = Tile(curX(0), curY(0)).X And Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).Y = Tile(curX(0), curY(0)).Y) Or (Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).X = Tile(curX(1), curY(1)).X And Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).Y = Tile(curX(1), curY(1)).Y) Or (Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).X = Tile(curX(2), curY(2)).X And Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).Y = Tile(curX(2), curY(2)).Y) Or (Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).X = Tile(curX(3), curY(3)).X And Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).Y = Tile(curX(3), curY(3)).Y) Or Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).coinEnabled = True
        intRand = randInt(0, tileCount - 1)
    Loop
    Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).coinEnabled = True
    Dim intType As Integer
    intType = randInt(1, 100)
    If intType <= 65 Then
        Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).coinType = "Y"
    ElseIf intType > 65 And intType <= 90 Then
        Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).coinType = "R"
    ElseIf intType > 90 Then
        Tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).coinType = "B"
    End If
    coinTileCount = coinTileCount + 1
End If
End Sub

Private Sub gameStart()
tmrChar(0).Enabled = True
tmrChar(1).Enabled = True
blnPlayerMoveable = True
tmrCoinEvent.Enabled = True
tmrCoin.Enabled = True
tmrCPUMove(1).Enabled = True
curX(0) = (mapWidth - 1) \ 2
curY(0) = (mapHeight - 1) \ 2
curX(1) = 0
curY(1) = 0
frameLimit(0) = 10
frameLimit(1) = 10
frameLimit(2) = 10
frameLimit(3) = 10
'cpu 1 moves every (counterLimit + 1) seconds
counterLimit(1) = 1
End Sub
