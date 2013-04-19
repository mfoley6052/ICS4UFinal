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
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   500
      Left            =   7680
      Top             =   360
   End
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   500
      Left            =   7200
      Top             =   360
   End
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   500
      Left            =   6720
      Top             =   360
   End
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   6240
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
         Picture         =   "frmMain.frx":09B8
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
         Picture         =   "frmMain.frx":0BBE
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
         Picture         =   "frmMain.frx":0DB1
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
         Picture         =   "frmMain.frx":0FBB
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
         Picture         =   "frmMain.frx":107A
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
         Picture         =   "frmMain.frx":1137
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
         Picture         =   "frmMain.frx":11FA
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
         Picture         =   "frmMain.frx":1408
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
         Picture         =   "frmMain.frx":14C8
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
         Picture         =   "frmMain.frx":16C0
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
         Picture         =   "frmMain.frx":177F
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
         Picture         =   "frmMain.frx":183A
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
         Picture         =   "frmMain.frx":18F7
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
         Picture         =   "frmMain.frx":1AFE
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
         Picture         =   "frmMain.frx":1CEA
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
         Picture         =   "frmMain.frx":1EE8
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
         Picture         =   "frmMain.frx":2025
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
         Picture         =   "frmMain.frx":2164
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
         Picture         =   "frmMain.frx":22A5
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
         Picture         =   "frmMain.frx":23E8
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
         Picture         =   "frmMain.frx":252A
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
         Picture         =   "frmMain.frx":269E
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
         Picture         =   "frmMain.frx":2800
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
         Picture         =   "frmMain.frx":298E
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
         Picture         =   "frmMain.frx":2AF9
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
         Picture         =   "frmMain.frx":2C58
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
         Picture         =   "frmMain.frx":2D15
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
         Picture         =   "frmMain.frx":2DD9
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
         Picture         =   "frmMain.frx":2EAD
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
         Picture         =   "frmMain.frx":2F64
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
         Picture         =   "frmMain.frx":3013
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
         Picture         =   "frmMain.frx":30B3
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
         Picture         =   "frmMain.frx":3146
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
         Picture         =   "frmMain.frx":31DE
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
         Picture         =   "frmMain.frx":3271
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
         Picture         =   "frmMain.frx":330E
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
         Picture         =   "frmMain.frx":33B8
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
         Picture         =   "frmMain.frx":346E
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
         Picture         =   "frmMain.frx":3542
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
         Picture         =   "frmMain.frx":3624
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
         Picture         =   "frmMain.frx":39A5
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
         Picture         =   "frmMain.frx":3D2D
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
         Picture         =   "frmMain.frx":40B7
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
         Picture         =   "frmMain.frx":443E
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
         Picture         =   "frmMain.frx":47C8
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
         Picture         =   "frmMain.frx":4B4E
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
         Picture         =   "frmMain.frx":4ED5
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
         Picture         =   "frmMain.frx":525B
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
         Picture         =   "frmMain.frx":55E0
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
         Picture         =   "frmMain.frx":5965
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
         Picture         =   "frmMain.frx":5CEE
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
         Picture         =   "frmMain.frx":6078
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
         Picture         =   "frmMain.frx":6404
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
         Picture         =   "frmMain.frx":678B
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
         Picture         =   "frmMain.frx":6848
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
         Picture         =   "frmMain.frx":690C
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
         Picture         =   "frmMain.frx":69E0
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
         Picture         =   "frmMain.frx":6A97
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
         Picture         =   "frmMain.frx":6B46
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
         Picture         =   "frmMain.frx":6BE6
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
         Picture         =   "frmMain.frx":6C79
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
         Picture         =   "frmMain.frx":6D11
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
         Picture         =   "frmMain.frx":6DA4
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
         Picture         =   "frmMain.frx":6E41
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
         Picture         =   "frmMain.frx":6EEB
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
         Picture         =   "frmMain.frx":6FA1
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
         Picture         =   "frmMain.frx":7075
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
         Picture         =   "frmMain.frx":7157
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
         Picture         =   "frmMain.frx":74AA
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
         Picture         =   "frmMain.frx":7803
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
         Picture         =   "frmMain.frx":7859
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
         Picture         =   "frmMain.frx":789E
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
         Picture         =   "frmMain.frx":78EE
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
         Picture         =   "frmMain.frx":7943
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
         Picture         =   "frmMain.frx":7999
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
         Picture         =   "frmMain.frx":7CEC
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
         Picture         =   "frmMain.frx":8046
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
         Picture         =   "frmMain.frx":83A6
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
         Picture         =   "frmMain.frx":8707
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
         Picture         =   "frmMain.frx":8757
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
         Picture         =   "frmMain.frx":879C
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
         Picture         =   "frmMain.frx":887E
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
         Picture         =   "frmMain.frx":8952
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
         Picture         =   "frmMain.frx":8A08
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
         Picture         =   "frmMain.frx":8AB2
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
         Picture         =   "frmMain.frx":8B4F
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
         Picture         =   "frmMain.frx":8BE2
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
         Picture         =   "frmMain.frx":8C7A
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
         Picture         =   "frmMain.frx":8D0D
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
         Picture         =   "frmMain.frx":8DAD
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
         Picture         =   "frmMain.frx":8E5C
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
         Picture         =   "frmMain.frx":8F13
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
         Picture         =   "frmMain.frx":8FE7
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
         Picture         =   "frmMain.frx":90AB
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
         Picture         =   "frmMain.frx":9168
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
         Picture         =   "frmMain.frx":92CA
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
         Picture         =   "frmMain.frx":9429
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
         Picture         =   "frmMain.frx":9594
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
         Picture         =   "frmMain.frx":9722
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
         Picture         =   "frmMain.frx":9896
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
         Picture         =   "frmMain.frx":99D8
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
         Picture         =   "frmMain.frx":9B1B
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
         Picture         =   "frmMain.frx":9C5C
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
         Picture         =   "frmMain.frx":9D9B
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
         Picture         =   "frmMain.frx":9ED8
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
         Picture         =   "frmMain.frx":B14E
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
         Picture         =   "frmMain.frx":B96B
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
         Picture         =   "frmMain.frx":C0DE
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

Private Function getJump(ByVal Index As Integer, ByVal strDirJ As String)
If tmrJump(Index).Enabled = False Then
    strDir(Index) = strDirJ
    tmrJump(Index).Enabled = True
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

Private Function evalMove(Index As Integer, ByVal strDirMove As String) As Boolean
If strDirMove = "L" Then
    If (curY(Index) + 1) Mod 2 = 0 Then
        If curX(Index) = 0 Then
            evalMove = False
        ElseIf curY(Index) < mapWidth Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curY(Index) > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "U" Then
    If (curY(Index) + 1) Mod 2 = 0 Then
        If curX(Index) < mapWidth Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curX(Index) < (mapWidth) And curY(Index) > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "R" Then
    If (curY(Index) + 1) Mod 2 = 0 Then
        If curX(Index) = mapWidth Then
            evalMove = False
        ElseIf curY(Index) > 0 Then
            evalMove = True
        End If
    ElseIf curY(Index) < (mapWidth - 1) And curX(Index) < mapWidth Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "D" Then
    If (curY(Index) + 1) Mod 2 = 0 Then
        If curX(Index) > 0 And curY(Index) < mapHeight Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curY(Index) < (mapHeight - 1) Then
        evalMove = True
    Else
        evalMove = False
    End If
End If
End Function

Private Sub tmrHurt_Timer(Index As Integer)
Call getHurt(Index)
curX(Index) = prevX(Index)
curY(Index) = prevY(Index)
blnPlayerMoveable = True
tmrHurt(Index).Enabled = False
End Sub

Private Sub getHurt(ByVal enemyIndex As Integer)
MsgBox ("You got hurt.")
End Sub

Private Sub tmrCPUMove_Timer(Index As Integer)
Static intCounter As Integer
'if counter limit is not reached by counter
If intCounter < counterLimit(Index) Then
    'increase counter
    intCounter = intCounter + 1
'if counter limit is reached
Else
    'initiate cpu movement
    Call cpuAI(Index)
    intCounter = 0
End If
End Sub
Private Function nearestCoin(ByVal Index As Integer) As Long
smallestX = 0
smallestY = 0
For X = 0 To mapWidth
    For Y = 0 To mapHeight
    'if the tile has a coin then check if it is closer than the closest one, if it is make it the closest one
        If Tile(X, Y).coinEnabled = True Then
            If Abs(curX(Index) - Tile(X, Y).X) < smallestX And Abs(curY(Index) - Tile(X, Y).Y) < smallestY Then
                smallestX = Abs(curX(Index) - Tile(X, Y).X)
                smallestY = Abs(curY(Index) - Tile(X, Y).Y)
            End If
        End If
    Next Y
Next X
End Function
Private Function aiDecideMove(ByVal Index As Integer) As Integer
'is it better to get rid of coins or chase player
Call nearestCoin(Index)
If Abs(curX(Index) - curX(0)) >= Abs(curX(Index) - smallestX) And Abs(curY(Index) - curY(0)) >= Abs(curY(Index) - smallestY) Then
    aiDecideMove = 1
Else
    aiDecideMove = 0
End If
End Function

Private Function cpuAI(ByVal Index As Integer)
'if the player score is over a certain amount(set in Form_load) then activate smart ai
'could add code to enable more modes based on different limits
If intScore >= limswitch Then
    mode = aiDecideMove(Index)
Else
    mode = 1
End If
frmDbg.txtAiMode(Index).Text = mode
'chase player
If mode = 1 Then
    'player x further than cpu x
    If curX(Index) < curX(0) Then
        'player y lower than cpu y
        If curY(Index) <= curY(0) Then
            If evalMove(Index, "R") = True Then
                Call getJump(Index, "R")
            End If
        'cpu y lower than player y
        ElseIf curY(Index) > curY(0) Then
            If evalMove(Index, "U") = True Then
                Call getJump(Index, "U")
            End If
        End If
    'cpu x matches player x
    ElseIf curX(Index) = curX(0) Then
        'ai y lower than player y
        If curY(Index) < curY(0) Then
            'if y row is even
            If (curY(Index) + 1) Mod 2 = 1 Then
                If evalMove(Index, "D") = True Then
                    Call getJump(Index, "D")
                End If
            'if y row is odd
            Else
                If evalMove(Index, "R") = True Then
                    Call getJump(Index, "R")
                End If
            End If
        'cpu y lower than player y
        ElseIf curY(Index) > curY(0) Then
            'if y row is even
            If (curY(Index) + 1) Mod 2 = 1 Then
                If evalMove(Index, "L") = True Then
                    Call getJump(Index, "L")
                End If
            'if y row is odd
            Else
                If evalMove(Index, "U") = True Then
                    Call getJump(Index, "U")
                End If
            End If
        End If
    'cpu x further than player x
    ElseIf curX(Index) > curX(0) Then
        'player y lower than cpu y
        If curY(Index) <= curY(0) Then
            If evalMove(Index, "D") = True Then
                Call getJump(Index, "D")
            End If
        'cpu y lower than player y
        ElseIf curY(Index) > curY(0) Then
            If evalMove(Index, "L") = True Then
                Call getJump(Index, "L")
            End If
        End If
    End If
'stomp nearest coin
ElseIf mode = 0 Then
    If curX(Index) < smallestX Then
        If curY(Index) <= smallestY Then
            If evalMove(Index, "R") Then
                Call getJump(Index, "R")
            End If
        ElseIf curY(Index) > smallestY Then
            If evalMove(Index, "U") Then
                Call getJump(Index, "U")
            End If
        End If
    ElseIf curX(Index) = smallestX Then
        If curY(Index) < smallestY Then
            'if y row is even
            If (curY(Index) + 1) Mod 2 = 1 Then
                If evalMove(Index, "D") = True Then
                    Call getJump(Index, "D")
                End If
            'if y row is odd
            Else
                If evalMove(Index, "R") = True Then
                    Call getJump(Index, "R")
                End If
            End If
        End If
    ElseIf curX(Index) > smallestX Then
        If curY(Index) <= smallestY Then
            If evalMove(Index, "D") = True Then
                Call getJump(Index, "D")
            End If
        'cpu y lower than coin
        ElseIf curY(Index) > smallestY Then
            If evalMove(Index, "L") = True Then
                Call getJump(Index, "L")
            End If
        End If
    End If
End If
End Function

Private Sub tmrJump_Timer(Index As Integer)
Dim pScore As Integer
prevX(Index) = curX(Index)
prevY(Index) = curY(Index)
If (Index = 0 And blnPlayerMoveable = True) Or Index > 0 Then
    If strDir(Index) = "L" Then
        'if y row is odd
        If (curY(Index) + 1) Mod 2 = 0 Then
            curX(Index) = curX(Index) - 1
        End If
        curY(Index) = curY(Index) - 1
    ElseIf strDir(Index) = "U" Then
        'if y row is even
        If (curY(Index) + 1) Mod 2 = 1 Then
            curX(Index) = curX(Index) + 1
        End If
        curY(Index) = curY(Index) - 1
    ElseIf strDir(Index) = "R" Then
        'if y row is even
        If (curY(Index) + 1) Mod 2 = 1 Then
            curX(Index) = curX(Index) + 1
        End If
        curY(Index) = curY(Index) + 1
    ElseIf strDir(Index) = "D" Then
        'if y row is odd
        If (curY(Index) + 1) Mod 2 = 0 Then
            curX(Index) = curX(Index) - 1
        End If
        curY(Index) = curY(Index) + 1
    End If
End If
If Index > 0 Then
    If curX(Index) = curX(0) And curY(Index) = curY(0) Then
        blnPlayerMoveable = False
        tmrHurt(Index).Enabled = True
    End If
Else
End If
If Index = 0 Then
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
If Tile(curX(Index), curY(Index)).coinEnabled = True Then
    Tile(curX(Index), curY(Index)).coinEnabled = False
    Tile(curX(Index), curY(Index)).coinTimer = 0
    coinTileCount = coinTileCount - 1
End If
tmrJump(Index).Enabled = False
blnClearPrevTile(Index) = True
strState(Index) = "I"
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

Private Function PaintSelector(ByVal Index As Integer, ByVal imgIndex As Integer) As Integer
'paint over last sel
If blnClearPrevTile(Index) = True Then
    'prevX, prevX
    Call clearTile(prevX(Index), prevY(Index), True)
    If prevY(Index) > 0 Then
        Dim prevXclear As Integer
        Dim prevYclear As Integer
        'prevX, prevY - 1
        prevYclear = prevY(Index) - 1
        'even y
        If prevX(Index) > 0 And prevY(Index) + 1 Mod 2 = 1 Then
            'prevX - 1, prevY - 1
            prevXclear = prevX(Index) - 1
            Call clearTile(prevXclear, prevYclear, True)
        'odd y
        ElseIf prevX(Index) < mapWidth And prevY(Index) + 1 Mod 2 = 0 Then
            'prevX + 1, prevY - 1
            prevXclear = prevX(Index) + 1
            Call clearTile(prevXclear, prevYclear, True)
        End If
    End If
    blnClearPrevTile(Index) = False
End If
'paint over frame
Call clearTile(curX(Index), curY(Index), True)
If curY(Index) > 0 Then
    If curX(Index) > 0 Then
        'curX - 1, curY - 1
        Call clearTile(curX(Index) - 1, curY(Index) - 1, True)
    End If
    'curX, curY - 1
    Call clearTile(curX(Index), curY(Index) - 1, True)
    'odd y
    If curY(Index) + 1 Mod 2 = 1 Then
        If curX(Index) < mapWidth - 1 Then
            'curX + 1, curY - 1
            Call clearTile(curX(Index) + 1, curY(Index) - 1, True)
        End If
    'even y
    Else
        If curX(Index) < mapWidth Then
            'curX + 1, curY - 1
            Call clearTile(curX(Index) + 1, curY(Index) - 1, True)
        End If
    End If
End If
If curY(Index) < mapHeight - 1 Then
    If curX(Index) > 0 Then
        'curX - 1, curY + 1
        Call clearTile(curX(Index) - 1, curY(Index) + 1, True)
    End If
    'curX, curY + 1
    Call clearTile(curX(Index), curY(Index) + 1, True)
    'odd y
    If curY(Index) + 1 Mod 2 = 1 Then
        If curX(Index) < mapWidth - 1 Then
            'curX + 1, curY + 1
            Call clearTile(curX(Index) + 1, curY(Index) + 1, True)
        End If
    'even y
    Else
        If curX(Index) < mapWidth Then
            'curX + 1, curY + 1
            Call clearTile(curX(Index) + 1, curY(Index) + 1, True)
        End If
    End If
End If
'paint sel
frmMain.PaintPicture picSelMask(imgIndex + 5 * Index).Image, Tile(curX(Index), curY(Index)).X, Tile(curX(Index), curY(Index)).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
frmMain.PaintPicture picSel(imgIndex + 5 * Index).Image, Tile(curX(Index), curY(Index)).X, Tile(curX(Index), curY(Index)).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
Call PaintCharSprite(Index, spriteX(Index), spriteY(Index))
End Function

Private Sub PaintCharSprite(ByVal Index As Integer, ByVal charX As Integer, ByVal charY As Integer)
Static counterC As Integer
If strState(Index) = "I" Then
    If strDir(Index) = "L" Then
        frmMain.PaintPicture picCharMaskIR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "U" Then
        frmMain.PaintPicture picCharMaskID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "R" Then
        frmMain.PaintPicture picCharMaskIR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "D" Then
        frmMain.PaintPicture picCharMaskID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    End If
ElseIf strState(Index) = "C" Then
    If counterC < 3 Then
        counterC = counterC + 1
    End If
    If strDir(Index) = "L" Then
        frmMain.PaintPicture picCharMaskCR(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1CR(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1CR(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "U" Then
        frmMain.PaintPicture picCharMaskCD(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1CD(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1CD(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "R" Then
        frmMain.PaintPicture picCharMaskCR(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1CR(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1CR(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "D" Then
        frmMain.PaintPicture picCharMaskCD(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1CD(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1CD(counterC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    End If
ElseIf strState(Index) = "J" Then
    counterC = 0
    If strDir(Index) = "L" Then
        frmMain.PaintPicture picCharMaskJR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "U" Then
        frmMain.PaintPicture picCharMaskJD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "R" Then
        frmMain.PaintPicture picCharMaskJR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "D" Then
        frmMain.PaintPicture picCharMaskJD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            frmMain.PaintPicture picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            frmMain.PaintPicture picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    End If
End If
End Sub

Private Sub tmrChar_Timer(Index As Integer)
'reverse boolean for select
Static blnRev(0 To 3) As Boolean
Static intFrame(0 To 3) As Integer
Static frameLimit(0 To 3) As Integer
'call selection paint
Call PaintSelector(Index, picCount(Index))
'if jump timer is started
If tmrJump(Index).Enabled = True Then
    'show each jump frame for p1
    'If Index = 0 Then
        'MsgBox (intFrame(0))
    'End If
    intFrame(Index) = intFrame(Index) + 1
    If strDir(Index) = "L" Then
        Call getCharJumpAnim(Index, intFrame(Index), Tile(curX(Index), curY(Index)).X - 25, Tile(curX(Index), curY(Index)).Y - 115, Tile(curX(Index), curY(Index)).X + 25, Tile(curX(Index), curY(Index)).Y - 15)
    ElseIf strDir(Index) = "R" Then
        Call getCharJumpAnim(Index, intFrame(Index), Tile(curX(Index), curY(Index)).X + 75, Tile(curX(Index), curY(Index)).Y + 85, Tile(curX(Index), curY(Index)).X + 25, Tile(curX(Index), curY(Index)).Y - 15)
    ElseIf strDir(Index) = "U" Then
        Call getCharJumpAnim(Index, intFrame(Index), Tile(curX(Index), curY(Index)).X + 75, Tile(curX(Index), curY(Index)).Y - 115, Tile(curX(Index), curY(Index)).X + 25, Tile(curX(Index), curY(Index)).Y - 15)
    ElseIf strDir(Index) = "D" Then
        Call getCharJumpAnim(Index, intFrame(Index), Tile(curX(Index), curY(Index)).X - 25, Tile(curX(Index), curY(Index)).Y + 85, Tile(curX(Index), curY(Index)).X + 25, Tile(curX(Index), curY(Index)).Y - 15)
    End If
    If intFrame(Index) = 1 Then
        strState(Index) = "C"
    ElseIf intFrame(Index) = 5 Then
        strState(Index) = "J"
    ElseIf intFrame(Index) = 10 Then
        strState(Index) = "I"
        intFrame(Index) = 0
    End If
Else
    strState(Index) = "I"
    intFrame(Index) = 0
    spriteX(Index) = Tile(curX(Index), curY(Index)).X + 25
    spriteY(Index) = Tile(curX(Index), curY(Index)).Y - 15
    'Call PaintCharSprite(Index, Tile(curX(Index), curY(Index)).X + 25, Tile(curX(Index), curY(Index)).Y - 15)
End If
If blnRev(Index) = False Then
    picCount(Index) = picCount(Index) + 1
End If
If blnRev(Index) = True Then
    picCount(Index) = picCount(Index) - 1
End If
If picCount(Index) >= 4 Or picCount(Index) <= 0 Then
    If blnRev(Index) = False Then
        blnRev(Index) = True
    Else
        blnRev(Index) = False
    End If
End If
End Sub

Private Function getCharJumpAnim(ByVal Index As Integer, ByVal intFrame As Integer, ByVal IntNewX As Integer, ByVal intNewY As Integer, ByVal intOldX As Integer, ByVal intOldY As Integer)
spriteX(Index) = intOldX + (intFrame * Int((IntNewX - intOldX) / 10))
If intFrame <= 5 Then
    spriteY(Index) = (intOldY + (intFrame * Int((intNewY - intOldY) / 10))) - (4 * intFrame)
    'Call PaintCharSprite(Index, intOldX + (intFrame * ((IntNewX - intOldX) / 10)), (intOldY + (intFrame * ((intNewY - intOldY) / 10))) - (4 * intFrame))
Else
    spriteY(Index) = (intOldY + (intFrame * Int((intNewY - intOldY) / 10))) - (4 * (10 - intFrame))
    'Call PaintCharSprite(Index, intOldX + (intFrame * ((IntNewX - intOldX) / 10)), (intOldY + (intFrame * ((intNewY - intOldY) / 10))) - (4 * (10 - intFrame)))
End If
End Function

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
'cpu 1 moves every (counterLimit + 1) seconds
counterLimit(1) = 1
End Sub
