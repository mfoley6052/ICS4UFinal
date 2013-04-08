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
   Begin VB.Timer tmrSel 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   125
      Left            =   6120
      Top             =   1320
   End
   Begin VB.Timer tmrSel 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   125
      Left            =   5640
      Top             =   1320
   End
   Begin VB.Timer tmrSel 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   125
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
      Left            =   6720
      Picture         =   "frmMain.frx":0000
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
      Interval        =   66
      Left            =   5040
      Top             =   840
   End
   Begin VB.Timer tmrTileAnimDelay 
      Enabled         =   0   'False
      Interval        =   125
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
      Interval        =   250
      Left            =   7680
      Top             =   360
   End
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   250
      Left            =   7200
      Top             =   360
   End
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   6720
      Top             =   360
   End
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   250
      Left            =   6240
      Top             =   360
   End
   Begin VB.Timer tmrSel 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   125
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
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   9
         Left            =   7920
         Picture         =   "frmMain.frx":0363
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
         Picture         =   "frmMain.frx":04A0
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
         Picture         =   "frmMain.frx":05DF
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
         Picture         =   "frmMain.frx":0720
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
         Picture         =   "frmMain.frx":0863
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
         Picture         =   "frmMain.frx":09A5
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
         Picture         =   "frmMain.frx":0B19
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
         Picture         =   "frmMain.frx":0C7B
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
         Picture         =   "frmMain.frx":0E09
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
         Picture         =   "frmMain.frx":0F74
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
         Picture         =   "frmMain.frx":10D3
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
         Picture         =   "frmMain.frx":1190
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
         Picture         =   "frmMain.frx":1254
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
         Picture         =   "frmMain.frx":1328
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
         Picture         =   "frmMain.frx":13DF
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
         Picture         =   "frmMain.frx":148E
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
         Picture         =   "frmMain.frx":152E
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
         Picture         =   "frmMain.frx":15C1
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
         Picture         =   "frmMain.frx":1659
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
         Picture         =   "frmMain.frx":16EC
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
         Picture         =   "frmMain.frx":1789
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
         Picture         =   "frmMain.frx":1833
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
         Picture         =   "frmMain.frx":18E9
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
         Picture         =   "frmMain.frx":19BD
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
         Picture         =   "frmMain.frx":1A9F
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
         Picture         =   "frmMain.frx":1E20
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
         Picture         =   "frmMain.frx":21A8
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
         Picture         =   "frmMain.frx":2532
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
         Picture         =   "frmMain.frx":28B9
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
         Picture         =   "frmMain.frx":2C43
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
         Picture         =   "frmMain.frx":2FC9
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
         Picture         =   "frmMain.frx":3350
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
         Picture         =   "frmMain.frx":36D6
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
         Picture         =   "frmMain.frx":3A5B
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
         Picture         =   "frmMain.frx":3DE0
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
         Picture         =   "frmMain.frx":4169
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
         Picture         =   "frmMain.frx":44F3
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
         Picture         =   "frmMain.frx":487F
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
         Picture         =   "frmMain.frx":4C06
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
         Picture         =   "frmMain.frx":4CC3
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
         Picture         =   "frmMain.frx":4D87
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
         Picture         =   "frmMain.frx":4E5B
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
         Picture         =   "frmMain.frx":4F12
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
         Picture         =   "frmMain.frx":4FC1
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
         Picture         =   "frmMain.frx":5061
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
         Picture         =   "frmMain.frx":50F4
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
         Picture         =   "frmMain.frx":518C
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
         Picture         =   "frmMain.frx":521F
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
         Picture         =   "frmMain.frx":52BC
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
         Picture         =   "frmMain.frx":5366
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
         Picture         =   "frmMain.frx":541C
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
         Picture         =   "frmMain.frx":54F0
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
         Left            =   7440
         Picture         =   "frmMain.frx":55D2
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
         Left            =   7080
         Picture         =   "frmMain.frx":5925
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
         Left            =   6720
         Picture         =   "frmMain.frx":5C7E
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
         Left            =   7440
         Picture         =   "frmMain.frx":5CD4
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
         Left            =   7080
         Picture         =   "frmMain.frx":5D19
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
         Left            =   6360
         Picture         =   "frmMain.frx":5D69
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
         Left            =   6000
         Picture         =   "frmMain.frx":5DBE
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
         Picture         =   "frmMain.frx":5E14
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
         Left            =   5640
         Picture         =   "frmMain.frx":6167
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
         Left            =   6360
         Picture         =   "frmMain.frx":64C1
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
         Left            =   6000
         Picture         =   "frmMain.frx":6821
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
         Left            =   5640
         Picture         =   "frmMain.frx":6B82
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
         Picture         =   "frmMain.frx":6BD2
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
         Picture         =   "frmMain.frx":6C17
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
         Picture         =   "frmMain.frx":6CF9
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
         Picture         =   "frmMain.frx":6DCD
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
         Picture         =   "frmMain.frx":6E83
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
         Picture         =   "frmMain.frx":6F2D
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
         Picture         =   "frmMain.frx":6FCA
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
         Picture         =   "frmMain.frx":705D
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
         Picture         =   "frmMain.frx":70F5
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
         Picture         =   "frmMain.frx":7188
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
         Picture         =   "frmMain.frx":7228
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
         Picture         =   "frmMain.frx":72D7
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
         Picture         =   "frmMain.frx":738E
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
         Picture         =   "frmMain.frx":7462
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
         Picture         =   "frmMain.frx":7526
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
         Picture         =   "frmMain.frx":75E3
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
         Picture         =   "frmMain.frx":7745
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
         Picture         =   "frmMain.frx":78A4
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
         Picture         =   "frmMain.frx":7A0F
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
         Picture         =   "frmMain.frx":7B9D
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
         Picture         =   "frmMain.frx":7D11
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
         Picture         =   "frmMain.frx":7E53
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
         Picture         =   "frmMain.frx":7F96
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
         Picture         =   "frmMain.frx":80D7
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
         Picture         =   "frmMain.frx":8216
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
         Picture         =   "frmMain.frx":8353
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
         Picture         =   "frmMain.frx":95C9
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
         Picture         =   "frmMain.frx":9DE6
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
         Picture         =   "frmMain.frx":A559
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   1
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Label lblScore 
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
Dim dirJump(0 To 3) As String
Dim counterLimit(1 To 3) As Integer
Dim intScore As Integer
Dim curX(0 To 3) As Integer
Dim curY(0 To 3) As Integer
Dim prevX(0 To 3) As Integer
Dim prevY(0 To 3) As Integer
Dim coinTileCount As Integer
Dim blnClearPrevTile(0 To 3) As Boolean
Dim tileSwitch(0 To 100) As Boolean

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

Private Function getJump(ByVal index As Integer, ByVal strDir As String)
dirJump(index) = strDir
tmrJump(index).Enabled = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
End Sub

Private Function evalMove(index As Integer, ByVal strDir As String) As Boolean
If strDir = "L" Then
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
ElseIf strDir = "U" Then
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
ElseIf strDir = "R" Then
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
ElseIf strDir = "D" Then
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

Private Sub tmrCPUMove_Timer(index As Integer)
Static intCounter As Integer
If intCounter < counterLimit(index) Then
    intCounter = intCounter + 1
Else
    Call cpuAI(index)
    intCounter = 0
End If
End Sub

Private Function cpuAI(ByVal index As Integer)
'player x further than cpu x
If curX(index) < curX(0) Then
    'player y lower than cpu y
    If curY(index) < curY(0) Then
        If evalMove(index, "R") = True Then
            Call getJump(index, "R")
        End If
    'player y matches cpu y
    ElseIf curY(index) = curY(0) Then
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
    'player y lower than cpu y
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
    If curY(index) < curY(0) Then
        If evalMove(index, "D") = True Then
            Call getJump(index, "D")
        End If
    'player y matches cpu y
    ElseIf curY(index) = curY(0) Then
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
End Function

Private Sub tmrJump_Timer(index As Integer)
Dim pScore As Integer
prevX(index) = curX(index)
prevY(index) = curY(index)
If dirJump(index) = "L" Then
    'if y row is odd
    If (curY(index) + 1) Mod 2 = 0 Then
        curX(index) = curX(index) - 1
    End If
    curY(index) = curY(index) - 1
ElseIf dirJump(index) = "U" Then
    'if y row is even
    If (curY(index) + 1) Mod 2 = 1 Then
        curX(index) = curX(index) + 1
    End If
    curY(index) = curY(index) - 1
ElseIf dirJump(index) = "R" Then
    'if y row is even
    If (curY(index) + 1) Mod 2 = 1 Then
        curX(index) = curX(index) + 1
    End If
    curY(index) = curY(index) + 1
ElseIf dirJump(index) = "D" Then
    'if y row is odd
    If (curY(index) + 1) Mod 2 = 0 Then
        curX(index) = curX(index) - 1
    End If
    curY(index) = curY(index) + 1
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
tmrJump(index).Enabled = False
blnClearPrevTile(index) = True
End Sub

Private Sub Form_Load()
Call DrawMap(1)
curX(0) = 2
curY(0) = 2
curX(1) = 0
curY(1) = 0
blnClearPrevTile(0) = False
blnClearPrevTile(1) = False
counterLimit(1) = 2
For t = 0 To 100
    tileSwitch(t) = False
Next t
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
            Call clearTile(getTileFromInt(True, c), getTileFromInt(False, c))
            coinTileCount = coinTileCount - 1
        End If
    End If
Next c
frameCount = frameCount + 1
If frameCount = 28 Then
    frameCount = 0
End If
End Sub

Private Function clearTile(ByVal intX As Integer, ByVal intY As Integer)
'paint over last coin
picBackground.PaintPicture frmMain.picScene(0).Image, Tile(intX, intY).X, Tile(intX, intY).Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
picBuffer.PaintPicture frmMain.picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
frmMain.PaintPicture frmMain.picMask.Image, Tile(intX, intY).X, Tile(intX, intY).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
frmMain.PaintPicture picBuffer.Image, Tile(intX, intY).X, Tile(intX, intY).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
'paint over frame
picBackground.PaintPicture frmMain.picScene(0).Image, Tile(intX, intY).X, Tile(intX, intY).Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
picBuffer.PaintPicture frmMain.picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
frmMain.PaintPicture frmMain.picMask.Image, Tile(intX, intY).X, Tile(intX, intY).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
frmMain.PaintPicture picBuffer.Image, Tile(intX, intY).X, Tile(intX, intY).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
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
Call clearTile(intCoinX, intCoinY)
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

Private Sub tmrSel_Timer(index As Integer)
Static blnRev(0 To 3) As Boolean
Call PaintSelector(index, picCount(index))
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

Private Function PaintSelector(ByVal index As Integer, ByVal imgIndex As Integer) As Integer
'paint over last sel
If blnClearPrevTile(index) = True Then
    picBackground.PaintPicture frmMain.picScene(0).Image, Tile(prevX(index), prevY(index)).X, Tile(prevX(index), prevY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
    picBuffer.PaintPicture frmMain.picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
    frmMain.PaintPicture frmMain.picMask.Image, Tile(prevX(index), prevY(index)).X, Tile(prevX(index), prevY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    frmMain.PaintPicture picBuffer.Image, Tile(prevX(index), prevY(index)).X, Tile(prevX(index), prevY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    blnClearPrevTile(index) = False
End If
'paint over frame
picBackground.PaintPicture frmMain.picScene(0).Image, Tile(curX(index), curY(index)).X, Tile(curX(index), curY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
picBuffer.PaintPicture frmMain.picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
frmMain.PaintPicture frmMain.picMask.Image, Tile(curX(index), curY(index)).X, Tile(curX(index), curY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
frmMain.PaintPicture picBuffer.Image, Tile(curX(index), curY(index)).X, Tile(curX(index), curY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
'paint sel
frmMain.PaintPicture picSelMask(imgIndex + 5 * index).Image, Tile(curX(index), curY(index)).X, Tile(curX(index), curY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
frmMain.PaintPicture picSel(imgIndex + 5 * index).Image, Tile(curX(index), curY(index)).X, Tile(curX(index), curY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
frmDbg.txtTest(0).Text = "(" & curX(0) & ", " & curY(0) & ")"
End Function

Private Function getTileAnim(ByVal intFrame As Integer, ByVal intX As Integer, ByVal intY As Integer, ByVal counter As Integer)
'picBackground.PaintPicture frmMain.picScene(0).Image, Tile(prevX(index), prevY(index)).X, Tile(prevX(index), prevY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
picBuffer.PaintPicture frmMain.picBackground.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
'frmMain.PaintPicture frmMain.picMask.Image, Tile(prevX(index), prevY(index)).X, Tile(prevX(index), prevY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
frmMain.PaintPicture picBuffer.Image, Tile(intX, intY).X, (Tile(intX, intY).Y - 200) + (intFrame - 1) * 25, 100, 100, 0, 0, 100, 100, vbSrcCopy
For X = 0 To ((mapWidth * mapHeight) + (Int(mapHeight / 2))) - 1
    If tileSwitch(X) = True And counter - (4 * X) > 8 Then
        PaintPicture frmMain.picMask.Picture, Tile(getTileFromInt(True, X), getTileFromInt(False, X)).X, Tile(getTileFromInt(True, X), getTileFromInt(False, X)).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
        PaintPicture frmMain.picScene(0).Picture, Tile(getTileFromInt(True, X), getTileFromInt(False, X)).X, Tile(getTileFromInt(True, X), getTileFromInt(False, X)).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    End If
Next X
'paint tile with new y
PaintPicture frmMain.picMask.Picture, Tile(intX, intY).X, (Tile(intX, intY).Y - 200) + intFrame * 25, 100, 100, 0, 0, 100, 100, vbSrcAnd
PaintPicture frmMain.picScene(0).Picture, Tile(intX, intY).X, (Tile(intX, intY).Y - 200) + intFrame * 25, 100, 100, 0, 0, 100, 100, vbSrcPaint
End Function

Private Sub tmrTileAnim_Timer()
Static intCounter As Integer
frmDbg.txtTest(1).Text = intCounter
For X = 0 To tileCount - 1
    If tileSwitch(X) = True And intCounter - (4 * X) <= 8 Then
        Call getTileAnim(intCounter - (4 * X), getTileFromInt(True, X), getTileFromInt(False, X), intCounter)
    End If
Next X
intCounter = intCounter + 1
If intCounter = 6 * (((mapWidth * mapHeight) + (Int(mapHeight / 2))) - 1) Then
    tmrTileAnim.Enabled = False
    tmrSel(0).Enabled = True
    tmrSel(1).Enabled = True
    tmrCoinEvent.Enabled = True
    tmrCoin.Enabled = True
    tmrCPUMove(1).Enabled = True
End If
End Sub

Private Sub tmrTileAnimDelay_Timer()
Static intCounter As Integer
Static intX As Integer
Static intY As Integer
frmMain.tmrTileAnim.Enabled = True
For z = 0 To intCounter - getAbs(intCounter - ((mapWidth * mapHeight) + (Int(mapHeight / 2))) - 1)
    tileSwitch(z) = True
Next z
frmDbg.lstMap.AddItem (intCounter & ": " & z - 1 & ", " & getAbs(intCounter - 15) & ", " & intCounter - getAbs(intCounter - (((mapWidth * mapHeight) + (Int(mapHeight / 2)))) - 1))
If intX = mapWidth - ((intY + 1) Mod 2) Then
    intY = intY + 1
    intX = 0
ElseIf intX < mapWidth + ((mapHeight + 1) Mod 2) Then
    intX = intX + 1
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
