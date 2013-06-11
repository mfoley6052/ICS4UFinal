VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dino Hopper"
   ClientHeight    =   9555
   ClientLeft      =   4245
   ClientTop       =   1230
   ClientWidth     =   12000
   LinkTopic       =   "Dino Hopper"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   637
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   10080
      Top             =   1080
   End
   Begin VB.Timer tmrStun 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   1000
      Left            =   11520
      Top             =   4440
   End
   Begin VB.Timer tmrStun 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   1000
      Left            =   11520
      Top             =   3960
   End
   Begin VB.Timer tmrStun 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   1000
      Left            =   11520
      Top             =   3480
   End
   Begin VB.Timer tmrStun 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   11520
      Top             =   3000
   End
   Begin VB.Timer tmrAlternate 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   11040
      Top             =   2520
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Index           =   4
      Left            =   240
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   201
      Top             =   2520
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picTileL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Index           =   4
      Left            =   240
      Picture         =   "frmMain.frx":1268
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   200
      Top             =   2640
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picTileR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Index           =   4
      Left            =   240
      Picture         =   "frmMain.frx":24AF
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   199
      Top             =   2760
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picTileLR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Index           =   4
      Left            =   240
      Picture         =   "frmMain.frx":36ED
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   198
      Top             =   2880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Timer tmrPow 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   250
      Left            =   10560
      Top             =   2520
   End
   Begin VB.Timer tmrPow 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   250
      Left            =   10560
      Top             =   2040
   End
   Begin VB.Timer tmrPow 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   250
      Left            =   10560
      Top             =   1560
   End
   Begin VB.Timer tmrPow 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   250
      Left            =   10560
      Top             =   1080
   End
   Begin VB.Timer tmrChar 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   40
      Left            =   10560
      Top             =   4440
   End
   Begin VB.Timer tmrChar 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   40
      Left            =   10560
      Top             =   3960
   End
   Begin VB.Timer tmrChar 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   40
      Left            =   10560
      Top             =   3480
   End
   Begin VB.Timer tmrCPUMove 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   300
      Left            =   11520
      Top             =   2040
   End
   Begin VB.Timer tmrCPUMove 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   300
      Left            =   11520
      Top             =   1560
   End
   Begin VB.Timer tmrCPUMove 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   300
      Left            =   11520
      Top             =   1080
   End
   Begin VB.Timer tmrObjEvent 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   11520
      Top             =   600
   End
   Begin VB.PictureBox picSparkleMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   4
      Left            =   6240
      Picture         =   "frmMain.frx":491B
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   38
      Top             =   6480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer tmrObj 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   10080
      Top             =   600
   End
   Begin VB.Timer tmrTileAnimDelay 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   11040
      Top             =   600
   End
   Begin VB.Timer tmrTileAnim 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   10560
      Top             =   600
   End
   Begin VB.Timer tmrChar 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   40
      Left            =   10560
      Top             =   3000
   End
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9555
      Left            =   0
      Picture         =   "frmMain.frx":4C7E
      ScaleHeight     =   637
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   12000
      Begin VB.PictureBox picMaskTop 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   4920
         Picture         =   "frmMain.frx":2BB09
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   266
         Top             =   7320
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picMaskSides 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1125
         Left            =   1800
         Picture         =   "frmMain.frx":2BBE9
         ScaleHeight     =   75
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   265
         Top             =   7320
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picPowSpeed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   9120
         Picture         =   "frmMain.frx":2BCE9
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   264
         Top             =   480
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSpacer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   240
         Picture         =   "frmMain.frx":2BD6E
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   263
         Top             =   7320
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picMaskTopLR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   3360
         Picture         =   "frmMain.frx":2BE85
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   262
         Top             =   7320
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picP4ID 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   1800
         Picture         =   "frmMain.frx":2BF67
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   261
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4JD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   4680
         Picture         =   "frmMain.frx":2C460
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   260
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   3240
         Picture         =   "frmMain.frx":2C942
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   259
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   3960
         Picture         =   "frmMain.frx":2CE2D
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   258
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   2520
         Picture         =   "frmMain.frx":2D304
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   257
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4IR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   1800
         Picture         =   "frmMain.frx":2D7F5
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   256
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4JR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   4680
         Picture         =   "frmMain.frx":2DCE1
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   255
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   2520
         Picture         =   "frmMain.frx":2E1B2
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   254
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   3960
         Picture         =   "frmMain.frx":2E69A
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   253
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   3240
         Picture         =   "frmMain.frx":2EB63
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   252
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4IL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   5520
         Picture         =   "frmMain.frx":2F03F
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   251
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4JL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   8400
         Picture         =   "frmMain.frx":2F4E5
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   250
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":2F982
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   249
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   7680
         Picture         =   "frmMain.frx":2FE10
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   248
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":302A6
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   247
         Top             =   3600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4IU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   5520
         Picture         =   "frmMain.frx":30741
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   246
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4JU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   8400
         Picture         =   "frmMain.frx":30BD6
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   245
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":31071
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   244
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   7680
         Picture         =   "frmMain.frx":314FD
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   243
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP4CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":3197F
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   242
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3ID 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   1800
         Picture         =   "frmMain.frx":31E03
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   241
         Top             =   3120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3JD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   4680
         Picture         =   "frmMain.frx":322FC
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   240
         Top             =   3120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   3240
         Picture         =   "frmMain.frx":327DE
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   239
         Top             =   3120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   3960
         Picture         =   "frmMain.frx":32CC9
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   238
         Top             =   3120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   2520
         Picture         =   "frmMain.frx":331A0
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   237
         Top             =   3120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3IR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   1800
         Picture         =   "frmMain.frx":33691
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   236
         Top             =   2880
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3JR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   4680
         Picture         =   "frmMain.frx":33B7D
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   235
         Top             =   2880
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   2520
         Picture         =   "frmMain.frx":3404E
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   234
         Top             =   2880
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   3960
         Picture         =   "frmMain.frx":34536
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   233
         Top             =   2880
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   3240
         Picture         =   "frmMain.frx":349FF
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   232
         Top             =   2880
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3IL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   5520
         Picture         =   "frmMain.frx":34EDB
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   231
         Top             =   3120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3JL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   8400
         Picture         =   "frmMain.frx":35381
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   230
         Top             =   3120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":3581E
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   229
         Top             =   3120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   7680
         Picture         =   "frmMain.frx":35CAC
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   228
         Top             =   3120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":36142
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   227
         Top             =   3120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3IU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   5520
         Picture         =   "frmMain.frx":365DD
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   226
         Top             =   2880
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3JU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   8400
         Picture         =   "frmMain.frx":36A72
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   225
         Top             =   2880
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":36F0D
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   224
         Top             =   2880
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   7680
         Picture         =   "frmMain.frx":37399
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   223
         Top             =   2880
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP3CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":3781B
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   222
         Top             =   2880
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2ID 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   1800
         Picture         =   "frmMain.frx":37C9F
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   221
         Top             =   2640
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2JD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   4680
         Picture         =   "frmMain.frx":38198
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   220
         Top             =   2640
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   3240
         Picture         =   "frmMain.frx":3867A
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   219
         Top             =   2640
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   3960
         Picture         =   "frmMain.frx":38B65
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   218
         Top             =   2640
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CD 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   2520
         Picture         =   "frmMain.frx":3903C
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   217
         Top             =   2640
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2IR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   1800
         Picture         =   "frmMain.frx":3952D
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   216
         Top             =   2400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2JR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   4680
         Picture         =   "frmMain.frx":39A19
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   215
         Top             =   2400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   2520
         Picture         =   "frmMain.frx":39EEA
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   214
         Top             =   2400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   3960
         Picture         =   "frmMain.frx":3A3D2
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   213
         Top             =   2400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   3240
         Picture         =   "frmMain.frx":3A89B
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   212
         Top             =   2400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2IL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   5520
         Picture         =   "frmMain.frx":3AD77
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   211
         Top             =   2640
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2JL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   8400
         Picture         =   "frmMain.frx":3B21D
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   210
         Top             =   2640
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":3B6BA
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   209
         Top             =   2640
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   7680
         Picture         =   "frmMain.frx":3BB48
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   208
         Top             =   2640
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":3BFDE
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   207
         Top             =   2640
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2IU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   5520
         Picture         =   "frmMain.frx":3C479
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   206
         Top             =   2400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2JU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   8400
         Picture         =   "frmMain.frx":3C90E
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   205
         Top             =   2400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":3CDA9
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   204
         Top             =   2400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   7680
         Picture         =   "frmMain.frx":3D235
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   203
         Top             =   2400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP2CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":3D6B7
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   202
         Top             =   2400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":3DB3B
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   119
         Top             =   2160
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   7680
         Picture         =   "frmMain.frx":3DD02
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   118
         Top             =   2160
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":3DEC2
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   117
         Top             =   2160
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1JL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   8400
         Picture         =   "frmMain.frx":3E07C
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   116
         Top             =   2160
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1IL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   5520
         Picture         =   "frmMain.frx":3E243
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   114
         Top             =   2160
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
         Left            =   2520
         Picture         =   "frmMain.frx":3E413
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   99
         Top             =   2160
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
         Left            =   3960
         Picture         =   "frmMain.frx":3E627
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   98
         Top             =   2160
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
         Left            =   3240
         Picture         =   "frmMain.frx":3E820
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   97
         Top             =   2160
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
         Left            =   4680
         Picture         =   "frmMain.frx":3EA2A
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   96
         Top             =   2160
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
         Left            =   1800
         Picture         =   "frmMain.frx":3EC30
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   94
         Top             =   2160
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picPowFreezeMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   10740
         Picture         =   "frmMain.frx":3EE42
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   197
         Top             =   4320
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picPowFreeze 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   6
         Left            =   11280
         Picture         =   "frmMain.frx":3EF1F
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   196
         Top             =   7650
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picPowFreeze 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   5
         Left            =   11280
         Picture         =   "frmMain.frx":3F059
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   195
         Top             =   7095
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picPowFreeze 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   4
         Left            =   11280
         Picture         =   "frmMain.frx":3F193
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   194
         Top             =   6540
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picPowFreeze 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   3
         Left            =   11280
         Picture         =   "frmMain.frx":3F2CD
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   193
         Top             =   5985
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picPowFreeze 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   2
         Left            =   11280
         Picture         =   "frmMain.frx":3F3F5
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   192
         Top             =   5430
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picPowFreeze 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   1
         Left            =   11280
         Picture         =   "frmMain.frx":3F530
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   191
         Top             =   4875
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picPowFreeze 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   0
         Left            =   11280
         Picture         =   "frmMain.frx":3F665
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   190
         Top             =   4320
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox picEggG 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   0
         Left            =   360
         Picture         =   "frmMain.frx":3F79B
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   189
         Top             =   3840
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picEggG 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   1
         Left            =   360
         Picture         =   "frmMain.frx":3F880
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   188
         Top             =   4320
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picEggG 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   2
         Left            =   360
         Picture         =   "frmMain.frx":3F95C
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   187
         Top             =   4800
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picTileLR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":3FA3A
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   186
         Top             =   360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   6600
         Picture         =   "frmMain.frx":409D0
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   184
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTileL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   6600
         Picture         =   "frmMain.frx":419D2
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   183
         Top             =   120
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTileR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   6600
         Picture         =   "frmMain.frx":4298E
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   182
         Top             =   240
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTileLR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   3
         Left            =   5040
         Picture         =   "frmMain.frx":4397E
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   181
         Top             =   360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTileR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   3
         Left            =   5040
         Picture         =   "frmMain.frx":44BB0
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   180
         Top             =   240
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTileL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   3
         Left            =   5040
         Picture         =   "frmMain.frx":45DE2
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   179
         Top             =   120
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   3
         Left            =   5040
         Picture         =   "frmMain.frx":47028
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   178
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   2
         Left            =   3480
         Picture         =   "frmMain.frx":4827B
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   176
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTileL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   2
         Left            =   3480
         Picture         =   "frmMain.frx":494D9
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   175
         Top             =   120
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTileR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   2
         Left            =   3480
         Picture         =   "frmMain.frx":4BDDD
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   174
         Top             =   240
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picMaskLR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   1800
         Picture         =   "frmMain.frx":4E6EE
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   173
         Top             =   360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picMaskR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   1800
         Picture         =   "frmMain.frx":4E807
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   172
         Top             =   240
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picMaskL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   1800
         Picture         =   "frmMain.frx":4E91D
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   171
         Top             =   120
         Visible         =   0   'False
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
         Picture         =   "frmMain.frx":4EA37
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   170
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTileR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":4EB4C
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   168
         Top             =   240
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTileL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":4FAFA
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   169
         Top             =   120
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picPowSpeed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   9360
         Picture         =   "frmMain.frx":50AB6
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   167
         Top             =   480
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   8880
         Picture         =   "frmMain.frx":50B39
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   166
         Top             =   480
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   8640
         Picture         =   "frmMain.frx":50BBE
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   165
         Top             =   480
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   8400
         Picture         =   "frmMain.frx":50C46
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   164
         Top             =   480
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeedMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   9360
         Picture         =   "frmMain.frx":50CCE
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   163
         Top             =   1440
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeedMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   9120
         Picture         =   "frmMain.frx":50D3B
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   162
         Top             =   1440
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeedMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   8880
         Picture         =   "frmMain.frx":50DAA
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   161
         Top             =   1440
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeedMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   8640
         Picture         =   "frmMain.frx":50E1B
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   160
         Top             =   1440
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeedMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   8400
         Picture         =   "frmMain.frx":50E8E
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   159
         Top             =   1440
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   9360
         Picture         =   "frmMain.frx":50F01
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   158
         Top             =   0
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   9120
         Picture         =   "frmMain.frx":50F89
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   157
         Top             =   0
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   8880
         Picture         =   "frmMain.frx":51010
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   156
         Top             =   0
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   8640
         Picture         =   "frmMain.frx":51095
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   155
         Top             =   0
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeed 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   8400
         Picture         =   "frmMain.frx":51118
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   154
         Top             =   0
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeedMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   9360
         Picture         =   "frmMain.frx":51199
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   153
         Top             =   960
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeedMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   9120
         Picture         =   "frmMain.frx":5120C
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   152
         Top             =   960
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeedMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   8880
         Picture         =   "frmMain.frx":5127D
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   151
         Top             =   960
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeedMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   8640
         Picture         =   "frmMain.frx":512ED
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   150
         Top             =   960
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picPowSpeedMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   8400
         Picture         =   "frmMain.frx":5135D
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   149
         Top             =   960
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picEgg 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   2
         Left            =   720
         Picture         =   "frmMain.frx":513CA
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   148
         Top             =   4800
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picEgg 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   1
         Left            =   720
         Picture         =   "frmMain.frx":514BE
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   147
         Top             =   4320
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picEgg 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   0
         Left            =   720
         Picture         =   "frmMain.frx":515B5
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   146
         Top             =   3840
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picEggMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   2
         Left            =   1200
         Picture         =   "frmMain.frx":516AD
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   145
         Top             =   4800
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picEggMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   1
         Left            =   1200
         Picture         =   "frmMain.frx":51713
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   144
         Top             =   4320
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picEggMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   0
         Left            =   1200
         Picture         =   "frmMain.frx":5177E
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   143
         Top             =   3840
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picPowScareMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   0
         Left            =   9720
         Picture         =   "frmMain.frx":517EB
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   142
         Top             =   2160
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox picPowScareMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   1
         Left            =   9720
         Picture         =   "frmMain.frx":51858
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   141
         Top             =   2640
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox picPowScareMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   2
         Left            =   9720
         Picture         =   "frmMain.frx":518C9
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   140
         Top             =   3120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox picPowScareMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   3
         Left            =   9720
         Picture         =   "frmMain.frx":5193D
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   139
         Top             =   3600
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox picPowScareMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   4
         Left            =   9720
         Picture         =   "frmMain.frx":519B0
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   138
         Top             =   4080
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox picPowScare 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   0
         Left            =   9240
         Picture         =   "frmMain.frx":51A1B
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   137
         Top             =   2160
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox picPowScare 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   1
         Left            =   9240
         Picture         =   "frmMain.frx":51B17
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   136
         Top             =   2640
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox picPowScare 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   2
         Left            =   9240
         Picture         =   "frmMain.frx":51C24
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   135
         Top             =   3120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox picPowScare 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   3
         Left            =   9240
         Picture         =   "frmMain.frx":51D3D
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   134
         Top             =   3600
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox picPowScare 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   4
         Left            =   9240
         Picture         =   "frmMain.frx":51E53
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   133
         Top             =   4080
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox picCharMaskIU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   5520
         Picture         =   "frmMain.frx":51F68
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   132
         Top             =   4320
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":5201B
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   131
         Top             =   1920
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   7680
         Picture         =   "frmMain.frx":52193
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   130
         Top             =   1920
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1CU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":52341
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   129
         Top             =   1920
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":524C3
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   128
         Top             =   4320
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   7680
         Picture         =   "frmMain.frx":5256F
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   127
         Top             =   4320
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":52618
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   126
         Top             =   4320
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1JU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   8400
         Picture         =   "frmMain.frx":526C7
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   125
         Top             =   1920
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskJU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   8400
         Picture         =   "frmMain.frx":52889
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   124
         Top             =   4320
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picP1IU 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   5520
         Picture         =   "frmMain.frx":5293F
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   123
         Top             =   1920
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":52ACA
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   122
         Top             =   4440
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   3
         Left            =   7680
         Picture         =   "frmMain.frx":52B7A
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   121
         Top             =   4440
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskCL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":52C26
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   120
         Top             =   4440
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskJL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   8400
         Picture         =   "frmMain.frx":52CD2
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   115
         Top             =   4440
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskIL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   5520
         Picture         =   "frmMain.frx":52D85
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   113
         Top             =   4440
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskIR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   1800
         Picture         =   "frmMain.frx":52E38
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   112
         Top             =   4320
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
         Left            =   3240
         Picture         =   "frmMain.frx":52EFB
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   111
         Top             =   1920
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
         Left            =   3960
         Picture         =   "frmMain.frx":530F9
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   110
         Top             =   1920
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
         Left            =   2520
         Picture         =   "frmMain.frx":532E5
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   109
         Top             =   1920
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
         Left            =   3240
         Picture         =   "frmMain.frx":534EC
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   108
         Top             =   4320
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
         Left            =   3960
         Picture         =   "frmMain.frx":535AA
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   107
         Top             =   4320
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
         Left            =   2520
         Picture         =   "frmMain.frx":53665
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   106
         Top             =   4320
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
         Left            =   4680
         Picture         =   "frmMain.frx":53727
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   105
         Top             =   1920
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
         Left            =   4680
         Picture         =   "frmMain.frx":5391F
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   104
         Top             =   4320
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
         Left            =   1800
         Picture         =   "frmMain.frx":539DF
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   103
         Top             =   1920
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
         Left            =   2520
         Picture         =   "frmMain.frx":53BED
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   102
         Top             =   4440
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
         Left            =   3960
         Picture         =   "frmMain.frx":53CB1
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   101
         Top             =   4440
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
         Left            =   3240
         Picture         =   "frmMain.frx":53D6E
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   100
         Top             =   4440
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
         Left            =   4680
         Picture         =   "frmMain.frx":53E2D
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   95
         Top             =   4440
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picCharMaskID 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   1800
         Picture         =   "frmMain.frx":53EEF
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   93
         Top             =   4440
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   9
         Left            =   7800
         Picture         =   "frmMain.frx":53FB4
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   87
         Top             =   5760
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
         Left            =   7680
         Picture         =   "frmMain.frx":54128
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   86
         Top             =   5760
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
         Left            =   7560
         Picture         =   "frmMain.frx":5428A
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   85
         Top             =   5760
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
         Left            =   7440
         Picture         =   "frmMain.frx":54418
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   84
         Top             =   5760
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
         Left            =   7320
         Picture         =   "frmMain.frx":54583
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   83
         Top             =   5760
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
         Picture         =   "frmMain.frx":546E2
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
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":5479F
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
         Index           =   2
         Left            =   960
         Picture         =   "frmMain.frx":54863
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
         Index           =   3
         Left            =   1320
         Picture         =   "frmMain.frx":54937
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
         Index           =   4
         Left            =   1680
         Picture         =   "frmMain.frx":549EE
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
         Index           =   5
         Left            =   2040
         Picture         =   "frmMain.frx":54A9D
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
         Index           =   6
         Left            =   2400
         Picture         =   "frmMain.frx":54B3D
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
         Index           =   7
         Left            =   2760
         Picture         =   "frmMain.frx":54BD0
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
         Index           =   8
         Left            =   3120
         Picture         =   "frmMain.frx":54C68
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
         Index           =   9
         Left            =   3480
         Picture         =   "frmMain.frx":54CFB
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   73
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
         Picture         =   "frmMain.frx":54D98
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   72
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
         Picture         =   "frmMain.frx":54E42
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   71
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
         Picture         =   "frmMain.frx":54EF8
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   70
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
         Picture         =   "frmMain.frx":54FCC
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   69
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
         Picture         =   "frmMain.frx":550AE
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   68
         Top             =   5760
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
         Picture         =   "frmMain.frx":5542F
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   67
         Top             =   5760
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
         Picture         =   "frmMain.frx":557B7
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   66
         Top             =   5760
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
         Picture         =   "frmMain.frx":55B41
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   65
         Top             =   5760
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
         Picture         =   "frmMain.frx":55EC8
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   64
         Top             =   5760
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
         Picture         =   "frmMain.frx":56252
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   63
         Top             =   5760
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
         Picture         =   "frmMain.frx":565D8
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   62
         Top             =   5760
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
         Picture         =   "frmMain.frx":5695F
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   61
         Top             =   5760
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
         Picture         =   "frmMain.frx":56CE5
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   60
         Top             =   5760
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
         Picture         =   "frmMain.frx":5706A
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   59
         Top             =   5760
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
         Picture         =   "frmMain.frx":573EF
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   58
         Top             =   5760
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
         Picture         =   "frmMain.frx":57778
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   57
         Top             =   5760
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
         Picture         =   "frmMain.frx":57B02
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   56
         Top             =   5760
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
         Picture         =   "frmMain.frx":57E8E
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   55
         Top             =   5760
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
         Picture         =   "frmMain.frx":58215
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   54
         Top             =   6240
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
         Picture         =   "frmMain.frx":582D2
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   53
         Top             =   6240
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
         Picture         =   "frmMain.frx":58396
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   52
         Top             =   6240
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
         Picture         =   "frmMain.frx":5846A
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   51
         Top             =   6240
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
         Picture         =   "frmMain.frx":58521
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   50
         Top             =   6240
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
         Picture         =   "frmMain.frx":585D0
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   49
         Top             =   6240
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
         Picture         =   "frmMain.frx":58670
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   48
         Top             =   6240
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
         Picture         =   "frmMain.frx":58703
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   47
         Top             =   6240
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
         Picture         =   "frmMain.frx":5879B
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   46
         Top             =   6240
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
         Picture         =   "frmMain.frx":5882E
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   45
         Top             =   6240
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
         Picture         =   "frmMain.frx":588CB
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   44
         Top             =   6240
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
         Picture         =   "frmMain.frx":58975
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   43
         Top             =   6240
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
         Picture         =   "frmMain.frx":58A2B
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   42
         Top             =   6240
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
         Picture         =   "frmMain.frx":58AFF
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   41
         Top             =   6240
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
         Picture         =   "frmMain.frx":58BE1
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   40
         Top             =   5760
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
         Picture         =   "frmMain.frx":58F34
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   39
         Top             =   5760
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
         Picture         =   "frmMain.frx":5928D
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   37
         Top             =   5280
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
         Picture         =   "frmMain.frx":592E3
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   36
         Top             =   5280
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
         Picture         =   "frmMain.frx":59328
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   35
         Top             =   5280
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
         Picture         =   "frmMain.frx":59378
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   32
         Top             =   5760
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
         Picture         =   "frmMain.frx":596CB
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   31
         Top             =   5760
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
         Picture         =   "frmMain.frx":59A25
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   30
         Top             =   5760
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
         Picture         =   "frmMain.frx":59D85
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   28
         Top             =   5280
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
         Picture         =   "frmMain.frx":59DD5
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   27
         Top             =   5280
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
         Picture         =   "frmMain.frx":59E1A
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   26
         Top             =   5280
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
         Picture         =   "frmMain.frx":59EFC
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   25
         Top             =   5280
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
         Picture         =   "frmMain.frx":59FD0
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   24
         Top             =   5280
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
         Picture         =   "frmMain.frx":5A086
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   23
         Top             =   5280
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
         Picture         =   "frmMain.frx":5A130
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   22
         Top             =   5280
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
         Picture         =   "frmMain.frx":5A1CD
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   21
         Top             =   5280
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
         Picture         =   "frmMain.frx":5A260
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   20
         Top             =   5280
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
         Picture         =   "frmMain.frx":5A2F8
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   19
         Top             =   5280
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
         Picture         =   "frmMain.frx":5A38B
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   18
         Top             =   5280
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
         Picture         =   "frmMain.frx":5A42B
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   17
         Top             =   5280
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
         Picture         =   "frmMain.frx":5A4DA
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   16
         Top             =   5280
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
         Picture         =   "frmMain.frx":5A591
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   15
         Top             =   5280
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
         Picture         =   "frmMain.frx":5A665
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   14
         Top             =   5280
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
         Picture         =   "frmMain.frx":5A729
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   13
         Top             =   5280
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
         Left            =   9720
         Picture         =   "frmMain.frx":5A7E6
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   12
         Top             =   5760
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
         Left            =   9600
         Picture         =   "frmMain.frx":5A948
         ScaleHeight     =   100
         ScaleMode       =   0  'User
         ScaleWidth      =   104.166
         TabIndex        =   11
         Top             =   5760
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
         Left            =   9480
         Picture         =   "frmMain.frx":5AAA7
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   10
         Top             =   5760
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
         Left            =   9360
         Picture         =   "frmMain.frx":5AC12
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   9
         Top             =   5760
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
         Left            =   9240
         Picture         =   "frmMain.frx":5ADA0
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   8
         Top             =   5760
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
         Left            =   9240
         Picture         =   "frmMain.frx":5AF14
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   7
         Top             =   5040
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
         Left            =   9360
         Picture         =   "frmMain.frx":5B056
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   6
         Top             =   5040
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
         Left            =   9480
         Picture         =   "frmMain.frx":5B199
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   5
         Top             =   5040
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
         Left            =   9600
         Picture         =   "frmMain.frx":5B2DA
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   4
         Top             =   5040
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
         Left            =   9720
         Picture         =   "frmMain.frx":5B419
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   3
         Top             =   5040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":5B556
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   2
         Top             =   0
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
         Left            =   3480
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   1
         Top             =   120
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   5
         Left            =   7320
         Picture         =   "frmMain.frx":5C581
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   88
         Top             =   5040
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
         Left            =   7440
         Picture         =   "frmMain.frx":5C6C3
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   89
         Top             =   5040
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
         Left            =   7560
         Picture         =   "frmMain.frx":5C806
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   90
         Top             =   5040
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
         Left            =   7680
         Picture         =   "frmMain.frx":5C947
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   91
         Top             =   5040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picSelMask 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   9
         Left            =   7800
         Picture         =   "frmMain.frx":5CA86
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   92
         Top             =   5040
         Visible         =   0   'False
         Width           =   1500
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
         Picture         =   "frmMain.frx":5CBC3
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   29
         Top             =   5760
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
         Picture         =   "frmMain.frx":5CF24
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   33
         Top             =   5280
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
         Left            =   5760
         Picture         =   "frmMain.frx":5CF7A
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   34
         Top             =   5280
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picTileLR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   2
         Left            =   3480
         Picture         =   "frmMain.frx":5CFCF
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   177
         Top             =   360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picTileLR 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   6600
         Picture         =   "frmMain.frx":5E1FB
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   185
         Top             =   360
         Visible         =   0   'False
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2677 lines as of June 6
Dim formW As Integer
Dim formH As Integer
Dim frmX As Integer
Dim frmY As Integer

Private Sub Form_GotFocus()
Dim curMouse As POINTAPI
' Get the current mouse cursor coordinates:
Call GetCursorPos(curMouse)
curMouse.x = curMouse.x * 15
curMouse.y = curMouse.y * 15
If curMouse.x < frmMain.Left + 50 Or curMouse.x > frmMain.Left + frmMain.Width - 50 Or curMouse.y < frmMain.Top + 370 Or curMouse.y > frmMain.Top + frmMain.Height - 50 Then
    Call keyUp(27, False)
    tmrRefresh.Enabled = True
    If frmX <> frmMain.Left Or frmY <> frmMain.Top Then
        frmX = frmMain.Left
        frmY = frmMain.Top
        frmPause.Left = frmX + 50
        frmPause.Top = frmY + 370
        frmGUI.Left = frmX + 50
        frmGUI.Top = frmY + 370
        tmrRefresh.Enabled = False
    End If
End If
End Sub

Private Sub tmrRefresh_Timer()
frmPause.Left = frmMain.Left + 50
frmPause.Top = frmMain.Top + 370
frmGUI.Left = frmMain.Left + 50
frmGUI.Top = frmMain.Top + 370
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call keyUp(KeyCode, Shift)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call keyHandler(KeyCode, Shift)
End Sub

Private Sub tmrAlternate_Timer()
Static blnAlt As Boolean
If blnAlt Then
    blnAlt = False
Else
    blnAlt = True
End If
tmrAlternate.Tag = blnAlt
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
    If blnPlayerMoveable(Index) Then
        Call cpuAI(Index)
    End If
    intCounter = 0
End If
End Sub

Private Sub Form_Resize()
frmMain.Width = formW
frmMain.Height = formH
End Sub

Private Sub Form_Load()
frmGUI.Show
frmGUI.SetFocus
frmGUI.BackColor = vbCyan
formW = frmMain.Width
formH = frmMain.Height
frmGUI.Top = frmMain.Top + 370
frmGUI.Left = frmMain.Left + 50
PaintPicture picBackground.Image, 0, 0, 800, 637, 0, 0, 800, 637, vbSrcCopy
Set picBG = picBackground
Call DrawMap(7, 7)
blnClearPrevTile(0) = False
blnClearPrevTile(1) = False
blnClearPrevTile(2) = False
blnClearPrevTile(3) = False
blnPlayerMoveable(0) = False
blnPlayerMoveable(1) = False
blnPlayerMoveable(2) = False
blnPlayerMoveable(3) = False
limswitch = 1000
End Sub

Private Sub tmrObj_Timer()
Dim frameCount As Integer
Dim frameLim As Integer
Dim intObjTimer As Long
Dim curTile As terrain
frmDbg.lstCoin.Clear
For o = 0 To tileCount - 1
    curTile = tile(getTileFromInt(True, o), getTileFromInt(False, o))
    'frmDbg.lstCoin.AddItem (o & ": " & Tile(getTileFromInt(True, o), getTileFromInt(False, o)).hasObj & ", " & Tile(getTileFromInt(True, c), getTileFromInt(False, c)).objTimer)
    If curTile.hasObj Then
        If curTile.objType(0) = "Coin" Then
            frameLim = 28 'frame limit for coin
        ElseIf curTile.objType(0) = "Pow" Then
            If curTile.objType(1) = "Scare" Then
                frameLim = 10 'frame limit for scare power-up
            ElseIf curTile.objType(1) = "Speed" Then
                frameLim = 10 'frame limit for speed power-up
            ElseIf curTile.objType(1) = "Freeze" Then
                frameLim = 20 'frame limit for freeze power-up
            End If
        ElseIf curTile.objType(0) = "Egg" Then
            If curTile.objType(1) = "M" Then
                frameLim = 7
            ElseIf curTile.objType(1) = "G" Then
                frameLim = 21
            End If
        End If
        intObjTimer = curTile.objTimer
        If curTile.objType(0) <> "Terrain" Then 'if obj is not a terrain, set frame count of object and paint object
            If gameMode = 0 Then
                frameCount = intObjTimer - ((intObjTimer \ frameLim) * frameLim)
                tile(getTileFromInt(True, o), getTileFromInt(False, o)).objFrame = frameCount
                Call PaintObj(curTile.objType(0), curTile.objType(1), frameCount, curTile.Xc, curTile.Yc, False)
            Else
                If curTile.objFrame >= frameLim - 1 Then
                    frameCount = 0
                Else
                    frameCount = curTile.objFrame + 1
                End If
                Call PaintObj(curTile.objType(0), curTile.objType(1), curTile.objFrame, curTile.Xc, curTile.Yc, False)
                tile(getTileFromInt(True, o), getTileFromInt(False, o)).objFrame = frameCount
            End If
        Else
            Call clearTile(curTile, False, 0, "ObjXY")
        End If
        'if objTimer (frame advancements on obj) is under objExpire, add 1 to it
        If intObjTimer < objExpire Then
            If gameMode = 0 Then
                tile(getTileFromInt(True, o), getTileFromInt(False, o)).objTimer = intObjTimer + 1
            End If
        Else 'if objTimer is objExpire (or greater), disable obj and clear tile
            Call killObj(tile(getTileFromInt(True, o), getTileFromInt(False, o)))
        End If
    End If
Next o
End Sub

Private Sub tmrChar_Timer(Index As Integer)
'reverse boolean for select
Static blnRev(0 To 3) As Boolean
'call selection paint
Call PaintSelector(Index, picCount(Index))
If frameCounter(Index) = 0 Then
    Call PaintCharSprite(Index, spriteX(Index), spriteY(Index))
End If
If frameCounter(Index) > 0 Then 'if jump timer is started
    Call charAction(Index, tile(nextX(Index), nextY(Index)))
Else
    spriteX(Index) = tile(curX(Index), curY(Index)).x + 25
    spriteY(Index) = tile(curX(Index), curY(Index)).y - 15
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

Private Sub tmrStun_Timer(Index As Integer)
If isPlayer(Index) Then
    blnPlayerMoveable(Index) = True
Else
    tmrCPUMove(Index).Enabled = True
End If
tmrStun(Index).Enabled = False
End Sub

Private Sub tmrTileAnim_Timer()
Static intCounter As Integer
For x = 0 To tileCount - 1
    If tileSwitch(x) = True And intCounter - (1 * ((tileCount - 1) - x)) <= 8 Then
        Call getTileAnim(intCounter - (1 * ((tileCount - 1) - x)), tile(getTileFromInt(True, x), getTileFromInt(False, x)))
    End If
Next x
intCounter = intCounter + 1
End Sub

Private Sub tmrTileAnimDelay_Timer()
Static intCounter As Integer
Static intX As Integer
Static intY As Integer
frmMain.tmrTileAnim.Enabled = True
For z = 0 To intCounter - getAbs(intCounter - ((mapWidth * mapHeight) + (Int(mapHeight / 2))) - 1)
    tileSwitch((tileCount - 1) - z) = True
Next z
If intX = mapWidth - ((intY + 1) Mod 2) Then
    intY = intY + 1
    intX = 0
ElseIf intX < mapWidth + ((mapHeight + 1) Mod 2) Then
    intX = intX + 1
End If
If intCounter >= ((mapWidth * mapHeight) + (Int(mapHeight / 2))) - 1 Then
    intCounter = 0
    intX = 0
    intY = 0
    tmrTileAnimDelay.Enabled = False
End If
intCounter = intCounter + 1
End Sub

Public Sub tmrPow_Timer(Index As Integer)
Call getPowTick(Index)
End Sub

Public Sub tmrObjEvent_Timer()
Dim intRand As Integer
If objTileCount < tileCount - 4 Then
    intRand = randInt(1, 100)
    Dim intType As Integer
    intType = randInt(1, 100)
    If intRand < 85 Then 'Coin: 85%
        intRand = randInt(0, tileCount - 1)
        Do Until Not tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).hasObj And Not tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).hasChar
            intRand = randInt(0, tileCount - 1)
        Loop
        tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).hasObj = True
        tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).objType(0) = "Coin"
        If intType <= 65 Then 'Yellow: 55.25%
            tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).objType(1) = "Y"
        ElseIf intType > 65 And intType <= 90 Then 'Red: 21.25%
            tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).objType(1) = "R"
        ElseIf intType > 90 Then 'Blue: 8.5%
            tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).objType(1) = "B"
        End If
    ElseIf intRand >= 85 And intRand < 95 Then 'Power-up: 15%
        intRand = randInt(0, tileCount - 1)
        Do Until Not tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).hasObj And Not tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).hasChar
            intRand = randInt(0, tileCount - 1)
        Loop
        tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).hasObj = True
        tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).objType(0) = "Pow"
        If intType <= 60 Then 'Speed: 9%
            tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).objType(1) = "Speed"
        ElseIf intType > 60 And intType <= 90 Then 'Scare: 4.5%
            tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).objType(1) = "Scare"
        ElseIf intType > 90 Then 'Freeze: 1.5%
            tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).objType(1) = "Freeze"
        End If
    ElseIf intRand >= 95 Then 'Egg: 5%
        intRand = randInt(0, tileCount - 1)
        Do Until Not tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).hasObj And Not tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).hasChar
            intRand = randInt(0, tileCount - 1)
        Loop
        tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).hasObj = True
        tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).objType(0) = "Egg"
        If intType > 10 Then 'Multi Egg: 4.5%
            tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).objType(1) = "M"
        ElseIf intType <= 10 Then 'Gold Egg (1up): 0.5%
            tile(getTileFromInt(True, intRand), getTileFromInt(False, intRand)).objType(1) = "G"
        End If
    End If
    objTileCount = objTileCount + 1
End If
End Sub

