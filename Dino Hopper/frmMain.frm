VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Dino Hopper"
   ClientHeight    =   9000
   ClientLeft      =   4260
   ClientTop       =   570
   ClientWidth     =   12000
   LinkTopic       =   "Dino Hopper"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   209
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
      TabIndex        =   208
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
      TabIndex        =   207
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
      TabIndex        =   206
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
   Begin VB.PictureBox picSpacerMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1800
      Picture         =   "frmMain.frx":491B
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   135
      Top             =   2040
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Timer tmrHurt 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   11520
      Top             =   3000
   End
   Begin VB.Timer tmrHurt 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   500
      Left            =   11520
      Top             =   4440
   End
   Begin VB.Timer tmrHurt 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   500
      Left            =   11520
      Top             =   3960
   End
   Begin VB.Timer tmrHurt 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   500
      Left            =   11520
      Top             =   3480
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
      Interval        =   750
      Left            =   11520
      Top             =   2040
   End
   Begin VB.Timer tmrCPUMove 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   750
      Left            =   11520
      Top             =   1560
   End
   Begin VB.Timer tmrCPUMove 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   750
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
      Picture         =   "frmMain.frx":49AF
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   39
      Top             =   6360
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer tmrObj 
      Enabled         =   0   'False
      Interval        =   40
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
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   0
      Picture         =   "frmMain.frx":4D12
      ScaleHeight     =   601
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
      Begin VB.PictureBox picP4ID 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   1800
         Picture         =   "frmMain.frx":2939A
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   269
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
         Picture         =   "frmMain.frx":29893
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   268
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
         Picture         =   "frmMain.frx":29D75
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   267
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
         Picture         =   "frmMain.frx":2A260
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   266
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
         Picture         =   "frmMain.frx":2A737
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   265
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
         Picture         =   "frmMain.frx":2AC28
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   264
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
         Picture         =   "frmMain.frx":2B114
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   263
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
         Picture         =   "frmMain.frx":2B5E5
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   262
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
         Picture         =   "frmMain.frx":2BACD
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   261
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
         Picture         =   "frmMain.frx":2BF96
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   260
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
         Picture         =   "frmMain.frx":2C472
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   259
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
         Picture         =   "frmMain.frx":2C918
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   258
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
         Picture         =   "frmMain.frx":2CDB5
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   257
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
         Picture         =   "frmMain.frx":2D243
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   256
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
         Picture         =   "frmMain.frx":2D6D9
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   255
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
         Picture         =   "frmMain.frx":2DB74
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   254
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
         Picture         =   "frmMain.frx":2E009
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   253
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
         Picture         =   "frmMain.frx":2E4A4
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   252
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
         Picture         =   "frmMain.frx":2E930
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   251
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
         Picture         =   "frmMain.frx":2EDB2
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   250
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
         Picture         =   "frmMain.frx":2F236
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   249
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
         Picture         =   "frmMain.frx":2F72F
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   248
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
         Picture         =   "frmMain.frx":2FC11
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   247
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
         Picture         =   "frmMain.frx":300FC
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   246
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
         Picture         =   "frmMain.frx":305D3
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   245
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
         Picture         =   "frmMain.frx":30AC4
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   244
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
         Picture         =   "frmMain.frx":30FB0
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   243
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
         Picture         =   "frmMain.frx":31481
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   242
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
         Picture         =   "frmMain.frx":31969
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   241
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
         Picture         =   "frmMain.frx":31E32
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   240
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
         Picture         =   "frmMain.frx":3230E
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   239
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
         Picture         =   "frmMain.frx":327B4
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   238
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
         Picture         =   "frmMain.frx":32C51
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   237
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
         Picture         =   "frmMain.frx":330DF
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   236
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
         Picture         =   "frmMain.frx":33575
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   235
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
         Picture         =   "frmMain.frx":33A10
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   234
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
         Picture         =   "frmMain.frx":33EA5
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   233
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
         Picture         =   "frmMain.frx":34340
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   232
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
         Picture         =   "frmMain.frx":347CC
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   231
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
         Picture         =   "frmMain.frx":34C4E
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   230
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
         Picture         =   "frmMain.frx":350D2
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   229
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
         Picture         =   "frmMain.frx":355CB
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   228
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
         Picture         =   "frmMain.frx":35AAD
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   227
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
         Picture         =   "frmMain.frx":35F98
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   226
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
         Picture         =   "frmMain.frx":3646F
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   225
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
         Picture         =   "frmMain.frx":36960
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   224
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
         Picture         =   "frmMain.frx":36E4C
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   223
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
         Picture         =   "frmMain.frx":3731D
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   222
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
         Picture         =   "frmMain.frx":37805
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   221
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
         Picture         =   "frmMain.frx":37CCE
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   220
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
         Picture         =   "frmMain.frx":381AA
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   219
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
         Picture         =   "frmMain.frx":38650
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   218
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
         Picture         =   "frmMain.frx":38AED
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   217
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
         Picture         =   "frmMain.frx":38F7B
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   216
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
         Picture         =   "frmMain.frx":39411
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   215
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
         Picture         =   "frmMain.frx":398AC
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   214
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
         Picture         =   "frmMain.frx":39D41
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   213
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
         Picture         =   "frmMain.frx":3A1C3
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   212
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
         Picture         =   "frmMain.frx":3A64F
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   211
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
         Picture         =   "frmMain.frx":3AAD1
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   210
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
         Picture         =   "frmMain.frx":3AF55
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   120
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
         Picture         =   "frmMain.frx":3B11C
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
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":3B2DC
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   118
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
         Picture         =   "frmMain.frx":3B496
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   117
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
         Picture         =   "frmMain.frx":3B65D
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   115
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
         Picture         =   "frmMain.frx":3B82D
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   100
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
         Picture         =   "frmMain.frx":3BA41
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
         Index           =   2
         Left            =   3240
         Picture         =   "frmMain.frx":3BC3A
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   98
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
         Picture         =   "frmMain.frx":3BE44
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   97
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
         Picture         =   "frmMain.frx":3C04A
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   95
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
         Picture         =   "frmMain.frx":3C25C
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   205
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
         Picture         =   "frmMain.frx":3C339
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   204
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
         Picture         =   "frmMain.frx":3C473
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   203
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
         Picture         =   "frmMain.frx":3C5AD
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   202
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
         Picture         =   "frmMain.frx":3C6E7
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   201
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
         Picture         =   "frmMain.frx":3C80F
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   200
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
         Picture         =   "frmMain.frx":3C94A
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   199
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
         Picture         =   "frmMain.frx":3CA7F
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   198
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
         Picture         =   "frmMain.frx":3CBB5
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   196
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
         Picture         =   "frmMain.frx":3CC9A
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   195
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
         Picture         =   "frmMain.frx":3CD76
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   194
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
         Picture         =   "frmMain.frx":3CE54
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   193
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
         Picture         =   "frmMain.frx":3DDEA
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   191
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
         Picture         =   "frmMain.frx":3EDEC
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   190
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
         Picture         =   "frmMain.frx":3FDA8
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   189
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
         Picture         =   "frmMain.frx":40D98
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   188
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
         Picture         =   "frmMain.frx":41FCA
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   187
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
         Picture         =   "frmMain.frx":431FC
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   186
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
         Picture         =   "frmMain.frx":44442
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   185
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
         Picture         =   "frmMain.frx":45695
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   183
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
         Picture         =   "frmMain.frx":468F3
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   182
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
         Picture         =   "frmMain.frx":491F7
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   181
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
         Picture         =   "frmMain.frx":4BB08
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   180
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
         Picture         =   "frmMain.frx":4BC21
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   179
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
         Picture         =   "frmMain.frx":4BD37
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   178
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
         Picture         =   "frmMain.frx":4BE51
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   177
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
         Picture         =   "frmMain.frx":4BF66
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   175
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
         Picture         =   "frmMain.frx":4CF14
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   176
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
         Picture         =   "frmMain.frx":4DED0
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   174
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
         Index           =   8
         Left            =   9120
         Picture         =   "frmMain.frx":4DF53
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   173
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
         Picture         =   "frmMain.frx":4DFD8
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   172
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
         Picture         =   "frmMain.frx":4E05D
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   171
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
         Picture         =   "frmMain.frx":4E0E5
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   170
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
         Picture         =   "frmMain.frx":4E16D
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   169
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
         Picture         =   "frmMain.frx":4E1DA
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   168
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
         Picture         =   "frmMain.frx":4E249
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   167
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
         Picture         =   "frmMain.frx":4E2BA
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   166
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
         Picture         =   "frmMain.frx":4E32D
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   165
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
         Picture         =   "frmMain.frx":4E3A0
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   164
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
         Picture         =   "frmMain.frx":4E428
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   163
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
         Picture         =   "frmMain.frx":4E4AF
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   162
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
         Picture         =   "frmMain.frx":4E534
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   161
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
         Picture         =   "frmMain.frx":4E5B7
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   160
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
         Picture         =   "frmMain.frx":4E638
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   159
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
         Picture         =   "frmMain.frx":4E6AB
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   158
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
         Picture         =   "frmMain.frx":4E71C
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   157
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
         Picture         =   "frmMain.frx":4E78C
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   156
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
         Picture         =   "frmMain.frx":4E7FC
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   155
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
         Picture         =   "frmMain.frx":4E869
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   153
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
         Picture         =   "frmMain.frx":4E95D
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   152
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
         Picture         =   "frmMain.frx":4EA54
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   151
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
         Picture         =   "frmMain.frx":4EB4C
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   150
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
         Picture         =   "frmMain.frx":4EBB2
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   149
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
         Picture         =   "frmMain.frx":4EC1D
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   148
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
         Picture         =   "frmMain.frx":4EC8A
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   145
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
         Picture         =   "frmMain.frx":4ECF7
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   144
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
         Picture         =   "frmMain.frx":4ED68
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   143
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
         Picture         =   "frmMain.frx":4EDDC
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   142
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
         Picture         =   "frmMain.frx":4EE4F
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   141
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
         Picture         =   "frmMain.frx":4EEBA
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   140
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
         Picture         =   "frmMain.frx":4EFB6
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   139
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
         Picture         =   "frmMain.frx":4F0C3
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   138
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
         Picture         =   "frmMain.frx":4F1DC
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   137
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
         Picture         =   "frmMain.frx":4F2F2
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   136
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
         Picture         =   "frmMain.frx":4F407
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   133
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
         Picture         =   "frmMain.frx":4F4BA
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   132
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
         Picture         =   "frmMain.frx":4F632
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
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":4F7E0
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   130
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
         Picture         =   "frmMain.frx":4F962
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   129
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
         Picture         =   "frmMain.frx":4FA0E
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
         Index           =   1
         Left            =   6240
         Picture         =   "frmMain.frx":4FAC1
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   127
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
         Picture         =   "frmMain.frx":4FB70
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   126
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
         Picture         =   "frmMain.frx":4FD32
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   125
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
         Picture         =   "frmMain.frx":4FDE8
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   124
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
         Picture         =   "frmMain.frx":4FF73
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   123
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
         Picture         =   "frmMain.frx":50023
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
         Index           =   2
         Left            =   6960
         Picture         =   "frmMain.frx":500CF
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   121
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
         Picture         =   "frmMain.frx":5017B
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   116
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
         Picture         =   "frmMain.frx":5022E
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   114
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
         Picture         =   "frmMain.frx":502E1
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   113
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
         Picture         =   "frmMain.frx":503A4
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   112
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
         Picture         =   "frmMain.frx":505A2
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
         Index           =   1
         Left            =   2520
         Picture         =   "frmMain.frx":5078E
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   110
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
         Picture         =   "frmMain.frx":50995
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   109
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
         Picture         =   "frmMain.frx":50A53
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
         Index           =   1
         Left            =   2520
         Picture         =   "frmMain.frx":50B0E
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   107
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
         Picture         =   "frmMain.frx":50BD0
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   106
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
         Picture         =   "frmMain.frx":50DC8
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   105
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
         Picture         =   "frmMain.frx":50E88
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   104
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
         Picture         =   "frmMain.frx":51096
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   103
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
         Picture         =   "frmMain.frx":5115A
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
         Index           =   2
         Left            =   3240
         Picture         =   "frmMain.frx":51217
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   101
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
         Picture         =   "frmMain.frx":512D6
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   96
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
         Picture         =   "frmMain.frx":51398
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   94
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
         Picture         =   "frmMain.frx":5145D
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   88
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
         Picture         =   "frmMain.frx":515D1
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
         Index           =   8
         Left            =   7560
         Picture         =   "frmMain.frx":51733
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
         Index           =   7
         Left            =   7440
         Picture         =   "frmMain.frx":518C1
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
         Index           =   6
         Left            =   7320
         Picture         =   "frmMain.frx":51A2C
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   84
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
         Picture         =   "frmMain.frx":51B8B
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
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":51C48
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
         Index           =   2
         Left            =   960
         Picture         =   "frmMain.frx":51D0C
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
         Index           =   3
         Left            =   1320
         Picture         =   "frmMain.frx":51DE0
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
         Index           =   4
         Left            =   1680
         Picture         =   "frmMain.frx":51E97
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
         Index           =   5
         Left            =   2040
         Picture         =   "frmMain.frx":51F46
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
         Index           =   6
         Left            =   2400
         Picture         =   "frmMain.frx":51FE6
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
         Index           =   7
         Left            =   2760
         Picture         =   "frmMain.frx":52079
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
         Index           =   8
         Left            =   3120
         Picture         =   "frmMain.frx":52111
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
         Index           =   9
         Left            =   3480
         Picture         =   "frmMain.frx":521A4
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
         Index           =   10
         Left            =   3840
         Picture         =   "frmMain.frx":52241
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
         Index           =   11
         Left            =   4200
         Picture         =   "frmMain.frx":522EB
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
         Index           =   12
         Left            =   4560
         Picture         =   "frmMain.frx":523A1
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
         Index           =   13
         Left            =   4920
         Picture         =   "frmMain.frx":52475
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   70
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
         Picture         =   "frmMain.frx":52557
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   69
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
         Picture         =   "frmMain.frx":528D8
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
         Index           =   2
         Left            =   960
         Picture         =   "frmMain.frx":52C60
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
         Index           =   3
         Left            =   1320
         Picture         =   "frmMain.frx":52FEA
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
         Index           =   4
         Left            =   1680
         Picture         =   "frmMain.frx":53371
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
         Index           =   5
         Left            =   2040
         Picture         =   "frmMain.frx":536FB
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
         Index           =   6
         Left            =   2400
         Picture         =   "frmMain.frx":53A81
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
         Index           =   7
         Left            =   2760
         Picture         =   "frmMain.frx":53E08
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
         Index           =   8
         Left            =   3120
         Picture         =   "frmMain.frx":5418E
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
         Index           =   9
         Left            =   3480
         Picture         =   "frmMain.frx":54513
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
         Index           =   10
         Left            =   3840
         Picture         =   "frmMain.frx":54898
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
         Index           =   11
         Left            =   4200
         Picture         =   "frmMain.frx":54C21
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
         Index           =   12
         Left            =   4560
         Picture         =   "frmMain.frx":54FAB
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
         Index           =   13
         Left            =   4920
         Picture         =   "frmMain.frx":55337
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   56
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
         Picture         =   "frmMain.frx":556BE
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   55
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
         Picture         =   "frmMain.frx":5577B
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
         Index           =   2
         Left            =   960
         Picture         =   "frmMain.frx":5583F
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
         Index           =   3
         Left            =   1320
         Picture         =   "frmMain.frx":55913
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
         Index           =   4
         Left            =   1680
         Picture         =   "frmMain.frx":559CA
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
         Index           =   5
         Left            =   2040
         Picture         =   "frmMain.frx":55A79
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
         Index           =   6
         Left            =   2400
         Picture         =   "frmMain.frx":55B19
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
         Index           =   7
         Left            =   2760
         Picture         =   "frmMain.frx":55BAC
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
         Index           =   8
         Left            =   3120
         Picture         =   "frmMain.frx":55C44
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
         Index           =   9
         Left            =   3480
         Picture         =   "frmMain.frx":55CD7
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
         Index           =   10
         Left            =   3840
         Picture         =   "frmMain.frx":55D74
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
         Index           =   11
         Left            =   4200
         Picture         =   "frmMain.frx":55E1E
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
         Index           =   12
         Left            =   4560
         Picture         =   "frmMain.frx":55ED4
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
         Index           =   13
         Left            =   4920
         Picture         =   "frmMain.frx":55FA8
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   42
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
         Picture         =   "frmMain.frx":5608A
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   41
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
         Picture         =   "frmMain.frx":563DD
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   40
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
         Picture         =   "frmMain.frx":56736
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   38
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
         Picture         =   "frmMain.frx":5678C
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
         Index           =   5
         Left            =   6480
         Picture         =   "frmMain.frx":567D1
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   36
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
         Picture         =   "frmMain.frx":56821
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   33
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
         Picture         =   "frmMain.frx":56B74
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
         Index           =   3
         Left            =   6000
         Picture         =   "frmMain.frx":56ECE
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   31
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
         Picture         =   "frmMain.frx":5722E
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   29
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
         Picture         =   "frmMain.frx":5727E
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   28
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
         Picture         =   "frmMain.frx":572C3
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
         Index           =   12
         Left            =   4560
         Picture         =   "frmMain.frx":573A5
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
         Index           =   11
         Left            =   4200
         Picture         =   "frmMain.frx":57479
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
         Index           =   10
         Left            =   3840
         Picture         =   "frmMain.frx":5752F
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
         Index           =   9
         Left            =   3480
         Picture         =   "frmMain.frx":575D9
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
         Index           =   8
         Left            =   3120
         Picture         =   "frmMain.frx":57676
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
         Index           =   7
         Left            =   2760
         Picture         =   "frmMain.frx":57709
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
         Index           =   6
         Left            =   2400
         Picture         =   "frmMain.frx":577A1
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
         Index           =   5
         Left            =   2040
         Picture         =   "frmMain.frx":57834
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
         Index           =   4
         Left            =   1680
         Picture         =   "frmMain.frx":578D4
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
         Index           =   3
         Left            =   1320
         Picture         =   "frmMain.frx":57983
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
         Index           =   2
         Left            =   960
         Picture         =   "frmMain.frx":57A3A
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
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":57B0E
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
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":57BD2
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   14
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
         Picture         =   "frmMain.frx":57C8F
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
         Picture         =   "frmMain.frx":57DF1
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
         Picture         =   "frmMain.frx":57F50
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
         Picture         =   "frmMain.frx":580BB
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
         Picture         =   "frmMain.frx":58249
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
         Picture         =   "frmMain.frx":583BD
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
         Picture         =   "frmMain.frx":584FF
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
         Picture         =   "frmMain.frx":58642
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
         Picture         =   "frmMain.frx":58783
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
         Picture         =   "frmMain.frx":588C2
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
         Picture         =   "frmMain.frx":589FF
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
         Picture         =   "frmMain.frx":59A2A
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
         Index           =   6
         Left            =   7440
         Picture         =   "frmMain.frx":59B6C
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
         Index           =   7
         Left            =   7560
         Picture         =   "frmMain.frx":59CAF
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
         Index           =   8
         Left            =   7680
         Picture         =   "frmMain.frx":59DF0
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   92
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
         Picture         =   "frmMain.frx":59F2F
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   93
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
         Picture         =   "frmMain.frx":5A06C
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
         Index           =   2
         Left            =   6000
         Picture         =   "frmMain.frx":5A3CD
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   34
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
         Picture         =   "frmMain.frx":5A423
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   35
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
         Picture         =   "frmMain.frx":5A478
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   184
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
         Picture         =   "frmMain.frx":5B6A4
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   192
         Top             =   360
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.PictureBox picSpacer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   240
      Picture         =   "frmMain.frx":5C651
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   134
      Top             =   1920
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblLives 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x3"
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
      Left            =   11640
      TabIndex        =   197
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblMulti 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multiplier: 1x"
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
      Left            =   150
      TabIndex        =   154
      Top             =   150
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lblTest2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   147
      Top             =   360
      Width           =   555
   End
   Begin VB.Label lblTest 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   146
      Top             =   0
      Width           =   555
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
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
      Left            =   9240
      TabIndex        =   13
      Top             =   150
      Visible         =   0   'False
      Width           =   1950
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
ElseIf KeyCode = 122 Then
        If frmSettings.Visible = False Then
        frmSettings.Visible = True
    Else
        frmSettings.Visible = False
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If blnPlayerMoveable = True Then
    If KeyCode = key(0) Then 'Left
        If evalMove(0, "L") Then
            If gameMode <> 1 Or (gameMode = 1 And strDir(0) = "L") Then
                Call getJump(0, "L")
            ElseIf gameMode = 1 Then
                strDir(0) = "L"
            End If
        End If
    ElseIf KeyCode = key(1) Then 'Up
        If evalMove(0, "U") Then
            If gameMode <> 1 Or (gameMode = 1 And strDir(0) = "U") Then
                Call getJump(0, "U")
            ElseIf gameMode = 1 Then
                strDir(0) = "U"
            End If
        End If
    ElseIf KeyCode = key(2) Then 'Right
        If evalMove(0, "R") Then
            If gameMode <> 1 Or (gameMode = 1 And strDir(0) = "R") Then
                Call getJump(0, "R")
            ElseIf gameMode = 1 Then
                strDir(0) = "R"
            End If
        End If
    ElseIf KeyCode = key(3) Then 'Down
        If evalMove(0, "D") Then
            If gameMode <> 1 Or (gameMode = 1 And strDir(0) = "D") Then
                Call getJump(0, "D")
            ElseIf gameMode = 1 Then
                strDir(0) = "D"
            End If
        End If
    ElseIf KeyCode = key(4) Then 'Action
        If evalMove(0, strDir(0)) Then
            Call getJump(0, strDir(0))
        End If
    End If
End If
End Sub

Private Sub tmrHurt_Timer(index As Integer)
Call getHurt(index)
curX(index) = prevX(index)
curY(index) = prevY(index)
blnPlayerMoveable = True
tmrHurt(index).Enabled = False
End Sub

Private Sub tmrCPUMove_Timer(index As Integer)
Static intCounter As Integer
If cpuInterval > 50 Then
    cpuInterval = cpuInterval - 5
    For x = 1 To 3
        If tmrCPUMove(x).Enabled Then
            tmrCPUMove(x).Interval = cpuInterval
        End If
    Next x
End If
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

Private Sub Form_Load()
cpuInterval = 750
Call DrawMap(1)
blnClearPrevTile(0) = False
blnClearPrevTile(1) = False
blnPlayerMoveable = False
For t = 0 To 100
    tileSwitch(t) = False
Next t
limswitch = 1000
End Sub

Private Sub Form_Resize()
Call DrawMap(1)
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
            frameCount = intObjTimer - ((intObjTimer \ frameLim) * frameLim)
            Call PaintObj(curTile.objType(0), curTile.objType(1), frameCount, curTile.Xc, curTile.Yc, False)
        Else
            Call clearTile(curTile, False, 0, "ObjXY")
        End If
        'if objTimer (frame advancements on obj) is under objExpire, add 1 to it
        If intObjTimer < objExpire Then
            tile(getTileFromInt(True, o), getTileFromInt(False, o)).objTimer = intObjTimer + 1
        Else 'if objTimer is objExpire (or greater), disable obj and clear tile
            Call killObj(tile(getTileFromInt(True, o), getTileFromInt(False, o)))
        End If
    End If
Next o
End Sub

Private Sub tmrChar_Timer(index As Integer)
'reverse boolean for select
Static blnRev(0 To 3) As Boolean
'call selection paint
Call PaintSelector(index, picCount(index))
'if jump timer is started
If frameCounter(index) > 0 Then
    'show each jump frame for p1
    'If Index = 0 Then
        'MsgBox (intFrame(0))
    'End If
    Call getCharJumpAnim(index, frameCounter(index), tile(curX(index), curY(index)), tile(nextX(index), nextY(index)))
    If frameCounter(index) = 1 Then
        strState(index) = "C"
    ElseIf frameCounter(index) = 5 Then
        strState(index) = "J"
    ElseIf frameCounter(index) = 10 Then
        strState(index) = "I"
    End If
    If frameCounter(index) = frameLimit(index) Then
        frameCounter(index) = 0
        blnPlayerMoveable = True
        Call getJumpComplete(index)
    Else
        frameCounter(index) = frameCounter(index) + 1
    End If
Else
    spriteX(index) = tile(curX(index), curY(index)).x + 25
    spriteY(index) = tile(curX(index), curY(index)).y - 15
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
    tmrTileAnimDelay.Enabled = False
End If
intCounter = intCounter + 1
End Sub

Public Sub tmrPow_Timer(index As Integer)
tmrPowCounter = tmrPowCounter + 1
If tmrPowCounter >= 20 Then
    Call getPowExpire(index, tmrPow(index).Tag)
    tmrPowCounter = 0
    tmrPow(index).Tag = ""
    tmrPow(index).Enabled = False
End If
End Sub

Private Sub tmrObjEvent_Timer()
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

