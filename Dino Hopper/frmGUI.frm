VERSION 5.00
Begin VB.Form frmGUI 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1305
   ClientLeft      =   4095
   ClientTop       =   615
   ClientWidth     =   12000
   LinkTopic       =   "GUI"
   MDIChild        =   -1  'True
   Picture         =   "frmGUI.frx":0000
   ScaleHeight     =   87
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
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
      Left            =   9090
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1950
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
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1665
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
      Left            =   11490
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "frmGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
