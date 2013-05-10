VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Dino Hopper"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTitleMainMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   9  'Not Mask Pen
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   1470
      Picture         =   "frmStart.frx":0000
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   604
      TabIndex        =   1
      Top             =   1500
      Width           =   9060
   End
   Begin VB.PictureBox picTitleMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   1470
      Picture         =   "frmStart.frx":09C7
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   604
      TabIndex        =   0
      Top             =   1500
      Width           =   9060
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
