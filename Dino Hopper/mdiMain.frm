VERSION 5.00
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Dino Hopper"
   ClientHeight    =   9555
   ClientLeft      =   4215
   ClientTop       =   660
   ClientWidth     =   12000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
frmMain.Top = 0
frmMain.Show
frmGUI.Top = 0
frmGUI.Left = 0
frmGUI.BackColor = vbCyan
frmGUI.Show
End Sub
