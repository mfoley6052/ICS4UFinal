VERSION 5.00
Begin VB.Form frmBg 
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pictest 
      Height          =   9015
      Left            =   0
      ScaleHeight     =   8955
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmBg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
pictest.Picture = frmMain.picBackground.Image
End Sub
