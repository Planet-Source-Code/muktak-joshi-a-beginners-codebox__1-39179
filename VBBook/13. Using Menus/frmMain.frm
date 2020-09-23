VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Using Menus"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFrmClr 
      Caption         =   "Form Color"
      Begin VB.Menu mnuRed 
         Caption         =   "Red"
      End
      Begin VB.Menu mnuGreen 
         Caption         =   "Green"
      End
      Begin VB.Menu mnuBlue 
         Caption         =   "Blue"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuBlue_Click()
frmMain.BackColor = vbBlue
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuGreen_Click()
frmMain.BackColor = vbGreen
End Sub

Private Sub mnuRed_Click()
frmMain.BackColor = vbRed
End Sub
