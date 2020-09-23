VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color with Timer"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ObjTimer 
      Interval        =   10
      Left            =   1680
      Top             =   1440
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00000000&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRed As Integer
Dim iGreen As Integer
Dim iBlue As Integer


Private Sub ObjTimer_Timer()
If iRed < 255 Then
    iRed = iRed + 1
Else
    If iGreen < 255 Then
        iGreen = iGreen + 1
    Else
        If iBlue < 255 Then
        iBlue = iBlue + 1
        End If
    End If
End If
lblColor.BackColor = RGB(iRed, iGreen, iBlue)
End Sub
