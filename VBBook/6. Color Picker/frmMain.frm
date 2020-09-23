VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Picker"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HSBlue 
      Height          =   375
      Left            =   1920
      Max             =   255
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.HScrollBar HSGreen 
      Height          =   375
      Left            =   1920
      Max             =   255
      TabIndex        =   2
      Top             =   1524
      Width           =   2655
   End
   Begin VB.HScrollBar HSRed 
      Height          =   375
      Left            =   1920
      Max             =   255
      TabIndex        =   1
      Top             =   528
      Width           =   2655
   End
   Begin VB.Label lblBlue 
      AutoSize        =   -1  'True
      Caption         =   "Blue : 0"
      Height          =   195
      Left            =   1920
      TabIndex        =   6
      Top             =   2115
      Width           =   540
   End
   Begin VB.Label lblGreen 
      AutoSize        =   -1  'True
      Caption         =   "Green : 0"
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   1110
      Width           =   660
   End
   Begin VB.Label lblRed 
      AutoSize        =   -1  'True
      Caption         =   "Red : 0 "
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   570
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00000000&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HSBlue_Change()
lblBlue.Caption = "Blue : " & HSBlue.Value
lblColor.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
End Sub

Private Sub HSGreen_Change()
lblGreen.Caption = "Green : " & HSGreen.Value
lblColor.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
End Sub

Private Sub HSRed_Change()
lblRed.Caption = "Red : " & HSRed.Value
lblColor.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
End Sub
