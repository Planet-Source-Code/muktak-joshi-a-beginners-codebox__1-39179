VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Simple Timer"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer objTimer 
      Interval        =   1000
      Left            =   2160
      Top             =   0
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "      "
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   270
   End
   Begin VB.Label lblCurrentTime 
      AutoSize        =   -1  'True
      Caption         =   "Current Time :"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
lblValue.Caption = Time
End Sub

Private Sub objTimer_Timer()
lblValue.Caption = Time
End Sub
