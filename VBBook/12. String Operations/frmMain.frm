VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "String Operations"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optReverse 
      Caption         =   "Reverse the String"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "Do It"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.OptionButton optMid 
      Caption         =   "Get 3 Letters from 2nd Letter"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.OptionButton optRight 
      Caption         =   "Get 3 Letters from Right"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.OptionButton optLeft 
      Caption         =   "Get 3 Letters from Left"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.TextBox txtString 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblString 
      AutoSize        =   -1  'True
      Caption         =   "Enter Text :"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   825
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDoIt_Click()
If Len(txtString.Text) < 6 Then
MsgBox "The String must contain at least 6 characters."
Exit Sub
End If
If optLeft.Value = True Then
MsgBox Left$(txtString.Text, 3)
End If

If optRight.Value = True Then
MsgBox Right(txtString.Text, 3)
End If

If optMid.Value = True Then
MsgBox Mid(txtString.Text, 2, 3)
End If

If optReverse.Value = True Then
MsgBox StrReverse(txtString.Text)
End If
End Sub
