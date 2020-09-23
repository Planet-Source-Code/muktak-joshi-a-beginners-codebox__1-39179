VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create Tables"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstTable 
      Height          =   2010
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   ">>"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtNumber 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "Enter a Number :"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()
Dim X As Integer

If txtNumber.Text = "" Then
MsgBox "Please Enter a Number."
txtNumber.SetFocus
Exit Sub
End If
lstTable.Clear
For X = 1 To 10
lstTable.AddItem vbTab & txtNumber.Text & "   X    " & X & vbTab & "=" & vbTab & txtNumber.Text * X
Next
End Sub

