VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hello World"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHello 
      Caption         =   "Say Hello"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblEnterName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Name :"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1305
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHello_Click()
If txtName.Text = "" Then
    MsgBox "Hey, You must enter name.", vbCritical
    Exit Sub
End If

MsgBox "Hello " & txtName.Text, vbExclamation
End Sub

