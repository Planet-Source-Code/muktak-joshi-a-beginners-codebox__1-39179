VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ListBox Sample"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">>"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.ListBox lstList2 
      Height          =   1620
      Left            =   2280
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdMoveL 
      Caption         =   "<"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdMoveR 
      Caption         =   ">"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.ListBox lstList1 
      Height          =   1620
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtData 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblEnterText 
      AutoSize        =   -1  'True
      Caption         =   "Enter Text : "
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
If txtData.Text = "" Then
MsgBox "Please Enter Some Text"
Exit Sub
End If
lstList1.AddItem txtData.Text
txtData.Text = ""
End Sub

Private Sub cmdMoveR_Click()
If lstList1.ListIndex = -1 Then
MsgBox "Please select an item from ListBox"
Exit Sub
End If
lstList2.AddItem lstList1.Text
lstList1.RemoveItem lstList1.ListIndex
End Sub
Private Sub cmdMovel_Click()
If lstList2.ListIndex = -1 Then
MsgBox "Please select an item from ListBox"
Exit Sub
End If
lstList1.AddItem lstList2.Text
lstList2.RemoveItem lstList2.ListIndex
End Sub

Private Sub txtData_Change()

End Sub
