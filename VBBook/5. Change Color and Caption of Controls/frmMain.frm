VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Controls"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContents 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   3600
      List            =   "frmMain.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdTarget 
      Caption         =   "Color Button"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtTarget 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "Color TextBox"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblContents 
      AutoSize        =   -1  'True
      Caption         =   "Set Contents :"
      Height          =   195
      Left            =   2520
      TabIndex        =   5
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      Caption         =   "Select Color :"
      Height          =   195
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label lblTarget 
      Caption         =   "Color Label"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cboColor_Click()
Dim lColor As Long
Select Case cboColor.ListIndex
    Case 0
        lColor = vbBlue
    Case 1
        lColor = vbRed
    Case 2
        lColor = vbYellow
    Case 3
        lColor = vbBlack
    Case 4
        lColor = vbWhite
End Select

txtTarget.BackColor = lColor
lblTarget.BackColor = lColor
cmdTarget.BackColor = lColor
End Sub

Private Sub txtContents_Change()
txtTarget.Text = txtContents.Text
lblTarget.Caption = txtContents.Text
cmdTarget.Caption = txtContents.Text
End Sub
