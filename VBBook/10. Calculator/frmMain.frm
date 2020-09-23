VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Advance Calculator"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2040
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   2040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOff 
      Caption         =   "OFF"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CE"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtScreen 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdEqualto 
      Caption         =   "="
      Height          =   375
      Left            =   1080
      TabIndex        =   15
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "."
      Height          =   375
      Index           =   10
      Left            =   600
      TabIndex        =   14
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdOperation 
      Caption         =   "+"
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   13
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdOperation 
      Caption         =   "-"
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   12
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdOperation 
      Caption         =   "*"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdOperation 
      Caption         =   "/"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   1080
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   600
      TabIndex        =   8
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sLastOperation As String
Dim dTotal As Single
Dim bNewEntry As Boolean

Private Sub cmdClear_Click()
txtScreen.Text = "0"
dTotal = 0
sLastOperation = ""
bNewEntry = True
End Sub

Private Sub cmdEqualto_Click()
cmdOperation_Click 0
txtScreen.Text = dTotal
dTotal = 0
bNewEntry = True
sLastOperation = ""
End Sub

Private Sub cmdKey_Click(Index As Integer)
If bNewEntry Then
txtScreen.Text = cmdKey(Index).Caption
bNewEntry = False
Else
txtScreen.Text = txtScreen.Text & cmdKey(Index).Caption
End If
End Sub


Private Sub cmdOff_Click()
End
End Sub

Private Sub cmdOperation_Click(Index As Integer)
If dTotal = 0 Then
dTotal = txtScreen.Text
sLastOperation = cmdOperation(Index).Caption
bNewEntry = True
Exit Sub
End If

Select Case sLastOperation
Case "-"
    dTotal = dTotal - txtScreen.Text
Case "+"
    dTotal = dTotal + txtScreen.Text
Case "/"
    dTotal = dTotal / txtScreen.Text
Case "*"
    dTotal = dTotal * txtScreen.Text
End Select
txtScreen.Text = dTotal
sLastOperation = cmdOperation(Index).Caption
bNewEntry = True
End Sub


Private Sub Form_Load()
bNewEntry = True
End Sub
