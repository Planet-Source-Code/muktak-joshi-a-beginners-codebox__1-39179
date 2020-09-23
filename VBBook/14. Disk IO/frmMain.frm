VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disk I/O"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txtContents 
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   5295
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Path of File :"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1305
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOpen_Click()
Dim FileNum
Dim InpStr As String
If Dir(txtPath.Text) = "" Then
MsgBox "File not Found", vbCritical
Exit Sub
End If
FileNum = FreeFile
Open txtPath.Text For Input As FileNum
While Not EOF(FileNum)
Line Input #FileNum, InpStr
txtContents.Text = txtContents.Text & InpStr & vbCrLf
Wend
Close FileNum

End Sub

Private Sub cmdSave_Click()
Dim FileNum
FileNum = FreeFile
Open txtPath.Text For Output As FileNum
Print #FileNum, txtContents.Text
Close FileNum
MsgBox "Done", vbExclamation
End Sub

Private Sub Form_Load()
txtPath.Text = App.Path + "\Test.txt"
End Sub

