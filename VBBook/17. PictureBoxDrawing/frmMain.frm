VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "PictureBox Drawing"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   ScaleHeight     =   5805
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fTools 
      Caption         =   "Select Tool"
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton optCircle 
         Caption         =   "Circle"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton optRectangle 
         Caption         =   "Rectangle"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton optLine 
         Caption         =   "Line"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton optFreeHand 
         Caption         =   "FreeHand"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   2640
      MousePointer    =   2  'Cross
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   2040
      Width           =   2175
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   1080
         X2              =   1320
         Y1              =   600
         Y2              =   840
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   3  'Dot
         Height          =   495
         Left            =   240
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00808080&
      Caption         =   " Draw Free Hand Drawing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tool As String
Dim StartX As Long
Dim StartY As Long

Private Sub Form_Load()
Tool = "FreeHand"
optFreeHand.Value = True
End Sub

Private Sub Form_Resize()
lblStatus.Move 2000, 0, Me.ScaleWidth - 2000, 480
fTools.Move 0, 0, 2000, Me.ScaleHeight
picDraw.Move 2000, 480, Me.ScaleWidth - 2000, Me.ScaleHeight - 480
End Sub



Private Sub optCircle_Click()
Tool = "Circle"
lblStatus.Caption = "  Select Centre"
End Sub

Private Sub optFreeHand_Click()
Tool = "FreeHand"
lblStatus.Caption = "  Draw Free Hand Drawing."
End Sub

Private Sub optLine_Click()
Tool = "Line"
lblStatus.Caption = "  Select Starting Point."
End Sub

Private Sub optRectangle_Click()
Tool = "Rectangle"
lblStatus.Caption = "  Select Starting Point."
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartX = X
StartY = Y
lblStatus.Caption = "Drag the Mouse while pressing button"
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Tool = "FreeHand" Then
picDraw.PSet (X, Y), vbRed
Else
lblStatus.Caption = "Leave the button to Draw"
If Tool = "Rectangle" Then
    Shape1.Move StartX, StartY
    Shape1.Width = Replace(X - StartX, "-", "")
    Shape1.Height = Replace(Y - StartY, "-", "")
    Shape1.Visible = True
End If
If Tool = "Line" Then
Line1.X1 = StartX
Line1.X2 = X
Line1.Y1 = StartY
Line1.Y2 = Y
Line1.Visible = True
End If
End If
End If
End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
Line1.Visible = False
Select Case Tool
Case "Rectangle"
picDraw.Line (StartX, StartY)-(X, StartY)
picDraw.Line (X, StartY)-(X, Y)
picDraw.Line (X, Y)-(StartX, Y)
picDraw.Line (StartX, Y)-(StartX, StartY)

Case "Circle"
picDraw.Circle (StartX, StartY), Replace(Y - StartY, "-", "")
Case "Line"
picDraw.Line (StartX, StartY)-(X, Y)
End Select

End Sub
