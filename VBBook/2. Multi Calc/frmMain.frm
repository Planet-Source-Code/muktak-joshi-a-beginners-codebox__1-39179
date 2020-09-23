VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Multi Calculator"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Frame fOperation 
      Caption         =   "Operation"
      Height          =   2775
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton cmdDoIt 
         Caption         =   "Do It"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   1335
      End
      Begin VB.OptionButton optDivide 
         Caption         =   "Divide"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton optMultiply 
         Caption         =   "Multiply"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optSubstract 
         Caption         =   "Substract"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fInputs 
      Caption         =   "Inputs"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtNumber2 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtNumber1 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblNumber2 
         AutoSize        =   -1  'True
         Caption         =   "Number 2 :"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   780
      End
      Begin VB.Label lblNumber1 
         AutoSize        =   -1  'True
         Caption         =   "Number 1 :"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   780
      End
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      Caption         =   "Ouput :"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   525
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdDoIt_Click()
If txtNumber1.Text = "" Then
MsgBox "Please Enter Number 1"
txtNumber1.SetFocus
Exit Sub
End If

If txtNumber2.Text = "" Then
MsgBox "Please Enter Number 2"
txtNumber2.SetFocus
Exit Sub
End If

If optAdd.Value = True Then
txtOutput.Text = Val(txtNumber1) + Val(txtNumber2)
End If
If optSubstract.Value = True Then
txtOutput.Text = Val(txtNumber1) - Val(txtNumber2)
End If
If optMultiply.Value = True Then
txtOutput.Text = Val(txtNumber1) * Val(txtNumber2)
End If
If optDivide.Value = True Then
txtOutput.Text = Val(txtNumber1) / Val(txtNumber2)
End If

End Sub
