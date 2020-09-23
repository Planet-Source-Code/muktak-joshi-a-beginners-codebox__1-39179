VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture Browser"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   120
      Pattern         =   "*.jpg"
      TabIndex        =   0
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Image imgBrowser 
      Height          =   3495
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboType_Click()
Select Case cboType.ListIndex
Case 0
File1.Pattern = "*.bmp"

Case 1
File1.Pattern = "*.gif"

Case 2
File1.Pattern = "*.jpg"

End Select
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
imgBrowser.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
End Sub

Private Sub Form_Load()
cboType.ListIndex = 0
File1.Pattern = "*.bmp"
End Sub
