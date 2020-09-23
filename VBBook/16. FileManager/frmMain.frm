VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "File Manager"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File2 
      Height          =   2040
      Left            =   3360
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdRen 
      Caption         =   "Rename"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.DirListBox Dir2 
      Height          =   1440
      Left            =   3360
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.DriveListBox Drive2 
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentFileBox As Object
Private Sub cmdCopy_Click()
If Right(File1.Path, 1) <> "\" Then
FileCopy File1.Path & "\" & File1.FileName, File2.Path & "\" & File1.FileName
Else
FileCopy File1.Path & File1.FileName, File2.Path & File1.FileName
End If
File2.Path = File2.Path
File2.Refresh
End Sub

Private Sub cmdDel_Click()
Kill CurrentFileBox.Path & "\" & CurrentFileBox.FileName
CurrentFileBox.Path = CurrentFileBox.Path
CurrentFileBox.Refresh
End Sub

Private Sub cmdRen_Click()
Dim Rento As String
Rento = InputBox("New File Name:", "Rename", CurrentFileBox.FileName)
FileCopy CurrentFileBox.Path & "\" & CurrentFileBox.FileName, CurrentFileBox.Path & "\" & Rento
Kill CurrentFileBox.Path & "\" & CurrentFileBox.FileName
CurrentFileBox.Path = CurrentFileBox.Path
CurrentFileBox.Refresh
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Drive2_Change()
Dir2.Path = Drive2.Drive
End Sub
Private Sub Dir2_Change()
File2.Path = Dir2.Path
End Sub

Private Sub File1_Click()
Set CurrentFileBox = File1
End Sub

Private Sub File1_PathChange()
If File1.ListCount = 0 Then
    
    cmdCopy.Enabled = False
    cmdRen.Enabled = False
    cmdDel.Enabled = False
Else
    File1.ListIndex = 0
    
    cmdCopy.Enabled = True
    cmdRen.Enabled = True
    cmdDel.Enabled = True
End If
End Sub

Private Sub File2_Click()
Set CurrentFileBox = File2
End Sub



Private Sub Form_Load()
If File1.ListCount <> 0 Then File1.ListIndex = 0
If File2.ListCount <> 0 Then File2.ListIndex = 0
CurrentFileBox = File1
End Sub
