VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Format Date/Time"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSeconds 
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtMinutes 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtHours 
      Height          =   375
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtYear 
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtMonth 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtDay 
      Height          =   375
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Height          =   375
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Seconds"
      Height          =   195
      Left            =   3120
      TabIndex        =   12
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Minutes"
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Hours"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Year "
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Month "
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Day "
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date / Time :"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
txtDate.Text = Now
txtDay.Text = Format(txtDate.Text, "dd")
txtMonth.Text = Format(txtDate.Text, "mm")
txtYear.Text = Format(txtDate.Text, "yyyy")
txtHours.Text = Hour(Time)
txtMinutes.Text = Minute(Time)
txtSeconds.Text = Second(Time)

End Sub

