VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "PhoneBook Using DAO"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add New"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   3840
      Width           =   975
   End
   Begin VB.Data ObjData 
      BOFAction       =   1  'BOF
      Caption         =   "PhoneBook"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Muktak Joshi\My Documents\VBBook\18. PhoneBook Using DAO\PhoneBook.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   495
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PhoneBook"
      Top             =   3120
      Width           =   4575
   End
   Begin VB.TextBox txtEmail 
      DataField       =   "Email"
      DataSource      =   "ObjData"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "Telephone"
      DataSource      =   "ObjData"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataSource      =   "ObjData"
      Height          =   1095
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      DataField       =   "Name"
      DataSource      =   "ObjData"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail :"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   525
   End
   Begin VB.Label lblPhone 
      AutoSize        =   -1  'True
      Caption         =   "Telephone :"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      Caption         =   "Address :"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   660
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name :"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   510
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddNew_Click()
ObjData.Recordset.AddNew
ObjData.Recordset.Update
ObjData.Recordset.MoveLast
txtName.SetFocus
End Sub

Private Sub cmdDelete_Click()
ObjData.Recordset.Delete
txtName.Text = ""
txtAddress.Text = ""
txtEmail.Text = ""
txtPhone.Text = ""
End Sub



Private Sub cmdSave_Click()
ObjData.Recordset.MoveLast
End Sub

Private Sub Form_Load()
ObjData.DatabaseName = App.Path + "\PhoneBook.mdb"
End Sub



