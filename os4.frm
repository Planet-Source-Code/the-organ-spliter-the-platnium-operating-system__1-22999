VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Delete Folder"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form4"
   ScaleHeight     =   2805
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Empty"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Restore"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "os4.frx":0000
      Left            =   0
      List            =   "os4.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
List1.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.del.Visible = False
Form1.del2.Visible = False
Form1.del3.Visible = False
End Sub

