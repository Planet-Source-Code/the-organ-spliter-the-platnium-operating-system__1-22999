VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Folder"
   ClientHeight    =   3255
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   2160
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files *.*|*.*"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "All Files"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "os2.frx":0000
      Left            =   0
      List            =   "os2.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu f 
      Caption         =   "File"
      Begin VB.Menu c 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CD.ShowOpen
List1.AddItem (CD.FileName)

End Sub

Private Sub Command2_Click()
file = Shell(List1, vbNormalFocus)
AppActivate (file)

End Sub

Private Sub List1_DblClick()
file = Shell(List1, vbNormalFocus)
AppActivate (file)

End Sub
