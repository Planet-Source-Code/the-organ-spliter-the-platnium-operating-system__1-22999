VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form7 
   Caption         =   "Platnium Pad"
   ClientHeight    =   6525
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7245
   LinkTopic       =   "Form7"
   ScaleHeight     =   6525
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   6495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2280
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu fff 
      Caption         =   "File"
      Begin VB.Menu OPEN 
         Caption         =   "Open"
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
      Begin VB.Menu l 
         Caption         =   "-"
      End
      Begin VB.Menu close 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CharCount As Boolean
Private Sub close_Click()
Unload Form7
End Sub

Private Sub OPEN_Click()
On Error GoTo err
Dim i As Long
Form1.Caption = "Platnium Pad Opening"
Text1.Text = ""
CD1.Filter = "Txt (*.txt)|*.txt|Any File (*.*)|*.*"
CD1.ShowOpen
If CD1.FileName <> "" Then
Dim t As Long
i = FreeFile
Open CD1.FileName For Input As #i
CharCount = True


Text1.Text = Input(LOF(i), i)
Close #i
Else
Form1.Caption = "Platnium Pad"
Exit Sub
End If


Form1.Caption = "Platnium Pad"

Exit Sub



err:

Close #i
Open CD1.FileName For Binary As #i
Text1.Text = Input(LOF(i), i)
Close #i
Form1.Caption = "Platnium Pad"
Exit Sub
End Sub

Private Sub save_Click()
CD1.Filter = "Txt (*.txt)|*.txt|Any File (*.*)|*.*"
CD1.ShowSave
Open CD1.FileName For Output As #1
Print #1, Text1.Text
Close #1
End Sub
