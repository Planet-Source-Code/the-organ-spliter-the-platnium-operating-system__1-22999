VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Platnium OS By CP Productions"
   ClientHeight    =   3255
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "os1.frx":0000
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label lblsc 
      Alignment       =   2  'Center
      Caption         =   "Shortcut #1"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgsc 
      Height          =   480
      Index           =   0
      Left            =   2640
      Picture         =   "os1.frx":0342
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LblText 
      Alignment       =   2  'Center
      Caption         =   "Text File #1"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image ImgText 
      Height          =   480
      Index           =   0
      Left            =   3600
      Picture         =   "os1.frx":0784
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Delete Folder"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "os1.frx":0BC6
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblfolder 
      Alignment       =   2  'Center
      Caption         =   "New Folder #1"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image ImgFolder 
      Height          =   480
      Index           =   0
      Left            =   1800
      Picture         =   "os1.frx":1008
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu g 
      Caption         =   "File"
      Begin VB.Menu n 
         Caption         =   "New"
         Begin VB.Menu fl 
            Caption         =   "Folder"
         End
         Begin VB.Menu tf 
            Caption         =   "Text File"
         End
      End
      Begin VB.Menu run 
         Caption         =   "Run"
      End
      Begin VB.Menu gu 
         Caption         =   "Get Updates"
      End
      Begin VB.Menu c 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu del 
      Caption         =   "Delete"
      Visible         =   0   'False
   End
   Begin VB.Menu del2 
      Caption         =   "Delete"
      Visible         =   0   'False
   End
   Begin VB.Menu del3 
      Caption         =   "Delete"
      Visible         =   0   'False
   End
   Begin VB.Menu ch 
      Caption         =   "Change Name"
      Begin VB.Menu fold 
         Caption         =   "Folder"
      End
      Begin VB.Menu d 
         Caption         =   "Text"
      End
      Begin VB.Menu scc 
         Caption         =   "Shortcut"
      End
   End
   Begin VB.Menu stt 
      Caption         =   "Shortcut To..."
      Begin VB.Menu hdd 
         Caption         =   "Hard Drive"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX As Long, OldY As Long, IsMoving As Boolean
Dim Selected As Integer, Stuffin As Boolean
Dim ChangingCaption As Boolean
Private Sub Command1_Click()

End Sub

Private Sub c_Click()
End
End Sub

Private Sub f_Click()
Static i As Integer
    i = i + 1

    Load ImgFolder(i)
    Load lblfolder(i)
    
    ImgFolder(i).Left = ImgFolder(i - 1).Left + 200
    ImgFolder(i).Top = ImgFolder(i - 1).Top + 600

    lblfolder(i).Left = lblfolder(i - 1).Left + 200
    lblfolder(i).Top = lblfolder(i - 1).Top + 600

    lblfolder(i).Caption = "New Folder #" & i
    ImgFolder(i).Visible = True
    lblfolder(i).Visible = True
End Sub

Private Sub d_Click()
Dim newName As String
newName$ = InputBox("What is the new name for the text file", "New Name")
LblText(Selected).Caption = newName$
End Sub

Private Sub ctf_Click()

End Sub

Private Sub del_Click()
ImgFolder(Selected).Visible = False
lblfolder(Selected).Visible = False
Form4.List1.AddItem (lblfolder(Selected).Caption)
End Sub

Private Sub don_Click()
Dim newName As String
newName$ = InputBox("What is the files new name", "New Name")
LblText(Selected).Caption = newName$
End Sub

Private Sub del2_Click()
ImgText(Selected).Visible = False
LblText(Selected).Visible = False
Form4.List1.AddItem (LblText(Selected).Caption)
End Sub

Private Sub del3_Click()
imgsc(Selected).Visible = False
lblsc(Selected).Visible = False
Form4.List1.AddItem (lblsc(Selected).Caption)
End Sub

Private Sub fl_Click()
Static i As Integer
    i = i + 1

    Load ImgFolder(i)
    Load lblfolder(i)
    
    ImgFolder(i).Left = ImgFolder(i - 1).Left + 200
    ImgFolder(i).Top = ImgFolder(i - 1).Top + 600

    lblfolder(i).Left = lblfolder(i - 1).Left + 200
    lblfolder(i).Top = lblfolder(i - 1).Top + 600

    lblfolder(i).Caption = "New Folder #" & i
    ImgFolder(i).Visible = True
    lblfolder(i).Visible = True
End Sub

Private Sub fold_Click()
Dim newName As String
newName$ = InputBox("What is the new name of the folder", "New Name")
lblfolder(Selected).Caption = newName$

End Sub

Private Sub Form_Click()
del.Visible = False
del3.Visible = False
del2.Visible = False
End Sub

Private Sub gu_Click()
Load Form5
Form5.Show
End Sub

Private Sub hdd_Click()
Static i As Integer
    i = i + 1

    Load imgsc(i)
    Load lblsc(i)
    
    imgsc(i).Left = imgsc(i - 1).Left + 200
    imgsc(i).Top = imgsc(i - 1).Top + 600

    lblsc(i).Left = lblsc(i - 1).Left + 200
    lblsc(i).Top = lblsc(i - 1).Top + 600

    lblsc(i).Caption = "Shortcut #" & i
    imgsc(i).Visible = True
    lblsc(i).Visible = True
End Sub

Private Sub Image1_DblClick()
Form4.Show
End Sub

Private Sub ImgFolder_Click(Index As Integer)
del2.Visible = False
del3.Visible = False
del.Visible = True
End Sub

Private Sub ImgFolder_DblClick(Index As Integer)
Load Form2
Form2.Show
End Sub

Private Sub ImgFolder_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
 Selected = ImgFolder(Index).Index
    Select Case Index
        Case Selected: OldX = X: OldY = Y: IsMoving = True
    End Select
End Sub

Private Sub ImgFolder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Selected = ImgFolder(Index).Index
    Select Case Index
        Case Selected
            If IsMoving Then
                ImgFolder(Selected).Top = ImgFolder(Selected).Top - (OldY - Y)
                ImgFolder(Selected).Left = ImgFolder(Selected).Left - (OldX - X)
        
                lblfolder(Selected).Top = lblfolder(Selected).Top - (OldY - Y)
                lblfolder(Selected).Left = lblfolder(Selected).Left - (OldX - X)
            End If
    End Select
End Sub

Private Sub ImgFolder_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Selected = ImgFolder(Index).Index
    Select Case Index
        Case Selected: IsMoving = False
    End Select
End Sub

Private Sub imgsc_Click(Index As Integer)
del.Visible = False
del2.Visible = False
del3.Visible = True
End Sub

Private Sub imgsc_DblClick(Index As Integer)
Load Form8
Form8.WebBrowser1.Navigate ("C:\")
Form8.Show
End Sub

Private Sub imgsc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Selected = imgsc(Index).Index
    Select Case Index
        Case Selected: OldX = X: OldY = Y: IsMoving = True
    End Select
End Sub

Private Sub imgsc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Selected = imgsc(Index).Index
    Select Case Index
        Case Selected
            If IsMoving Then
                imgsc(Selected).Top = imgsc(Selected).Top - (OldY - Y)
                imgsc(Selected).Left = imgsc(Selected).Left - (OldX - X)
        
                lblsc(Selected).Top = lblsc(Selected).Top - (OldY - Y)
                lblsc(Selected).Left = lblsc(Selected).Left - (OldX - X)
            End If
    End Select
End Sub

Private Sub imgsc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Selected = imgsc(Index).Index
    Select Case Index
        Case Selected: IsMoving = False
    End Select
End Sub

Private Sub ImgText_Click(Index As Integer)
del.Visible = False
del3.Visible = False
del2.Visible = True
End Sub

Private Sub ImgText_DblClick(Index As Integer)
Load Form7
Form7.Show
End Sub

Private Sub ImgText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Selected = ImgText(Index).Index
    Select Case Index
        Case Selected: OldX = X: OldY = Y: IsMoving = True
    End Select
End Sub

Private Sub ImgText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Selected = ImgText(Index).Index
    Select Case Index
        Case Selected
            If IsMoving Then
                ImgText(Selected).Top = ImgText(Selected).Top - (OldY - Y)
                ImgText(Selected).Left = ImgText(Selected).Left - (OldX - X)
        
                LblText(Selected).Top = LblText(Selected).Top - (OldY - Y)
                LblText(Selected).Left = LblText(Selected).Left - (OldX - X)
            End If
    End Select
End Sub

Private Sub ImgText_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Selected = ImgText(Index).Index
    Select Case Index
        Case Selected: IsMoving = False
    End Select
End Sub

Private Sub LblText_Click(Index As Integer)
Text1.Visible = True
Text1.Top = LblText(Index).Top
Text1.Left = LblText(Index).Left
Text1.SetFocus
End Sub

Private Sub run_Click()
Load Form6
Form6.Show
End Sub

Private Sub Text1_Change()
 ChangingCaption = True
    ChrNum = Len(Text1)
    Select Case ChrNum
        Case 13: Text1.Height = 525
        Case 26: Text1.Height = 765
        Case 39: Text1.Height = 1005
    End Select
End Sub

Private Sub scc_Click()
Dim newName As String
newName$ = InputBox("What is the new name for the Shortcut", "New Name")
lblsc(Selected).Caption = newName$

End Sub

Private Sub tf_Click()
Static i As Integer
    i = i + 1

    Load ImgText(i)
    Load LblText(i)
    
    ImgText(i).Left = ImgText(i - 1).Left + 200
    ImgText(i).Top = ImgText(i - 1).Top + 600

    LblText(i).Left = LblText(i - 1).Left + 200
    LblText(i).Top = LblText(i - 1).Top + 600

    LblText(i).Caption = "Text File #" & i
    ImgText(i).Visible = True
    LblText(i).Visible = True
End Sub
