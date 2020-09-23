VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Updater"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   1800
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar p2 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   300
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3480
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar p1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "C A N C E L"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Finding Site"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form5
End Sub

Private Sub Timer1_Timer()
p1.Value = p1.Value + "1"
If p1.Value = "63" Then
Label1.Caption = "Retriving Information"
End If
If p1.Value = "100" Then
p1.Value = "0"
p1.Value = p1.Value + "0"
p1.Visible = False
p2.Visible = True
Timer2.Enabled = True
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Command1.Enabled = False
p2.Value = p2.Value + "1"
If p2.Value = "1" Then
Label1.Caption = "Installing Updates"
End If
If p2.Value = "300" Then
p2.Value = "0"
p2.Value = p2.Value + "0"
p2.Visible = False
Unload Form5
Timer2.Enabled = False
Timer1.Enabled = False
End If
End Sub
