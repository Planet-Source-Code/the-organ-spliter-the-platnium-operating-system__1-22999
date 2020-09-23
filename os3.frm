VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Loading"
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form3"
   ScaleHeight     =   1515
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2640
      Top             =   960
   End
   Begin MSComctlLib.ProgressBar P 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   400
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "os3.frx":0000
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
P.Value = P.Value + "2"
If P.Value = "400" Then
Unload Me
Form1.Show
End If
End Sub
