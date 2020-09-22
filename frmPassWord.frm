VERSION 5.00
Begin VB.Form frmPassWord 
   Caption         =   "RijnDael Password"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmPassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPass As String

Public Function GetPassWord(Prompt As String, Optional Title As String = "", Optional Default As String = "") As String
Label1.Caption = Prompt
Me.Caption = Title
If Len(Title) = 0 Then
    Title = App.Title & " Password Input"
End If
Me.Show 1
GetPassWord = mPass
End Function

Private Sub Command1_Click()
mPass = Text1.Text
Unload Me
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Unload Me
End Sub
