VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form4"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4035
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2070
      TabIndex        =   2
      Top             =   495
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rename"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   675
      TabIndex        =   1
      Top             =   495
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   142
      MaxLength       =   80
      TabIndex        =   0
      Top             =   90
      Width           =   3750
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileName As String
Private Sub Command1_Click()
If Text1.Text = "" Then Exit Sub
Form1.ReNameCodeEntry Left(Form1.Combo1.ListIndex, 1), FileName, Text1.Text
Me.Hide
Form1.LoadMyTopicList
Call Form1.SaveList(App.Path & "\Data\CodeLib.cod", Form1.List1)
End Sub
Private Sub Command2_Click()
Me.Hide
End Sub
Private Sub Form_Load()
btnFlat Command1
btnFlat Command2
Call FormOnTop(Me.hWnd, True)
End Sub
