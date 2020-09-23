VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Catagory To Reassign To !"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3420
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3420
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
      Left            =   1883
      TabIndex        =   2
      Top             =   555
      Width           =   885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reassign"
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
      Left            =   653
      TabIndex        =   1
      Top             =   555
      Width           =   885
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   90
      Width           =   3015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
X = 0
Do While X < Form1.List2.ListCount
If Form1.List2.Selected(X) = True Then 'Found a selected item
Form1.ReAsignCodeEntry Left(Form1.Combo1.ListIndex, 1), Form1.List2.List(X), Left(Combo1.ListIndex, 1)
Else 'This is very important. without this, it would generate an error:
End If
X = X + 1
DoEvents
Loop
Form1.Combo1.ListIndex = Left(Combo1.ListIndex, 1)
Call Form1.SaveList(App.Path & "\Data\CodeLib.cod", Form1.List1)
'Me.Hide
Unload Me
End Sub
Private Sub Command2_Click()
'Me.Hide
Unload Me
End Sub
Private Sub Form_Load()
btnFlat Command1
btnFlat Command2
Call FormOnTop(Me.hWnd, True)
For X = 0 To Form1.Combo1.ListCount - 1
Combo1.AddItem Form1.Combo1.List(X)
Next X
Combo1.ListIndex = 0
End Sub
