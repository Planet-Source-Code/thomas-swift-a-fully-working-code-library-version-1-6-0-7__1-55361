VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add A New Main Catagory"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3705
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
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
      Left            =   547
      TabIndex        =   2
      Top             =   525
      Width           =   1215
   End
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
      Left            =   1942
      TabIndex        =   1
      Top             =   525
      Width           =   1215
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
      Height          =   345
      Left            =   90
      MaxLength       =   20
      TabIndex        =   0
      Top             =   60
      Width           =   3480
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For X = 0 To Form1.Combo1.ListIndex
 If Form1.Combo1.List(X) = Text1.Text Then
  MsgBox "That name allready exsists !"
  Exit Sub
  End If
Next X

If Text1.Text = "" Then Exit Sub
Form1.List3.AddItem Text1.Text
Form1.Combo1.Clear
For X = 0 To Form1.List3.ListCount - 1
Form1.Combo1.AddItem Form1.List3.List(X)
Next X
Form1.Combo1.ListIndex = Form1.Combo1.ListCount - 1
Call Form1.SaveList(App.Path & "\Data\CodeLibCatagorys.cod", Form1.List3)
Me.Hide
End Sub
Private Sub Command2_Click()
Me.Hide
End Sub
Private Sub Form_Load()
btnFlat Command1
btnFlat Command2
Call FormOnTop(Me.hWnd, True)
End Sub
