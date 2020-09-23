VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T.A.S. Code Library Add New Code"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4140
      Left            =   135
      TabIndex        =   5
      Top             =   60
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   7303
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1.00000e5
      TextRTF         =   $"Form2.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Left            =   3225
      TabIndex        =   4
      Top             =   5460
      Width           =   840
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
      Left            =   5205
      TabIndex        =   3
      Top             =   5460
      Width           =   840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Max Title Legnth 80"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2460
      TabIndex        =   0
      Top             =   4320
      Width           =   4215
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
         Left            =   150
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   555
         Width           =   3930
      End
      Begin VB.TextBox Text2 
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
         Height          =   285
         Left            =   150
         MaxLength       =   80
         TabIndex        =   1
         Text            =   "Type Or Paste Title Here"
         Top             =   240
         Width           =   3930
      End
   End
   Begin VB.Menu MenuY 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu MnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "Paste"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If RichTextBox1.Text = "" Then Exit Sub
If Text2.Text = "Type Name Here" Then Exit Sub
Form1.CheckCodeEntry Text2.Text, Left(Combo1.ListIndex, 1)
If Form1.Testing = True Then
MsgBox "That name allready exsists !"
Exit Sub
End If
Dim memoryhog As String
Form1.RichTextBox1.Visible = False
Me.Hide
memoryhog = RichTextBox1.Text & vbCr
Open App.Path & "\Data\CodeLib.cod" For Append As #1
Print #1, "*-*-*-*-*" & Left(Combo1.ListIndex, 1) & "*-*-*-*-*"
Print #1, "/*-*-*-*-*" & Text2.Text & "*-*-*-*-*\"
Print #1, memoryhog
Print #1, "*-*-*-*-*"
Close #1
Form1.Show
Form1.List1.Clear
Call Form1.LoadListFromFile(App.Path & "\Data\CodeLib.cod", Form1.List1)
Call Form1.LoadMyTopicList
If Form1.Combo1.ListIndex = Combo1.ListIndex Then
  For z = 0 To Form1.List2.ListCount - 1
  If (Form1.List2.List(z)) = (Text2.Text) Then 'LCase
  Form1.List2.Selected(z) = True
  Form1.OpenCodeEntry (Form1.List2.Text)
  Exit For
  End If
  Next
End If
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
Form1.Show
End Sub
Private Sub Form_Load()
btnFlat Command1
btnFlat Command2
Call FormOnTop(Me.hWnd, True)
For X = 0 To Form1.Combo1.ListCount - 1
Combo1.AddItem Form1.Combo1.List(X)
Next X
Combo1.ListIndex = Form1.Combo1.ListIndex
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
Form1.Show
End Sub
Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
        Case 1 'Left Click
         'Text2.Text = ""
        Case 2 'Right Click
         PopupMenu MenuY
End Select
End Sub
Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
        Case 1 'Left Click
         'Text2.Text = ""
        Case 2 'Right Click
         Text2.Text = ""
End Select
End Sub
Private Sub MnuCopy_Click()
'Clears the Clipboard to put text on it
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
End Sub
Private Sub MnuPaste_Click()
Colorize RichTextBox1, Clipboard.GetText 'DCode
End Sub
Private Sub MnuCut_Click()
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from rtfText onto the Clipboard
    Clipboard.SetText RichTextBox1.SelText
    'Deletes the Selected Text on rtfText
    RichTextBox1.SelText = ""
End Sub
