VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T.A.S. Code Library"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   8385
      Top             =   1395
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Copy To Clipboard"
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
      Left            =   5565
      TabIndex        =   12
      Top             =   6030
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   1080
      TabIndex        =   10
      Top             =   705
      Width           =   7080
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   105
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   6840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Catagorys"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2618
      TabIndex        =   6
      Top             =   15
      Width           =   4005
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
         Left            =   75
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   270
         Width           =   2100
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2235
         TabIndex        =   8
         Top             =   315
         Width           =   780
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3120
         TabIndex        =   7
         Top             =   315
         Width           =   780
      End
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   4410
      TabIndex        =   0
      Top             =   6690
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Changes"
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
      Left            =   3915
      TabIndex        =   5
      Top             =   6030
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3525
      Left            =   68
      TabIndex        =   4
      Top             =   2415
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   6218
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1.00000e5
      TextRTF         =   $"Form1.frx":030A
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
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   6750
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New Code"
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
      Left            =   2220
      TabIndex        =   2
      Top             =   6030
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Height          =   390
      Left            =   1605
      ScaleHeight     =   330
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   6750
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   2160
      TabIndex        =   13
      Text            =   "Working please be patient !"
      Top             =   4020
      Width           =   4920
   End
   Begin VB.Menu MenuY 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu MnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu MnuCut 
         Caption         =   "Cut"
      End
   End
   Begin VB.Menu MenuX 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu MnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu MnuReAsignCodeEntry 
         Caption         =   "Reassign Code Entry"
      End
      Begin VB.Menu MnuRenameCodeEntry 
         Caption         =   "Rename Code Entry"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Testing As Boolean
Public Sub SaveList(xPath As String, LListBox As ListBox)
'This Sub saves the contents of a ListBox control to a
'file on the hard drive
Dim TheFile, ListCount As Integer
TheFile = FreeFile() 'This assures that our file number
                     'is not already in use by another file
                     
Open xPath For Output As #TheFile
    For ListCount = 0 To LListBox.ListCount - 1
        Print #TheFile, LListBox.List(ListCount)
    DoEvents
    Next ListCount
Close #TheFile
        
'frmPriority.lstbxSystemDialog.AddItem Date & " " & Time & " : Saved Listbox Contents From " & LListBox
End Sub
Public Sub LoadListFromFile(xPath As String, LListBox As ListBox)
'This Sub will load any list of information
'from your hard drive into a ListBox control.
'xPath represents the file to be loaded.
Dim TheFile As Integer
Dim OurBuffer As String
If Dir(xPath) = "" Then Exit Sub 'MsgBox "File Not Found!", 16, "Error: File Not Found": Exit Sub
TheFile = FreeFile() 'This will assure that the
                     'File number is a good one
                     
Open xPath For Input As #TheFile
    Do While Not EOF(TheFile)
        Line Input #TheFile, OurBuffer
        'If the line being read is empty, do not add it
        If OurBuffer = "" Then GoTo SkipAdd
        LListBox.AddItem OurBuffer
SkipAdd: 'If the item is empty, instructions jump here
        DoEvents 'Allow other proccesses to run
    Loop
Close #TheFile
'frmPriority.lstbxSystemDialog.AddItem Date & " " & Time & " : Loaded Listbox Contents For " & LListBox

End Sub
Public Sub RemoveCodeEntry(Input1 As String, Input2 As String)
Dim X As Long
Dim Y As Long
'Input1 = "0"
'Input2 = "BoxGradient"
For X = 0 To List1.ListCount - 1
 If List1.List(X) = "*-*-*-*-*" & Input1 & "*-*-*-*-*" And List1.List(X + 1) = "/*-*-*-*-*" & Input2 & "*-*-*-*-*\" Then
  For Y = X To List1.ListCount - 1
  If List1.List(X) = "*-*-*-*-*" Then List1.RemoveItem (X): GoTo Finish
  List1.RemoveItem (X)
  Next Y
 End If
Next X
Finish:
Call SaveList(App.Path & "\Data\CodeLib.cod", Form1.List1)
End Sub
Public Sub ReAsignCodeEntry(Input1 As String, Input2 As String, Input3 As String)
Dim X As Long
'Input1 = "0"
'Input2 = "BoxGradient"
'Input3 = "3"
For X = 0 To List1.ListCount - 1
 If List1.List(X) = "*-*-*-*-*" & Input1 & "*-*-*-*-*" And List1.List(X + 1) = "/*-*-*-*-*" & Input2 & "*-*-*-*-*\" Then
  List1.List(X) = "*-*-*-*-*" & Input3 & "*-*-*-*-*"
  Call SaveList(App.Path & "\Data\CodeLib.cod", Form1.List1)
  Exit Sub
 End If
Next X
End Sub
Public Sub ReNameCodeEntry(Input1 As String, Input2 As String, Input3 As String)
Dim X As Long
'Input1 = "0"
'Input2 = "BoxGradient"
'Input3 = "Box Gradient"
For X = 0 To List1.ListCount - 1
 If List1.List(X) = "*-*-*-*-*" & Input1 & "*-*-*-*-*" And List1.List(X + 1) = "/*-*-*-*-*" & Input2 & "*-*-*-*-*\" Then
  List1.List(X + 1) = "/*-*-*-*-*" & Input3 & "*-*-*-*-*\"
  Exit Sub
 End If
Next X
Call SaveList(App.Path & "\Data\CodeLib.cod", Form1.List1)
End Sub
Private Sub Combo1_Click()
On Error Resume Next
Picture1.SetFocus
'Text1.Text = ""
Call LoadMyTopicList
If List2.ListCount <> 0 Then
List2.Selected(0) = True
End If
OpenCodeEntry (List2.List("0"))
End Sub
Private Sub Command1_Click()
On Error Resume Next
Picture1.SetFocus
Me.Hide
Form2.Show
End Sub
Private Sub Command2_Click()
On Error Resume Next
Picture1.SetFocus
Form5.Show
End Sub
Private Sub Command3_Click()
On Error Resume Next
Picture1.SetFocus
If Combo1 = "Unsorted" Then
Form6.Show
Exit Sub
End If
If List2.ListCount > 0 Then
CheckForSure = MsgBox( _
   "All files left uder this topic will be automaticly moved to Unsorted! Are you sure you want to proceed with this operation?", _
   vbOKCancel + vbCritical, _
   " ")

Select Case CheckForSure
   Case vbCancel
      Exit Sub
End Select
Else
CheckForSure = MsgBox( _
   "Are you sure you want to delete this catagory?", _
   vbOKCancel + vbCritical, _
   " ")

Select Case CheckForSure
   Case vbCancel
     Exit Sub
End Select
End If
For X = 0 To List2.ListCount - 1
ReAsignCodeEntry Left(Combo1.ListIndex, 1), List2.List(X), "0"
Next X

Dim MyStartPoint As String
Dim MyEndPoint As String
MyStartPoint = Left(Combo1.ListIndex, 1) + 1
MyEndPoint = Combo1.ListCount
For X = MyStartPoint To MyEndPoint
Combo1.ListIndex = X
 For Y = 0 To List2.ListCount - 1
  ReAsignCodeEntry Left(Combo1.ListIndex, 1), List2.List(Y), X - 1
 Next Y
Next X

List3.RemoveItem MyStartPoint - 1
Combo1.Clear
For X = 0 To List3.ListCount - 1
Combo1.AddItem List3.List(X)
Next X
Combo1.ListIndex = 0
Call SaveList(App.Path & "\Data\CodeLib.cod", Form1.List1)
Call SaveList(App.Path & "\Data\CodeLibCatagorys.cod", Form1.List3)
End Sub
Private Sub Command4_Click()
On Error Resume Next
Picture1.SetFocus
Dim TheFileName As String
X = 0
Do While X < List2.ListCount
If List2.Selected(X) = True Then 'Found a selected item
TheFileName = List2.List(X)
Else 'This is very important. without this, it would generate an error:
End If
X = X + 1
DoEvents
Loop
RemoveCodeEntry Left(Combo1.ListIndex, 1), TheFileName
DoEvents
Open App.Path & "\Data\CodeLib.cod" For Append As #1
Print #1, "*-*-*-*-*" & Left(Combo1.ListIndex, 1) & "*-*-*-*-*"
Print #1, "/*-*-*-*-*" & TheFileName & "*-*-*-*-*\"
Print #1, RichTextBox1.Text & vbCr
Print #1, "*-*-*-*-*"
Close #1
List1.Clear
Call LoadListFromFile(App.Path & "\Data\CodeLib.cod", Form1.List1)
OpenCodeEntry TheFileName
End Sub



Private Sub Command5_Click()
On Error Resume Next
Picture1.SetFocus
'Clears the Clipboard to put text on it
Clipboard.Clear
Clipboard.SetText RichTextBox1.Text
End Sub

Private Sub MnuRenameCodeEntry_Click()
X = 0
Do While X < List2.ListCount
If List2.Selected(X) = True Then 'Found a selected item
Form4.Show
Form4.Caption = "Renaming - " & List2.List(X)
Form4.FileName = List2.List(X)
Form4.Text1.Text = List2.List(X)
Else 'This is very important. without this, it would generate an error:
End If
X = X + 1
DoEvents
Loop
End Sub
Private Sub MnuReAsignCodeEntry_Click()
'Me.Hide
Form3.Show
End Sub
Private Sub List2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Select Case Button
        Case 1 'Left Click
          On Error Resume Next
          Picture1.SetFocus
          X = 0
          Do While X < List2.ListCount
          If List2.Selected(X) = True Then 'Found a selected item
          If List2.SelCount > 1 Then Exit Sub
          OpenCodeEntry (List2.List(X))
          Exit Sub
          Else 'This is very important. without this, it would generate an error:
          End If
          X = X + 1
          DoEvents
          Loop
        Case 2 'Right Click
          PopupMenu MenuX
    End Select
End Sub
Private Sub MnuDelete_Click()
CheckForSure = MsgBox( _
   "Are you sure you want to delete these files?", _
   vbOKCancel + vbCritical, _
   " ")

Select Case CheckForSure
   Case vbCancel
      Exit Sub
End Select
X = 0
Do While X < List2.ListCount
If List2.Selected(X) = True Then 'Found a selected item
RemoveCodeEntry Left(Combo1.ListIndex, 1), List2.List(X) 'OpenCodeEntry ()
List2.RemoveItem (X)
RichTextBox1.Text = ""
List2.Selected(0) = True
OpenCodeEntry (List2.List(0))
Else 'This is very important. without this, it would generate an error:
End If
X = X + 1
DoEvents
Loop
End Sub
Private Sub Form_Load()
If App.PrevInstance Then End
Me.Caption = Me.Caption & " - " & App.Major & "." & App.Minor & "." & App.Revision

btnFlat Command1
btnFlat Command2
btnFlat Command3
btnFlat Command4
btnFlat Command5

Call LoadListFromFile(App.Path & "\Data\CodeLibCatagorys.cod", Form1.List3)
For X = 0 To List3.ListCount - 1
Combo1.AddItem List3.List(X)
Next X
Combo1.ListIndex = GetSetting("T.A.S. Code Library", "Settings", "Combo1", "0")

On Error Resume Next
Picture1.SetFocus
Me.Visible = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call SaveList(App.Path & "\Data\CodeLib.cod", Form1.List1)
Call SaveList(App.Path & "\Data\CodeLibCatagorys.cod", Form1.List3)
SaveSetting "T.A.S. Code Library", "Settings", "Combo1", Combo1.ListIndex
End
End Sub
Public Sub LoadMyTopicList()
Dim X As Long
Dim Y As Long
Dim R As Long
Dim MyIndex As String
List2.Clear
MyIndex = Combo1.ListIndex
For X = 0 To List1.ListCount - 1
 If List1.List(X) = "*-*-*-*-*" & MyIndex & "*-*-*-*-*" Then
 R = Len(List1.List(X + 1)) - 10
 List2.AddItem Left(Right(List1.List(X + 1), R), R - 10)
 End If
Next X
End Sub
Public Sub OpenCodeEntry(Name As String)
Dim X As Long
Dim Y As Long
Dim DCode As String
If List2.ListCount = 0 Then GoTo NoCode
RichTextBox1.Visible = False
RichTextBox1.Text = ""
For X = 0 To List1.ListCount - 1
 If List1.List(X) = "*-*-*-*-*" & Left(Combo1.ListIndex, 1) & "*-*-*-*-*" And List1.List(X + 1) = "/*-*-*-*-*" & Name & "*-*-*-*-*\" Then
  For Y = X + 2 To List1.ListCount - 1
  If List1.List(Y) = "*-*-*-*-*" Then GoTo Finish
  DCode = DCode & List1.List(Y) & Chr(13) + Chr(10)
  Next Y
 End If
Next X
Exit Sub
Finish:
Colorize RichTextBox1, DCode
RichTextBox1.Visible = True
Exit Sub
NoCode:
RichTextBox1.Visible = True
RichTextBox1.Text = ""
Colorize RichTextBox1, "'Their is no code in this category ! Click on 'Add New Code' below to add code to this category !"
End Sub
Public Sub CheckCodeEntry(Name As String, ComListIndex As String)
Dim X As Long
Dim Y As Long

RichTextBox1.Text = ""
For X = 0 To List1.ListCount - 1
 If List1.List(X) = "*-*-*-*-*" & ComListIndex & "*-*-*-*-*" And List1.List(X + 1) = "/*-*-*-*-*" & Name & "*-*-*-*-*\" Then
   Testing = True: Exit Sub
 End If
Next X
Testing = False
End Sub
Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
        Case 1 'Left Click
        
        Case 2 'Right Click
         PopupMenu MenuY
End Select
End Sub
Private Sub MnuCopy_Click()
'Clears the Clipboard to put text on it
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
End Sub
Private Sub MnuPaste_Click()
RichTextBox1.SelText = Clipboard.GetText
Dim DCode As String
DCode = RichTextBox1.Text
RichTextBox1.Text = ""
Colorize RichTextBox1, DCode
End Sub
Private Sub MnuCut_Click()
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from rtfText onto the Clipboard
    Clipboard.SetText RichTextBox1.SelText
    'Deletes the Selected Text on rtfText
    RichTextBox1.SelText = ""
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Call LoadListFromFile(App.Path & "\Data\CodeLib.cod", Form1.List1)
Call LoadMyTopicList
If List2.ListCount <> 0 Then
List2.Selected(0) = True
OpenCodeEntry (List2.Text)
End If
End Sub
