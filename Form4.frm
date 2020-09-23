VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Screwed Text"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3495
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      MaxLength       =   5
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Columns         =   3
      Height          =   2010
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   1320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "to:"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s(52) As String
Dim t(52) As String
Dim li As Long

Public Sub Sav2()
Directory$ = App.Path & "\text.dat"
       Dim SaveList As Integer
       On Error Resume Next
       Open Directory$ For Output As #1
       For SaveList = 0 To List1.ListCount - 1
           Print #1, s(SaveList)
       Next SaveList
       Close #1
End Sub

Public Sub Opn2()
Directory$ = App.Path & "\text.dat"
If Directory$ <> "" Then
    Dim MyString As String
    Dim c As Long
    Dim i As Integer
    c = 0
       On Error Resume Next
       Open Directory$ For Input As #1
       While Not EOF(1)
           Input #1, MyString$
           DoEvents
               s(c) = MyString$
               a = 1
               c = c + 1
           Wend
           Close #1
List1.Clear
For i = 0 To c - 1
List1.AddItem s(i)
Next i
End If
End Sub

Function sReplaceCharacters(strMainString As String, strOld As String, strNew As String) As String
    sReplaceCharacters = ""
    Dim strNewString As String
    Dim i As Integer
    For i = 1 To Len(strMainString)
        If Mid(strMainString, i, Len(strOld)) = strOld Then
            strNewString = strNewString & strNew
            i = i + Len(strOld) - 1
        Else
        strNewString = strNewString & Mid(strMainString, i, 1)
        End If
    Next i
    sReplaceCharacters = strNewString
End Function

Private Sub Command1_Click()
Sav2
Form2.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
  li = List1.ListIndex
  If List1.ListIndex > List1.ListCount - 1 Or List1.ListIndex < 0 Then Exit Sub
  List1.RemoveItem li
  List1.AddItem Text2.Text, li
  s(li) = Text2.Text
  List1.ListIndex = li
Form1.a
End Sub

Private Sub Form_Load()
Dim i As Integer
t(0) = "ae"
t(1) = "a"
t(2) = "b"
t(3) = "c"
t(4) = "d"
t(5) = "e"
t(6) = "f"
t(7) = "g"
t(8) = "h"
t(9) = "i"
t(10) = "j"
t(11) = "k"
t(12) = "l"
t(13) = "m"
t(14) = "n"
t(15) = "o"
t(16) = "p"
t(17) = "q"
t(18) = "r"
t(19) = "s"
t(20) = "t"
t(21) = "u"
t(22) = "v"
t(23) = "w"
t(24) = "x"
t(25) = "y"
t(26) = "z"
t(27) = "A"
t(28) = "B"
t(29) = "C"
t(30) = "D"
t(31) = "E"
t(32) = "F"
t(33) = "G"
t(34) = "H"
t(35) = "I"
t(36) = "J"
t(37) = "K"
t(38) = "L"
t(39) = "M"
t(40) = "N"
t(41) = "O"
t(42) = "P"
t(43) = "Q"
t(44) = "R"
t(45) = "S"
t(46) = "T"
t(47) = "U"
t(48) = "V"
t(49) = "W"
t(50) = "X"
t(51) = "Y"
t(52) = "Z"
s(0) = "æ"
s(1) = "å"
s(2) = "b"
s(3) = "<"
s(4) = "c|"
s(5) = "ë"
s(6) = "f"
s(7) = "9"
s(8) = "h"
s(9) = "ï"
s(10) = "j"
s(11) = "|<"
s(12) = "|_"
s(13) = "/x\"
s(14) = "|\|"
s(15) = "0"
s(16) = "p"
s(17) = "q"
s(18) = "r"
s(19) = "_/¯"
s(20) = "-|-"
s(21) = "µ"
s(22) = "\/"
s(23) = "\/\/"
s(24) = "×"
s(25) = "ÿ"
s(26) = "¯/_"
s(27) = "Ä"
s(28) = "ß"
s(29) = "©"
s(30) = "|}"
s(31) = "È"
s(32) = "F"
s(33) = "G"
s(34) = "|-|"
s(35) = "I"
s(36) = "J"
s(37) = "]<"
s(38) = "]_"
s(39) = "/\/\"
s(40) = "|\|"
s(41) = "{}"
s(42) = "P"
s(43) = "¶"
s(44) = "|2"
s(45) = "§"
s(46) = "¯|¯"
s(47) = "|_|"
s(48) = "\/"
s(49) = "\x/"
s(50) = "><"
s(51) = "¥"
s(52) = "¯/_"
For i = 0 To 52
List1.AddItem s(i)
Next i
Opn2
List1.ListIndex = 0
End Sub

Private Sub Form_Terminate()
Sav2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Sav2
End Sub

Private Sub List1_Click()
If List1.ListIndex < 0 Then
Exit Sub
Else
Text2.Text = List1.Text
Label1.Caption = t(List1.ListIndex)
End If
End Sub

Private Sub List1_GotFocus()
If List1.ListIndex < 0 Then
Exit Sub
Else
Text2.Text = List1.Text
Label1.Caption = t(List1.ListIndex)
End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If List1.ListIndex < 0 Then
Exit Sub
Else
Text2.Text = List1.Text
Label1.Caption = t(List1.ListIndex)
End If
End Sub
